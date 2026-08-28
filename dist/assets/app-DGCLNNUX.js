var ka=Object.defineProperty;var ct=(t,e)=>()=>(t&&(e=t(t=0)),e);var dt=(t,e)=>{for(var r in e)ka(t,r,{get:e[r],enumerable:!0})};var Ee={};dt(Ee,{API:()=>Et,CLIENT_SIDE_MAX_ROWS:()=>we,IS_DEVELOPMENT:()=>Me,apiFetch:()=>w,clearToken:()=>Re,getToken:()=>Oe,getUser:()=>ce,setToken:()=>pt,setUser:()=>Ke});function Oe(){return localStorage.getItem("fm_token")}function pt(t){localStorage.setItem("fm_token",t)}function Re(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ce(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Ke(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function w(t,e={}){let r=Oe(),a={"Content-Type":"application/json",...r?{Authorization:`Bearer ${r}`}:{},...e.headers||{}};try{let i=`cb=${Date.now()}`,s=t.includes("?")?"&":"?",o=`${Et}${t}${s}${i}`,l=await fetch(o,{...e,headers:a}),n;try{let p=await l.text();try{n=JSON.parse(p)}catch{n={error:`Server Error (${l.status}): ${p.substring(0,80)}...`}}}catch{n={error:"Gagal membaca respon dari server"}}return l.status===401&&(Re(),window.location.hash="#/login"),{ok:l.ok,status:l.status,data:n}}catch(i){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${i.message})`}}}}var Me,Sa,Et,we,F=ct(()=>{Me=!1,Sa="https://fm-operations-api.facilitycare-audydental.workers.dev",Et=Sa,we=1e4});var Dt={};dt(Dt,{confirmDialog:()=>qe,createModal:()=>de});function de({title:t,content:e,onConfirm:r,onCancel:a,confirmText:i="Simpan",cancelText:s="Batal",size:o="md",confirmClass:l="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},p=document.createElement("div");p.className="modal-overlay",p.innerHTML=`
    <div class="modal" style="max-width:${n[o]||n.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${s}</button>
        ${r?`<button class="btn ${l} modal-confirm">${i}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&p.querySelector(".modal-body").appendChild(e);let u=()=>{p.classList.remove("show"),setTimeout(()=>p.remove(),250)};return p.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),u()}),p.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),u()}),r&&p.querySelector(".modal-confirm").addEventListener("click",()=>r(p,u)),p.addEventListener("click",d=>{d.target===p&&(a&&a(),u())}),document.body.appendChild(p),requestAnimationFrame(()=>p.classList.add("show")),{overlay:p,close:u}}function qe(t,e,r="Konfirmasi"){return de({title:r,content:`<p>${t}</p>`,onConfirm:(a,i)=>{e(),i()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var xe=ct(()=>{});var oe={};dt(oe,{downloadExcel:()=>I,parseExcel:()=>Je,renderExcelButtons:()=>xa});function Je(t){return new Promise((e,r)=>{let a=new FileReader;a.onload=i=>{try{let s=new Uint8Array(i.target.result),o=XLSX.read(s,{type:"array"}),l=o.SheetNames[0],n=o.Sheets[l];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${o.SheetNames.join(", ")}`),console.log(`Sheet Used: ${l}`);let p=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),u=p.e.r-p.s.r+1,d=p.e.c-p.s.c+1;console.log(`Total Rows (including empty): ${u}`),console.log(`Total Columns: ${d}`);let c=[];for(let f=p.s.c;f<=p.e.c;++f){let b=n[XLSX.utils.encode_cell({c:f,r:p.s.r})];b&&b.v&&c.push(b.v)}console.log(`Headers Found: ${c.join(", ")}`),console.log("---------------------------");let g=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(g,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(g,"__headers",{value:c,enumerable:!1}),e(g)}catch(s){r(s)}},a.onerror=i=>r(i),a.readAsArrayBuffer(t)})}function I(t,e){try{let r=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,r,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(r){throw console.error("Error generating Excel file:",r),r}}function xa(t){return`
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
  `}var R=ct(()=>{});F();var mt={},Qe=null;function U(t,e){mt[t]=e}function $e(t){window.location.hash=t}function $t(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[r,...a]=e.split("?"),i=mt[r];if(!i){for(let[o,l]of Object.entries(mt))if(o.endsWith("/*")&&r.startsWith(o.slice(0,-2))){i=l;break}}Qe&&(Qe(),Qe=null);let s=document.getElementById("main-content");if(s&&(s.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),i){let o=new URLSearchParams(a.join("?")),l=r.split("/").filter(Boolean),n=await i({path:r,params:o,segments:l,main:s});n&&(Qe=n)}else{let o=s||document.getElementById("app");o&&(o.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var je;function wa(){return je||(je=document.createElement("div"),je.id="toast-container",document.body.appendChild(je)),je}function Tt(t,e="info",r=3500){let a=wa(),i=document.createElement("div");i.className=`toast toast-${e}`;let s={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};i.innerHTML=`<span class="toast-icon">${s[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(i),requestAnimationFrame(()=>i.classList.add("show")),setTimeout(()=>{i.classList.remove("show"),setTimeout(()=>i.remove(),350)},r)}var W=t=>Tt(t,"success"),z=t=>Tt(t,"error");xe();F();F();xe();function Ye({columns:t,data:e,onEdit:r,onDelete:a,onView:i,actions:s=[],emptyText:o="Tidak ada data",bulkSelect:l=null}){let n=document.createElement("div");if(n.className="table-wrapper",!e||e.length===0)return n.innerHTML=`<div class="empty-state"><p>${o}</p></div>`,n;let p=document.createElement("table");p.className="data-table";let u=document.createElement("thead"),d=document.createElement("tr");if(l){let g=document.createElement("th");g.style.width="40px",g.style.textAlign="center";let f=document.createElement("input");f.type="checkbox",f.id="select-all-checkbox",f.title="Pilih semua",f.addEventListener("change",()=>{e.forEach(b=>{f.checked?l.selectedIds.add(b.id):l.selectedIds.delete(b.id)}),n.querySelectorAll(".row-checkbox").forEach(b=>b.checked=f.checked),l.onToggle()}),g.appendChild(f),d.appendChild(g)}if(t.forEach(g=>{let f=document.createElement("th");f.textContent=g.label,g.width&&(f.style.width=g.width),d.appendChild(f)}),r||a||i||s.length>0){let g=document.createElement("th");g.textContent="Aksi",g.style.width="120px",d.appendChild(g)}u.appendChild(d),p.appendChild(u);let c=document.createElement("tbody");return e.forEach(g=>{let f=document.createElement("tr");if(l){let b=document.createElement("td");b.style.textAlign="center",b.style.width="40px";let m=document.createElement("input");m.type="checkbox",m.className="row-checkbox",m.checked=l.selectedIds.has(g.id),m.addEventListener("change",()=>{if(m.checked)l.selectedIds.add(g.id);else{l.selectedIds.delete(g.id);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1)}l.onToggle()}),b.appendChild(m),f.appendChild(b)}if(t.forEach(b=>{let m=document.createElement("td");if(b.render){let y=b.render(g[b.key],g);y instanceof HTMLElement?m.appendChild(y):m.innerHTML=y||""}else m.textContent=g[b.key]!==null&&g[b.key]!==void 0&&g[b.key]!==""?g[b.key]:"";b.nowrap&&(m.style.whiteSpace="nowrap"),f.appendChild(m)}),r||a||i||s.length>0){let b=document.createElement("td");b.className="actions-cell";let m=document.createElement("div");if(m.className="btn-group",i){let y=document.createElement("button");y.className="btn btn-xs btn-ghost",y.innerHTML="\u{1F441}",y.title="Lihat",y.addEventListener("click",()=>i(g)),m.appendChild(y)}if(r){let y=document.createElement("button");y.className="btn btn-xs btn-secondary",y.innerHTML="\u270F\uFE0F",y.title="Edit",y.addEventListener("click",()=>r(g)),m.appendChild(y)}s.forEach(y=>{let S=document.createElement("button");S.className=`btn btn-xs ${y.class||"btn-ghost"}`,S.innerHTML=y.icon||y.label,S.title=y.label,S.addEventListener("click",()=>y.handler(g)),m.appendChild(S)}),b.appendChild(m),f.appendChild(b)}c.appendChild(f)}),p.appendChild(c),n.appendChild(p),n}function We({page:t,pages:e,total:r,limit:a,onPage:i}){if(e<=1)return null;let s=document.createElement("div");s.className="pagination";let o=document.createElement("span");o.className="pagination-info",o.textContent=`Total: ${r} data`,s.appendChild(o);let l=document.createElement("div");l.className="pagination-btns";let n=(d,c,g=!1,f=!1)=>{let b=document.createElement("button");b.className=`btn btn-sm ${f?"btn-primary":"btn-ghost"} pagination-btn`,b.textContent=d,b.disabled=g,b.addEventListener("click",()=>i(c)),l.appendChild(b)};n("\xAB",1,t===1),n("\u2039",t-1,t===1);let p=Math.max(1,t-2),u=Math.min(e,t+2);for(let d=p;d<=u;d++)n(d,d,!1,d===t);return n("\u203A",t+1,t===e),n("\xBB",e,t===e),s.appendChild(l),s}xe();function He(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${He(e.fields)}</div>`;let r=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",i="";switch(e.type){case"textarea":i=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${r} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let o=(e.options||[]).map(d=>{let c=typeof d=="object"?d.value:d,g=typeof d=="object"?d.label:d,f=e.value==c?"selected":"";return`<option value="${c}" ${f}>${g}</option>`}).join("");i=`<select name="${e.name}" class="form-control" ${r}><option value="">-- Pilih ${e.label||""} --</option>${o}</select>`;break;case"combobox":let l=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,n=(e.options||[]).map(d=>{let c=typeof d=="object"?d.value:d,g=typeof d=="object"?d.label||d.value||"":d||"";return(g==="undefined"||g==="[object Object]"||g==="null")&&(g=""),g?`<option value="${g}"></option>`:""}).join(""),p=e.value||"";if(e.value){let d=(e.options||[]).find(c=>(typeof c=="object"?c.value:c)==e.value);if(d){let c=typeof d=="object"?d.label||d.value||"":d||"";c&&c!=="undefined"&&c!=="[object Object]"&&c!=="null"&&(p=c)}}i=`
          <input type="text" name="${e.name}" list="${l}" class="form-control" value="${p}" placeholder="Pilih atau ketik baru..." ${r} autocomplete="off">
          <datalist id="${l}">${n}</datalist>
        `;break;case"checkbox":i=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let u=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";i=`<input type="date" name="${e.name}" class="form-control" value="${u}" ${r}>`;break;case"number":i=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${r}>`;break;case"email":i=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${r}>`;break;case"url":i=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${r}>`;break;default:i=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${r} autocomplete="off">`}let s=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${i}${s}</div>`}).join("")}function Xe(t){let e={},r=new FormData(t);for(let[a,i]of r.entries())e[a]=i===""?null:i;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function Ze(t,e){e&&Object.entries(e).forEach(([r,a])=>{let i=t.querySelector(`[name="${r}"]`);i&&(i.hasAttribute("list")||(i.type==="checkbox"?i.checked=!!a:i.type==="date"&&a&&window.parseFlexibleDate?i.value=window.parseFlexibleDate(a):i.value=a??""))})}R();var ke={},It={on(t,e){ke[t]||(ke[t]=new Set),ke[t].add(e)},off(t,e){ke[t]&&ke[t].delete(e)},emit(t,e){ke[t]&&ke[t].forEach(r=>{try{r(e)}catch(a){console.warn("[calendarBus] Handler error:",a)}})},clear(){Object.keys(ke).forEach(t=>delete ke[t])}},_a=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function ge(t){if(!t){It.emit("data:changed",{module:"unknown"});return}let e=String(t).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");It.emit("data:changed",{module:e,relevant:_a.has(e)})}function A({container:t,title:e,icon:r,apiPath:a,columns:i,formFields:s,filterFields:o,defaultFilters:l={},itemLabel:n="Data",canCreate:p=!0,canEdit:u=!0,canDelete:d=!0,onBeforeSubmit:c,onAfterLoad:g,onDataLoaded:f,extraActions:b=[],initialSearch:m="",exportOptions:y=null,bulkDelete:S=!1,paginationMode:_="server"}){let C=ce();C&&C.role==="viewer"&&(p=!1,u=!1,d=!1,S=!1,y=null);let x=1,T={...l};m&&(T.search=m);let E=new Set;t.innerHTML=`
    ${S?`
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
      <h1 class="page-title">${r} ${e}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${p?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
        ${y?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${y.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${y.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${y.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${y.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
        ${o.filter(h=>h.type==="search").map(h=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${h.placeholder||"Cari..."}" id="filter-search" value="${T.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${o.filter(h=>h.type!=="search").map(h=>{if(h.type==="select"||h.type==="combobox"){let k=(h.label||"").startsWith("Pilih")?h.label:`Pilih ${h.label||""}`;return`<select class="filter-select" name="${h.name}" id="filter-${h.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${k}</option>${(h.options||[]).map(v=>`<option value="${typeof v=="object"?v.value:v}" ${T[h.name]===(typeof v=="object"?v.value:v)?"selected":""}>${typeof v=="object"?v.label:v}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(h=>h.type!=="search").map(h=>{if(h.type==="select"||h.type==="combobox"){let k=(h.label||"").startsWith("Pilih")?h.label:`Pilih ${h.label||""}`;return`<select class="filter-select filter-select-sheet" name="${h.name}-sheet" id="filter-sheet-${h.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${k}</option>${(h.options||[]).map(v=>`<option value="${typeof v=="object"?v.value:v}" ${T[h.name]===(typeof v=="object"?v.value:v)?"selected":""}>${typeof v=="object"?v.label:v}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function $(){let h=document.getElementById("bulk-toolbar");if(!h)return;let k=document.getElementById("bulk-count"),v=document.getElementById("btn-bulk-delete"),M=document.getElementById("btn-bulk-cancel");k.textContent=`${E.size} item dipilih`,E.size>0?(h.style.display="flex",v.disabled=!1,M.disabled=!1):(h.style.display="none",v.disabled=!0,M.disabled=!0);let N=document.getElementById("select-all-checkbox");if(N){let J=document.querySelectorAll(".row-checkbox");if(J.length>0){let D=[...J].every(ie=>ie.checked),L=[...J].some(ie=>ie.checked);N.checked=D,N.indeterminate=L&&!D}else N.checked=!1,N.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{E.clear(),document.querySelectorAll(".row-checkbox").forEach(k=>k.checked=!1);let h=document.getElementById("select-all-checkbox");h&&(h.checked=!1),$()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(E.size===0)return;let h=[...E],k=document.createElement("div");k.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",k.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${h.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${h.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(k),k.querySelector("#bulk-cancel-btn").addEventListener("click",()=>k.remove()),k.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let v=k.querySelector("#bulk-confirm-btn");v.disabled=!0,v.textContent="Menghapus...";let M=await w(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:h})});k.remove(),M.ok?(W(`${h.length} ${n} berhasil dihapus.`),E.clear(),$(),ge(a),j()):z(M.data?.error||"Gagal menghapus data.")})});let B=document.getElementById("filter-search"),P;B?.addEventListener("input",h=>{clearTimeout(P),P=setTimeout(()=>{T.search=h.target.value,x=1,E.clear(),$(),j()},400)}),o?.forEach(h=>{(h.type==="select"||h.type==="combobox")&&(document.getElementById(`filter-${h.name}`)?.addEventListener("change",k=>{T[h.name]=k.target.value;let v=document.getElementById(`filter-sheet-${h.name}`);v&&(v.value=k.target.value),x=1,E.clear(),$(),j()}),document.getElementById(`filter-sheet-${h.name}`)?.addEventListener("change",k=>{T[h.name]=k.target.value;let v=document.getElementById(`filter-${h.name}`);v&&(v.value=k.target.value),x=1,E.clear(),$(),j(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{T={},B&&(B.value=""),o?.forEach(h=>{let k=document.getElementById(`filter-${h.name}`);k&&(k.value="");let v=document.getElementById(`filter-sheet-${h.name}`);v&&(v.value="")}),x=1,E.clear(),$(),j()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{T={},B&&(B.value=""),o?.forEach(h=>{let k=document.getElementById(`filter-${h.name}`);k&&(k.value="");let v=document.getElementById(`filter-sheet-${h.name}`);v&&(v.value="")}),x=1,E.clear(),$(),j(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>ye(null)),y&&document.addEventListener("click",function(h){let k=document.getElementById("aksi-menu-main"),v=document.getElementById("btn-aksi-main");k&&v&&!v.contains(h.target)&&!k.contains(h.target)&&k.classList.remove("show-aksi-menu")});let Q=document.getElementById("btn-mobile-filter"),ee=document.getElementById("filter-options-wrapper"),re=document.getElementById("btn-close-filter-sheet");if(Q&&ee&&(Q.addEventListener("click",h=>{h.preventDefault(),ee.classList.add("sheet-open")}),re&&re.addEventListener("click",h=>{h.preventDefault(),ee.classList.remove("sheet-open")})),y){document.getElementById(`btn-export-${y.moduleName}`)?.addEventListener("click",async k=>{let v=k.target,M=v.innerHTML;v.innerHTML="\u23F3 Loading...",v.disabled=!0;try{await y.onExport()}catch{z("Gagal export data")}finally{v.innerHTML=M,v.disabled=!1}}),document.getElementById(`btn-template-${y.moduleName}`)?.addEventListener("click",()=>{y.onTemplate()});let h=document.getElementById(`input-import-${y.moduleName}`);h?.addEventListener("change",async k=>{let v=k.target.files[0];if(!v)return;let M=document.getElementById(`label-import-${y.moduleName}`),N=M?M.querySelector(".import-text"):null,J=N?N.innerText:"";N&&(N.innerText="\u231B Memproses..."),M&&(M.style.pointerEvents="none"),h.disabled=!0;try{let D=await Je(v);if(D.length===0)throw new Error("File kosong atau format salah");await y.onImport(D),W("Import berhasil!"),ge(a),j()}catch(D){z(D.message||"Gagal import data")}finally{N&&(N.innerText=J),M&&(M.style.pointerEvents="auto"),h.disabled=!1,h.value=""}})}async function j(){$();let h=document.getElementById("table-container");if(!h)return;h.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let k=_==="client",v=k?1:x,M=k?we:20,N=new URLSearchParams({page:v,limit:M,...Object.fromEntries(Object.entries(T).filter(([,q])=>q))}),J=await w(`${a}?${N}`);if(!J.ok){h.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${J.data?.error||"Error"}</p></div>`;return}let D=J.data?.data||J.data||[],L=J.data?.pagination,ie=D.length;if(k){D=f(D);let q=D.length,V=20,te=Math.ceil(q/V);x>te&&te>0&&(x=te);let O=(x-1)*V,le=x*V;D=D.slice(O,le),L={page:x,limit:V,total:q,pages:te}}!1,g&&g(D);let fe=Ye({columns:i,data:D,onEdit:u?q=>ye(q):null,actions:b.map(q=>({...q,handler:V=>q.handler(V,j)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:S?{selectedIds:E,onToggle:$}:null});h.innerHTML="",h.appendChild(fe);let se=document.getElementById("pagination-container");if(se&&(se.innerHTML="",L&&L.pages>1)){let q=We({page:L.page,pages:L.pages,total:L.total,limit:L.limit,onPage:V=>{x=V,j()}});q&&se.appendChild(q)}}function Ce(h){let k=typeof s=="function"?s(h):s;return He(k)}function ye(h){let k=!!h,v=document.createElement("form");if(v.noValidate=!0,v.innerHTML=Ce(h),k){let N=typeof s=="function"?s(h):s;Ze(v,h)}let{close:M}=de({title:k?`Edit ${n}`:`Tambah ${n}`,content:v,size:"lg",confirmText:k?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(N,J)=>{if(!v.reportValidity())return;let D=N.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let L=Xe(v),ie=typeof s=="function"?s(h):s,fe=async te=>{for(let O of te)if(O.type==="row")await fe(O.fields);else if(O.type==="combobox"&&L[O.name]){let le=L[O.name],ve=(O.options||[]).find(Y=>{let ae=String(typeof Y=="object"?Y.value:Y),st=String(typeof Y=="object"?Y.label:Y);return ae===le||st===le});if(ve)L[O.name]=typeof ve=="object"?ve.value:ve;else if(O.createApi){let Y={};Y[O.createApi.field]=le,O.createApi.extra&&Object.assign(Y,O.createApi.extra);let ae=await w(O.createApi.path,{method:"POST",body:JSON.stringify(Y)});if(ae.ok&&ae.data?.id)L[O.name]=ae.data.id;else if(ae.ok&&!ae.data?.id)L[O.name]=le;else throw new Error(`Gagal membuat master data: ${ae.data?.error||"Unknown error"}`)}}};try{await fe(ie)}catch(te){z(te.message),D.disabled=!1,D.textContent=k?"Simpan Perubahan":`Tambah ${n}`;return}c&&(L=await c(L,h));let se=k?"PUT":"POST",q=k?`${a}/${h.id}`:a,V=await w(q,{method:se,body:JSON.stringify(L)});V.ok?(W(k?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),J(),ge(a),j()):(z(V.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=k?"Simpan Perubahan":`Tambah ${n}`)}})}function va(h){qe(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let k=await w(`${a}/${h.id}`,{method:"DELETE"});k.ok?(W(`${n} berhasil dihapus.`),ge(a),j()):z(k.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return j(),j}F();F();var Te=null,et=null,tt=null;async function De(t=!1){if(Te&&!t)return console.log("Employees Raw (Cache Hit)",Te.slice(0,5)),Te;let e=await w(`/api/employees?limit=${we}&status=Aktif`);return Te=(e.data?.data||[]).map(r=>({value:r.id,label:r.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",Te.slice(0,5)),Te}async function X(t=!1){let r=(await De(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",r.slice(0,5)),r}async function K(t=!1){return et&&!t||(et=((await w("/api/branches?all=1")).data?.data||[]).map(r=>({value:r.id,label:r.full_name}))),et}async function _e(t=!1){return tt&&!t||(tt=((await w("/api/vendors?all=1")).data?.data||[]).map(r=>({value:r.id,label:r.name}))),tt}function H(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function ut(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function Ie(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function gt(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function pe(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}R();function bt(t,e){let r=new Date,a=`${r.getFullYear()}-${String(r.getMonth()+1).padStart(2,"0")}`;if(!(t.completion_date||t.target_date||t.opening_date||"").startsWith(a))return!1;let s=String(t.status||"").toLowerCase();if(s!=="selesai"&&s!=="completed"&&s!=="done")return!1;let o=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?o.includes("inspeksi"):e==="gcdc"?o.includes("general cleaning")||o.includes("deep cleaning"):!1}F();R();function Bt(t,e){let r=String(t.status||"").toLowerCase();return e==="active"?r==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&r==="aktif":!1}F();R();function ht(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let i=new Date(a);i.setDate(a.getDate()+30);let s=new Date(t.end_date);return s.setHours(0,0,0,0),s>=a&&s<=i}return!1}F();R();function Pt(t,e){let r=String(t.status||"").toLowerCase();return e==="open"?r==="open":!1}F();function Lt(t,e){let r=String(t.status||"").toLowerCase();return e==="pending"?r==="pending":!1}var Se={};function Le(t){if(Se[t]){try{Se[t].destroy()}catch{}delete Se[t]}}function Ca(){Object.keys(Se).forEach(Le)}var be=(t,e=0)=>{let r=Number(t);return isNaN(r)||t===null||t===void 0?e:r},Be=(t,e="\u2014")=>{if(t==null||t==="")return e;let r=String(t).trim();return r===""||r==="[object Object]"?e:r};function Ft(t,e,r=900){if(!t)return;let a=Math.max(0,Math.round(be(e)));if(a===0){t.textContent="0";return}let i=Date.now(),s=()=>{let o=Math.min((Date.now()-i)/r,1),l=1-Math.pow(1-o,3);t.textContent=Math.round(l*a).toLocaleString("id-ID"),o<1?requestAnimationFrame(s):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(s)}var Ea={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},$a=t=>{let e=Be(t,"\u2014");return`<span class="status-pill ${Ea[e]||"pill-neutral"}">${e}</span>`};var me={family:"Inter",size:11},he="#94A3B8",Pe="#F1F5F9",yt=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],Ta=()=>window.innerWidth<768;function at(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:Ta()?"bottom":"top",labels:{font:me,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:me,titleFont:{...me,weight:"700"}}},scales:{x:{grid:{color:Pe},ticks:{font:me,color:he,maxRotation:0}},y:{grid:{color:Pe},ticks:{font:me,color:he},beginAtZero:!0}},...t}}var Da=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join("");function At(t=3){return Array(t).fill(0).map((e,r)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${r<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function ne(t,e,r=8e3){try{let a=new AbortController,i=setTimeout(()=>a.abort(),r),s=await w(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(i),!s||!s.ok)return e;let o=s.data;return o?o.data!==void 0?o.data??e:o:e}catch{return e}}function Ba(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(a=>{let i=document.getElementById(a);i&&(i.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(a=>{let i=document.getElementById(a);if(i&&i.style.display==="none"){i.style.display="block";let s=i.parentElement;if(s&&!s.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent="Belum ada data",i.style.display="none",s.appendChild(o)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Mt({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Ot({}),["table-contracts","table-issues"].forEach(a=>{let i=document.getElementById(a);i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada data</div>')});let r=document.getElementById("activity-log");r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Nt(t){Ca(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">


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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${At(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${At(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>ft(t)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async r=>{let a=r.target.value,i=document.getElementById("jadwal-year-label");i&&(i.textContent=a);let s=document.getElementById("skel-jadwal"),o=document.getElementById("chart-jadwal");s&&(s.style.display="block",s.style.position="absolute"),o&&(o.style.display="none");let l=await ne(`/api/dashboard/schedule-chart?year=${a}`,{},8e3);try{Rt(l)}catch(n){console.warn("ScheduleChart render:",n),ue("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async r=>{let a=r.target.value,i=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",s=document.getElementById("skel-insp"),o=document.getElementById("chart-insp");s&&(s.style.display="block",s.style.position="absolute"),o&&(o.style.display="none");let l=await ne(i,{},8e3);try{Kt(l)}catch(n){console.warn("InspBar render:",n),ue("skel-insp","chart-insp")}}),t._skelTimeout=setTimeout(()=>Ba(),5e3),await ft(t),t._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?ft(t):clearInterval(t._dashRefresh)},6e4)}async function ft(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,r,a,i,s,o,l,n,p,u,d,c,g,f]=await Promise.all([ne("/api/dashboard/kpi",{},8e3),ne("/api/dashboard/issues-trend",{},8e3),ne("/api/dashboard/issues-summary",{},8e3),ne("/api/dashboard/stats",{},8e3),ne("/api/dashboard/calendar",[],8e3),ne("/api/schedule?limit=10000",{data:[]},8e3),ne("/api/employees?limit=10000",{data:[]},8e3),ne("/api/contracts?limit=10000",{data:[]},8e3),ne("/api/issues?limit=10000",{data:[]},8e3),ne("/api/one-on-one?limit=10000",{data:[]},8e3),ne("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),ne(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),ne("/api/relievers?limit=10000",{data:[]},8e3),ne("/api/reports/fogging?limit=10000",{data:[]},8e3)]),b=document.getElementById("filter-insp-month"),m=b?b.value:"",y=m?`/api/dashboard/inspection-bar?month=${m}`:"/api/dashboard/inspection-bar",S=await ne(y,{},8e3);if(e){let _=Array.isArray(o?.data)?o.data:Array.isArray(o)?o:[];window.dashboardSchedules=_;let C=Array.isArray(l?.data)?l.data:Array.isArray(l)?l:[],x=Array.isArray(n?.data)?n.data:Array.isArray(n)?n:[],T=Array.isArray(p?.data)?p.data:Array.isArray(p)?p:[],E=Array.isArray(u?.data)?u.data:Array.isArray(u)?u:[],$=Array.isArray(g?.data)?g.data:Array.isArray(g)?g:[];window.dashboardRelievers=$;let B=Array.isArray(f?.data)?f.data:Array.isArray(f)?f:[];if(window.dashboardFogging=B,e.employees&&(e.employees.current=C.filter(P=>Bt(P,"active")).length),e.contracts&&(e.contracts.current=x.filter(P=>ht(P,"active")).length),e.expiring30&&(e.expiring30={current:x.filter(P=>ht(P,"expiring30")).length}),e.issues&&(e.issues.current=T.filter(P=>Pt(P,"open")).length),e.one_on_one&&(e.one_on_one.current=E.filter(P=>Lt(P,"pending")).length),e.schedule){let P=`Q${Math.ceil((new Date().getMonth()+1)/3)}`;e.schedule.current=_.filter(Q=>{if(Q.period===P)return!0;if(Q.target_date){let ee=Q.target_date.split("-");if(ee.length>=2){let re=parseInt(ee[1],10);return re&&`Q${Math.ceil(re/3)}`===P}}return!1}).length}e.inspection_month&&(e.inspection_month.current=_.filter(P=>bt(P,"inspeksi")).length),e.cleaning_month&&(e.cleaning_month.current=_.filter(P=>bt(P,"gcdc")).length)}try{Mt(e)}catch(_){console.warn("KPI render:",_)}try{Ot(e)}catch(_){console.warn("MiniStats render:",_)}try{Rt(c)}catch(_){console.warn("ScheduleChart render:",_),ue("skel-jadwal","chart-jadwal")}try{Pa(Array.isArray(a?.by_category)?a.by_category:[])}catch(_){console.warn("Donut render:",_),ue("skel-donut","chart-donut")}try{La(r)}catch(_){console.warn("Trend render:",_),ue("skel-trend","chart-trend")}try{Kt(S)}catch(_){console.warn("InspBar render:",_),ue("skel-insp","chart-insp")}try{let _=Array.isArray(i)?i:Array.isArray(i?.recent_issues)?i.recent_issues:[];Fa(_)}catch(_){console.warn("IssuesTable render:",_)}try{let _=Array.isArray(i?.expiring_contracts)?i.expiring_contracts:[];Aa(d)}catch(_){console.warn("ContractsTable render:",_)}try{Na(Array.isArray(s)?s:[])}catch(_){console.warn("Agenda render:",_)}try{Ma()}catch(_){console.warn("Quick Actions render:",_)}}function Mt(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let r=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=r.map(a=>{let i=be(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${i}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${i}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{Ft(a,parseInt(a.dataset.target)||0)})}function Ot(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let r=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,a=new Date().getFullYear(),i=String(new Date().getMonth()+1).padStart(2,"0"),s=`${a}-${i}`,o=d=>`
    <select id="${d}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.preventDefault(); event.stopPropagation();">
      <option value="${a}-01" ${s===`${a}-01`?"selected":""}>Jan</option>
      <option value="${a}-02" ${s===`${a}-02`?"selected":""}>Feb</option>
      <option value="${a}-03" ${s===`${a}-03`?"selected":""}>Mar</option>
      <option value="${a}-04" ${s===`${a}-04`?"selected":""}>Apr</option>
      <option value="${a}-05" ${s===`${a}-05`?"selected":""}>Mei</option>
      <option value="${a}-06" ${s===`${a}-06`?"selected":""}>Jun</option>
      <option value="${a}-07" ${s===`${a}-07`?"selected":""}>Jul</option>
      <option value="${a}-08" ${s===`${a}-08`?"selected":""}>Agu</option>
      <option value="${a}-09" ${s===`${a}-09`?"selected":""}>Sep</option>
      <option value="${a}-10" ${s===`${a}-10`?"selected":""}>Okt</option>
      <option value="${a}-11" ${s===`${a}-11`?"selected":""}>Nov</option>
      <option value="${a}-12" ${s===`${a}-12`?"selected":""}>Des</option>
    </select>
  `,l=[{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.preventDefault(); event.stopPropagation();">
          <option value="Q1" ${r==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${r==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${r==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${r==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:t.schedule?.current,href:`#/timeline?dash_filter=period_${r.toLowerCase()}`,color:"mini-blue"},{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:o("dash-inspeksi-month"),val:t.inspection_month?.current,href:`#/timeline?dash_filter=inspeksi&month=${s}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:o("dash-gcdc-month"),val:t.cleaning_month?.current,href:`#/timeline?dash_filter=gcdc&month=${s}`,color:"mini-green"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:o("dash-reliefer-month"),val:t.reliever_completed?.current,href:`#/relievers?dash_filter=reliever&month=${s}`,color:"mini-teal"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:o("dash-fogging-month"),val:t.fogging_month?.current,href:`#/reports/fogging?dash_filter=fogging&month=${s}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=l.map(d=>`
    <a href="${d.href}" class="mini-stat ${d.color}" style="text-decoration:none" id="${d.id||""}">
      <div class="mini-stat-icon">${d.icon}</div>
      <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
        <div style="display:flex; align-items:baseline; gap:3px;">
          <div class="mini-stat-value" data-target="${be(d.val)}">0</div>
          ${d.dropdown?d.dropdown:""}
        </div>
        <div class="mini-stat-text">${d.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(d=>Ft(d,parseInt(d.dataset.target)||0,700));let n=document.getElementById("dash-jadwal-period");n&&n.addEventListener("change",d=>{let c=d.target.value,g=(window.dashboardSchedules||[]).filter(m=>{if(m.period===c)return!0;if(m.target_date){let y=m.target_date.split("-");if(y.length>=2){let S=parseInt(y[1],10);return S&&`Q${Math.ceil(S/3)}`===c}}return!1}).length,f=document.querySelector("#mini-jadwal .mini-stat-value");f&&(f.dataset.target=g,f.textContent=g);let b=document.getElementById("mini-jadwal");b&&(b.href=`#/timeline?dash_filter=period_${c.toLowerCase()}`)});let p=(d,c,g,f,b)=>{let m=document.getElementById(d);if(m){let y=S=>{let _=(g||[]).filter(T=>f(T,S)).length,C=document.querySelector(`#${c} .mini-stat-value`);C&&(C.dataset.target=_,C.textContent=_);let x=document.getElementById(c);x&&(x.href=`${b}&month=${S}`)};y(m.value),m.addEventListener("change",S=>y(S.target.value))}},u=d=>{let c=String(d.status||"").toLowerCase();return c==="done"||c==="selesai"||c==="completed"};p("dash-reliefer-month","mini-reliefer",window.dashboardRelievers,(d,c)=>window.parseFlexibleDate(d.backup_date).startsWith(c)&&u(d),"#/relievers?dash_filter=reliever"),p("dash-inspeksi-month","mini-inspeksi",window.dashboardSchedules,(d,c)=>d.activity_type==="Inspeksi Hygiene"&&u(d)&&window.parseFlexibleDate(d.completion_date||d.target_date).startsWith(c),"#/timeline?dash_filter=inspeksi"),p("dash-gcdc-month","mini-gcdc",window.dashboardSchedules,(d,c)=>(d.activity_type==="General Cleaning"||d.activity_type==="Deep Cleaning")&&u(d)&&window.parseFlexibleDate(d.completion_date||d.target_date).startsWith(c),"#/timeline?dash_filter=gcdc"),p("dash-fogging-month","mini-fogging",window.dashboardFogging,(d,c)=>u(d)&&window.parseFlexibleDate(d.activity_date).startsWith(c),"#/reports/fogging?dash_filter=fogging")}function Pa(t){ue("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),r=document.getElementById("donut-legend");if(!e||!r)return;Le("donut");let a=(t||[]).filter(n=>be(n.count)>0);if(!a.length){Ue(e,"Belum ada data permasalahan");return}let i=a.map(n=>`${Be(n.category,"Lainnya")}`),s=a.map(n=>be(n.count)),o=s.reduce((n,p)=>n+p,0);r.innerHTML=a.map((n,p)=>{let u=yt[p%yt.length],d=o>0?Math.round(n.count/o*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${u}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${d}%)</span></div>
          <div class="donut-legend-label">${i[p]}</div>
        </div>
      </div>
    `}).join("");let l={id:"centerText",beforeDraw:function(n){let p=n.width,u=n.height,d=n.ctx;d.restore();let c=(u/80).toFixed(2);d.font="bold "+c+"em Inter",d.textBaseline="middle",d.fillStyle="#1E293B";let g=o.toString(),f=Math.round((p-d.measureText(g).width)/2),b=u/2;d.fillText(g,f,b-4),d.font="600 "+(c*.35).toFixed(2)+"em Inter",d.fillStyle="#64748B";let m="Total",y=Math.round((p-d.measureText(m).width)/2);d.fillText(m,y,b+10),d.save()}};Se.donut=new Chart(e,{type:"doughnut",data:{labels:i,datasets:[{data:s,backgroundColor:yt,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:me,titleFont:{...me,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutout:"75%"},plugins:[l]})}function La(t){ue("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;Le("trend"),t=t||{};let r=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(t.labels||[]).map(o=>{if(!o||typeof o!="string")return"";try{let[l,n]=o.split("-");return(r[Number(n)-1]||n)+" "+String(l).slice(-2)}catch{return o}}),i=(t.open||[]).map(o=>be(o)),s=(t.closed||[]).map(o=>be(o));if(!a.length){Ue(e,"Belum ada data trend");return}Se.trend=new Chart(e,{type:"line",data:{labels:a,datasets:[{label:"Open",data:i,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:s,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:at({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:he,maxRotation:45,autoSkip:!0}},y:{grid:{color:Pe},ticks:{font:{family:"Inter",size:9},color:he},beginAtZero:!0}}})})}function Rt(t){ue("skel-jadwal","chart-jadwal");let e=document.getElementById("chart-jadwal");if(!e)return;Le("jadwal"),t=t||{};let r=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(t).some(n=>Array.isArray(n)&&n.some(p=>p>0))){Ue(e,"Belum ada data jadwal");return}let i=t["Inspeksi Hygiene"]||Array(12).fill(0),s=t["General Cleaning"]||Array(12).fill(0),o=t["Deep Cleaning"]||Array(12).fill(0),l=t.Fogging||Array(12).fill(0);Se.jadwal=new Chart(e,{type:"bar",data:{labels:r,datasets:[{label:"Inspeksi",data:i,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:s,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:o,backgroundColor:"#F59E0B"},{label:"Fogging",data:l,backgroundColor:"#EF4444"}]},options:at({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:he,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Pe},ticks:{font:{family:"Inter",size:9},color:he},min:0}}})})}function Kt(t){ue("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;Le("inspBar"),t=t||{};let r=t.labels||[],a=(t.fc||[]).map(s=>be(s)),i=(t.spv||[]).map(s=>be(s));if(!r.length){Ue(e,"Belum ada data inspeksi");return}Se.inspBar=new Chart(e,{type:"bar",data:{labels:r,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:i,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:at({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:me,color:he,maxRotation:45,minRotation:45}},y:{grid:{color:Pe},ticks:{font:me,color:he},min:0,max:100}}})})}function Aa(t){ue("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;Le("contractMiniBar"),t=t||{};let r={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(o=>{let l=o.split("-")[1];return r[l]||o}),i=(t.data||[]).map(o=>be(o));if(!a.length){Ue(e,"Belum ada data");return}let s=e.getContext("2d");Se.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:i,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:at({onClick:(o,l)=>{if(l&&l.length>0){let n=l[0].index,p=(t.labels||[])[n];p&&(window.location.hash="#/contracts?month_expiry="+p)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:me,color:he,maxRotation:0,autoSkip:!1}},y:{grid:{color:Pe,borderDash:[4,4],drawBorder:!1},ticks:{font:me,color:he,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function Fa(t){let e=document.getElementById("table-issues");if(!e)return;let r=(t||[]).slice(0,8);if(!r.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${r.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${$a(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Be(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Be(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function Na(t){let e=document.getElementById("widget-agenda");if(!e)return;let r=new Date,a=`${r.getFullYear()}-${String(r.getMonth()+1).padStart(2,"0")}-${String(r.getDate()).padStart(2,"0")}`,s=(t||[]).filter(o=>(o.event_date||"").startsWith(a)).slice(0,10);if(!s.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${s.map(o=>{let l="#3B82F6",n="#EFF6FF",p="Agenda",u=(o.title||"").toLowerCase();return u.includes("inspeksi")?(l="#10B981",n="#ECFDF5",p="Inspeksi"):u.includes("cleaning")||u.includes("gcdc")?(l="#3B82F6",n="#EFF6FF",p="Cleaning"):u.includes("reliefer")?(l="#F59E0B",n="#FFFBEB",p="Reliefer"):u.includes("fogging")&&(l="#8B5CF6",n="#F5F3FF",p="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(o.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${l};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Be(o.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Be(o.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${l}">${p}</div>
        </div>
      `}).join("")}
    </div>
  `}function Ma(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(r=>`
    <a href="${r.href}" class="action-btn">
      <div class="action-icon" style="background:${r.bg}">${r.icon}</div>
      ${r.label}
    </a>
  `).join("")}function ue(t,e){let r=document.getElementById(t),a=document.getElementById(e);if(r&&(r.style.display="none",r.style.position=""),a){a.style.display="block";let i=a.parentElement;if(i){let s=i.querySelector(".chart-empty");s&&s.remove()}}}function Ue(t,e="Belum ada data"){if(!t)return;t.style.display="none";let r=t.parentElement;if(!r)return;if(!r.querySelector(".chart-empty")){let i=document.createElement("div");i.className="chart-empty",i.textContent=e,r.appendChild(i)}}F();async function jt(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),r=document.getElementById("login-error"),a=document.getElementById("login-btn"),i=document.getElementById("toggle-password"),s=document.getElementById("login-password");i?.addEventListener("click",()=>{let o=s.type==="text";s.type=o?"password":"text",i.style.color=o?"":"var(--primary)"}),e?.addEventListener("submit",async o=>{o.preventDefault(),r.style.display="none";let l=e.username.value.trim(),n=e.password.value;if(!l||!n){r.textContent="Username dan password wajib diisi.",r.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let p=await w("/api/auth/login",{method:"POST",body:JSON.stringify({username:l,password:n})});p.ok&&p.data.success?(pt(p.data.data.token),Ke(p.data.data.user),W("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(r.textContent=p.data.error||"Username atau password salah.",r.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{r.textContent="Gagal terhubung ke server. Periksa koneksi internet.",r.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}F();R();async function Oa(){return await K()}function Ra(t,e){let r=String(t.status||"").toLowerCase();return e==="active"?r==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&r==="aktif":!1}async function qt(t,e){let r=await Oa(),a=e?e.get("dash_filter"):null;A({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",enableMobileFilterSheet:!0,itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:i=>a?i.filter(s=>Ra(s,a)):i,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:i=>Ie(i)},{key:"phone",label:"No. HP",render:i=>i?`<a href="tel:${i}">${i}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:i=>window.formatDate(i)},{key:"status",label:"Status",render:i=>H(i)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:i=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:i?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:i?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:r,value:i?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:i?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:i?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:i?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:i?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let i=await w(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let s=i.data.data.map(o=>({"Nama Lengkap":o.full_name,Cabang:o.branch_name||"",Divisi:o.division||"","No. HP":o.phone||"","Tgl Masuk":o.join_date||"",Status:o.status||""}));I(s,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{I([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async i=>{let s=n=>{if(!n)return null;let p=String(n||"").toLowerCase(),u=r.find(d=>String(d.label||"").toLowerCase()===p);return u?u.value:null},o=i.map(n=>({full_name:String(n["Nama Lengkap"]||"").trim(),branch_id:s(String(n.Cabang||"").trim()),division:String(n.Divisi||"").trim()||"FACILITY CARE",phone:String(n["No. HP"]||"").trim(),join_date:String(n["Tgl Masuk"]||"").trim(),status:String(n.Status||"").trim(),notes:String(n.Catatan||"").trim()})).filter(n=>n.full_name),l=await w("/api/import/employees",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}F();R();var kt=[],Ht=[];async function Ka(){kt=await K(),Ht=await De()}var vt=async t=>{let e=[],r=1;for(;;){let i=await(await Promise.resolve().then(()=>(F(),Ee))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${r}`);if(!i.ok)break;let s=i.data?.data||i.data||[],o=Array.isArray(s)?s:[];if(e=e.concat(o),o.length<100||i.data?.pagination&&r>=i.data.pagination.pages)break;r++}return e};async function nt(t,e){await Ka(),A({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",defaultFilters:{},onDataLoaded:a=>a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>Ie(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,i)=>i.end_date&&String(i.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':ut(a)},{key:"status",label:"Status",render:a=>H(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:kt},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[s,o]=await Promise.all([vt("/api/employees?status=Aktif"),vt("/api/contracts")]);if(s.length>0){let l=o.filter(d=>d.status==="Aktif"),n=new Set(l.map(d=>d.employee_id)),p=s.filter(d=>!n.has(d.id)),u=`<p style="margin-bottom:12px">Data yang terbaca: <b>${s.length}</b> Karyawan Aktif, dan <b>${l.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${p.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;p.forEach(d=>{let c=o.filter(f=>f.employee_id===d.id),g='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(c.length>0){let f=c[0];g=`Pernah ada kontrak (Status: <b style="color:#EF4444">${f.status}</b>, Selesai: ${window.formatDate(f.end_date)})`}u+=`<li style="margin-bottom:8px"><b>${d.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${d.branch_name||"-"} | ${g}</span></li>`}),u+="</ul>",Promise.resolve().then(()=>(xe(),Dt)).then(d=>d.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:u,cancelText:"Tutup"}))}}catch(s){console.error(s)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let i=document.querySelector(".page-actions");i&&i.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:Ht,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:kt,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await w(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let i=a.data.data.map(s=>({"Nama Lengkap":s.employee_name,Cabang:s.branch_name||"","Div / Bagian":s.division||"","Tanggal Mulai":s.start_date||"","Tanggal Selesai":s.end_date&&String(s.end_date).startsWith("2099")?"":s.end_date||"","Sisa Kontrak":s.end_date&&String(s.end_date).startsWith("2099")?"Tetap":s.days_remaining!==null&&s.days_remaining!==void 0?`${s.days_remaining} Hari`:"",Status:s.status||""}));I(i,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{I([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[i,s]=await Promise.all([w("/api/branches?limit=10000"),vt("/api/employees")]),o=i.data?.data||[],l=s||[];console.log(`Total employee yang berhasil dimuat dari database : ${l.length}`),l.length>0&&(console.log("Contoh 5 employee pertama:"),l.slice(0,5).forEach((m,y)=>{console.log(`${y+1}. ID: ${m.id}, Name: ${m.full_name}, Status: ${m.status}`)}));let n=m=>{if(!m)return null;let y=String(m||"").replace(/\s+/g," ").toLowerCase().trim(),S=o.find(_=>String(_.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===y||String(_.code||"").replace(/\s+/g," ").toLowerCase().trim()===y||String(_.name||"").replace(/\s+/g," ").toLowerCase().trim()===y);return S?S.id:null},p=(m,y)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${y}`),console.log(`Nama dari Excel : "${m}"`),!m)return console.log("Alasan gagal mapping : Nama kosong"),null;let S=String(m||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${S}"`),console.log(`Jumlah employee di database : ${l.length}`);let _=l.find(C=>String(C.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===S);return _?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${_.id}`),_.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},u=m=>{if(!m)return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let y=String(m).trim();if(/^\d{4,5}(\.\d+)?$/.test(y)){let _=Math.floor(Number(y));if(_>2e4&&_<99999){let C=new Date(Date.UTC(1899,11,30)+_*864e5);return isNaN(C.getTime())?"":C.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);let S=y.split(/[\/\-\.]/);if(S.length===3){let[_,C,x]=S.map(T=>T.trim());if(_.length===4&&C.length<=2&&x.length<=2)return`${_}-${C.padStart(2,"0")}-${x.padStart(2,"0")}`;if(x.length===4&&C.length<=2&&_.length<=2)return`${x}-${C.padStart(2,"0")}-${_.padStart(2,"0")}`}return y},d=a.map((m,y)=>{let S=y+2,_=String(m["Nama Lengkap"]||"").trim(),C=m["Tanggal Mulai"],x=u(C);if(!x){let $=a.__worksheet,B=a.__headers||[],P=B.indexOf("Tanggal Mulai"),Q="N/A",ee="N/A",re="N/A";if(P!==-1&&$&&window.XLSX){let Ce=window.XLSX.utils.encode_cell({c:P,r:S-1});re=Ce;let ye=$[Ce];ye?(Q=ye.t||"undefined",ee=ye.w||"undefined"):Q="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let j="Unknown";C==null||C===""?j="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":C instanceof Date&&isNaN(C.getTime())?j="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":j="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${S}`),console.log(`Employee Name : ${_}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${P})`),console.log(`Raw Cell Value : "${C}"`),console.log(`JavaScript Type : ${typeof C}`),console.log(`SheetJS Cell Type : ${Q}`),console.log(`SheetJS Formatted Value : "${ee}"`),console.log(`Value After Trim : "${String(C||"").trim()}"`),console.log(`Value After Date Parser : "${x}"`),console.log(`Is Empty : ${!C}`),console.log(`Is Invalid Date : ${C instanceof Date?isNaN(C.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${j}`),console.log(`Workbook Sheet : ${$?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${re}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(m,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(B)),console.log(`==========================
`)}let T=p(_,S),E=null;return T?x||(E="Tanggal Mulai kosong atau tidak berformat tanggal"):E="Karyawan tidak ditemukan di Database",{isValid:!!(T&&x),invalidReason:E,rowNum:S,data:{employee_id:T,branch_id:n(String(m.Cabang||"").trim()),division:String(m["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:x,end_date:u(m["Tanggal Selesai"])||"2099-12-31",status:String(m.Status||"").trim(),_rawName:_}}}),c=[],g=[];if(d.forEach(m=>{m.isValid?c.push(m.data):g.push({rowNum:m.rowNum,name:m.data._rawName,reason:m.invalidReason})}),console.log(`Split Validation - Valid: ${c.length}, Invalid: ${g.length}`),c.length===0){let m=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${a.length}
Valid: 0
Invalid: ${g.length}

Daftar Kegagalan (Contoh):
`;g.slice(0,10).forEach(y=>{m+=`- Row ${y.rowNum} | Nama: ${y.name} | Alasan: ${y.reason}
`}),g.length>10&&(m+=`- ... dan ${g.length-10} lainnya.
`),alert(m);return}let f=await w("/api/contracts/import",{method:"POST",body:JSON.stringify(c)}),b=`IMPORT SUMMARY
======================
`;b+=`Total Baris Excel : ${a.length}
`,b+=`Baris Valid       : ${c.length}
`,b+=`Baris Invalid     : ${g.length}

`,f&&f.data&&f.data.metrics?(b+=`Berhasil INSERT   : ${f.data.metrics.inserted}
`,b+=`Berhasil UPDATE   : ${f.data.metrics.updated}
`):b+=`Berhasil diproses : ${c.length}
`,g.length>0&&(b+=`
DAFTAR DATA DILEWATI:
`,g.forEach(m=>{b+=`- Row ${m.rowNum} | ${m.name} | ${m.reason}
`})),alert(b),typeof nt=="function"&&nt()}}})}F();R();var St=[],Ge=[];function ja(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let r of e)if(t.some(a=>a.period===r))return r;return"Q3"}async function Jt(t,e){St=await K();let r=await X();Ge=["Berlin Ariansyah","Ade Surahman"];let a=m=>m&&!Ge.find(y=>String(typeof y=="object"?y.value:y).toLowerCase()===String(m).toLowerCase())?[...Ge,m]:Ge,i=await w(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),s=m=>{if(!m||m==="-"||String(m).trim()==="")return"";let y=String(m).split("-");return y.length===3&&y[0].length===4?`${y[2]}-${y[1]}-${y[0]}`:m},o=i.data?.data||[],l=ja(o),n=e?e.get("dash_filter"):null,p=new Date,u=`${p.getFullYear()}-${String(p.getMonth()+1).padStart(2,"0")}`,d={},c=e&&e.get("month")?e.get("month"):null;n==="inspeksi"?d={status:"Done",activity_type:"Inspeksi Hygiene",month:c}:n==="gcdc"?d={status:"Done",activity_type:"GCDC",month:c}:n&&n.startsWith("period_")&&(d={period:n.replace("period_","").toUpperCase()});let g=new Date().getFullYear(),b=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((m,y)=>{let S=String(y+1).padStart(2,"0");return{value:`${g}-${S}`,label:`${m} ${g}`}});A({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:d,onDataLoaded:m=>m.sort((y,S)=>{let _=y.opening_date?new Date(y.opening_date).getTime():0;return(S.opening_date?new Date(S.opening_date).getTime():0)-_}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:m=>gt(m)},{key:"period",label:"Periode",render:m=>pe(m)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:m=>s(m)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:m=>s(m)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:m=>s(m)},{key:"status",label:"Status",render:m=>H(m)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:St},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:b},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Ge}],formFields:m=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:St,value:m?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:m?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:m?.period},{name:"pic",label:"PIC",type:"combobox",options:a(m?.pic),value:m?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:m?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:m?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:m?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:m?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:m?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let m=await w(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(m.ok){let y=m.data.data.map(S=>({Cabang:S.branch_name||"",Kegiatan:S.activity_type||"",Periode:S.period||"",PIC:S.pic||"","Tgl Opening":S.opening_date||"","Tgl Target":S.target_date||"","Tgl Selesai":S.completion_date||"",Status:S.status||""}));I(y,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{I([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async m=>{let S=(await w("/api/branches?all=1")).data?.data||[],_=E=>{if(!E)return null;let $=String(E||"").toLowerCase(),B=S.find(P=>String(P.full_name||"").toLowerCase()===$||String(P.code||"").toLowerCase()===$||String(P.name||"").toLowerCase()===$);return B?B.id:null},C=E=>{if(E==null||E==="")return"";if(E instanceof Date&&!isNaN(E.getTime()))return E.toISOString().slice(0,10);let $=String(E).trim();if($===""||$==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test($))return $.slice(0,10);if(/^\d{4,5}$/.test($)){let P=Number($);if(P>2e4&&P<99999){let Q=new Date(Date.UTC(1899,11,30)+P*864e5);return isNaN(Q.getTime())?"":Q.toISOString().slice(0,10)}}let B=$.split(/[\/\-\.]/);if(B.length===3){let[P,Q,ee]=B.map(re=>re.trim());if(P.length===4&&Q.length<=2&&ee.length<=2)return`${P}-${Q.padStart(2,"0")}-${ee.padStart(2,"0")}`;if(ee.length===4&&Q.length<=2&&P.length<=2)return`${ee}-${Q.padStart(2,"0")}-${P.padStart(2,"0")}`}return $},x=m.map(E=>({branch_id:_(String(E.Cabang||"").trim()),activity_type:String(E.Kegiatan||"").trim(),period:String(E.Periode||"").trim(),pic:String(E.PIC||E.Pic||"").trim(),opening_date:C(E["Tgl Opening"]||E["Tanggal Opening"]||E["Tgl Openir"]),target_date:C(E["Tgl Target"]||E["Tanggal Target"]),completion_date:C(E["Tgl Selesai"]||E["Tanggal Selesai"]),status:String(E.Status||"").trim(),notes:String(E.Catatan||E.Keterangan||"").trim()})).filter(E=>E.activity_type&&E.period),T=await w("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:x,onDuplicate:"update"})});if(!T.ok)throw new Error(T.data?.error||"Import gagal");return T.data}}})}F();R();var wt=[],it=[],Ut=[];function qa(t,e){let r=String(t.status||"").toLowerCase();return e==="open"?r==="open":!1}async function Gt(t,e){let r=e?e.get("dash_filter"):null;wt=await K(),it=await X(),Ut=await _e();let a=n=>n&&!it.find(p=>p.value===n)?[...it,{value:n,label:n}]:it,i=new Date().getFullYear(),s=["2025","2026","2027","2028","2029","2030"],l=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((n,p)=>{let u=String(p+1).padStart(2,"0");return{value:`${i}-${u}`,label:`${n} ${i}`}});A({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:n=>r?n.filter(p=>qa(p,r)):n,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:n=>`<span class="badge badge-secondary">${n}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:n=>`<span title="${n}">${n?.length>50?n.slice(0,50)+"\u2026":n}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:n=>`<span title="${n||""}">${n?.length>40?n.slice(0,40)+"\u2026":n||"-"}</span>`},{key:"status",label:"Status",render:n=>H(n)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"day_count",label:"Hari",render:n=>n??"-"},{key:"inspection_id",label:"Inspection",render:n=>n?'<span class="badge badge-info">Finding</span>':"-"},{key:"vendor_name",label:"Vendor",render:n=>n||"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:wt},{type:"select",name:"month",label:"Bulan",options:l},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done","OPEN","ASSIGNED","IN_PROGRESS","RESOLVED","CLOSED"]},{type:"select",name:"year",label:"Tahun",options:s},{type:"select",name:"inspection_id",label:"Type",options:[{value:"",label:"All"},{value:"null",label:"Permasalahan"},{value:"notnull",label:"Finding"}]}],formFields:n=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:n?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:wt,value:n?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:n?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:n?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:n?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(n?.employee_name),value:n?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(n?.fc_specialist),value:n?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:n?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:n?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let n=await w(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let p=n.data.data.map(u=>({Tanggal:u.report_date||"",Cabang:u.branch_name||"",Kategori:u.category||"",Sumber:u.source||"",Keluhan:u.complaint||"","Nama FC":u.employee_name||"","FC Spesialis":u.fc_specialist||"",Solusi:u.solution||"","Tgl Selesai":u.completion_date||"",Status:u.status||"",Vendor:u.vendor_name||""}));I(p,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{I([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done",Vendor:"Vendor A"}],"Template_Import_Permasalahan")},onImport:async n=>{let u=(await w("/api/branches?all=1")).data?.data||[],d=b=>{if(!b)return null;let m=String(b||"").toLowerCase(),y=u.find(S=>String(S.full_name||"").toLowerCase()===m||String(S.code||"").toLowerCase()===m||String(S.name||"").toLowerCase()===m);return y?y.id:null},c=b=>{if(!b)return null;let m=String(b||"").toLowerCase(),y=Ut.find(S=>String(S.label||"").toLowerCase()===m);return y?y.value:null},g=n.map(b=>({branch_id:d(String(b.Cabang||"").trim()),report_date:String(b.Tanggal||"").trim(),category:String(b.Kategori||"").trim(),source:String(b.Sumber||"").trim(),complaint:String(b.Keluhan||"").trim(),employee_name:String(b["Nama FC"]||"").trim(),fc_specialist:String(b["FC Spesialis"]||"").trim(),solution:String(b.Solusi||"").trim(),completion_date:String(b["Tgl Selesai"]||"").trim(),status:String(b.Status||"").trim(),vendor_id:b.Vendor?c(String(b.Vendor||"").trim()):null})).filter(b=>b.report_date&&b.complaint&&b.category),f=await w("/api/import/issues",{method:"POST",body:JSON.stringify({rows:g,onDuplicate:"update"})});if(!f.ok)throw new Error(f.data?.error||"Import gagal");return f.data}}})}F();var Ae=[];function Ha(t,e){let r=String(t.status||"").toLowerCase();return e==="pending"?r==="pending":!1}async function Vt(t,e){let r=e?e.get("dash_filter"):null;Ae=await K();let a=await X(),i=["Ade","Berlin"],s=l=>l&&!a.find(n=>n.value===l)?[...a,{value:l,label:l}]:a,o=l=>l&&!i.find(n=>(typeof n=="object"?n.value:n)===l)?[...i,l]:i;A({container:t,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:l=>r?l.filter(n=>Ha(n,r)):l,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:l=>`<span title="${l||""}">${l?.length>50?l.slice(0,50)+"\u2026":l||"-"}</span>`},{key:"solution",label:"Solusi",render:l=>`<span title="${l||""}">${l?.length>40?l.slice(0,40)+"\u2026":l||"-"}</span>`},{key:"status",label:"Status",render:l=>H(l)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:l=>window.formatDate(l)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:l=>l?`<a href="${l}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:Ae},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async l=>{let n=new URLSearchParams(l||{}).toString(),p=await w(`/api/one-on-one?limit=10000&${n}`);if(p.ok){let u=p.data.data.map(c=>({Tanggal:c.meeting_date||"",Cabang:c.branch_name||"","Nama Karyawan":c.employee_name||"",PIC:c.pic||"",Masalah:c.problem||"",Solusi:c.solution||"",Status:c.status||"","Tgl Selesai":c.completion_date||"",Dokumen:c.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(R(),oe));d(u,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let l=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(R(),oe));n(l,"Template_Import_OneOnOne")},onImport:async l=>{let n=c=>{if(!c)return null;let g=String(c||"").toLowerCase(),f=Ae.find(b=>String(b.label||"").toLowerCase()===g);return f?f.value:null},p=c=>{if(!c)return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let g=String(c).trim();if(/^\d{4,5}$/.test(g)){let b=Number(g);if(b>2e4&&b<99999){let m=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);let f=g.split(/[\/\-\.]/);if(f.length===3){let[b,m,y]=f.map(S=>S.trim());if(b.length===4&&m.length<=2&&y.length<=2)return`${b}-${m.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&m.length<=2&&b.length<=2)return`${y}-${m.padStart(2,"0")}-${b.padStart(2,"0")}`}return g},u=l.map(c=>({meeting_date:p(c.Tanggal),employee_name:String(c["Nama Karyawan"]||"").trim(),branch_id:n(String(c.Cabang||"").trim()),pic:String(c.PIC||"").trim(),problem:String(c.Masalah||"").trim(),solution:String(c.Solusi||"").trim(),status:String(c.Status||"").trim(),completion_date:p(c["Tgl Selesai"]),document_link:String(c.Dokumen||"").trim()})).filter(c=>c.meeting_date&&c.employee_name&&c.branch_id),d=await w("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:u,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:l=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:l?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:l?.branch_id&&!Ae.find(n=>n.value==l.branch_id)?[...Ae,{value:l.branch_id,label:l.branch_name||l.branch_id}]:Ae,createApi:{path:"/api/branches",field:"full_name"},value:l?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:s(l?.employee_name),value:l?.employee_name},{name:"pic",label:"PIC",type:"select",options:o(l?.pic),createApi:{path:"/api/pic",field:"name"},value:l?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:l?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:l?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:l?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:l?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:l?.document_link}]})}F();async function zt(t){let e=await K(),r=await X(),a=["Ade Surahman","Berlin Ariansyah","Mizwar","Fajar","Ade","Berlin"],i=n=>n&&!r.find(p=>p.value===n)?[...r,{value:n,label:n}]:r,s=n=>n&&!a.find(p=>(typeof p=="object"?p.value:p)===n)?[...a,n]:a,o=Array.from({length:5},(n,p)=>String(new Date().getFullYear()-p));A({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:n=>{try{let p=JSON.parse(n);return Array.isArray(p)?p.join(", "):n||"-"}catch{return n||"-"}}},{key:"score",label:"Nilai",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer / peserta..."},{type:"select",name:"batch",label:"Batch",options:["Batch 1","Batch 2","Batch 3","Batch 4","Batch 5"]},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"trainer",label:"Trainer",options:["Berlin Ariansyah"]},{type:"select",name:"year",label:"Tahun",options:o}],exportOptions:{moduleName:"training",onExport:async n=>{let p=new URLSearchParams(n||{}).toString(),u=await w(`/api/training?limit=10000&${p}`);if(u.ok){let d=u.data.data.map(g=>{let f=g.participants||"";try{let b=JSON.parse(f);f=Array.isArray(b)?b.join(", "):f}catch{}return{Tanggal:g.training_date||"",Batch:g.batch||"",Materi:g.subject||"",Cabang:g.branch_name||"",Trainer:g.trainer||"",Peserta:f,Nilai:g.score!==null&&g.score!==void 0?g.score:"",Dokumen:g.document_link||""}}),{downloadExcel:c}=await Promise.resolve().then(()=>(R(),oe));c(d,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:p}=await Promise.resolve().then(()=>(R(),oe));p(n,"Template_Import_Training")},onImport:async n=>{let p=g=>{if(!g)return null;let f=String(g||"").toLowerCase(),b=e.find(m=>String(m.label||"").toLowerCase()===f);return b?b.value:null},u=g=>{if(!g)return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let f=String(g).trim();if(/^\d{4,5}$/.test(f)){let m=Number(f);if(m>2e4&&m<99999){let y=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(f))return f.slice(0,10);let b=f.split(/[\/\-\.]/);if(b.length===3){let[m,y,S]=b.map(_=>_.trim());if(m.length===4&&y.length<=2&&S.length<=2)return`${m}-${y.padStart(2,"0")}-${S.padStart(2,"0")}`;if(S.length===4&&y.length<=2&&m.length<=2)return`${S}-${y.padStart(2,"0")}-${m.padStart(2,"0")}`}return f},d=n.map(g=>({training_date:u(g.Tanggal),batch:String(g.Batch||"").trim(),subject:String(g.Materi||"").trim(),branch_id:p(String(g.Cabang||"").trim()),trainer:String(g.Trainer||"").trim(),participants:String(g.Peserta||"").trim(),score:g.Nilai?Number(g.Nilai):null,document_link:String(g.Dokumen||"").trim()})).filter(g=>g.training_date&&g.subject&&g.branch_id),c=await w("/api/training/import",{method:"POST",body:JSON.stringify(d)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}},formFields:n=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:n?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:n?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:n?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:n?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:s(n?.trainer),value:n?.trainer}]},{name:"participants",label:"Peserta",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let p=JSON.parse(n?.participants);return Array.isArray(p)?p.join(", "):n?.participants||""}catch{return n?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:n?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:n?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],onBeforeSubmit:async n=>n})}F();xe();R();function Qt({container:t,title:e,icon:r,apiPath:a,columns:i,formFields:s,filterFields:o,defaultFilters:l={},itemLabel:n="Data",canCreate:p=!0,canEdit:u=!0,canDelete:d=!0,onBeforeSubmit:c,onAfterLoad:g,onDataLoaded:f,extraActions:b=[],initialSearch:m="",exportOptions:y=null,bulkDelete:S=!1,paginationMode:_="server"}){let C=ce();C&&C.role==="viewer"&&(p=!1,u=!1,d=!1,S=!1,y=null);let x=1,T={...l};m&&(T.search=m);let E=new Set;t.innerHTML=`
    ${S?`
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
      <h1 class="page-title">${r} ${e}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${p?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
        ${y?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${y.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${y.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${y.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${y.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
        ${o.filter(h=>h.type==="search").map(h=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${h.placeholder||"Cari..."}" id="filter-search" value="${T.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${o.filter(h=>h.type!=="search").map(h=>{if(h.type==="select"||h.type==="combobox"){let k=(h.label||"").startsWith("Pilih")?h.label:`Pilih ${h.label||""}`;return`<select class="filter-select" name="${h.name}" id="filter-${h.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${k}</option>${(h.options||[]).map(v=>`<option value="${typeof v=="object"?v.value:v}" ${T[h.name]===(typeof v=="object"?v.value:v)?"selected":""}>${typeof v=="object"?v.label:v}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(h=>h.type!=="search").map(h=>{if(h.type==="select"||h.type==="combobox"){let k=(h.label||"").startsWith("Pilih")?h.label:`Pilih ${h.label||""}`;return`<select class="filter-select filter-select-sheet" name="${h.name}-sheet" id="filter-sheet-${h.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${k}</option>${(h.options||[]).map(v=>`<option value="${typeof v=="object"?v.value:v}" ${T[h.name]===(typeof v=="object"?v.value:v)?"selected":""}>${typeof v=="object"?v.label:v}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function $(){let h=document.getElementById("bulk-toolbar");if(!h)return;let k=document.getElementById("bulk-count"),v=document.getElementById("btn-bulk-delete"),M=document.getElementById("btn-bulk-cancel");k.textContent=`${E.size} item dipilih`,E.size>0?(h.style.display="flex",v.disabled=!1,M.disabled=!1):(h.style.display="none",v.disabled=!0,M.disabled=!0);let N=document.getElementById("select-all-checkbox");if(N){let J=document.querySelectorAll(".row-checkbox");if(J.length>0){let D=[...J].every(ie=>ie.checked),L=[...J].some(ie=>ie.checked);N.checked=D,N.indeterminate=L&&!D}else N.checked=!1,N.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{E.clear(),document.querySelectorAll(".row-checkbox").forEach(k=>k.checked=!1);let h=document.getElementById("select-all-checkbox");h&&(h.checked=!1),$()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(E.size===0)return;let h=[...E],k=document.createElement("div");k.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",k.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${h.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${h.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(k),k.querySelector("#bulk-cancel-btn").addEventListener("click",()=>k.remove()),k.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let v=k.querySelector("#bulk-confirm-btn");v.disabled=!0,v.textContent="Menghapus...";let M=await w(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:h})});k.remove(),M.ok?(W(`${h.length} ${n} berhasil dihapus.`),E.clear(),$(),ge(a),j()):z(M.data?.error||"Gagal menghapus data.")})});let B=document.getElementById("filter-search"),P;B?.addEventListener("input",h=>{clearTimeout(P),P=setTimeout(()=>{T.search=h.target.value,x=1,E.clear(),$(),j()},400)}),o?.forEach(h=>{(h.type==="select"||h.type==="combobox")&&(document.getElementById(`filter-${h.name}`)?.addEventListener("change",k=>{T[h.name]=k.target.value;let v=document.getElementById(`filter-sheet-${h.name}`);v&&(v.value=k.target.value),x=1,E.clear(),$(),j()}),document.getElementById(`filter-sheet-${h.name}`)?.addEventListener("change",k=>{T[h.name]=k.target.value;let v=document.getElementById(`filter-${h.name}`);v&&(v.value=k.target.value),x=1,E.clear(),$(),j(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{T={},B&&(B.value=""),o?.forEach(h=>{let k=document.getElementById(`filter-${h.name}`);k&&(k.value="");let v=document.getElementById(`filter-sheet-${h.name}`);v&&(v.value="")}),x=1,E.clear(),$(),j()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{T={},B&&(B.value=""),o?.forEach(h=>{let k=document.getElementById(`filter-${h.name}`);k&&(k.value="");let v=document.getElementById(`filter-sheet-${h.name}`);v&&(v.value="")}),x=1,E.clear(),$(),j(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>ye(null)),y&&document.addEventListener("click",function(h){let k=document.getElementById("aksi-menu-main"),v=document.getElementById("btn-aksi-main");k&&v&&!v.contains(h.target)&&!k.contains(h.target)&&k.classList.remove("show-aksi-menu")});let Q=document.getElementById("btn-mobile-filter"),ee=document.getElementById("filter-options-wrapper"),re=document.getElementById("btn-close-filter-sheet");if(Q&&ee&&(Q.addEventListener("click",h=>{h.preventDefault(),ee.classList.add("sheet-open")}),re&&re.addEventListener("click",h=>{h.preventDefault(),ee.classList.remove("sheet-open")})),y){document.getElementById(`btn-export-${y.moduleName}`)?.addEventListener("click",async k=>{let v=k.target,M=v.innerHTML;v.innerHTML="\u23F3 Loading...",v.disabled=!0;try{await y.onExport()}catch{z("Gagal export data")}finally{v.innerHTML=M,v.disabled=!1}}),document.getElementById(`btn-template-${y.moduleName}`)?.addEventListener("click",()=>{y.onTemplate()});let h=document.getElementById(`input-import-${y.moduleName}`);h?.addEventListener("change",async k=>{let v=k.target.files[0];if(!v)return;let M=document.getElementById(`label-import-${y.moduleName}`),N=M?M.querySelector(".import-text"):null,J=N?N.innerText:"";N&&(N.innerText="\u231B Memproses..."),M&&(M.style.pointerEvents="none"),h.disabled=!0;try{let D=await Je(v);if(D.length===0)throw new Error("File kosong atau format salah");await y.onImport(D),W("Import berhasil!"),ge(a),j()}catch(D){z(D.message||"Gagal import data")}finally{N&&(N.innerText=J),M&&(M.style.pointerEvents="auto"),h.disabled=!1,h.value=""}})}async function j(){$();let h=document.getElementById("table-container");if(!h)return;h.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let k=_==="client",v=k?1:x,M=k?we:20,N=new URLSearchParams({page:v,limit:M,...Object.fromEntries(Object.entries(T).filter(([,q])=>q))}),J=await w(`${a}?${N}`);if(!J.ok){h.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${J.data?.error||"Error"}</p></div>`;return}let D=J.data?.data||J.data||[],L=J.data?.pagination,ie=D.length;if(k){D=f(D);let q=D.length,V=20,te=Math.ceil(q/V);x>te&&te>0&&(x=te);let O=(x-1)*V,le=x*V;D=D.slice(O,le),L={page:x,limit:V,total:q,pages:te}}!1,g&&g(D);let fe=Ye({columns:i,data:D,onEdit:u?q=>ye(q):null,actions:b.map(q=>({...q,handler:V=>q.handler(V,j)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:S?{selectedIds:E,onToggle:$}:null});h.innerHTML="",h.appendChild(fe);let se=document.getElementById("pagination-container");if(se&&(se.innerHTML="",L&&L.pages>1)){let q=We({page:L.page,pages:L.pages,total:L.total,limit:L.limit,onPage:V=>{x=V,j()}});q&&se.appendChild(q)}}function Ce(h){let k=typeof s=="function"?s(h):s;return He(k)}function ye(h){let k=!!h,v=document.createElement("form");if(v.noValidate=!0,v.innerHTML=Ce(h),k){let N=typeof s=="function"?s(h):s;Ze(v,h)}let{close:M}=de({title:k?`Edit ${n}`:`Tambah ${n}`,content:v,size:"lg",confirmText:k?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(N,J)=>{if(!v.reportValidity())return;let D=N.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let L=Xe(v),ie=typeof s=="function"?s(h):s,fe=async te=>{for(let O of te)if(O.type==="row")await fe(O.fields);else if(O.type==="combobox"&&L[O.name]){let le=L[O.name],ve=(O.options||[]).find(Y=>{let ae=String(typeof Y=="object"?Y.value:Y),st=String(typeof Y=="object"?Y.label:Y);return ae===le||st===le});if(ve)L[O.name]=typeof ve=="object"?ve.value:ve;else if(O.createApi){let Y={};Y[O.createApi.field]=le,O.createApi.extra&&Object.assign(Y,O.createApi.extra);let ae=await w(O.createApi.path,{method:"POST",body:JSON.stringify(Y)});if(ae.ok&&ae.data?.id)L[O.name]=ae.data.id;else if(ae.ok&&!ae.data?.id)L[O.name]=le;else throw new Error(`Gagal membuat master data: ${ae.data?.error||"Unknown error"}`)}}};try{await fe(ie)}catch(te){z(te.message),D.disabled=!1,D.textContent=k?"Simpan Perubahan":`Tambah ${n}`;return}c&&(L=await c(L,h));let se=k?"PUT":"POST",q=k?`${a}/${h.id}`:a,V=await w(q,{method:se,body:JSON.stringify(L)});V.ok?(W(k?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),J(),ge(a),j()):(z(V.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=k?"Simpan Perubahan":`Tambah ${n}`)}})}function va(h){qe(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let k=await w(`${a}/${h.id}`,{method:"DELETE"});k.ok?(W(`${n} berhasil dihapus.`),ge(a),j()):z(k.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return j(),j}F();R();async function Yt(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let r=await K(),a=await X(),i=e?e.get("dash_filter"):null,s={};if(i==="reliever"){let c=new Date,g=`${c.getFullYear()}-${String(c.getMonth()+1).padStart(2,"0")}`;s={status:"Done",month:e&&e.get("month")?e.get("month"):g}}console.log("RAW",await De()),console.log("OPTIONS",a);let o=c=>c&&!a.find(g=>g.value===c)?[...a,{value:c,label:c}]:a,l=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],n=c=>c&&!l.includes(c)?[...l,c]:l,p=new Date().getFullYear(),d=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((c,g)=>{let f=String(g+1).padStart(2,"0");return{value:`${p}-${f}`,label:`${c} ${p}`}});Qt({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:s,onDataLoaded:c=>c.sort((g,f)=>{let b=g.backup_date?new Date(g.backup_date).getTime():0;return(f.backup_date?new Date(f.backup_date).getTime():0)-b}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:c=>pe(c)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:c=>window.formatDate(c)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:c=>window.formatDate(c)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:c=>c?`<span class="badge badge-info">${c}</span>`:"-"},{key:"status",label:"Status",render:c=>H(c)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:l},{type:"select",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"month",label:"Bulan",options:d},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:c=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:r,value:c?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:c?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:o(c?.original_fc_name),value:c?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:n(c?.reliever_name),value:c?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:c?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:c?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:c?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:c?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:c?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let c=await w(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(c.ok){let g=c.data.data.map(f=>({Cabang:f.branch_name||"","Nama Facility care":f.original_fc_name||"",Periode:f.period||"",Relifer:f.reliever_name||"","Tanggal Back Up":f.backup_date||"","Tanggal Selesai":f.completion_date||"",Keterangan:f.reason||"",Shift:f.shift||"",Status:f.status||""}));g.length===0&&g.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),I(g,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{I([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async c=>{let f=(await w("/api/branches?all=1")).data?.data||[],b=S=>{if(!S)return null;let _=String(S||"").toLowerCase(),C=f.find(x=>String(x.full_name||"").toLowerCase()===_||String(x.code||"").toLowerCase()===_||String(x.name||"").toLowerCase()===_);return C?C.id:null},m=c.map(S=>({branch_name:String(S.Cabang||"").trim(),backup_date:String(S["Tanggal Back Up"]||S["Tanggal Backup"]||"").trim(),original_fc_name:String(S["Nama Facility care"]||S["FC Digantikan"]||"").trim(),reliever_name:String(S.Relifer||S.Reliefer||"").trim(),period:String(S.Periode||"").trim(),reason:String(S.Keterangan||"").trim(),shift:String(S.Shift||"").trim(),completion_date:String(S["Tanggal Selesai"]||"").trim(),status:String(S.Status||"").trim()})).filter(S=>S.reliever_name&&S.backup_date),y=await w("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!y.ok)throw new Error(y.data?.error||"Import gagal");return y.data}}})}F();R();async function Wt(t){let e=await K(),r=await _e(),a=Array.from({length:4},(i,s)=>String(new Date().getFullYear()-s));A({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:i=>pe(i)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"fc_score",label:"Point FC",render:i=>i!=null?`<strong class="${i>=80?"text-success":i>=60?"text-warning":"text-danger"}">${i}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:i=>i!=null?`<strong>${i}</strong>`:"-"},{key:"status",label:"Status",render:i=>H(i)},{key:"document_link",label:"Dokumen",render:i=>i?`<a href="${i}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"vendor_name",label:"Vendor",render:i=>i||"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:i=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:i?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:i?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:i?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:i?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:i?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:i?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:i?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:i?.notes},{name:"vendor_id",label:"Vendor",type:"select",options:r,value:i?.vendor_id,help:"Vendor yang menyediakan jasa inspeksi"}],exportOptions:{moduleName:"inspection_reports",onExport:async i=>{let s=new URLSearchParams(i||{}).toString(),o=await w(`/api/reports/inspection?limit=10000&${s}`);if(o.ok){let l=o.data.data.map(n=>({Cabang:n.branch_name||"",Periode:n.period||"",Tanggal:n.inspection_date||"","Point FC":n.fc_score!==null&&n.fc_score!==void 0?n.fc_score:"","Point SPV":n.spv_score!==null&&n.spv_score!==void 0?n.spv_score:"",Status:n.status||"","Link Dokumen":n.document_link||"",Vendor:n.vendor_name||""}));I(l,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{I([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Vendor:"Vendor A",Catatan:""}],"Template_Import_Inspeksi")},onImport:async i=>{let s=u=>{if(!u)return null;let d=String(u||"").toLowerCase(),c=e.find(g=>String(g.label||"").toLowerCase()===d);return c?c.value:null},o=u=>{if(!u)return null;let d=String(u||"").toLowerCase(),c=r.find(g=>String(g.label||"").toLowerCase()===d);return c?c.value:null},l=u=>{if(u==null||u==="")return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let d=String(u).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let g=Number(d);if(g>2e4&&g<99999){let f=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}let c=d.split(/[\/\-\.]/);if(c.length===3){let[g,f,b]=c.map(m=>m.trim());if(g.length===4&&f.length<=2&&b.length<=2)return`${g}-${f.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&f.length<=2&&g.length<=2)return`${b}-${f.padStart(2,"0")}-${g.padStart(2,"0")}`}return d},n=i.map(u=>({branch_id:s(String(u.Cabang||"").trim()),period:String(u.Periode||"").trim(),inspection_date:l(u.Tanggal),fc_score:u["Point FC"]!==void 0&&u["Point FC"]!==""?Number(u["Point FC"]):null,spv_score:u["Point SPV"]!==void 0&&u["Point SPV"]!==""?Number(u["Point SPV"]):null,status:String(u.Status||"").trim(),document_link:String(u["Link Dokumen"]||"").trim(),vendor_id:o(String(u.Vendor||"").trim()),notes:String(u.Catatan||u.Keterangan||"").trim()})).filter(u=>u.branch_id&&u.period&&u.inspection_date),p=await w("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:n,onDuplicate:"update"})});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}}})}F();R();async function Xt(t){let e=await K(),r=Array.from({length:4},(a,i)=>String(new Date().getFullYear()-i));A({container:t,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>pe(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>H(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let i=new URLSearchParams(a||{}).toString(),s=await w(`/api/reports/cleaning?limit=10000&${i}`);if(s.ok){let o=s.data.data.map(l=>({Cabang:l.branch_name||"",Jenis:l.activity_type||"",Periode:l.period||"",Tanggal:l.activity_date||"",Status:l.status||"","Link Dokumen":l.document_link||""}));I(o,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{I([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let i=n=>{if(!n)return null;let p=String(n||"").toLowerCase(),u=e.find(d=>String(d.label||"").toLowerCase()===p);return u?u.value:null},s=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let p=String(n).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let d=Number(p);if(d>2e4&&d<99999){let c=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(c.getTime())?"":c.toISOString().slice(0,10)}}let u=p.split(/[\/\-\.]/);if(u.length===3){let[d,c,g]=u.map(f=>f.trim());if(d.length===4&&c.length<=2&&g.length<=2)return`${d}-${c.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&c.length<=2&&d.length<=2)return`${g}-${c.padStart(2,"0")}-${d.padStart(2,"0")}`}return p},o=a.map(n=>({branch_id:i(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:s(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),l=await w("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}F();R();async function Zt(t,e){let r=await K(),a=Array.from({length:4},(o,l)=>String(new Date().getFullYear()-l)),i=e?e.get("dash_filter"):null,s={};if(i==="fogging"){let o=new Date,l=String(o.getMonth()+1).padStart(2,"0"),n=String(o.getFullYear()),p=e?e.get("month"):null;p&&p.length===7&&(n=p.split("-")[0],l=p.split("-")[1]),s={status:"Done",month:l,year:n}}A({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:s,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:o=>`<span class="badge badge-warning">${o}</span>`},{key:"period",label:"Periode",render:o=>pe(o)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>H(o)},{key:"document_link",label:"Dokumen",render:o=>o?`<a href="${o}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:o=>o||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:r,value:o?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:o?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:o?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:o?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:o?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async o=>{let l=new URLSearchParams(o||{}).toString(),n=await w(`/api/reports/fogging?limit=10000&${l}`);if(n.ok){let p=n.data.data.map(u=>({Cabang:u.branch_name||"",Jenis:u.activity_type||"Fogging",Periode:u.period||"",Tanggal:u.activity_date||"",Status:u.status||"","Link Dokumen":u.document_link||""}));I(p,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{I([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async o=>{let l=d=>{if(!d)return null;let c=String(d||"").toLowerCase(),g=r.find(f=>String(f.label||"").toLowerCase()===c);return g?g.value:null},n=d=>{if(d==null||d==="")return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let c=String(d).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let f=Number(c);if(f>2e4&&f<99999){let b=new Date(Date.UTC(1899,11,30)+f*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}let g=c.split(/[\/\-\.]/);if(g.length===3){let[f,b,m]=g.map(y=>y.trim());if(f.length===4&&b.length<=2&&m.length<=2)return`${f}-${b.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&b.length<=2&&f.length<=2)return`${m}-${b.padStart(2,"0")}-${f.padStart(2,"0")}`}return c},p=o.map(d=>({branch_id:l(String(d.Cabang||"").trim()),activity_type:String(d.Jenis||d.Kegiatan||"Fogging").trim(),period:String(d.Periode||"").trim(),activity_date:n(d.Tanggal),status:String(d.Status||"").trim(),document_link:String(d["Link Dokumen"]||"").trim(),notes:String(d.Catatan||d.Keterangan||"").trim()})).filter(d=>d.branch_id&&d.period&&d.activity_date),u=await w("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(p)});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}F();R();async function ea(t){let e=await K(),r=await X(),a=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],i=Array.from({length:4},(l,n)=>String(new Date().getFullYear()-n)),s=l=>l&&!r.find(n=>n.value===l)?[...r,{value:l,label:l}]:r,o=l=>l&&!a.find(n=>n.value===l)?[...a,{value:l,label:l}]:a;A({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:l=>`<span title="${l||""}">${l?.length>60?l.slice(0,60)+"\u2026":l||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>H(l)},{key:"notes",label:"Keterangan",render:l=>l?.length>40?l.slice(0,40)+"\u2026":l||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:l=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:l?.branch_id},{name:"pic",label:"PIC",type:"select",options:o(l?.pic),value:l?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:l?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:l?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:l?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:l?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async l=>{let n=new URLSearchParams(l||{}).toString(),p=await w(`/api/reports/basecamp?limit=10000&${n}`);if(p.ok){let u=p.data.data.map(d=>({"Tgl Info":d.info_date||"",Cabang:d.branch_name||"",Permasalahan:d.problem||"",PIC:d.pic||"","Tgl Done":d.done_date||"",Status:d.status||"",Keterangan:d.notes||""}));I(u,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{I([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async l=>{let n=c=>{if(!c)return null;let g=String(c||"").toLowerCase(),f=e.find(b=>String(b.label||"").toLowerCase()===g);return f?f.value:null},p=c=>{if(c==null||c==="")return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let g=String(c).trim();if(g===""||g==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);if(/^\d{4,5}$/.test(g)){let b=Number(g);if(b>2e4&&b<99999){let m=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}let f=g.split(/[\/\-\.]/);if(f.length===3){let[b,m,y]=f.map(S=>S.trim());if(b.length===4&&m.length<=2&&y.length<=2)return`${b}-${m.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&m.length<=2&&b.length<=2)return`${y}-${m.padStart(2,"0")}-${b.padStart(2,"0")}`}return g},u=l.map(c=>({info_date:p(c["Tgl Info"]||c["Tanggal Info"]),branch_id:n(String(c.Cabang||"").trim()),problem:String(c.Permasalahan||"").trim(),pic:String(c.PIC||"").trim(),done_date:p(c["Tgl Done"]||c["Tanggal Done"]),status:String(c.Status||"").trim(),notes:String(c.Keterangan||c.Catatan||"").trim()})).filter(c=>c.info_date&&c.branch_id&&c.problem),d=await w("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(u)});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}}})}F();R();async function ta(t){A({container:t,title:"Vendor Master",icon:"\u{1F3E2}",apiPath:"/api/vendors",enableMobileFilterSheet:!0,itemLabel:"Vendor",bulkDelete:!0,paginationMode:"client",columns:[{key:"id",label:"ID",render:e=>`<span class="badge badge-primary">#${e}</span>`},{key:"code",label:"Code"},{key:"name",label:"Name",render:e=>`<strong>${e}</strong>`},{key:"category",label:"Category"},{key:"contact_person",label:"Contact Person"},{key:"phone",label:"Phone"},{key:"email",label:"Email"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-danger">Nonaktif</span>'},{key:"created_at",label:"Created",render:e=>e?window.formatDate(e):"-"}],filterFields:[{type:"search",placeholder:"Cari code / name / contact..."},{type:"select",name:"active",label:"Status",options:[{value:"",label:"Semua"},{value:"true",label:"Aktif"},{value:"false",label:"Nonaktif"}]}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Code",type:"text",placeholder:"OPR-001",value:e?.code},{name:"name",label:"Name *",type:"text",required:!0,placeholder:"Nama Vendor",value:e?.name}]},{type:"row",fields:[{name:"category",label:"Category",type:"select",options:["Cleaning","Fogging","Maintenance","Security","Other"],value:e?.category},{name:"contact_person",label:"Contact Person",type:"text",value:e?.contact_person}]},{type:"row",fields:[{name:"phone",label:"Phone",type:"text",placeholder:"0812-3456-7890",value:e?.phone},{name:"email",label:"Email",type:"email",placeholder:"vendor@company.com",value:e?.email}]},{name:"address",label:"Address",type:"textarea",rows:2,placeholder:"Alamat lengkap",value:e?.address},{type:"row",fields:[{name:"is_active",label:"Active",type:"checkbox",checkLabel:"Vendor aktif",value:e?.is_active!==void 0?e.is_active:!0}]}],exportOptions:{moduleName:"vendors",onExport:async()=>{let e=await w(`/api/vendors${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok){let r=e.data.data.map(a=>({Code:a.code||"",Name:a.name||"",Category:a.category||"","Contact Person":a.contact_person||"",Phone:a.phone||"",Email:a.email||"",Address:a.address||"",Status:a.is_active?"Aktif":"Nonaktif"}));I(r,`Master_Vendor_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{I([{Code:"OPR-001",Name:"Vendor Cleaning A",Category:"Cleaning","Contact Person":"John Doe",Phone:"0812-3456-7890",Email:"john@vendor.com",Address:"Jakarta",Status:"Aktif"}],"Template_Import_Vendor")},onImport:async e=>{let r=e.map(i=>({code:String(i.Code||"").trim(),name:String(i.Name||"").trim(),category:String(i.Category||"").trim(),contact_person:String(i["Contact Person"]||"").trim(),phone:String(i.Phone||"").trim(),email:String(i.Email||"").trim(),address:String(i.Address||"").trim(),is_active:String(i.Status||"").toLowerCase()==="aktif"})).filter(i=>i.name),a=await w("/api/vendors/import",{method:"POST",body:JSON.stringify({rows:r})});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}F();R();var xt=[],Ja=[],_t=[];async function aa(t){xt=await K(),Ja=await X(),_t=await _e(),A({container:t,title:"Corrective Action",icon:"\u{1F527}",apiPath:"/api/corrective-actions",enableMobileFilterSheet:!0,itemLabel:"Corrective Action",paginationMode:"client",columns:[{key:"id",label:"ID",render:e=>`<span class="badge badge-primary">#${e}</span>`},{key:"branch_name",label:"Cabang"},{key:"vendor_name",label:"Vendor",render:e=>e||"-"},{key:"finding_id",label:"Finding",render:e=>e?`<a href="#/issues/${e}" target="_self">#${e}</a>`:"-"},{key:"description",label:"Deskripsi",render:e=>`<span title="${e||""}">${e?.length>50?e.slice(0,50)+"\u2026":e||"-"}</span>`},{key:"assigned_to",label:"Ditugaskan ke"},{key:"due_date",label:"Due Date",nowrap:!0,render:e=>e?window.formatDate(e):"-"},{key:"status",label:"Status",render:e=>H(e)},{key:"is_overdue",label:"Overdue",render:e=>e?'<span class="text-danger">\u26A0\uFE0F Ya</span>':"-"},{key:"evidence_link",label:"Evidence",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4}</a>`:"-"},{key:"approved_by",label:"Disetujui oleh",render:e=>e||"-"},{key:"approved_at",label:"Disetujui pada",render:e=>e?window.formatDate(e):"-"}],filterFields:[{type:"search",placeholder:"Cari deskripsi / assigned_to..."},{type:"select",name:"branch_id",label:"Cabang",options:xt},{type:"select",name:"vendor_id",label:"Vendor",options:_t},{type:"select",name:"status",label:"Status",options:["ASSIGNED","IN_PROGRESS","REVIEW_PENDING","APPROVED","REJECTED"]},{type:"select",name:"overdue",label:"Overdue",options:[{value:"",label:"Semua"},{value:"true",label:"Ya"},{value:"false",label:"Tidak"}]}],formFields:e=>[{type:"row",fields:[{name:"finding_id",label:"Finding ID",type:"number",required:!0,value:e?.finding_id,hint:"ID dari Finding (inspection)"},{name:"branch_id",label:"Cabang",type:"select",options:xt,value:e?.branch_id}]},{type:"row",fields:[{name:"vendor_id",label:"Vendor",type:"select",options:_t,value:e?.vendor_id,readonly:!0,hint:"Vendor diambil dari Inspection terkait (read-only)"},{name:"assigned_to",label:"Ditugaskan ke",type:"text",value:e?.assigned_to,hint:"Nama vendor/employee"}]},{name:"description",label:"Deskripsi",type:"textarea",required:!0,rows:3,value:e?.description,hint:"Deskripsi tindakan korektif"},{type:"row",fields:[{name:"due_date",label:"Due Date",type:"date",value:e?.due_date},{name:"status",label:"Status",type:"select",required:!0,options:["ASSIGNED","IN_PROGRESS","REVIEW_PENDING","APPROVED","REJECTED"],value:e?.status||"ASSIGNED"}]},{name:"evidence_link",label:"Link Evidence",type:"url",placeholder:"https://drive.google.com/...",value:e?.evidence_link},{name:"cost_estimate",label:"Estimasi Biaya",type:"number",step:"0.01",value:e?.cost_estimate},{name:"actual_cost",label:"Biaya Sebenarnya",type:"number",step:"0.01",value:e?.actual_cost},{name:"completed_at",label:"Selesai pada",type:"date",value:e?.completed_at},{name:"reject_reason",label:"Alasan Reject",type:"textarea",rows:2,value:e?.reject_reason}],exportOptions:{moduleName:"corrective_actions",onExport:async()=>{let e=await w(`/api/corrective-actions${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok){let r=e.data.data.map(a=>({ID:a.id||"",Cabang:a.branch_name||"",Vendor:a.vendor_name||"","Finding ID":a.finding_id||"",Deskripsi:a.description||"","Ditugaskan ke":a.assigned_to||"","Due Date":a.due_date||"",Status:a.status||"",Overdue:a.is_overdue?"Ya":"Tidak","Link Evidence":a.evidence_link||"","Disetujui oleh":a.approved_by||"","Disetujui pada":a.approved_at||""}));I(r,`Corrective_Action_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{I([{"Finding ID":"1",Cabang:"001. Pondok Bambu",Vendor:"Vendor A",Deskripsi:"Perbaiki lantai kotor","Ditugaskan ke":"Vendor A","Due Date":"2026-09-01",Status:"ASSIGNED","Link Evidence":"https://drive.google.com/..."}],"Template_Import_Corrective_Action")}}})}async function na(t){A({container:t,title:"SOP",icon:"\u{1F4DA}",apiPath:"/api/sop",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(F(),Ee)),i=await a(`/api/sop?limit=10000&${r}`);if(i.ok){let s=i.data.data.map(l=>({"Nama SOP":l.name||"",Kategori:l.category||"",Dokumen:l.document_link||"",Catatan:l.notes||l.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(R(),oe));o(s,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:r}=await Promise.resolve().then(()=>(R(),oe));r(e,"Template_Import_SOP")},onImport:async e=>{let r=e.map(s=>({name:String(s["Nama SOP"]||"").trim(),category:String(s.Kategori||"").trim(),document_link:String(s.Dokumen||"").trim(),description:String(s.Catatan||"").trim()})).filter(s=>s.name),{apiFetch:a}=await Promise.resolve().then(()=>(F(),Ee)),i=await a("/api/sop/import",{method:"POST",body:JSON.stringify(r)});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function ia(t){A({container:t,title:"Master Checklist",icon:"\u{1F4CB}",apiPath:"/api/checklist",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(F(),Ee)),i=await a(`/api/checklist?limit=10000&${r}`);if(i.ok){let s=i.data.data.map(l=>({"Nama Checklist":l.name||"",Kategori:l.category||"",Dokumen:l.document_link||"",Deskripsi:l.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(R(),oe));o(s,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:r}=await Promise.resolve().then(()=>(R(),oe));r(e,"Template_Import_Checklist")},onImport:async e=>{let r=e.map(s=>({name:String(s["Nama Checklist"]||"").trim(),category:String(s.Kategori||"").trim(),document_link:String(s.Dokumen||"").trim(),description:String(s.Deskripsi||"").trim()})).filter(s=>s.name),{apiFetch:a}=await Promise.resolve().then(()=>(F(),Ee)),i=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(r)});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}F();xe();R();async function Ct(t,e="forms"){if(e==="supply")return Ga(t);Ua(t)}function Ua(t){A({container:t,title:"Master Form",icon:"\u{1F4D1}",apiPath:"/api/forms",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await w(`/api/forms?limit=10000&${r}`);a.data?.data?I(a.data.data,"Data_Master_Form"):z("Gagal export data master form")},onImport:async e=>{let r=await w("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!r.ok)throw new Error(r.data?.error||"Import failed");return r.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function Ga(t){let r=((await w("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));A({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,i)=>i.branch_name_ref||i.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let i=JSON.parse(a);return Array.isArray(i)?i.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let i=JSON.parse(a);return Array.isArray(i)?i.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>H(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let i=a?.tools_items;try{i=Array.isArray(JSON.parse(i))?JSON.parse(i).join(", "):i}catch{}let s=a?.chemical_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:a?.branch_id&&!r.find(o=>o.value==a.branch_id)?[...r,{value:a.branch_id,label:a.branch_name||a.branch_id}]:r,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:i},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:s},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let i=new URLSearchParams(a||{}).toString(),s=await w(`/api/reports/supply?limit=10000&${i}`);if(s.ok){let o=s.data.data.map(l=>{let n=l.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let p=l.chemical_items;try{p=Array.isArray(JSON.parse(p))?JSON.parse(p).join(", "):p}catch{}return{Waktu:l.submitted_at||"",Pengirim:l.submitter_name||"",Cabang:l.branch_name_ref||l.branch_name||"","Alat/Barang":n||"",Chemical:p||"",Catatan:l.additional_notes||"",Status:l.status||"","Diproses Oleh":l.processed_by||""}});I(o,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{I([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let s=(await w("/api/branches?all=1")).data?.data||[],o=u=>{if(!u)return null;let d=String(u||"").toLowerCase(),c=s.find(g=>String(g.full_name||"").toLowerCase()===d||String(g.code||"").toLowerCase()===d||String(g.name||"").toLowerCase()===d);return c?c.id:null},l=u=>{if(u==null||u==="")return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let d=String(u).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let g=Number(d);if(g>2e4&&g<99999){let f=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}let c=d.split(/[\/\-\.]/);if(c.length===3){let[g,f,b]=c.map(m=>m.trim());if(g.length===4&&f.length<=2&&b.length<=2)return`${g}-${f.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&f.length<=2&&g.length<=2)return`${b}-${f.padStart(2,"0")}-${g.padStart(2,"0")}`}return d},n=a.map(u=>({submitted_at:l(u.Waktu||u.Tanggal),submitter_name:String(u.Pengirim||"").trim(),branch_id:o(String(u.Cabang||"").trim()),tools_items:String(u["Alat/Barang"]||u.Alat||"").trim(),chemical_items:String(u.Chemical||"").trim(),additional_notes:String(u.Catatan||u.Keterangan||"").trim(),status:String(u.Status||"").trim(),processed_by:String(u["Diproses Oleh"]||u.PIC||"").trim()})).filter(u=>u.submitted_at&&u.submitter_name&&u.branch_id),p=await w("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,i)=>{let s=de({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(o,l)=>{let n=o.querySelector("#supply-status").value,p=o.querySelector("#supply-processed-by").value;(await w(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:p})})).ok?(W("Status diperbarui."),l(),i()):z("Gagal update status.")}})}}]})}F();R();async function ra(t){let e=ce();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}A({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:r=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[r]||"badge-neutral"}">${r}</span>`},{key:"is_active",label:"Status",render:r=>r?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:r=>r?new Date(r).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:r=>{let a=!!r;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:r?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:r?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:r?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:r?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?r?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let r=await w(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let a=r.data.data.map(i=>({"Nama Lengkap":i.full_name||"",Username:i.username||"",Email:i.email||"",Role:i.role||"",Status:i.is_active?"Aktif":"Nonaktif"}));I(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{I([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async r=>{let a=r.map(s=>({full_name:String(s["Nama Lengkap"]||"").trim(),username:String(s.Username||"").trim(),email:String(s.Email||"").trim(),role:String(s.Role||"").trim()||"viewer",password:String(s.Password||"").trim()})).filter(s=>s.username&&s.password&&s.email&&s.full_name),i=await w("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}F();R();async function la(t){A({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await w(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)I(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{I([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let r=e.map(i=>({code:String(i["Kode Cabang"]||"").trim(),name:String(i["Nama Pendek"]||"").trim(),full_name:String(i["Nama Lengkap"]||"").trim(),city:String(i.Kota||"").trim()})).filter(i=>i.code&&i.name),a=await w("/api/branches/import",{method:"POST",body:JSON.stringify(r)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}F();async function oa(t){let e=new Date,r=[];t.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),i()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),i()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(s=>s.addEventListener("change",i));async function a(){try{let s=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;r=(await w(`/api/dashboard/calendar?month=${s}`)).data?.data||[]}catch(s){console.warn("[Calendar] Failed to load events, rendering empty grid:",s),r=[]}}async function i(){let s=document.getElementById("calendar-grid");if(s){s.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let o=e.getFullYear(),l=e.getMonth(),n=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),p=document.getElementById("cal-month-label");p&&(p.textContent=n);let u=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(C=>C.value)),d=r.filter(C=>u.has(C.type)),c={};d.forEach(C=>{let x=(C.event_date||"").slice(0,10);c[x]||(c[x]=[]),c[x].push(C)});let g=new Date(o,l,1).getDay(),f=new Date(o,l+1,0).getDate(),b=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],m=new Date().toISOString().slice(0,10),y='<div class="calendar-grid">';b.forEach(C=>{y+=`<div class="cal-day-header">${C}</div>`});for(let C=0;C<g;C++)y+='<div class="cal-cell cal-cell-empty"></div>';for(let C=1;C<=f;C++){let x=`${o}-${String(l+1).padStart(2,"0")}-${String(C).padStart(2,"0")}`,T=c[x]||[],E=x===m;y+=`
          <div class="cal-cell ${E?"cal-today":""} ${T.length?"cal-has-events":""}"
               data-date="${x}" tabindex="0" role="button" aria-label="${x}">
            <div class="cal-day-num ${E?"today-num":""}">${C}</div>
            <div class="cal-events-preview">
              ${T.slice(0,3).map($=>`
                <div class="cal-event-dot cal-color-${$.color||"gray"}" title="${rt($.title||$.type)}">
                  <span class="cal-event-dot-label">${Va($.title||$.branch_name||$.type,18)}</span>
                </div>
              `).join("")}
              ${T.length>3?`<div class="cal-more">+${T.length-3} lagi</div>`:""}
            </div>
          </div>`}let _=(g+f)%7;if(_!==0)for(let C=0;C<7-_;C++)y+='<div class="cal-cell cal-cell-empty"></div>';y+="</div>",s.innerHTML=y,s.querySelectorAll(".cal-cell[data-date]").forEach(C=>{C.addEventListener("click",()=>{let x=C.dataset.date,T=c[x]||[];if(!T.length)return;let E=document.getElementById("cal-event-list"),$=new Date(x+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=$,document.getElementById("cal-event-items").innerHTML=T.map(B=>`
            <div class="cal-event-item cal-color-border-${B.color||"gray"}">
              <div class="cal-event-type">${za(B.type)}</div>
              <div class="cal-event-title">${rt(B.title||"-")}</div>
              <div class="cal-event-branch">${rt(B.branch_name||"")}</div>
              ${B.status?`<div class="cal-event-status">${rt(B.status)}</div>`:""}
              ${B.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${B.days_remaining} hari</div>`:""}
            </div>
          `).join(""),E.style.display="block"})})}catch(o){console.error("[Calendar] Render error:",o),s&&(s.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}i()}function Va(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function rt(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function za(t){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[t]||t}F();async function sa(t){let e=ce(),r=(e?.full_name||e?.username||"U")[0].toUpperCase(),i={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${i},${i}99)">
            ${r}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${e?.full_name||"\u2014"}</div>
            <div class="profile-username">@${e?.username||"\u2014"}</div>
            <span class="badge badge-info" style="background:${i}18;color:${i};margin-top:6px">
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
            <span class="info-value" style="color:${i};font-weight:700">${e?.role||"\u2014"}</span>
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
  `;let s=localStorage.getItem("fm_token"),o=document.getElementById("session-info");if(s&&o)try{let l=JSON.parse(atob(s.split(".")[1])),n=new Date(l.exp*1e3);o.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{o.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async l=>{l.preventDefault();let n=document.getElementById("pwd-error"),p=document.getElementById("pwd-success"),u=document.getElementById("btn-save-pwd");n.style.display="none",p.style.display="none";let d=l.target,c=d.current_password.value,g=d.new_password.value,f=d.confirm_password.value;if(g!==f){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(g.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}u.disabled=!0,u.textContent="\u23F3 Menyimpan...";let b=await w("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:c,new_password:g})});u.disabled=!1,u.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',b.ok?(p.textContent="\u2705 Password berhasil diubah.",p.style.display="block",d.reset(),W("Password berhasil diubah.")):(n.textContent=b.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}F();var lt={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function Z(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let i=Number(e);if(i>2e4&&i<99999){let s=new Date(Date.UTC(1899,11,30)+i*864e5);return isNaN(s.getTime())?null:s.toISOString().slice(0,10)}}let r=e.split(/[\/\-\.]/);if(r.length===3){let[i,s,o]=r.map(u=>u.trim()),l=Number(i),n=Number(s),p=Number(o);if(i.length===4&&l>1900)return`${i}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`;if(o.length===4&&p>1900)return l>12?`${o}-${s.padStart(2,"0")}-${i.padStart(2,"0")}`:n>12?`${o}-${i.padStart(2,"0")}-${s.padStart(2,"0")}`:`${o}-${s.padStart(2,"0")}-${i.padStart(2,"0")}`;if(o.length===2&&!isNaN(p)){let u=p>=50?`19${o}`:`20${o}`;return l>12?`${u}-${s.padStart(2,"0")}-${i.padStart(2,"0")}`:`${u}-${s.padStart(2,"0")}-${i.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function ca(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Qa={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",phone:t["No. Hp"],join_date:Z(t["Tanggal Masuk"]),status:t.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:Z(t["Tanggal Mulai"]),end_date:Z(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:Z(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:Z(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:Z(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:Z(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:Z(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:Z(t["Tanggal Target"]||t["Tgl Target"]),completion_date:Z(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:Z(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:Z(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:Z(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:Z(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:Z(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:Z(t["Tanggal Back Up"]),completion_date:Z(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:Z(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:Z(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Ya(t,e){let r=lt[t];if(!r)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Qa[r.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let i=[],s=[],o=[];return e.filter(n=>!ca(n)).forEach((n,p)=>{let u=e.indexOf(n)+2,d=[];a.required.forEach(({key:g,label:f})=>{let b=n[g];if(b==null||String(b).trim()===""){let m=Object.keys(n).filter(y=>y.trim()).join(", ");d.push({column:f,originalValue:b||"",reason:`Kolom "${f}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${m.slice(0,120)}`})}});let c=a.map(n);d.length>0?s.push({row:u,data:c,raw:n,errors:d}):(i.push(n),o.push(c))}),{valid:i,errors:s,mapped:o}}function da(t){let e=[];return t.SheetNames.forEach(r=>{let a=lt[r];if(!a)return;let i=t.Sheets[r],s=window.XLSX.utils.sheet_to_json(i,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),o=Ya(r,s),l=s.filter(n=>!ca(n));e.push({sheetName:r,module:a.module,label:a.label,total:l.length,valid:o.mapped.length,errorCount:o.errors.length,errors:o.errors,mapped:o.mapped,skipped:!1})}),e}function pa(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,i])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(i),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function ma(t){let e=window.XLSX,r=e.utils.book_new(),a=!1;return t.forEach(i=>{if(!i.errors||i.errors.length===0)return;a=!0;let s=i.errors.map(l=>({"No. Baris":l.row,"Kolom Gagal":(l.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(l.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(l.data||{}).map(([n,p])=>[n,p??""]))})),o=e.utils.json_to_sheet(s);e.utils.book_append_sheet(r,o,i.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(r,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Wa=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function ua(t){t.innerHTML=`
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
              ${Object.entries(lt).map(([b,{label:m}])=>`<span class="import-sheet-tag">\u{1F4C4} ${b} \u2192 ${m}</span>`).join("")}
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
  `;let e=null,r=null,a=0,i={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function s(b){Object.entries(i).forEach(([m,y])=>{y.style.display=m===b?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let b=document.getElementById("btn-backup-db");b.disabled=!0,b.textContent="\u23F3 Memproses Backup...";try{let m=await w("/api/import/backup");if(m.ok){if(!window.XLSX){z("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let y=window.XLSX,S=y.utils.book_new();Object.entries(m.data.database).forEach(([_,C])=>{let x=C.length>0?C:[{}],T=y.utils.json_to_sheet(x);y.utils.book_append_sheet(S,T,_.substring(0,31))}),y.writeFile(S,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),W("Backup berhasil diunduh!")}else z("Gagal memproses backup: "+(m.data?.error||"Unknown error"))}catch(m){z("Gagal memproses backup: "+m.message)}finally{b.disabled=!1,b.textContent="\u{1F4E6} Backup Database"}});let o=document.getElementById("btn-sync-google");o&&o.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let b=o.innerHTML;o.innerHTML='<span class="spinner"></span> Menyinkronkan...',o.disabled=!0;try{let m=await w("/api/sync/google-sheets",{method:"POST"});m.ok?alert("Sinkronisasi Berhasil: "+(m.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(m.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{o.innerHTML=b,o.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{pa(),W("Template Excel berhasil didownload!")});let l=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",b=>{b.stopPropagation(),l.click()}),l.addEventListener("change",b=>{b.target.files[0]&&p(b.target.files[0])}),n.addEventListener("dragover",b=>{b.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",b=>{b.preventDefault(),n.classList.remove("drag-over");let m=b.dataTransfer.files[0];m&&m.name.match(/\.xlsx?$/i)?p(m):z("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,l.value="",document.getElementById("file-info").style.display="none",n.style.display="",s("upload")});async function p(b){e=b,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${b.name} (${(b.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await u(b)}async function u(b){s("validating");let m=document.getElementById("validation-status"),y=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");m.textContent="Membaca file Excel...",y.style.width="20%",await Ve(200);let S=await b.arrayBuffer(),_=window.XLSX.read(S,{type:"array",cellDates:!0});m.textContent=`Memvalidasi ${_.SheetNames.length} sheet...`,y.style.width="50%",await Ve(100),r=da(_),y.style.width="100%",m.textContent="Validasi selesai!",await Ve(300),d()}catch(S){s("upload"),z("Gagal memproses file: "+S.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function d(){s("preview");let b=r.filter($=>!$.skipped).length,m=r.reduce(($,B)=>$+B.total,0),y=r.reduce(($,B)=>$+B.valid,0),S=r.reduce(($,B)=>$+B.errorCount,0),_=m>0?Math.round(y/m*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${b} sheet</span>
      <span class="badge badge-secondary">${m} baris</span>
      <span class="badge badge-success">${y} valid (${_}%)</span>
      ${S>0?`<span class="badge badge-danger">${S} error</span>`:""}
    `;let C=document.getElementById("preview-table-container");C.innerHTML=`
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
          ${r.map(($,B)=>`
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
                ${$.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${B}">\u{1F50D} ${$.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,C.querySelectorAll(".btn-detail-error").forEach($=>{$.addEventListener("click",()=>{let B=r[Number($.dataset.idx)];c(B)})});let x=document.getElementById("error-detail-section"),T=document.getElementById("error-detail-container");T.innerHTML="",x.style.display="none";let E=document.getElementById("btn-start-import");y===0?(E.disabled=!0,E.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(E.disabled=!1,S>0?(E.innerHTML=`\u{1F680} Import ${y} Data Valid (${S} dilewati)`,E.title="Baris error akan dilewati, baris valid tetap diimport"):E.innerHTML=`\u{1F680} Mulai Import ${y} Data`)}function c(b){let m=document.getElementById("error-detail-section"),y=document.getElementById("error-detail-container");m.style.display="";let S=b.errors.slice(0,100).map(_=>(Array.isArray(_.errors)?_.errors:[]).map(x=>{let T=typeof x=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${_.row}</span></td>
            <td><strong>${T?x.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${T&&x.originalValue!==void 0?x.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${T?x.reason:x}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${T&&x.aliases?`Gunakan salah satu nama kolom:<br><em>${x.aliases}</em>`:T&&x.hint?x.hint:""}
            </td>
          </tr>
        `}).join("")).join("");y.innerHTML=`
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
            <tbody>${S||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${b.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,m.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{s("upload"),document.getElementById("file-info").style.display="none",n.style.display="",e=null,l.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!r)return;ma(r)?W("Log error berhasil didownload."):W("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let b=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";g(b)});async function g(b){s("importing"),a=Date.now();let m=[];Wa.forEach(x=>{let T=r?.find(E=>E.module===x&&E.mapped?.length>0);T&&m.push(T)});let y=document.getElementById("import-steps-list");y.innerHTML=m.map(x=>`
      <div class="import-step-item" id="step-item-${x.module}">
        <span class="step-item-icon" id="step-icon-${x.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${x.label} <span class="step-item-count">(${x.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${x.module}"></span>
      </div>
    `).join("");let S=document.getElementById("import-bar"),_=document.getElementById("import-current-status"),C={totalSheets:m.length,totalRows:m.reduce((x,T)=>x+T.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let x=0;x<m.length;x++){let T=m[x],E=document.getElementById(`step-icon-${T.module}`),$=document.getElementById(`step-status-${T.module}`);E.textContent="\u{1F504}",$.textContent="Mengimport...",_.textContent=`Mengimport ${T.label}...`,S.style.width=`${Math.round(x/m.length*100)}%`;try{let B=await w(`/api/import/${T.module}`,{method:"POST",body:JSON.stringify({rows:T.mapped,onDuplicate:b})});if(B.ok){let P=B.data;C.inserted+=P.inserted||0,C.skipped+=P.skipped||0,C.moduleResults.push({label:T.label,inserted:P.inserted||0,skipped:P.skipped||0,status:"ok"}),E.textContent="\u2705",$.innerHTML=`<span class="badge badge-success">${P.inserted||0} berhasil</span>${P.skipped>0?` <span class="badge badge-neutral">${P.skipped} skip</span>`:""}`}else C.failed++,C.moduleResults.push({label:T.label,inserted:0,skipped:0,status:"error",error:B.data?.error}),E.textContent="\u274C",$.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(B){C.failed++,C.moduleResults.push({label:T.label,inserted:0,skipped:0,status:"error",error:B.message}),E.textContent="\u274C",$.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Ve(150)}S.style.width="100%",_.textContent="Selesai!",await Ve(400),f(C)}function f(b){s("summary");let m=((Date.now()-a)/1e3).toFixed(1),y=b.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${y?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${y?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
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
        <div class="stat-value">${m}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${b.moduleResults.map(S=>`
            <tr>
              <td>${S.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${S.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${S.skipped}</span></td>
              <td style="text-align:center">
                ${S.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${S.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,r=null,l.value="",document.getElementById("file-info").style.display="none",n.style.display="",s("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Ve(t){return new Promise(e=>setTimeout(e,t))}F();var ot=[],ga=[];async function ba(t){ot=await K(),ga=await X(),A({container:t,title:"Data SP (Surat Peringatan)",icon:"\u{1F4DC}",apiPath:"/api/sp",enableMobileFilterSheet:!0,itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:ot}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await w(`/api/sp?limit=10000&${r}`);if(a.ok){let i=a.data.data.map(o=>({"Nama Karyawan":o.employee_name||"",Divisi:o.division||"",Cabang:o.branch_name||"","Tanggal Sp":o.tanggal||"","Akhir Sp":o.akhir_sp||"","Jenis Sp":o.sp_type||"","Link Document / Foto":o.document_link||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(R(),oe));s(i,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(R(),oe));r(e,"Template_Import_SP")},onImport:async e=>{let r=o=>{if(!o)return null;let l=String(o||"").toLowerCase(),n=ot.find(p=>String(p.label||"").toLowerCase()===l);return n?n.value:null},a=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let l=String(o).trim();if(/^\d{4,5}$/.test(l)){let p=Number(l);if(p>2e4&&p<99999){let u=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);let n=l.split(/[\/\-\.]/);if(n.length===3){let[p,u,d]=n.map(c=>c.trim());if(p.length===4&&u.length<=2&&d.length<=2)return`${p}-${u.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&u.length<=2&&p.length<=2)return`${d}-${u.padStart(2,"0")}-${p.padStart(2,"0")}`}return l},i=e.map(o=>({employee_name:String(o["Nama Karyawan"]||"").trim(),division:String(o.Divisi||"").trim(),branch_id:r(String(o.Cabang||"").trim()),tanggal:a(o["Tanggal Sp"]),akhir_sp:a(o["Akhir Sp"]),sp_type:String(o["Jenis Sp"]||"").trim(),document_link:String(o["Link Document / Foto"]||"").trim()})).filter(o=>o.employee_name&&o.branch_id),s=await w("/api/import/sp",{method:"POST",body:JSON.stringify({rows:i,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:[{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:ga},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:ot,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}F();var Fe=[],ha=[];async function ya(t){Fe=await K(),ha=await X(),A({container:t,title:"Data Mutasi",icon:"\u{1F501}",apiPath:"/api/mutasi",enableMobileFilterSheet:!0,itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:Fe},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:Fe}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await w(`/api/mutasi?limit=10000&${r}`);if(a.ok){let i=a.data.data.map(o=>({Tanggal:o.tanggal||"","Nama Karyawan":o.employee_name||"","Cabang Asal":o.from_branch_name||"","Cabang Tujuan":o.to_branch_name||"",Status:o.status||"",Dokumen:o.document_link||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(R(),oe));s(i,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(R(),oe));r(e,"Template_Import_Mutasi")},onImport:async e=>{let r=o=>{if(!o)return null;let l=String(o||"").toLowerCase(),n=Fe.find(p=>String(p.label||"").toLowerCase()===l);return n?n.value:null},a=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let l=String(o).trim();if(/^\d{4,5}$/.test(l)){let p=Number(l);if(p>2e4&&p<99999){let u=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);let n=l.split(/[\/\-\.]/);if(n.length===3){let[p,u,d]=n.map(c=>c.trim());if(p.length===4&&u.length<=2&&d.length<=2)return`${p}-${u.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&u.length<=2&&p.length<=2)return`${d}-${u.padStart(2,"0")}-${p.padStart(2,"0")}`}return l},i=e.map(o=>({tanggal:a(o.Tanggal),employee_name:String(o["Nama Karyawan"]||"").trim(),from_branch_id:r(String(o["Cabang Asal"]||"").trim()),to_branch_id:r(String(o["Cabang Tujuan"]||"").trim()),status:String(o.Status||"").trim(),document_link:String(o.Dokumen||"").trim()})).filter(o=>o.tanggal&&o.employee_name&&o.from_branch_id&&o.to_branch_id),s=await w("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:i,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:ha},{type:"select",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Fe,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Fe,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let r=Math.floor(Number(t)-25569);return new Date(r*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let r=t.split(/[\/\-]/);return`${r[2]}-${r[1]}-${r[0]}`}let e=t.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(e){let r=e[1],a=parseInt(e[2],10),i=parseInt(e[3],10);if(a>12&&i<=12)return`${r}-${e[3]}-${e[2]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let r=e.split("-");if(r.length===3&&r[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],i=parseInt(r[2],10),s=a[parseInt(r[1],10)-1];return`${i} ${s} ${r[0]}`}return e};function G(t){return async e=>{if(!Oe()){$e("/login");return}return t(e)}}var ze=null;function Xa(){ze&&clearInterval(ze);let t=()=>{let e=new Date,r=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),i=document.getElementById("header-clock-time"),s=document.getElementById("header-clock-date");i&&(i.textContent=r),s&&(s.textContent=a)};t(),ze=setInterval(t,1e3)}async function Za(){try{let t=await w("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},r=(a,i)=>{let s=document.getElementById(a);s&&(s.textContent=i>0?i:"",s.style.display=i>0?"inline-flex":"none")};r("badge-issues",e.issues?.current||0),r("badge-contracts",e.expiring30?.current||0),r("badge-oo1",e.one_on_one?.current||0),r("badge-schedule",e.schedule?.current||0),r("badge-supply",e.supply?.current||0)}catch{}}var Ne=[];async function en(){try{let t=await w("/api/dashboard/notifications");if(!t.ok)return;Ne=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Ne.length>0?"block":"none",e.textContent=Ne.length)}catch{}}function tn(){if(!Ne.length){de({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,r)=>r()});return}let t=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${Ne.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;de({title:`Notifikasi (${Ne.length})`,content:t,confirmText:"Tutup",onConfirm:(e,r)=>r()})}function fa(){let t=ce(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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

          <!-- Vendor & Quality -->
          <div class="nav-section">
            <span class="nav-section-label">VENDOR & QUALITY</span>
            <a href="#/vendors" class="nav-item" data-route="/vendors">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M2 20a2 2 0 012-2h16a2 2 0 012 2v2H2v-2z"/><path d="M16 21V5a2 2 0 00-2-2h-4a2 2 0 00-2 2v16"/></svg>
              </span>
              <span class="nav-label">Vendor</span>
            </a>
            <a href="#/reports/inspection" class="nav-item" data-route="/reports/inspection">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              </span>
              <span class="nav-label">Inspection</span>
            </a>
            <a href="#/issues" class="nav-item" data-route="/issues">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
              </span>
              <span class="nav-label">Findings</span>
              <span class="nav-badge badge-danger" id="badge-issues-2"></span>
            </a>
            <a href="#/corrective-actions" class="nav-item" data-route="/corrective-actions">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 7l-5 5 5 5"/><path d="M20 4l-10 10"/></svg>
              </span>
              <span class="nav-label">Corrective Actions</span>
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
                <span class="topbar-greeting-time">${(()=>{let p=new Date().getHours();return p>=4&&p<11?"Selamat Pagi":p>=11&&p<15?"Selamat Siang":p>=15&&p<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${t?.full_name||t?.username||"Guest"}</span> \u{1F44B}
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
  `;let r=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),i=document.getElementById("topbar-menu-btn"),s=document.getElementById("sidebar-close"),o=()=>{r.classList.add("open"),a.classList.add("show")},l=()=>{r.classList.remove("open"),a.classList.remove("show")};i?.addEventListener("click",o),s?.addEventListener("click",l),a?.addEventListener("click",l),document.querySelectorAll(".nav-item").forEach(p=>p.addEventListener("click",l));function n(){let p=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(c=>{let g=c.dataset.route;c.classList.toggle("active",p===g||g!=="/dashboard"&&p.startsWith(g))});let u=document.getElementById("topbar-title"),d=document.querySelector(".nav-item.active .nav-label");u&&d&&(u.textContent=d.textContent)}window.addEventListener("hashchange",n),n(),Xa(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await w("/api/auth/logout",{method:"POST"}),Re(),ze&&clearInterval(ze),$e("/login")}),Za(),en(),document.getElementById("btn-notif")?.addEventListener("click",p=>{p.preventDefault(),tn()})}async function an(){U("/login",({main:e})=>jt(e)),U("/dashboard",G(({main:e})=>Nt(e))),U("/calendar",G(({main:e})=>oa(e))),U("/employees",G(({main:e,params:r})=>qt(e,r))),U("/contracts",G(({main:e,params:r})=>nt(e,r))),U("/sp",G(({main:e})=>ba(e))),U("/mutasi",G(({main:e})=>ya(e))),U("/timeline",G(({main:e,params:r})=>Jt(e,r))),U("/issues",G(({main:e,params:r})=>Gt(e,r))),U("/corrective-actions",G(({main:e,params:r})=>aa(e,r))),U("/one-on-one",G(({main:e,params:r})=>Vt(e,r))),U("/training",G(({main:e})=>zt(e))),U("/relievers",G(({main:e,params:r})=>Yt(e,r))),U("/reports/inspection",G(({main:e})=>Wt(e))),U("/reports/cleaning",G(({main:e})=>Xt(e))),U("/reports/fogging",G(({main:e})=>Zt(e))),U("/reports/basecamp",G(({main:e})=>ea(e))),U("/reports/supply",G(({main:e})=>Ct(e,"supply"))),U("/sop",G(({main:e})=>na(e))),U("/checklist",G(({main:e})=>ia(e))),U("/forms",G(({main:e})=>Ct(e))),U("/users",G(({main:e})=>ra(e))),U("/branches",G(({main:e})=>la(e))),U("/vendors",G(({main:e})=>ta(e))),U("/profile",G(({main:e})=>sa(e))),U("/settings/import",G(({main:e})=>ua(e)));let t=Oe();if(!t&&window.location.hash!=="#/login"&&$e("/login"),t){let e=await w("/api/auth/me");e.ok?(Ke(e.data.data),fa()):(Re(),$e("/login"))}window.addEventListener("fm:login",()=>{fa(),$e("/dashboard")}),$t()}an();
