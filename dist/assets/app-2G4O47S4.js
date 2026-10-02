var ba=Object.defineProperty;var ct=(t,e)=>()=>(t&&(e=t(t=0)),e);var pt=(t,e)=>{for(var i in e)ba(t,i,{get:e[i],enumerable:!0})};var Ee={};pt(Ee,{API:()=>Ct,CLIENT_SIDE_MAX_ROWS:()=>$e,IS_DEVELOPMENT:()=>Me,apiFetch:()=>x,clearToken:()=>Ke,getToken:()=>Oe,getUser:()=>ie,setToken:()=>mt,setUser:()=>Re});function Oe(){return localStorage.getItem("fm_token")}function mt(t){localStorage.setItem("fm_token",t)}function Ke(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ie(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Re(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function x(t,e={}){let i=Oe(),a={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...e.headers||{}};try{let r=`cb=${Date.now()}`,s=t.includes("?")?"&":"?",l=`${Ct}${t}${s}${r}`,o=await fetch(l,{...e,headers:a}),n;try{let d=await o.text();try{n=JSON.parse(d)}catch{n={error:`Server Error (${o.status}): ${d.substring(0,80)}...`}}}catch{n={error:"Gagal membaca respon dari server"}}return o.status===401&&(Ke(),window.location.hash="#/login"),{ok:o.ok,status:o.status,data:n}}catch(r){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${r.message})`}}}}var Me,ha,Ct,$e,A=ct(()=>{Me=!1,ha="https://fm-operations-api.facilitycare-audydental.workers.dev",Ct=ha,$e=1e4});var gt={};pt(gt,{confirmDialog:()=>qe,createModal:()=>de});function de({title:t,content:e,onConfirm:i,onCancel:a,confirmText:r="Simpan",cancelText:s="Batal",size:l="md",confirmClass:o="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},d=document.createElement("div");d.className="modal-overlay",d.innerHTML=`
    <div class="modal" style="max-width:${n[l]||n.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${s}</button>
        ${i?`<button class="btn ${o} modal-confirm">${r}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&d.querySelector(".modal-body").appendChild(e);let u=()=>{d.classList.remove("show"),setTimeout(()=>d.remove(),250)};return d.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),u()}),d.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),u()}),i&&d.querySelector(".modal-confirm").addEventListener("click",()=>i(d,u)),d.addEventListener("click",g=>{g.target===d&&(a&&a(),u())}),document.body.appendChild(d),requestAnimationFrame(()=>d.classList.add("show")),{overlay:d,close:u}}function qe(t,e,i="Konfirmasi"){return de({title:i,content:`<p>${t}</p>`,onConfirm:(a,r)=>{e(),r()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var _e=ct(()=>{});var re={};pt(re,{downloadExcel:()=>F,parseExcel:()=>Je,renderExcelButtons:()=>ya});function Je(t){return new Promise((e,i)=>{let a=new FileReader;a.onload=r=>{try{let s=new Uint8Array(r.target.result),l=XLSX.read(s,{type:"array"}),o=l.SheetNames[0],n=l.Sheets[o];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${l.SheetNames.join(", ")}`),console.log(`Sheet Used: ${o}`);let d=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),u=d.e.r-d.s.r+1,g=d.e.c-d.s.c+1;console.log(`Total Rows (including empty): ${u}`),console.log(`Total Columns: ${g}`);let p=[];for(let f=d.s.c;f<=d.e.c;++f){let h=n[XLSX.utils.encode_cell({c:f,r:d.s.r})];h&&h.v&&p.push(h.v)}console.log(`Headers Found: ${p.join(", ")}`),console.log("---------------------------");let c=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(c,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(c,"__headers",{value:p,enumerable:!1}),e(c)}catch(s){i(s)}},a.onerror=r=>i(r),a.readAsArrayBuffer(t)})}function F(t,e){try{let i=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,i,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function ya(t){return`
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
  `}var j=ct(()=>{});A();var ut={},Ve=null;function J(t,e){ut[t]=e}function xe(t){window.location.hash=t}function $t(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[i,...a]=e.split("?"),r=ut[i];if(!r){for(let[l,o]of Object.entries(ut))if(l.endsWith("/*")&&i.startsWith(l.slice(0,-2))){r=o;break}}Ve&&(Ve(),Ve=null);let s=document.getElementById("main-content");if(s&&(s.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),r){let l=new URLSearchParams(a.join("?")),o=i.split("/").filter(Boolean),n=await r({path:i,params:l,segments:o,main:s});n&&(Ve=n)}else{let l=s||document.getElementById("app");l&&(l.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var je;function fa(){return je||(je=document.createElement("div"),je.id="toast-container",document.body.appendChild(je)),je}function Tt(t,e="info",i=3500){let a=fa(),r=document.createElement("div");r.className=`toast toast-${e}`;let s={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};r.innerHTML=`<span class="toast-icon">${s[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(r),requestAnimationFrame(()=>r.classList.add("show")),setTimeout(()=>{r.classList.remove("show"),setTimeout(()=>r.remove(),350)},i)}var W=t=>Tt(t,"success"),V=t=>Tt(t,"error");_e();A();A();_e();function Ye({columns:t,data:e,onEdit:i,onDelete:a,onView:r,actions:s=[],emptyText:l="Tidak ada data",bulkSelect:o=null}){let n=document.createElement("div");if(n.className="table-wrapper",!e||e.length===0)return n.innerHTML=`<div class="empty-state"><p>${l}</p></div>`,n;let d=document.createElement("table");d.className="data-table";let u=document.createElement("thead"),g=document.createElement("tr");if(o){let c=document.createElement("th");c.style.width="40px",c.style.textAlign="center";let f=document.createElement("input");f.type="checkbox",f.id="select-all-checkbox",f.title="Pilih semua",f.addEventListener("change",()=>{e.forEach(h=>{f.checked?o.selectedIds.add(h.id):o.selectedIds.delete(h.id)}),n.querySelectorAll(".row-checkbox").forEach(h=>h.checked=f.checked),o.onToggle()}),c.appendChild(f),g.appendChild(c)}if(t.forEach(c=>{let f=document.createElement("th");f.textContent=c.label,c.width&&(f.style.width=c.width),g.appendChild(f)}),i||a||r||s.length>0){let c=document.createElement("th");c.textContent="Aksi",c.style.width="120px",g.appendChild(c)}u.appendChild(g),d.appendChild(u);let p=document.createElement("tbody");return e.forEach(c=>{let f=document.createElement("tr");if(o){let h=document.createElement("td");h.style.textAlign="center",h.style.width="40px";let m=document.createElement("input");m.type="checkbox",m.className="row-checkbox",m.checked=o.selectedIds.has(c.id),m.addEventListener("change",()=>{if(m.checked)o.selectedIds.add(c.id);else{o.selectedIds.delete(c.id);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1)}o.onToggle()}),h.appendChild(m),f.appendChild(h)}if(t.forEach(h=>{let m=document.createElement("td");if(h.render){let y=h.render(c[h.key],c);y instanceof HTMLElement?m.appendChild(y):m.innerHTML=y||""}else m.textContent=c[h.key]!==null&&c[h.key]!==void 0&&c[h.key]!==""?c[h.key]:"";h.nowrap&&(m.style.whiteSpace="nowrap"),f.appendChild(m)}),i||a||r||s.length>0){let h=document.createElement("td");h.className="actions-cell";let m=document.createElement("div");if(m.className="btn-group",r){let y=document.createElement("button");y.className="btn btn-xs btn-ghost",y.innerHTML="\u{1F441}",y.title="Lihat",y.addEventListener("click",()=>r(c)),m.appendChild(y)}if(i){let y=document.createElement("button");y.className="btn btn-xs btn-secondary",y.innerHTML="\u270F\uFE0F",y.title="Edit",y.addEventListener("click",()=>i(c)),m.appendChild(y)}s.forEach(y=>{let S=document.createElement("button");S.className=`btn btn-xs ${y.class||"btn-ghost"}`,S.innerHTML=y.icon||y.label,S.title=y.label,S.addEventListener("click",()=>y.handler(c)),m.appendChild(S)}),h.appendChild(m),f.appendChild(h)}p.appendChild(f)}),d.appendChild(p),n.appendChild(d),n}function We({page:t,pages:e,total:i,limit:a,onPage:r}){if(e<=1)return null;let s=document.createElement("div");s.className="pagination";let l=document.createElement("span");l.className="pagination-info",l.textContent=`Total: ${i} data`,s.appendChild(l);let o=document.createElement("div");o.className="pagination-btns";let n=(g,p,c=!1,f=!1)=>{let h=document.createElement("button");h.className=`btn btn-sm ${f?"btn-primary":"btn-ghost"} pagination-btn`,h.textContent=g,h.disabled=c,h.addEventListener("click",()=>r(p)),o.appendChild(h)};n("\xAB",1,t===1),n("\u2039",t-1,t===1);let d=Math.max(1,t-2),u=Math.min(e,t+2);for(let g=d;g<=u;g++)n(g,g,!1,g===t);return n("\u203A",t+1,t===e),n("\xBB",e,t===e),s.appendChild(o),s}_e();function He(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${He(e.fields)}</div>`;let i=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",r="";switch(e.type){case"textarea":r=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${i} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let l=(e.options||[]).map(p=>{let c=typeof p=="object"?p.value:p,f=typeof p=="object"?p.label:p,h=e.value==c?"selected":"";return`<option value="${c}" ${h}>${f}</option>`}).join("");r=`<select name="${e.name}" class="form-control" ${i}><option value="">-- Pilih ${e.label||""} --</option>${l}</select>`;break;case"combobox":let o=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,n=(e.options||[]).map(p=>{let c=typeof p=="object"?p.value:p,f=typeof p=="object"?p.label||p.value||"":p||"";return(f==="undefined"||f==="[object Object]"||f==="null")&&(f=""),f?`<option value="${f}"></option>`:""}).join(""),d=e.value||"";if(e.value){let p=(e.options||[]).find(c=>(typeof c=="object"?c.value:c)==e.value);if(p){let c=typeof p=="object"?p.label||p.value||"":p||"";c&&c!=="undefined"&&c!=="[object Object]"&&c!=="null"&&(d=c)}}r=`
          <input type="text" name="${e.name}" list="${o}" class="form-control" value="${d}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${o}">${n}</datalist>
        `;break;case"checkbox":r=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let u=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";r=`<input type="date" name="${e.name}" class="form-control" value="${u}" ${i}>`;break;case"number":let g=e.readonly?"readonly":"";r=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${i} ${g}>`;break;case"email":r=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i}>`;break;case"url":r=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${i}>`;break;default:r=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i} autocomplete="off">`}let s=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${r}${s}</div>`}).join("")}function Xe(t){let e={},i=new FormData(t);for(let[a,r]of i.entries())e[a]=r===""?null:r;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function Ze(t,e){e&&Object.entries(e).forEach(([i,a])=>{let r=t.querySelector(`[name="${i}"]`);r&&(r.hasAttribute("list")||(r.type==="checkbox"?r.checked=!!a:r.type==="date"&&a&&window.parseFlexibleDate?r.value=window.parseFlexibleDate(a):r.value=a??""))})}j();var Se={},et={on(t,e){Se[t]||(Se[t]=new Set),Se[t].add(e)},off(t,e){Se[t]&&Se[t].delete(e)},emit(t,e){Se[t]&&Se[t].forEach(i=>{try{i(e)}catch(a){console.warn("[calendarBus] Handler error:",a)}})},clear(){Object.keys(Se).forEach(t=>delete Se[t])}},va=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function oe(t){if(!t){et.emit("data:changed",{module:"unknown"});return}let e=String(t).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");et.emit("data:changed",{module:e,relevant:va.has(e)})}function P({container:t,title:e,icon:i,apiPath:a,columns:r,formFields:s,filterFields:l,defaultFilters:o={},itemLabel:n="Data",canCreate:d=!0,canEdit:u=!0,canDelete:g=!0,onBeforeSubmit:p,onAfterLoad:c,onDataLoaded:f,extraActions:h=[],initialSearch:m="",exportOptions:y=null,bulkDelete:S=!1,paginationMode:D="server"}){let T=ie();T&&T.role==="viewer"&&(d=!1,u=!1,g=!1,S=!1,y=null),T&&T.role==="editor_khusus"&&a!=="/api/issues"&&a!=="/api/overtime"&&(d=!1,u=!1,g=!1,S=!1,y=null),T&&T.role==="input_lembur"&&(a==="/api/overtime"?(g=!1,S=!1,y=null):(d=!1,u=!1,g=!1,S=!1,y=null));let w=1,$={...o};m&&($.search=m);let _=new Set;t.innerHTML=`
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
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${d?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
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
    

    ${l&&l.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${l.filter(b=>b.type==="search").map(b=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${b.placeholder||"Cari..."}" id="filter-search" value="${$.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${l.filter(b=>b.type!=="search").map(b=>{if(b.type==="select"||b.type==="combobox"){let k=(b.label||"").startsWith("Pilih")?b.label:`Pilih ${b.label||""}`;return`<select class="filter-select" name="${b.name}" id="filter-${b.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${k}</option>${(b.options||[]).map(v=>`<option value="${typeof v=="object"?v.value:v}" ${$[b.name]===(typeof v=="object"?v.value:v)?"selected":""}>${typeof v=="object"?v.label:v}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${l.filter(b=>b.type!=="search").map(b=>{if(b.type==="select"||b.type==="combobox"){let k=(b.label||"").startsWith("Pilih")?b.label:`Pilih ${b.label||""}`;return`<select class="filter-select filter-select-sheet" name="${b.name}-sheet" id="filter-sheet-${b.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${k}</option>${(b.options||[]).map(v=>`<option value="${typeof v=="object"?v.value:v}" ${$[b.name]===(typeof v=="object"?v.value:v)?"selected":""}>${typeof v=="object"?v.label:v}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function C(){let b=document.getElementById("bulk-toolbar");if(!b)return;let k=document.getElementById("bulk-count"),v=document.getElementById("btn-bulk-delete"),M=document.getElementById("btn-bulk-cancel");k.textContent=`${_.size} item dipilih`,_.size>0?(b.style.display="flex",v.disabled=!1,M.disabled=!1):(b.style.display="none",v.disabled=!0,M.disabled=!0);let N=document.getElementById("select-all-checkbox");if(N){let H=document.querySelectorAll(".row-checkbox");if(H.length>0){let I=[...H].every(ne=>ne.checked),L=[...H].some(ne=>ne.checked);N.checked=I,N.indeterminate=L&&!I}else N.checked=!1,N.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{_.clear(),document.querySelectorAll(".row-checkbox").forEach(k=>k.checked=!1);let b=document.getElementById("select-all-checkbox");b&&(b.checked=!1),C()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(_.size===0)return;let b=[..._],k=document.createElement("div");k.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",k.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${b.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${b.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(k),k.querySelector("#bulk-cancel-btn").addEventListener("click",()=>k.remove()),k.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let v=k.querySelector("#bulk-confirm-btn");v.disabled=!0,v.textContent="Menghapus...";let M=await x(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:b})});k.remove(),M.ok?(W(`${b.length} ${n} berhasil dihapus.`),_.clear(),C(),oe(a),R()):V(M.data?.error||"Gagal menghapus data.")})});let B=document.getElementById("filter-search"),E;B?.addEventListener("input",b=>{clearTimeout(E),E=setTimeout(()=>{$.search=b.target.value,w=1,_.clear(),C(),R()},400)}),l?.forEach(b=>{(b.type==="select"||b.type==="combobox")&&(document.getElementById(`filter-${b.name}`)?.addEventListener("change",k=>{$[b.name]=k.target.value;let v=document.getElementById(`filter-sheet-${b.name}`);v&&(v.value=k.target.value),w=1,_.clear(),C(),R()}),document.getElementById(`filter-sheet-${b.name}`)?.addEventListener("change",k=>{$[b.name]=k.target.value;let v=document.getElementById(`filter-${b.name}`);v&&(v.value=k.target.value),w=1,_.clear(),C(),R(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{$={},B&&(B.value=""),l?.forEach(b=>{let k=document.getElementById(`filter-${b.name}`);k&&(k.value="");let v=document.getElementById(`filter-sheet-${b.name}`);v&&(v.value="")}),w=1,_.clear(),C(),R()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{$={},B&&(B.value=""),l?.forEach(b=>{let k=document.getElementById(`filter-${b.name}`);k&&(k.value="");let v=document.getElementById(`filter-sheet-${b.name}`);v&&(v.value="")}),w=1,_.clear(),C(),R(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>ye(null)),y&&document.addEventListener("click",function(b){let k=document.getElementById("aksi-menu-main"),v=document.getElementById("btn-aksi-main");k&&v&&!v.contains(b.target)&&!k.contains(b.target)&&k.classList.remove("show-aksi-menu")});let z=document.getElementById("btn-mobile-filter"),ee=document.getElementById("filter-options-wrapper"),be=document.getElementById("btn-close-filter-sheet");if(z&&ee&&(z.addEventListener("click",b=>{b.preventDefault(),ee.classList.add("sheet-open")}),be&&be.addEventListener("click",b=>{b.preventDefault(),ee.classList.remove("sheet-open")})),y){document.getElementById(`btn-export-${y.moduleName}`)?.addEventListener("click",async k=>{let v=k.target,M=v.innerHTML;v.innerHTML="\u23F3 Loading...",v.disabled=!0;try{await y.onExport()}catch{V("Gagal export data")}finally{v.innerHTML=M,v.disabled=!1}}),document.getElementById(`btn-template-${y.moduleName}`)?.addEventListener("click",()=>{y.onTemplate()});let b=document.getElementById(`input-import-${y.moduleName}`);b?.addEventListener("change",async k=>{let v=k.target.files[0];if(!v)return;let M=document.getElementById(`label-import-${y.moduleName}`),N=M?M.querySelector(".import-text"):null,H=N?N.innerText:"";N&&(N.innerText="\u231B Memproses..."),M&&(M.style.pointerEvents="none"),b.disabled=!0;try{let I=await Je(v);if(I.length===0)throw new Error("File kosong atau format salah");await y.onImport(I),W("Import berhasil!"),oe(a),R()}catch(I){V(I.message||"Gagal import data")}finally{N&&(N.innerText=H),M&&(M.style.pointerEvents="auto"),b.disabled=!1,b.value=""}})}async function R(){C();let b=document.getElementById("table-container");if(!b)return;b.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let k=D==="client",v=k?1:w,M=k?$e:20,N=new URLSearchParams({page:v,limit:M,...Object.fromEntries(Object.entries($).filter(([,q])=>q))}),H=await x(`${a}?${N}`);if(!H.ok){b.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${H.data?.error||"Error"}</p></div>`;return}let I=H.data?.data||H.data||[],L=H.data?.pagination,ne=I.length;if(k){I=f(I);let q=I.length,Q=20,te=Math.ceil(q/Q);w>te&&te>0&&(w=te);let O=(w-1)*Q,le=w*Q;I=I.slice(O,le),L={page:w,limit:Q,total:q,pages:te}}!1,c&&c(I);let ve=Ye({columns:r,data:I,onEdit:u?q=>ye(q):null,actions:h.map(q=>({...q,handler:Q=>q.handler(Q,R)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:S?{selectedIds:_,onToggle:C}:null});b.innerHTML="",b.appendChild(ve);let se=document.getElementById("pagination-container");if(se&&(se.innerHTML="",L&&L.pages>1)){let q=We({page:L.page,pages:L.pages,total:L.total,limit:L.limit,onPage:Q=>{w=Q,R()}});q&&se.appendChild(q)}}function Te(b){let k=typeof s=="function"?s(b):s;return He(k)}function ye(b){let k=!!b,v=document.createElement("form");if(v.noValidate=!0,v.innerHTML=Te(b),k){let N=typeof s=="function"?s(b):s;Ze(v,b)}let{close:M}=de({title:k?`Edit ${n}`:`Tambah ${n}`,content:v,size:"lg",confirmText:k?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(N,H)=>{if(!v.reportValidity())return;let I=N.querySelector(".modal-confirm");I.disabled=!0,I.textContent="Menyimpan...";let L=Xe(v),ne=typeof s=="function"?s(b):s,ve=async te=>{for(let O of te)if(O.type==="row")await ve(O.fields);else if(O.type==="combobox"&&L[O.name]){let le=L[O.name],ke=(O.options||[]).find(Y=>{let ae=String(typeof Y=="object"?Y.value:Y),dt=String(typeof Y=="object"?Y.label:Y);return ae===le||dt===le});if(ke)L[O.name]=typeof ke=="object"?ke.value:ke;else if(O.createApi){let Y={};Y[O.createApi.field]=le,O.createApi.extra&&Object.assign(Y,O.createApi.extra);let ae=await x(O.createApi.path,{method:"POST",body:JSON.stringify(Y)});if(ae.ok&&ae.data?.id)L[O.name]=ae.data.id;else if(ae.ok&&!ae.data?.id)L[O.name]=le;else throw new Error(`Gagal membuat master data: ${ae.data?.error||"Unknown error"}`)}}};try{await ve(ne)}catch(te){V(te.message),I.disabled=!1,I.textContent=k?"Simpan Perubahan":`Tambah ${n}`;return}p&&(L=await p(L,b));let se=k?"PUT":"POST",q=k?`${a}/${b.id}`:a,Q=await x(q,{method:se,body:JSON.stringify(L)});Q.ok?(W(k?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),H(),oe(a),R()):(V(Q.data?.error||"Gagal menyimpan data."),I.disabled=!1,I.textContent=k?"Simpan Perubahan":`Tambah ${n}`)}})}function ga(b){qe(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let k=await x(`${a}/${b.id}`,{method:"DELETE"});k.ok?(W(`${n} berhasil dihapus.`),oe(a),R()):V(k.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return R(),R}A();A();var De=null,tt=null;async function Ce(t=!1){if(De&&!t)return console.log("Employees Raw (Cache Hit)",De.slice(0,5)),De;let e=await x(`/api/employees?limit=${$e}&status=Aktif`);return De=(e.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",De.slice(0,5)),De}async function Z(t=!1){let i=(await Ce(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",i.slice(0,5)),i}async function K(t=!1){return tt&&!t||(tt=((await x("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),tt}function U(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function bt(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function Ie(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function ht(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function ce(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}j();A();j();A();j();A();j();function at(t){if(t.target.name==="report_date"||t.target.name==="completion_date"){let e=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),a=document.querySelector('input[name="day_count"]');if(e&&i&&a)if(e.value&&i.value){let r=new Date(e.value),s=new Date(i.value),l=Math.floor((s-r)/864e5);a.value=isNaN(l)?"":l}else a.value=""}}document.body.removeEventListener("input",at);document.body.addEventListener("input",at);document.body.removeEventListener("change",at);document.body.addEventListener("change",at);A();var we={};function Pe(t){if(we[t]){try{we[t].destroy()}catch{}delete we[t]}}function ka(){Object.keys(we).forEach(Pe)}var he=(t,e=0)=>{let i=Number(t);return isNaN(i)||t===null||t===void 0?e:i},Be=(t,e="\u2014")=>{if(t==null||t==="")return e;let i=String(t).trim();return i===""||i==="[object Object]"?e:i};function It(t,e,i=900){if(!t)return;let a=Math.max(0,Math.round(he(e)));if(a===0){t.textContent="0";return}let r=Date.now(),s=()=>{let l=Math.min((Date.now()-r)/i,1),o=1-Math.pow(1-l,3);t.textContent=Math.round(o*a).toLocaleString("id-ID"),l<1?requestAnimationFrame(s):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(s)}var Sa={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},wa=t=>{let e=Be(t,"\u2014");return`<span class="status-pill ${Sa[e]||"pill-neutral"}">${e}</span>`};var me={family:"Inter",size:11},fe="#94A3B8",Le="#F1F5F9",ft=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],xa=()=>window.innerWidth<768;function nt(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:xa()?"bottom":"top",labels:{font:me,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:me,titleFont:{...me,weight:"700"}}},scales:{x:{grid:{color:Le},ticks:{font:me,color:fe,maxRotation:0}},y:{grid:{color:Le},ticks:{font:me,color:fe},beginAtZero:!0}},...t}}var _a=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),Ca=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Et(t=3){return Array(t).fill(0).map((e,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function pe(t,e,i=8e3){try{let a=new AbortController,r=setTimeout(()=>a.abort(),i),s=await x(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(r),!s||!s.ok)return e;let l=s.data;return l?l.data!==void 0?l.data??e:l:e}catch{return e}}function $a(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(a=>{let r=document.getElementById(a);r&&(r.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(a=>{let r=document.getElementById(a);if(r&&r.style.display==="none"){r.style.display="block";let s=r.parentElement;if(s&&!s.querySelector(".chart-empty")){let l=document.createElement("div");l.className="chart-empty",l.textContent="Belum ada data",r.style.display="none",s.appendChild(l)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Lt({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Pt({}),["table-contracts","table-issues"].forEach(a=>{let r=document.getElementById(a);r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Bt(t){ka(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">
      
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h2 style="margin:0; font-size:1.2rem; color:var(--text-1);">Dashboard Overview</h2>
        <button id="btn-dash-refresh" class="btn-primary" style="padding:6px 12px; font-size:0.8rem; display:flex; align-items:center; gap:6px; cursor:pointer; background:var(--primary); color:#fff; border:none; border-radius:6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
          Refresh Data
        </button>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${_a()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${Ca()}</div>

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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${Et(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${Et(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>Dt(t)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async i=>{let a=i.target.value,r=document.getElementById("jadwal-year-label");r&&(r.textContent=a);let s=document.getElementById("skel-jadwal"),l=document.getElementById("chart-jadwal");s&&(s.style.display="block",s.style.position="absolute"),l&&(l.style.display="none");let o=await pe(`/api/dashboard/schedule-chart?year=${a}`,{},8e3);try{Ft(o)}catch(n){console.warn("ScheduleChart render:",n),ue("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let a=i.target.value,r=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",s=document.getElementById("skel-insp"),l=document.getElementById("chart-insp");s&&(s.style.display="block",s.style.position="absolute"),l&&(l.style.display="none");let o=await pe(r,{},8e3);try{At(o)}catch(n){console.warn("InspBar render:",n),ue("skel-insp","chart-insp")}}),t._skelTimeout=setTimeout(()=>$a(),5e3),await Dt(t)}async function Dt(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,i,a,r,s,l,o,n]=await Promise.all([pe("/api/dashboard/kpi",{},8e3),pe("/api/dashboard/issues-trend",{},8e3),pe("/api/dashboard/issues-summary",{},8e3),pe("/api/dashboard/stats",{},8e3),pe("/api/dashboard/calendar",[],8e3),pe("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),pe(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),pe("/api/dashboard/dropdown-aggregates",{},8e3)]),d=document.getElementById("filter-insp-month"),u=d?d.value:"",g=u?`/api/dashboard/inspection-bar?month=${u}`:"/api/dashboard/inspection-bar",p=await pe(g,{},8e3);window.dashboardAggregates=n||{};try{Lt(e)}catch(c){console.warn("KPI render:",c)}try{Pt(e)}catch(c){console.warn("MiniStats render:",c)}try{Ft(o)}catch(c){console.warn("ScheduleChart render:",c),ue("skel-jadwal","chart-jadwal")}try{Ta(Array.isArray(a?.by_category)?a.by_category:[])}catch(c){console.warn("Donut render:",c),ue("skel-donut","chart-donut")}try{Ea(i)}catch(c){console.warn("Trend render:",c),ue("skel-trend","chart-trend")}try{At(p)}catch(c){console.warn("InspBar render:",c),ue("skel-insp","chart-insp")}try{let c=Array.isArray(r)?r:Array.isArray(r?.recent_issues)?r.recent_issues:[];Ia(c)}catch(c){console.warn("IssuesTable render:",c)}try{let c=Array.isArray(r?.expiring_contracts)?r.expiring_contracts:[];Da(l)}catch(c){console.warn("ContractsTable render:",c)}try{Ba(Array.isArray(s)?s:[])}catch(c){console.warn("Agenda render:",c)}try{La()}catch(c){console.warn("Quick Actions render:",c)}}function Lt(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let i=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=i.map(a=>{let r=he(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${r}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${r}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{It(a,parseInt(a.dataset.target)||0)})}function Pt(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let i=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,a=new Date().getFullYear(),r=String(new Date().getMonth()+1).padStart(2,"0"),s=`${a}-${r}`,l=u=>`
    <select id="${u}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.preventDefault(); event.stopPropagation();">
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
  `,o=[{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.preventDefault(); event.stopPropagation();">
          <option value="Q1" ${i==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${i==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${i==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${i==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:t.schedule?.current,href:`#/timeline?dash_filter=period_${i.toLowerCase()}`,color:"mini-blue"},{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:l("dash-inspeksi-month"),val:t.inspection_month?.current,href:`#/timeline?dash_filter=inspeksi&month=${s}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:l("dash-gcdc-month"),val:t.cleaning_month?.current,href:`#/timeline?dash_filter=gcdc&month=${s}`,color:"mini-green"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:l("dash-reliefer-month"),val:t.reliever_completed?.current,href:`#/relievers?dash_filter=reliever&month=${s}`,color:"mini-teal"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:l("dash-fogging-month"),val:t.fogging_month?.current,href:`#/reports/fogging?dash_filter=fogging&month=${s}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=o.map(u=>`
    <a href="${u.href}" class="mini-stat ${u.color}" style="text-decoration:none" id="${u.id||""}">
      <div class="mini-stat-icon">${u.icon}</div>
      <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
        <div style="display:flex; align-items:baseline; gap:3px;">
          <div class="mini-stat-value" data-target="${he(u.val)}">0</div>
          ${u.dropdown?u.dropdown:""}
        </div>
        <div class="mini-stat-text">${u.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(u=>It(u,parseInt(u.dataset.target)||0,700));let n=document.getElementById("dash-jadwal-period");if(n){let u=g=>{let c=(window.dashboardAggregates?.schedule_by_quarter||{})[g]||0,f=document.querySelector("#mini-jadwal .mini-stat-value");f&&(f.dataset.target=c,f.textContent=c);let h=document.getElementById("mini-jadwal");h&&(h.href=`#/timeline?dash_filter=period_${g.toLowerCase()}`)};u(n.value),n.addEventListener("change",g=>u(g.target.value))}let d=(u,g,p,c)=>{let f=document.getElementById(u);if(f){let h=m=>{let S=(window.dashboardAggregates?.[p]||{})[m]||0,D=document.querySelector(`#${g} .mini-stat-value`);D&&(D.dataset.target=S,D.textContent=S);let T=document.getElementById(g);T&&(T.href=`${c}&month=${m}`)};h(f.value),f.addEventListener("change",m=>h(m.target.value))}};d("dash-reliefer-month","mini-reliefer","relievers_by_month","#/relievers?dash_filter=reliever"),d("dash-inspeksi-month","mini-inspeksi","inspeksi_by_month","#/timeline?dash_filter=inspeksi"),d("dash-gcdc-month","mini-gcdc","gcdc_by_month","#/timeline?dash_filter=gcdc"),d("dash-fogging-month","mini-fogging","fogging_by_month","#/reports/fogging?dash_filter=fogging")}function Ta(t){ue("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!e||!i)return;Pe("donut");let a=(t||[]).filter(n=>he(n.count)>0);if(!a.length){Ue(e,"Belum ada data permasalahan");return}let r=a.map(n=>`${Be(n.category,"Lainnya")}`),s=a.map(n=>he(n.count)),l=s.reduce((n,d)=>n+d,0);i.innerHTML=a.map((n,d)=>{let u=ft[d%ft.length],g=l>0?Math.round(n.count/l*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${u}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${g}%)</span></div>
          <div class="donut-legend-label">${r[d]}</div>
        </div>
      </div>
    `}).join("");let o={id:"centerText",beforeDraw:function(n){let d=n.width,u=n.height,g=n.ctx;g.restore();let p=(u/80).toFixed(2);g.font="bold "+p+"em Inter",g.textBaseline="middle",g.fillStyle="#1E293B";let c=l.toString(),f=Math.round((d-g.measureText(c).width)/2),h=u/2;g.fillText(c,f,h-4),g.font="600 "+(p*.35).toFixed(2)+"em Inter",g.fillStyle="#64748B";let m="Total",y=Math.round((d-g.measureText(m).width)/2);g.fillText(m,y,h+10),g.save()}};we.donut=new Chart(e,{type:"doughnut",data:{labels:r,datasets:[{data:s,backgroundColor:ft,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:me,titleFont:{...me,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutout:"75%"},plugins:[o]})}function Ea(t){ue("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;Pe("trend"),t=t||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(t.labels||[]).map(l=>{if(!l||typeof l!="string")return"";try{let[o,n]=l.split("-");return(i[Number(n)-1]||n)+" "+String(o).slice(-2)}catch{return l}}),r=(t.open||[]).map(l=>he(l)),s=(t.closed||[]).map(l=>he(l));if(!a.length){Ue(e,"Belum ada data trend");return}we.trend=new Chart(e,{type:"line",data:{labels:a,datasets:[{label:"Open",data:r,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:s,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:nt({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:fe,maxRotation:45,autoSkip:!0}},y:{grid:{color:Le},ticks:{font:{family:"Inter",size:9},color:fe},beginAtZero:!0}}})})}function Ft(t){ue("skel-jadwal","chart-jadwal");let e=document.getElementById("chart-jadwal");if(!e)return;Pe("jadwal"),t=t||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(t).some(n=>Array.isArray(n)&&n.some(d=>d>0))){Ue(e,"Belum ada data jadwal");return}let r=t["Inspeksi Hygiene"]||Array(12).fill(0),s=t["General Cleaning"]||Array(12).fill(0),l=t["Deep Cleaning"]||Array(12).fill(0),o=t.Fogging||Array(12).fill(0);we.jadwal=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Inspeksi",data:r,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:s,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:l,backgroundColor:"#F59E0B"},{label:"Fogging",data:o,backgroundColor:"#EF4444"}]},options:nt({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:fe,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Le},ticks:{font:{family:"Inter",size:9},color:fe},min:0}}})})}function At(t){ue("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;Pe("inspBar"),t=t||{};let i=t.labels||[],a=(t.fc||[]).map(s=>he(s)),r=(t.spv||[]).map(s=>he(s));if(!i.length){Ue(e,"Belum ada data inspeksi");return}we.inspBar=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:r,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:nt({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:me,color:fe,maxRotation:45,minRotation:45}},y:{grid:{color:Le},ticks:{font:me,color:fe},min:0,max:100}}})})}function Da(t){ue("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;Pe("contractMiniBar"),t=t||{};let i={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(l=>{let o=l.split("-")[1];return i[o]||l}),r=(t.data||[]).map(l=>he(l));if(!a.length){Ue(e,"Belum ada data");return}let s=e.getContext("2d");we.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:r,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:nt({onClick:(l,o)=>{if(o&&o.length>0){let n=o[0].index,d=(t.labels||[])[n];d&&(window.location.hash="#/contracts?month_expiry="+d)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:me,color:fe,maxRotation:0,autoSkip:!1}},y:{grid:{color:Le,borderDash:[4,4],drawBorder:!1},ticks:{font:me,color:fe,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function Ia(t){let e=document.getElementById("table-issues");if(!e)return;let i=(t||[]).slice(0,8);if(!i.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${i.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${wa(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Be(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Be(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function Ba(t){let e=document.getElementById("widget-agenda");if(!e)return;let i=new Date,a=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,s=(t||[]).filter(l=>(l.event_date||"").startsWith(a)).slice(0,10);if(!s.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${s.map(l=>{let o="#3B82F6",n="#EFF6FF",d="Agenda",u=(l.title||"").toLowerCase();return u.includes("inspeksi")?(o="#10B981",n="#ECFDF5",d="Inspeksi"):u.includes("cleaning")||u.includes("gcdc")?(o="#3B82F6",n="#EFF6FF",d="Cleaning"):u.includes("reliefer")?(o="#F59E0B",n="#FFFBEB",d="Reliefer"):u.includes("fogging")&&(o="#8B5CF6",n="#F5F3FF",d="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(l.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${o};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Be(l.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Be(l.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${o}">${d}</div>
        </div>
      `}).join("")}
    </div>
  `}function La(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function ue(t,e){let i=document.getElementById(t),a=document.getElementById(e);if(i&&(i.style.display="none",i.style.position=""),a){a.style.display="block";let r=a.parentElement;if(r){let s=r.querySelector(".chart-empty");s&&s.remove()}}}function Ue(t,e="Belum ada data"){if(!t)return;t.style.display="none";let i=t.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent=e,i.appendChild(r)}}A();async function Nt(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),i=document.getElementById("login-error"),a=document.getElementById("login-btn"),r=document.getElementById("toggle-password"),s=document.getElementById("login-password");r?.addEventListener("click",()=>{let l=s.type==="text";s.type=l?"password":"text",r.style.color=l?"":"var(--primary)"}),e?.addEventListener("submit",async l=>{l.preventDefault(),i.style.display="none";let o=e.username.value.trim(),n=e.password.value;if(!o||!n){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let d=await x("/api/auth/login",{method:"POST",body:JSON.stringify({username:o,password:n})});d.ok&&d.data.success?(mt(d.data.data.token),Re(d.data.data.user),W("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=d.data.error||"Username atau password salah.",i.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}A();j();async function Pa(){return await K()}function Fa(t,e){let i=String(t.status||"").toLowerCase();return e==="active"?i==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&i==="aktif":!1}async function Mt(t,e){let i=await Pa(),a=e?e.get("dash_filter"):null;P({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",enableMobileFilterSheet:!0,itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:r=>a?r.filter(s=>Fa(s,a)):r,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:r=>Ie(r)},{key:"phone",label:"No. HP",render:r=>r?`<a href="tel:${r}">${r}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:r=>window.formatDate(r)},{key:"target_pindah_os",label:"Target Pindah OS",render:r=>window.formatDate(r)},{key:"target_selesai_os",label:"Target Selesai OS",render:r=>window.formatDate(r)},{key:"status",label:"Status",render:r=>U(r)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:r=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:r?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:r?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:i,value:r?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:r?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:r?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:r?.status||""}]},{type:"row",fields:[{name:"target_pindah_os",label:"Target Pindah OS",type:"date",value:r?.target_pindah_os},{name:"target_selesai_os",label:"Target Selesai OS",type:"date",value:r?.target_selesai_os}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let r=await x(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let s=r.data.data.map(l=>({"Nama Lengkap":l.full_name,Cabang:l.branch_name||"",Divisi:l.division||"","No. HP":l.phone||"","Tgl Masuk":l.join_date||"",Status:l.status||"","Target Pindah OS":l.target_pindah_os||"","Target Selesai OS":l.target_selesai_os||""}));F(s,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{F([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif","Target Pindah OS":"","Target Selesai OS":""},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif","Target Pindah OS":"","Target Selesai OS":""}],"Template_Import_Karyawan")},onImport:async r=>{let s=n=>{if(!n)return null;let d=String(n||"").toLowerCase(),u=i.find(g=>String(g.label||"").toLowerCase()===d);return u?u.value:null},l=r.map(n=>({full_name:String(n["Nama Lengkap"]||"").trim(),branch_id:s(String(n.Cabang||"").trim()),division:String(n.Divisi||n["Div / Bagian"]||"").trim()||"FACILITY CARE",phone:String(n["No. HP"]||n["No. Hp"]||"").trim(),join_date:String(n["Tgl Masuk"]||n["Tanggal Masuk"]||"").trim(),status:String(n.Status||"").trim(),target_pindah_os:String(n["Target Pindah OS"]||"").trim(),target_selesai_os:String(n["Target Selesai OS"]||"").trim(),notes:String(n.Catatan||"").trim()})).filter(n=>n.full_name),o=await x("/api/import/employees",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}}})}A();j();var vt=[],Ot=[];async function Aa(){vt=await K(),Ot=await Ce()}var yt=async t=>{let e=[],i=1;for(;;){let r=await(await Promise.resolve().then(()=>(A(),Ee))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${i}`);if(!r.ok)break;let s=r.data?.data||r.data||[],l=Array.isArray(s)?s:[];if(e=e.concat(l),l.length<100||r.data?.pagination&&i>=r.data.pagination.pages)break;i++}return e};async function it(t,e){await Aa(),P({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",defaultFilters:{},onDataLoaded:a=>a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>Ie(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,r)=>r.end_date&&String(r.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':bt(a)},{key:"status",label:"Status",render:a=>U(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:vt},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[s,l]=await Promise.all([yt("/api/employees?status=Aktif"),yt("/api/contracts")]);if(s.length>0){let o=l.filter(g=>g.status==="Aktif"),n=new Set(o.map(g=>g.employee_id)),d=s.filter(g=>!n.has(g.id)),u=`<p style="margin-bottom:12px">Data yang terbaca: <b>${s.length}</b> Karyawan Aktif, dan <b>${o.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${d.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;d.forEach(g=>{let p=l.filter(f=>f.employee_id===g.id),c='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(p.length>0){let f=p[0];c=`Pernah ada kontrak (Status: <b style="color:#EF4444">${f.status}</b>, Selesai: ${window.formatDate(f.end_date)})`}u+=`<li style="margin-bottom:8px"><b>${g.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${g.branch_name||"-"} | ${c}</span></li>`}),u+="</ul>",Promise.resolve().then(()=>(_e(),gt)).then(g=>g.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:u,cancelText:"Tutup"}))}}catch(s){console.error(s)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let r=document.querySelector(".page-actions");r&&r.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:Ot,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:vt,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await x(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let r=a.data.data.map(s=>({"Nama Lengkap":s.employee_name,Cabang:s.branch_name||"","Div / Bagian":s.division||"","Tanggal Mulai":s.start_date||"","Tanggal Selesai":s.end_date&&String(s.end_date).startsWith("2099")?"":s.end_date||"","Sisa Kontrak":s.end_date&&String(s.end_date).startsWith("2099")?"Tetap":s.days_remaining!==null&&s.days_remaining!==void 0?`${s.days_remaining} Hari`:"",Status:s.status||""}));F(r,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{F([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[r,s]=await Promise.all([x("/api/branches?limit=10000"),yt("/api/employees")]),l=r.data?.data||[],o=s||[];console.log(`Total employee yang berhasil dimuat dari database : ${o.length}`),o.length>0&&(console.log("Contoh 5 employee pertama:"),o.slice(0,5).forEach((m,y)=>{console.log(`${y+1}. ID: ${m.id}, Name: ${m.full_name}, Status: ${m.status}`)}));let n=m=>{if(!m)return null;let y=String(m||"").replace(/\s+/g," ").toLowerCase().trim(),S=l.find(D=>String(D.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===y||String(D.code||"").replace(/\s+/g," ").toLowerCase().trim()===y||String(D.name||"").replace(/\s+/g," ").toLowerCase().trim()===y);return S?S.id:null},d=(m,y)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${y}`),console.log(`Nama dari Excel : "${m}"`),!m)return console.log("Alasan gagal mapping : Nama kosong"),null;let S=String(m||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${S}"`),console.log(`Jumlah employee di database : ${o.length}`);let D=o.find(T=>String(T.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===S);return D?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${D.id}`),D.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},u=m=>{if(!m)return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let y=String(m).trim();if(/^\d{4,5}(\.\d+)?$/.test(y)){let D=Math.floor(Number(y));if(D>2e4&&D<99999){let T=new Date(Date.UTC(1899,11,30)+D*864e5);return isNaN(T.getTime())?"":T.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);let S=y.split(/[\/\-\.]/);if(S.length===3){let[D,T,w]=S.map($=>$.trim());if(D.length===4&&T.length<=2&&w.length<=2)return`${D}-${T.padStart(2,"0")}-${w.padStart(2,"0")}`;if(w.length===4&&T.length<=2&&D.length<=2)return`${w}-${T.padStart(2,"0")}-${D.padStart(2,"0")}`}return y},g=a.map((m,y)=>{let S=y+2,D=String(m["Nama Lengkap"]||"").trim(),T=m["Tanggal Mulai"],w=u(T);if(!w){let C=a.__worksheet,B=a.__headers||[],E=B.indexOf("Tanggal Mulai"),z="N/A",ee="N/A",be="N/A";if(E!==-1&&C&&window.XLSX){let Te=window.XLSX.utils.encode_cell({c:E,r:S-1});be=Te;let ye=C[Te];ye?(z=ye.t||"undefined",ee=ye.w||"undefined"):z="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let R="Unknown";T==null||T===""?R="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":T instanceof Date&&isNaN(T.getTime())?R="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":R="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${S}`),console.log(`Employee Name : ${D}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${E})`),console.log(`Raw Cell Value : "${T}"`),console.log(`JavaScript Type : ${typeof T}`),console.log(`SheetJS Cell Type : ${z}`),console.log(`SheetJS Formatted Value : "${ee}"`),console.log(`Value After Trim : "${String(T||"").trim()}"`),console.log(`Value After Date Parser : "${w}"`),console.log(`Is Empty : ${!T}`),console.log(`Is Invalid Date : ${T instanceof Date?isNaN(T.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${R}`),console.log(`Workbook Sheet : ${C?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${be}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(m,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(B)),console.log(`==========================
`)}let $=d(D,S),_=null;return $?w||(_="Tanggal Mulai kosong atau tidak berformat tanggal"):_="Karyawan tidak ditemukan di Database",{isValid:!!($&&w),invalidReason:_,rowNum:S,data:{employee_id:$,branch_id:n(String(m.Cabang||"").trim()),division:String(m["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:w,end_date:u(m["Tanggal Selesai"])||"2099-12-31",status:String(m.Status||"").trim(),_rawName:D}}}),p=[],c=[];if(g.forEach(m=>{m.isValid?p.push(m.data):c.push({rowNum:m.rowNum,name:m.data._rawName,reason:m.invalidReason})}),console.log(`Split Validation - Valid: ${p.length}, Invalid: ${c.length}`),p.length===0){let m=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${a.length}
Valid: 0
Invalid: ${c.length}

Daftar Kegagalan (Contoh):
`;c.slice(0,10).forEach(y=>{m+=`- Row ${y.rowNum} | Nama: ${y.name} | Alasan: ${y.reason}
`}),c.length>10&&(m+=`- ... dan ${c.length-10} lainnya.
`),alert(m);return}let f=await x("/api/contracts/import",{method:"POST",body:JSON.stringify(p)}),h=`IMPORT SUMMARY
======================
`;h+=`Total Baris Excel : ${a.length}
`,h+=`Baris Valid       : ${p.length}
`,h+=`Baris Invalid     : ${c.length}

`,f&&f.data&&f.data.metrics?(h+=`Berhasil INSERT   : ${f.data.metrics.inserted}
`,h+=`Berhasil UPDATE   : ${f.data.metrics.updated}
`):h+=`Berhasil diproses : ${p.length}
`,c.length>0&&(h+=`
DAFTAR DATA DILEWATI:
`,c.forEach(m=>{h+=`- Row ${m.rowNum} | ${m.name} | ${m.reason}
`})),alert(h),typeof it=="function"&&it()}}})}A();j();var kt=[],Ge=[];function Na(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let i of e)if(t.some(a=>a.period===i))return i;return"Q3"}async function Kt(t,e){kt=await K();let i=await Z();Ge=["Berlin Ariansyah","Ade Surahman"];let a=m=>m&&!Ge.find(y=>String(typeof y=="object"?y.value:y).toLowerCase()===String(m).toLowerCase())?[...Ge,m]:Ge,r=await x(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),s=m=>{if(!m||m==="-"||String(m).trim()==="")return"";let y=String(m).split("-");return y.length===3&&y[0].length===4?`${y[2]}-${y[1]}-${y[0]}`:m},l=r.data?.data||[],o=Na(l),n=e?e.get("dash_filter"):null,d=new Date,u=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`,g={},p=e&&e.get("month")?e.get("month"):null;n==="inspeksi"?g={status:"Done",activity_type:"Inspeksi Hygiene",month:p}:n==="gcdc"?g={status:"Done",activity_type:"GCDC",month:p}:n&&n.startsWith("period_")&&(g={period:n.replace("period_","").toUpperCase()});let c=new Date().getFullYear(),h=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((m,y)=>{let S=String(y+1).padStart(2,"0");return{value:`${c}-${S}`,label:`${m} ${c}`}});P({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:g,onDataLoaded:m=>m.sort((y,S)=>{let D=y.opening_date?new Date(y.opening_date).getTime():0;return(S.opening_date?new Date(S.opening_date).getTime():0)-D}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:m=>ht(m)},{key:"period",label:"Periode",render:m=>ce(m)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:m=>s(m)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:m=>s(m)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:m=>s(m)},{key:"status",label:"Status",render:m=>U(m)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:kt},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:h},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Ge}],formFields:m=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:kt,value:m?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:m?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:m?.period},{name:"pic",label:"PIC",type:"combobox",options:a(m?.pic),value:m?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:m?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:m?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:m?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:m?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:m?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let m=await x(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(m.ok){let y=m.data.data.map(S=>({Cabang:S.branch_name||"",Kegiatan:S.activity_type||"",Periode:S.period||"",PIC:S.pic||"","Tgl Opening":S.opening_date||"","Tgl Target":S.target_date||"","Tgl Selesai":S.completion_date||"",Status:S.status||""}));F(y,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{F([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async m=>{let S=(await x("/api/branches?all=1")).data?.data||[],D=_=>{if(!_)return null;let C=String(_||"").toLowerCase().trim(),B=C.match(/^(\d{3})/);if(B){let z=S.find(ee=>String(ee.code||"")===B[1]);if(z)return z.id}let E=S.find(z=>String(z.full_name||"").toLowerCase()===C||String(z.code||"").toLowerCase()===C||String(z.name||"").toLowerCase()===C);return E?E.id:null},T=_=>{if(_==null||_==="")return"";if(_ instanceof Date&&!isNaN(_.getTime()))return _.toISOString().slice(0,10);let C=String(_).trim();if(C===""||C==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(C))return C.slice(0,10);if(/^\d{4,5}$/.test(C)){let E=Number(C);if(E>2e4&&E<99999){let z=new Date(Date.UTC(1899,11,30)+E*864e5);return isNaN(z.getTime())?"":z.toISOString().slice(0,10)}}let B=C.split(/[\/\-\.]/);if(B.length===3){let[E,z,ee]=B.map(be=>be.trim());if(E.length===4&&z.length<=2&&ee.length<=2)return`${E}-${z.padStart(2,"0")}-${ee.padStart(2,"0")}`;if(ee.length===4&&z.length<=2&&E.length<=2)return`${ee}-${z.padStart(2,"0")}-${E.padStart(2,"0")}`}return C},w=m.map(_=>({branch_id:D(String(_.Cabang||"").trim()),branch_name:String(_.Cabang||"").trim(),activity_type:String(_.Kegiatan||"").trim(),period:String(_.Periode||"").trim(),pic:String(_.PIC||_.Pic||"").trim(),opening_date:T(_["Tgl Opening"]||_["Tanggal Opening"]||_["Tgl Openir"]),target_date:T(_["Tgl Target"]||_["Tanggal Target"]),completion_date:T(_["Tgl Selesai"]||_["Tanggal Selesai"]),status:String(_.Status||"").trim(),notes:String(_.Catatan||_.Keterangan||"").trim()})).filter(_=>_.activity_type&&_.period),$=await x("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:w,onDuplicate:"update"})});if(!$.ok)throw new Error($.data?.error||"Import gagal");return oe("schedule"),$.data}}})}A();j();var St=[],rt=[];function lt(t){if(t.target.name==="report_date"||t.target.name==="completion_date"){let e=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),a=document.querySelector('input[name="day_count"]');if(e&&i&&a)if(e.value&&i.value){let r=new Date(e.value),s=new Date(i.value),l=Math.floor((s-r)/864e5);a.value=isNaN(l)?"":l}else a.value=""}}document.body.removeEventListener("input",lt);document.body.addEventListener("input",lt);document.body.removeEventListener("change",lt);document.body.addEventListener("change",lt);function Ma(t,e){let i=String(t.status||"").toLowerCase();return e==="open"?i==="open":!1}async function Rt(t,e){let i=e?e.get("dash_filter"):null;St=await K(),rt=await Z();let a=n=>n&&!rt.find(d=>d.value===n)?[...rt,{value:n,label:n}]:rt,r=new Date().getFullYear(),s=["2025","2026","2027","2028","2029","2030"],o=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((n,d)=>{let u=String(d+1).padStart(2,"0");return{value:`${r}-${u}`,label:`${n} ${r}`}});P({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:n=>i?n.filter(d=>Ma(d,i)):n,columns:[{key:"report_date",label:"Tanggal Info",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:n=>`<span class="badge badge-secondary">${n}</span>`},{key:"source",label:"Sumber Laporan"},{key:"complaint",label:"Keluhan",render:n=>`<span title="${n}">${n?.length>50?n.slice(0,50)+"...":n}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:n=>`<span title="${n||""}">${n?.length>40?n.slice(0,40)+"...":n||"-"}</span>`},{key:"status",label:"Status",render:n=>U(n)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"day_count",label:"Day",render:n=>n??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:St},{type:"select",name:"month",label:"Bulan",options:o},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:s}],formFields:n=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:n?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:St,value:n?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:n?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:n?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:n?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(n?.employee_name),value:n?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(n?.fc_specialist),value:n?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:n?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:n?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date},{name:"day_count",label:"Day",type:"number",value:n?.day_count,readonly:!0}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let n=await x(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let d=n.data.data.map(u=>({"Tanggal Info":u.report_date||"",Cabang:u.branch_name||"",Kategori:u.category||"","Sumber Laporan":u.source||"",Keluhan:u.complaint||"","Nama FC":u.employee_name||"","FC Spesialis":u.fc_specialist||"",Solusi:u.solution||"",Status:u.status||"","Tanggal Selesai":u.completion_date||"",Day:u.day_count!==null?u.day_count:""}));F(d,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{F([{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan",Status:"Done","Tanggal Selesai":"2024-03-02",Day:1}],"Template_Import_Permasalahan")},onImport:async n=>{let u=(await x("/api/branches?all=1")).data?.data||[],g=f=>{if(!f)return null;let h=String(f||"").toLowerCase(),m=u.find(y=>String(y.full_name||"").toLowerCase()===h||String(y.code||"").toLowerCase()===h||String(y.name||"").toLowerCase()===h);return m?m.id:null},p=n.map(f=>({branch_id:g(String(f.Cabang||"").trim()),report_date:String(f["Tanggal Info"]||f.Tanggal||"").trim(),category:String(f.Kategori||"").trim(),source:String(f["Sumber Laporan"]||f.Sumber||"").trim(),complaint:String(f.Keluhan||"").trim(),employee_name:String(f["Nama FC"]||"").trim(),fc_specialist:String(f["FC Spesialis"]||"").trim(),solution:String(f.Solusi||"").trim(),completion_date:String(f["Tanggal Selesai"]||f["Tgl Selesai"]||"").trim(),status:String(f.Status||"").trim(),day_count:f.Day||f.Hari||null})).filter(f=>f.report_date&&f.complaint&&f.category),c=await x("/api/import/issues",{method:"POST",body:JSON.stringify({rows:p,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}}})}A();var Fe=[];function Oa(t,e){let i=String(t.status||"").toLowerCase();return e==="pending"?i==="pending":!1}async function jt(t,e){let i=e?e.get("dash_filter"):null;Fe=await K();let a=await Z(),r=["Ade","Berlin"],s=o=>o&&!a.find(n=>n.value===o)?[...a,{value:o,label:o}]:a,l=o=>o&&!r.find(n=>(typeof n=="object"?n.value:n)===o)?[...r,o]:r;P({container:t,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:o=>i?o.filter(n=>Oa(n,i)):o,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:o=>window.formatDate(o)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:o=>`<span title="${o||""}">${o?.length>50?o.slice(0,50)+"\u2026":o||"-"}</span>`},{key:"solution",label:"Solusi",render:o=>`<span title="${o||""}">${o?.length>40?o.slice(0,40)+"\u2026":o||"-"}</span>`},{key:"status",label:"Status",render:o=>U(o)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:o=>window.formatDate(o)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:o=>o?`<a href="${o}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:Fe},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async o=>{let n=new URLSearchParams(o||{}).toString(),d=await x(`/api/one-on-one?limit=10000&${n}`);if(d.ok){let u=d.data.data.map(p=>({Tanggal:p.meeting_date||"",Cabang:p.branch_name||"","Nama Karyawan":p.employee_name||"",PIC:p.pic||"",Masalah:p.problem||"",Solusi:p.solution||"",Status:p.status||"","Tgl Selesai":p.completion_date||"",Dokumen:p.document_link||""})),{downloadExcel:g}=await Promise.resolve().then(()=>(j(),re));g(u,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let o=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(j(),re));n(o,"Template_Import_OneOnOne")},onImport:async o=>{let n=p=>{if(!p)return null;let c=String(p||"").toLowerCase(),f=Fe.find(h=>String(h.label||"").toLowerCase()===c);return f?f.value:null},d=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let c=String(p).trim();if(/^\d{4,5}$/.test(c)){let h=Number(c);if(h>2e4&&h<99999){let m=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);let f=c.split(/[\/\-\.]/);if(f.length===3){let[h,m,y]=f.map(S=>S.trim());if(h.length===4&&m.length<=2&&y.length<=2)return`${h}-${m.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&m.length<=2&&h.length<=2)return`${y}-${m.padStart(2,"0")}-${h.padStart(2,"0")}`}return c},u=o.map(p=>({meeting_date:d(p.Tanggal),employee_name:String(p["Nama Karyawan"]||"").trim(),branch_id:n(String(p.Cabang||"").trim()),pic:String(p.PIC||"").trim(),problem:String(p.Masalah||"").trim(),solution:String(p.Solusi||"").trim(),status:String(p.Status||"").trim(),completion_date:d(p["Tgl Selesai"]),document_link:String(p.Dokumen||"").trim()})).filter(p=>p.meeting_date&&p.employee_name&&p.branch_id),g=await x("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:u,onDuplicate:"update"})});if(!g.ok)throw new Error(g.data?.error||"Import gagal");return g.data}},formFields:o=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:o?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:o?.branch_id&&!Fe.find(n=>n.value==o.branch_id)?[...Fe,{value:o.branch_id,label:o.branch_name||o.branch_id}]:Fe,createApi:{path:"/api/branches",field:"full_name"},value:o?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:s(o?.employee_name),value:o?.employee_name},{name:"pic",label:"PIC",type:"select",options:l(o?.pic),createApi:{path:"/api/pic",field:"name"},value:o?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:o?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:o?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:o?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:o?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:o?.document_link}]})}A();async function qt(t){let e=await K(),i=await Z(),a=["Ade Surahman","Berlin Ariansyah","Mizwar","Fajar","Ade","Berlin"],r=n=>n&&!i.find(d=>d.value===n)?[...i,{value:n,label:n}]:i,s=n=>n&&!a.find(d=>(typeof d=="object"?d.value:d)===n)?[...a,n]:a,l=Array.from({length:5},(n,d)=>String(new Date().getFullYear()-d));P({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:n=>{try{let d=JSON.parse(n);return Array.isArray(d)?d.join(", "):n||"-"}catch{return n||"-"}}},{key:"score",label:"Nilai",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer / peserta..."},{type:"select",name:"batch",label:"Batch",options:["Batch 1","Batch 2","Batch 3","Batch 4","Batch 5"]},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"trainer",label:"Trainer",options:["Berlin Ariansyah"]},{type:"select",name:"year",label:"Tahun",options:l}],exportOptions:{moduleName:"training",onExport:async n=>{let d=new URLSearchParams(n||{}).toString(),u=await x(`/api/training?limit=10000&${d}`);if(u.ok){let g=u.data.data.map(c=>{let f=c.participants||"";try{let h=JSON.parse(f);f=Array.isArray(h)?h.join(", "):f}catch{}return{Tanggal:c.training_date||"",Batch:c.batch||"",Materi:c.subject||"",Cabang:c.branch_name||"",Trainer:c.trainer||"",Peserta:f,Nilai:c.score!==null&&c.score!==void 0?c.score:"",Dokumen:c.document_link||""}}),{downloadExcel:p}=await Promise.resolve().then(()=>(j(),re));p(g,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:d}=await Promise.resolve().then(()=>(j(),re));d(n,"Template_Import_Training")},onImport:async n=>{let d=c=>{if(!c)return null;let f=String(c||"").toLowerCase(),h=e.find(m=>String(m.label||"").toLowerCase()===f);return h?h.value:null},u=c=>{if(!c)return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let f=String(c).trim();if(/^\d{4,5}$/.test(f)){let m=Number(f);if(m>2e4&&m<99999){let y=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(f))return f.slice(0,10);let h=f.split(/[\/\-\.]/);if(h.length===3){let[m,y,S]=h.map(D=>D.trim());if(m.length===4&&y.length<=2&&S.length<=2)return`${m}-${y.padStart(2,"0")}-${S.padStart(2,"0")}`;if(S.length===4&&y.length<=2&&m.length<=2)return`${S}-${y.padStart(2,"0")}-${m.padStart(2,"0")}`}return f},g=n.map(c=>({training_date:u(c.Tanggal),batch:String(c.Batch||"").trim(),subject:String(c.Materi||"").trim(),branch_id:d(String(c.Cabang||"").trim()),trainer:String(c.Trainer||"").trim(),participants:String(c.Peserta||"").trim(),score:c.Nilai?Number(c.Nilai):null,document_link:String(c.Dokumen||"").trim()})).filter(c=>c.training_date&&c.subject&&c.branch_id),p=await x("/api/training/import",{method:"POST",body:JSON.stringify(g)});if(!p.ok)throw new Error(p.data?.error||"Import gagal")}},formFields:n=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:n?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:n?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:n?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:n?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:s(n?.trainer),value:n?.trainer}]},{name:"participants",label:"Peserta",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let d=JSON.parse(n?.participants);return Array.isArray(d)?d.join(", "):n?.participants||""}catch{return n?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:n?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:n?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],onBeforeSubmit:async n=>n})}A();_e();j();function Ht({container:t,title:e,icon:i,apiPath:a,columns:r,formFields:s,filterFields:l,defaultFilters:o={},itemLabel:n="Data",canCreate:d=!0,canEdit:u=!0,canDelete:g=!0,onBeforeSubmit:p,onAfterLoad:c,onDataLoaded:f,extraActions:h=[],initialSearch:m="",exportOptions:y=null,bulkDelete:S=!1,paginationMode:D="server"}){let T=ie();T&&T.role==="viewer"&&(d=!1,u=!1,g=!1,S=!1,y=null),T&&T.role==="editor_khusus"&&a!=="/api/issues"&&a!=="/api/overtime"&&(d=!1,u=!1,g=!1,S=!1,y=null),T&&T.role==="input_lembur"&&(a==="/api/overtime"?(g=!1,S=!1,y=null):(d=!1,u=!1,g=!1,S=!1,y=null));let w=1,$={...o};m&&($.search=m);let _=new Set;t.innerHTML=`
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
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${d?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
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
    

    ${l&&l.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${l.filter(b=>b.type==="search").map(b=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${b.placeholder||"Cari..."}" id="filter-search" value="${$.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${l.filter(b=>b.type!=="search").map(b=>{if(b.type==="select"||b.type==="combobox"){let k=(b.label||"").startsWith("Pilih")?b.label:`Pilih ${b.label||""}`;return`<select class="filter-select" name="${b.name}" id="filter-${b.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${k}</option>${(b.options||[]).map(v=>`<option value="${typeof v=="object"?v.value:v}" ${$[b.name]===(typeof v=="object"?v.value:v)?"selected":""}>${typeof v=="object"?v.label:v}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${l.filter(b=>b.type!=="search").map(b=>{if(b.type==="select"||b.type==="combobox"){let k=(b.label||"").startsWith("Pilih")?b.label:`Pilih ${b.label||""}`;return`<select class="filter-select filter-select-sheet" name="${b.name}-sheet" id="filter-sheet-${b.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${k}</option>${(b.options||[]).map(v=>`<option value="${typeof v=="object"?v.value:v}" ${$[b.name]===(typeof v=="object"?v.value:v)?"selected":""}>${typeof v=="object"?v.label:v}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function C(){let b=document.getElementById("bulk-toolbar");if(!b)return;let k=document.getElementById("bulk-count"),v=document.getElementById("btn-bulk-delete"),M=document.getElementById("btn-bulk-cancel");k.textContent=`${_.size} item dipilih`,_.size>0?(b.style.display="flex",v.disabled=!1,M.disabled=!1):(b.style.display="none",v.disabled=!0,M.disabled=!0);let N=document.getElementById("select-all-checkbox");if(N){let H=document.querySelectorAll(".row-checkbox");if(H.length>0){let I=[...H].every(ne=>ne.checked),L=[...H].some(ne=>ne.checked);N.checked=I,N.indeterminate=L&&!I}else N.checked=!1,N.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{_.clear(),document.querySelectorAll(".row-checkbox").forEach(k=>k.checked=!1);let b=document.getElementById("select-all-checkbox");b&&(b.checked=!1),C()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(_.size===0)return;let b=[..._],k=document.createElement("div");k.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",k.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${b.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${b.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(k),k.querySelector("#bulk-cancel-btn").addEventListener("click",()=>k.remove()),k.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let v=k.querySelector("#bulk-confirm-btn");v.disabled=!0,v.textContent="Menghapus...";let M=await x(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:b})});k.remove(),M.ok?(W(`${b.length} ${n} berhasil dihapus.`),_.clear(),C(),oe(a),R()):V(M.data?.error||"Gagal menghapus data.")})});let B=document.getElementById("filter-search"),E;B?.addEventListener("input",b=>{clearTimeout(E),E=setTimeout(()=>{$.search=b.target.value,w=1,_.clear(),C(),R()},400)}),l?.forEach(b=>{(b.type==="select"||b.type==="combobox")&&(document.getElementById(`filter-${b.name}`)?.addEventListener("change",k=>{$[b.name]=k.target.value;let v=document.getElementById(`filter-sheet-${b.name}`);v&&(v.value=k.target.value),w=1,_.clear(),C(),R()}),document.getElementById(`filter-sheet-${b.name}`)?.addEventListener("change",k=>{$[b.name]=k.target.value;let v=document.getElementById(`filter-${b.name}`);v&&(v.value=k.target.value),w=1,_.clear(),C(),R(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{$={},B&&(B.value=""),l?.forEach(b=>{let k=document.getElementById(`filter-${b.name}`);k&&(k.value="");let v=document.getElementById(`filter-sheet-${b.name}`);v&&(v.value="")}),w=1,_.clear(),C(),R()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{$={},B&&(B.value=""),l?.forEach(b=>{let k=document.getElementById(`filter-${b.name}`);k&&(k.value="");let v=document.getElementById(`filter-sheet-${b.name}`);v&&(v.value="")}),w=1,_.clear(),C(),R(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>ye(null)),y&&document.addEventListener("click",function(b){let k=document.getElementById("aksi-menu-main"),v=document.getElementById("btn-aksi-main");k&&v&&!v.contains(b.target)&&!k.contains(b.target)&&k.classList.remove("show-aksi-menu")});let z=document.getElementById("btn-mobile-filter"),ee=document.getElementById("filter-options-wrapper"),be=document.getElementById("btn-close-filter-sheet");if(z&&ee&&(z.addEventListener("click",b=>{b.preventDefault(),ee.classList.add("sheet-open")}),be&&be.addEventListener("click",b=>{b.preventDefault(),ee.classList.remove("sheet-open")})),y){document.getElementById(`btn-export-${y.moduleName}`)?.addEventListener("click",async k=>{let v=k.target,M=v.innerHTML;v.innerHTML="\u23F3 Loading...",v.disabled=!0;try{await y.onExport()}catch{V("Gagal export data")}finally{v.innerHTML=M,v.disabled=!1}}),document.getElementById(`btn-template-${y.moduleName}`)?.addEventListener("click",()=>{y.onTemplate()});let b=document.getElementById(`input-import-${y.moduleName}`);b?.addEventListener("change",async k=>{let v=k.target.files[0];if(!v)return;let M=document.getElementById(`label-import-${y.moduleName}`),N=M?M.querySelector(".import-text"):null,H=N?N.innerText:"";N&&(N.innerText="\u231B Memproses..."),M&&(M.style.pointerEvents="none"),b.disabled=!0;try{let I=await Je(v);if(I.length===0)throw new Error("File kosong atau format salah");await y.onImport(I),W("Import berhasil!"),oe(a),R()}catch(I){V(I.message||"Gagal import data")}finally{N&&(N.innerText=H),M&&(M.style.pointerEvents="auto"),b.disabled=!1,b.value=""}})}async function R(){C();let b=document.getElementById("table-container");if(!b)return;b.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let k=D==="client",v=k?1:w,M=k?$e:20,N=new URLSearchParams({page:v,limit:M,...Object.fromEntries(Object.entries($).filter(([,q])=>q))}),H=await x(`${a}?${N}`);if(!H.ok){b.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${H.data?.error||"Error"}</p></div>`;return}let I=H.data?.data||H.data||[],L=H.data?.pagination,ne=I.length;if(k){I=f(I);let q=I.length,Q=20,te=Math.ceil(q/Q);w>te&&te>0&&(w=te);let O=(w-1)*Q,le=w*Q;I=I.slice(O,le),L={page:w,limit:Q,total:q,pages:te}}!1,c&&c(I);let ve=Ye({columns:r,data:I,onEdit:u?q=>ye(q):null,actions:h.map(q=>({...q,handler:Q=>q.handler(Q,R)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:S?{selectedIds:_,onToggle:C}:null});b.innerHTML="",b.appendChild(ve);let se=document.getElementById("pagination-container");if(se&&(se.innerHTML="",L&&L.pages>1)){let q=We({page:L.page,pages:L.pages,total:L.total,limit:L.limit,onPage:Q=>{w=Q,R()}});q&&se.appendChild(q)}}function Te(b){let k=typeof s=="function"?s(b):s;return He(k)}function ye(b){let k=!!b,v=document.createElement("form");if(v.noValidate=!0,v.innerHTML=Te(b),k){let N=typeof s=="function"?s(b):s;Ze(v,b)}let{close:M}=de({title:k?`Edit ${n}`:`Tambah ${n}`,content:v,size:"lg",confirmText:k?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(N,H)=>{if(!v.reportValidity())return;let I=N.querySelector(".modal-confirm");I.disabled=!0,I.textContent="Menyimpan...";let L=Xe(v),ne=typeof s=="function"?s(b):s,ve=async te=>{for(let O of te)if(O.type==="row")await ve(O.fields);else if(O.type==="combobox"&&L[O.name]){let le=L[O.name],ke=(O.options||[]).find(Y=>{let ae=String(typeof Y=="object"?Y.value:Y),dt=String(typeof Y=="object"?Y.label:Y);return ae===le||dt===le});if(ke)L[O.name]=typeof ke=="object"?ke.value:ke;else if(O.createApi){let Y={};Y[O.createApi.field]=le,O.createApi.extra&&Object.assign(Y,O.createApi.extra);let ae=await x(O.createApi.path,{method:"POST",body:JSON.stringify(Y)});if(ae.ok&&ae.data?.id)L[O.name]=ae.data.id;else if(ae.ok&&!ae.data?.id)L[O.name]=le;else throw new Error(`Gagal membuat master data: ${ae.data?.error||"Unknown error"}`)}}};try{await ve(ne)}catch(te){V(te.message),I.disabled=!1,I.textContent=k?"Simpan Perubahan":`Tambah ${n}`;return}p&&(L=await p(L,b));let se=k?"PUT":"POST",q=k?`${a}/${b.id}`:a,Q=await x(q,{method:se,body:JSON.stringify(L)});Q.ok?(W(k?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),H(),oe(a),R()):(V(Q.data?.error||"Gagal menyimpan data."),I.disabled=!1,I.textContent=k?"Simpan Perubahan":`Tambah ${n}`)}})}function ga(b){qe(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let k=await x(`${a}/${b.id}`,{method:"DELETE"});k.ok?(W(`${n} berhasil dihapus.`),oe(a),R()):V(k.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return R(),R}A();j();async function Jt(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await K(),a=await Z(),r=e?e.get("dash_filter"):null,s={};if(r==="reliever"){let p=new Date,c=`${p.getFullYear()}-${String(p.getMonth()+1).padStart(2,"0")}`;s={status:"Done",month:e&&e.get("month")?e.get("month"):c}}console.log("RAW",await Ce()),console.log("OPTIONS",a);let l=p=>p&&!a.find(c=>c.value===p)?[...a,{value:p,label:p}]:a,o=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],n=p=>p&&!o.includes(p)?[...o,p]:o,d=new Date().getFullYear(),g=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((p,c)=>{let f=String(c+1).padStart(2,"0");return{value:`${d}-${f}`,label:`${p} ${d}`}});Ht({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:s,onDataLoaded:p=>p.sort((c,f)=>{let h=c.backup_date?new Date(c.backup_date).getTime():0;return(f.backup_date?new Date(f.backup_date).getTime():0)-h}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:p=>ce(p)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:p=>window.formatDate(p)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:p=>window.formatDate(p)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:p=>p?`<span class="badge badge-info">${p}</span>`:"-"},{key:"status",label:"Status",render:p=>U(p)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:o},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"month",label:"Bulan",options:g},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:p=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:p?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:p?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:l(p?.original_fc_name),value:p?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:n(p?.reliever_name),value:p?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:p?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:p?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:p?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:p?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:p?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let p=await x(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(p.ok){let c=p.data.data.map(f=>({Cabang:f.branch_name||"","Nama Facility care":f.original_fc_name||"",Periode:f.period||"",Relifer:f.reliever_name||"","Tanggal Back Up":f.backup_date||"","Tanggal Selesai":f.completion_date||"",Keterangan:f.reason||"",Shift:f.shift||"",Status:f.status||""}));c.length===0&&c.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),F(c,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{F([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async p=>{let f=(await x("/api/branches?all=1")).data?.data||[],h=S=>{if(!S)return null;let D=String(S||"").toLowerCase(),T=f.find(w=>String(w.full_name||"").toLowerCase()===D||String(w.code||"").toLowerCase()===D||String(w.name||"").toLowerCase()===D);return T?T.id:null},m=p.map(S=>({branch_name:String(S.Cabang||"").trim(),backup_date:String(S["Tanggal Back Up"]||S["Tanggal Backup"]||"").trim(),original_fc_name:String(S["Nama Facility care"]||S["FC Digantikan"]||"").trim(),reliever_name:String(S.Relifer||S.Reliefer||"").trim(),period:String(S.Periode||"").trim(),reason:String(S.Keterangan||"").trim(),shift:String(S.Shift||"").trim(),completion_date:String(S["Tanggal Selesai"]||"").trim(),status:String(S.Status||"").trim()})).filter(S=>S.reliever_name&&S.backup_date),y=await x("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!y.ok)throw new Error(y.data?.error||"Import gagal");return y.data}}})}A();j();async function Ut(t){let e=await K(),i=Array.from({length:4},(a,r)=>String(new Date().getFullYear()-r));P({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>ce(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>U(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),s=await x(`/api/reports/inspection?limit=10000&${r}`);if(s.ok){let l=s.data.data.map(o=>({Cabang:o.branch_name||"",Periode:o.period||"",Tanggal:o.inspection_date||"","Point FC":o.fc_score!==null&&o.fc_score!==void 0?o.fc_score:"","Point SPV":o.spv_score!==null&&o.spv_score!==void 0?o.spv_score:"",Status:o.status||"","Link Dokumen":o.document_link||""}));F(l,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{F([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let r=n=>{if(!n)return null;let d=String(n||"").toLowerCase(),u=e.find(g=>String(g.label||"").toLowerCase()===d);return u?u.value:null},s=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let d=String(n).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let g=Number(d);if(g>2e4&&g<99999){let p=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let u=d.split(/[\/\-\.]/);if(u.length===3){let[g,p,c]=u.map(f=>f.trim());if(g.length===4&&p.length<=2&&c.length<=2)return`${g}-${p.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&p.length<=2&&g.length<=2)return`${c}-${p.padStart(2,"0")}-${g.padStart(2,"0")}`}return d},l=a.map(n=>({branch_id:r(String(n.Cabang||"").trim()),period:String(n.Periode||"").trim(),inspection_date:s(n.Tanggal),fc_score:n["Point FC"]!==void 0&&n["Point FC"]!==""?Number(n["Point FC"]):null,spv_score:n["Point SPV"]!==void 0&&n["Point SPV"]!==""?Number(n["Point SPV"]):null,status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.inspection_date),o=await x("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}}})}A();j();async function Gt(t){let e=await K(),i=Array.from({length:4},(a,r)=>String(new Date().getFullYear()-r));P({container:t,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>ce(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>U(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),s=await x(`/api/reports/cleaning?limit=10000&${r}`);if(s.ok){let l=s.data.data.map(o=>({Cabang:o.branch_name||"",Jenis:o.activity_type||"",Periode:o.period||"",Tanggal:o.activity_date||"",Status:o.status||"","Link Dokumen":o.document_link||""}));F(l,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{F([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let r=n=>{if(!n)return null;let d=String(n||"").toLowerCase(),u=e.find(g=>String(g.label||"").toLowerCase()===d);return u?u.value:null},s=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let d=String(n).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let g=Number(d);if(g>2e4&&g<99999){let p=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let u=d.split(/[\/\-\.]/);if(u.length===3){let[g,p,c]=u.map(f=>f.trim());if(g.length===4&&p.length<=2&&c.length<=2)return`${g}-${p.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&p.length<=2&&g.length<=2)return`${c}-${p.padStart(2,"0")}-${g.padStart(2,"0")}`}return d},l=a.map(n=>({branch_id:r(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:s(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),o=await x("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}}})}A();j();async function zt(t,e){let i=await K(),a=Array.from({length:4},(l,o)=>String(new Date().getFullYear()-o)),r=e?e.get("dash_filter"):null,s={};if(r==="fogging"){let l=new Date,o=String(l.getMonth()+1).padStart(2,"0"),n=String(l.getFullYear()),d=e?e.get("month"):null;d&&d.length===7&&(n=d.split("-")[0],o=d.split("-")[1]),s={status:"Done",month:o,year:n}}P({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:s,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:l=>`<span class="badge badge-warning">${l}</span>`},{key:"period",label:"Periode",render:l=>ce(l)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>U(l)},{key:"document_link",label:"Dokumen",render:l=>l?`<a href="${l}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:l=>l||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:l=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:i,value:l?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:l?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:l?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:l?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:l?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async l=>{let o=new URLSearchParams(l||{}).toString(),n=await x(`/api/reports/fogging?limit=10000&${o}`);if(n.ok){let d=n.data.data.map(u=>({Cabang:u.branch_name||"",Jenis:u.activity_type||"Fogging",Periode:u.period||"",Tanggal:u.activity_date||"",Status:u.status||"","Link Dokumen":u.document_link||""}));F(d,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{F([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async l=>{let o=g=>{if(!g)return null;let p=String(g||"").toLowerCase(),c=i.find(f=>String(f.label||"").toLowerCase()===p);return c?c.value:null},n=g=>{if(g==null||g==="")return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let p=String(g).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let f=Number(p);if(f>2e4&&f<99999){let h=new Date(Date.UTC(1899,11,30)+f*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}let c=p.split(/[\/\-\.]/);if(c.length===3){let[f,h,m]=c.map(y=>y.trim());if(f.length===4&&h.length<=2&&m.length<=2)return`${f}-${h.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&h.length<=2&&f.length<=2)return`${m}-${h.padStart(2,"0")}-${f.padStart(2,"0")}`}return p},d=l.map(g=>({branch_id:o(String(g.Cabang||"").trim()),activity_type:String(g.Jenis||g.Kegiatan||"Fogging").trim(),period:String(g.Periode||"").trim(),activity_date:n(g.Tanggal),status:String(g.Status||"").trim(),document_link:String(g["Link Dokumen"]||"").trim(),notes:String(g.Catatan||g.Keterangan||"").trim()})).filter(g=>g.branch_id&&g.period&&g.activity_date),u=await x("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(d)});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}A();j();async function Qt(t){let e=await K(),i=await Z(),a=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],r=Array.from({length:4},(o,n)=>String(new Date().getFullYear()-n)),s=o=>o&&!i.find(n=>n.value===o)?[...i,{value:o,label:o}]:i,l=o=>o&&!a.find(n=>n.value===o)?[...a,{value:o,label:o}]:a;P({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:o=>window.formatDate(o)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:o=>`<span title="${o||""}">${o?.length>60?o.slice(0,60)+"\u2026":o||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>U(o)},{key:"notes",label:"Keterangan",render:o=>o?.length>40?o.slice(0,40)+"\u2026":o||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:o?.branch_id},{name:"pic",label:"PIC",type:"select",options:l(o?.pic),value:o?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:o?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:o?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:o?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:o?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async o=>{let n=new URLSearchParams(o||{}).toString(),d=await x(`/api/reports/basecamp?limit=10000&${n}`);if(d.ok){let u=d.data.data.map(g=>({"Tgl Info":g.info_date||"",Cabang:g.branch_name||"",Permasalahan:g.problem||"",PIC:g.pic||"","Tgl Done":g.done_date||"",Status:g.status||"",Keterangan:g.notes||""}));F(u,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{F([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async o=>{let n=p=>{if(!p)return null;let c=String(p||"").toLowerCase(),f=e.find(h=>String(h.label||"").toLowerCase()===c);return f?f.value:null},d=p=>{if(p==null||p==="")return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let c=String(p).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let h=Number(c);if(h>2e4&&h<99999){let m=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}let f=c.split(/[\/\-\.]/);if(f.length===3){let[h,m,y]=f.map(S=>S.trim());if(h.length===4&&m.length<=2&&y.length<=2)return`${h}-${m.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&m.length<=2&&h.length<=2)return`${y}-${m.padStart(2,"0")}-${h.padStart(2,"0")}`}return c},u=o.map(p=>({info_date:d(p["Tgl Info"]||p["Tanggal Info"]),branch_id:n(String(p.Cabang||"").trim()),problem:String(p.Permasalahan||"").trim(),pic:String(p.PIC||"").trim(),done_date:d(p["Tgl Done"]||p["Tanggal Done"]),status:String(p.Status||"").trim(),notes:String(p.Keterangan||p.Catatan||"").trim()})).filter(p=>p.info_date&&p.branch_id&&p.problem),g=await x("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(u)});if(!g.ok)throw new Error(g.data?.error||"Import gagal");return g.data}}})}async function Vt(t){P({container:t,title:"SOP",icon:"\u{1F4DA}",apiPath:"/api/sop",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(A(),Ee)),r=await a(`/api/sop?limit=10000&${i}`);if(r.ok){let s=r.data.data.map(o=>({"Nama SOP":o.name||"",Kategori:o.category||"",Dokumen:o.document_link||"",Catatan:o.notes||o.description||""})),{downloadExcel:l}=await Promise.resolve().then(()=>(j(),re));l(s,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(j(),re));i(e,"Template_Import_SOP")},onImport:async e=>{let i=e.map(s=>({name:String(s["Nama SOP"]||"").trim(),category:String(s.Kategori||"").trim(),document_link:String(s.Dokumen||"").trim(),description:String(s.Catatan||"").trim()})).filter(s=>s.name),{apiFetch:a}=await Promise.resolve().then(()=>(A(),Ee)),r=await a("/api/sop/import",{method:"POST",body:JSON.stringify(i)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function Yt(t){P({container:t,title:"Master Checklist",icon:"\u{1F4CB}",apiPath:"/api/checklist",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(A(),Ee)),r=await a(`/api/checklist?limit=10000&${i}`);if(r.ok){let s=r.data.data.map(o=>({"Nama Checklist":o.name||"",Kategori:o.category||"",Dokumen:o.document_link||"",Deskripsi:o.description||""})),{downloadExcel:l}=await Promise.resolve().then(()=>(j(),re));l(s,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(j(),re));i(e,"Template_Import_Checklist")},onImport:async e=>{let i=e.map(s=>({name:String(s["Nama Checklist"]||"").trim(),category:String(s.Kategori||"").trim(),document_link:String(s.Dokumen||"").trim(),description:String(s.Deskripsi||"").trim()})).filter(s=>s.name),{apiFetch:a}=await Promise.resolve().then(()=>(A(),Ee)),r=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(i)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}A();_e();j();async function wt(t,e="forms"){if(e==="supply")return Ra(t);Ka(t)}function Ka(t){P({container:t,title:"Master Form",icon:"\u{1F4D1}",apiPath:"/api/forms",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await x(`/api/forms?limit=10000&${i}`);a.data?.data?F(a.data.data,"Data_Master_Form"):V("Gagal export data master form")},onImport:async e=>{let i=await x("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!i.ok)throw new Error(i.data?.error||"Import failed");return i.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function Ra(t){let i=((await x("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));P({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,r)=>r.branch_name_ref||r.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let r=JSON.parse(a);return Array.isArray(r)?r.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let r=JSON.parse(a);return Array.isArray(r)?r.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>U(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let r=a?.tools_items;try{r=Array.isArray(JSON.parse(r))?JSON.parse(r).join(", "):r}catch{}let s=a?.chemical_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:a?.branch_id&&!i.find(l=>l.value==a.branch_id)?[...i,{value:a.branch_id,label:a.branch_name||a.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:r},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:s},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),s=await x(`/api/reports/supply?limit=10000&${r}`);if(s.ok){let l=s.data.data.map(o=>{let n=o.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let d=o.chemical_items;try{d=Array.isArray(JSON.parse(d))?JSON.parse(d).join(", "):d}catch{}return{Waktu:o.submitted_at||"",Pengirim:o.submitter_name||"",Cabang:o.branch_name_ref||o.branch_name||"","Alat/Barang":n||"",Chemical:d||"",Catatan:o.additional_notes||"",Status:o.status||"","Diproses Oleh":o.processed_by||""}});F(l,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{F([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let s=(await x("/api/branches?all=1")).data?.data||[],l=u=>{if(!u)return null;let g=String(u||"").toLowerCase(),p=s.find(c=>String(c.full_name||"").toLowerCase()===g||String(c.code||"").toLowerCase()===g||String(c.name||"").toLowerCase()===g);return p?p.id:null},o=u=>{if(u==null||u==="")return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let g=String(u).trim();if(g===""||g==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);if(/^\d{4,5}$/.test(g)){let c=Number(g);if(c>2e4&&c<99999){let f=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}let p=g.split(/[\/\-\.]/);if(p.length===3){let[c,f,h]=p.map(m=>m.trim());if(c.length===4&&f.length<=2&&h.length<=2)return`${c}-${f.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&f.length<=2&&c.length<=2)return`${h}-${f.padStart(2,"0")}-${c.padStart(2,"0")}`}return g},n=a.map(u=>({submitted_at:o(u.Waktu||u.Tanggal),submitter_name:String(u.Pengirim||"").trim(),branch_id:l(String(u.Cabang||"").trim()),tools_items:String(u["Alat/Barang"]||u.Alat||"").trim(),chemical_items:String(u.Chemical||"").trim(),additional_notes:String(u.Catatan||u.Keterangan||"").trim(),status:String(u.Status||"").trim(),processed_by:String(u["Diproses Oleh"]||u.PIC||"").trim()})).filter(u=>u.submitted_at&&u.submitter_name&&u.branch_id),d=await x("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,r)=>{let s=de({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(l,o)=>{let n=l.querySelector("#supply-status").value,d=l.querySelector("#supply-processed-by").value;(await x(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:d})})).ok?(W("Status diperbarui."),o(),r()):V("Gagal update status.")}})}}]})}A();j();async function Wt(t){let e=ie();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}P({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",editor_khusus:"badge-warning",input_lembur:"badge-info",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let a=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"editor_khusus",label:"Editor (Hanya Lembur & Masalah)"},{value:"input_lembur",label:"Input Lembur (Hanya Isi Lembur)"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await x(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let a=i.data.data.map(r=>({"Nama Lengkap":r.full_name||"",Username:r.username||"",Email:r.email||"",Role:r.role||"",Status:r.is_active?"Aktif":"Nonaktif"}));F(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{F([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let a=i.map(s=>({full_name:String(s["Nama Lengkap"]||"").trim(),username:String(s.Username||"").trim(),email:String(s.Email||"").trim(),role:String(s.Role||"").trim()||"viewer",password:String(s.Password||"").trim()})).filter(s=>s.username&&s.password&&s.email&&s.full_name),r=await x("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}A();j();async function Xt(t){P({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await x(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)F(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{F([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let i=e.map(r=>({code:String(r["Kode Cabang"]||"").trim(),name:String(r["Nama Pendek"]||"").trim(),full_name:String(r["Nama Lengkap"]||"").trim(),city:String(r.Kota||"").trim()})).filter(r=>r.code&&r.name),a=await x("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}A();async function Zt(t){let e=new Date,i=[];t.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),r()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),r()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(l=>l.addEventListener("change",r));async function a(){try{let l=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;i=(await x(`/api/dashboard/calendar?month=${l}`)).data?.data||[]}catch(l){console.warn("[Calendar] Failed to load events, rendering empty grid:",l),i=[]}}async function r(){let l=document.getElementById("calendar-grid");if(l){l.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let o=e.getFullYear(),n=e.getMonth(),d=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),u=document.getElementById("cal-month-label");u&&(u.textContent=d);let g=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(w=>w.value)),p=i.filter(w=>g.has(w.type)),c={};p.forEach(w=>{let $=ja(w.event_date);$&&(c[$]||(c[$]=[]),c[$].push(w))});let f=new Date(o,n,1).getDay(),h=new Date(o,n+1,0).getDate(),m=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],y=new Date().toISOString().slice(0,10),S='<div class="calendar-grid">';m.forEach(w=>{S+=`<div class="cal-day-header">${w}</div>`});for(let w=0;w<f;w++)S+='<div class="cal-cell cal-cell-empty"></div>';for(let w=1;w<=h;w++){let $=`${o}-${String(n+1).padStart(2,"0")}-${String(w).padStart(2,"0")}`,_=c[$]||[],C=$===y;S+=`
          <div class="cal-cell ${C?"cal-today":""} ${_.length?"cal-has-events":""}"
               data-date="${$}" tabindex="0" role="button" aria-label="${$}">
            <div class="cal-day-num ${C?"today-num":""}">${w}</div>
            <div class="cal-events-preview">
              ${_.slice(0,3).map(B=>`
                <div class="cal-event-dot cal-color-${B.color||"gray"}" title="${ge(B.title||B.type)}">
                  <span class="cal-event-dot-label">${qa(B.title||B.branch_name||B.type,18)}</span>
                </div>
              `).join("")}
              ${_.length>3?`<div class="cal-more">+${_.length-3} lagi</div>`:""}
            </div>
          </div>`}let T=(f+h)%7;if(T!==0)for(let w=0;w<7-T;w++)S+='<div class="cal-cell cal-cell-empty"></div>';S+="</div>",l.innerHTML=S,l.querySelectorAll(".cal-cell[data-date]").forEach(w=>{w.addEventListener("click",()=>{let $=w.dataset.date,_=c[$]||[];if(!_.length)return;let C=document.getElementById("cal-event-list"),B=new Date($+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=B,document.getElementById("cal-event-items").innerHTML=_.map(E=>`
            <div class="cal-event-item cal-color-border-${E.color||"gray"}">
              <div class="cal-event-type">
                ${Ha(E.type)}
                ${E.date_type?` <span class="badge" style="font-size:0.75rem;padding:2px 6px;background:rgba(0,0,0,0.06);margin-left:4px;">${ge(E.date_type)}</span>`:""}
              </div>
              <div class="cal-event-title" style="font-weight:600;font-size:0.95rem;margin-top:2px;">${ge(E.title||"-")}</div>
              ${E.branch_name?`<div class="cal-event-branch" style="color:var(--text-2);font-size:0.85rem;margin-top:2px;">\u{1F4CD} ${ge(E.branch_name)}</div>`:""}
              ${E.pic||E.period||E.opening_date||E.target_date||E.completion_date||E.notes?`
                <div class="cal-event-meta" style="margin-top:6px;padding-top:6px;border-top:1px dashed var(--border);font-size:0.8rem;color:var(--text-2);display:flex;flex-direction:column;gap:3px;">
                  ${E.pic?`<div>\u{1F464} <b>PIC:</b> ${ge(E.pic)}</div>`:""}
                  ${E.period?`<div>\u{1F4C5} <b>Periode:</b> ${ge(E.period)}</div>`:""}
                  ${E.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${ge(xt(E.opening_date))}</div>`:""}
                  ${E.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${ge(xt(E.target_date))}</div>`:""}
                  ${E.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${ge(xt(E.completion_date))}</div>`:""}
                  ${E.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${ge(E.notes)}</div>`:""}
                </div>
              `:""}
              ${E.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${E.status==="Done"?"badge-success":"badge-warning"}">${ge(E.status)}</span></div>`:""}
              ${E.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${E.days_remaining} hari</div>`:""}
            </div>
          `).join(""),C.style.display="block"})})}catch(o){console.error("[Calendar] Render error:",o),l&&(l.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let s=()=>{a().then(()=>r())};et.on("data:changed",s),r()}function ja(t){if(!t||t==="-"||String(t).trim()==="")return"";let e=String(t).trim();if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);let i=e.split(/[\/\-\.]/);if(i.length===3){let[a,r,s]=i.map(l=>l.trim());if(s.length===4)return`${s}-${r.padStart(2,"0")}-${a.padStart(2,"0")}`;if(a.length===4)return`${a}-${r.padStart(2,"0")}-${s.padStart(2,"0")}`}return e.slice(0,10)}function xt(t){if(!t||t==="-"||String(t).trim()==="")return"";let i=String(t).trim().split(/[\/\-\.]/);if(i.length===3){let[a,r,s]=i.map(l=>l.trim());if(a.length===4)return`${s.padStart(2,"0")}-${r.padStart(2,"0")}-${a}`;if(s.length===4)return`${a.padStart(2,"0")}-${r.padStart(2,"0")}-${s}`}return t}function qa(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function ge(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Ha(t){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging"}[t]||t}A();async function ea(t){let e=ie(),i=(e?.full_name||e?.username||"U")[0].toUpperCase(),r={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
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
  `;let s=localStorage.getItem("fm_token"),l=document.getElementById("session-info");if(s&&l)try{let o=JSON.parse(atob(s.split(".")[1])),n=new Date(o.exp*1e3);l.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{l.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async o=>{o.preventDefault();let n=document.getElementById("pwd-error"),d=document.getElementById("pwd-success"),u=document.getElementById("btn-save-pwd");n.style.display="none",d.style.display="none";let g=o.target,p=g.current_password.value,c=g.new_password.value,f=g.confirm_password.value;if(c!==f){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(c.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}u.disabled=!0,u.textContent="\u23F3 Menyimpan...";let h=await x("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:p,new_password:c})});u.disabled=!1,u.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',h.ok?(d.textContent="\u2705 Password berhasil diubah.",d.style.display="block",g.reset(),W("Password berhasil diubah.")):(n.textContent=h.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}A();var ot={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function X(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let r=Number(e);if(r>2e4&&r<99999){let s=new Date(Date.UTC(1899,11,30)+r*864e5);return isNaN(s.getTime())?null:s.toISOString().slice(0,10)}}let i=e.split(/[\/\-\.]/);if(i.length===3){let[r,s,l]=i.map(u=>u.trim()),o=Number(r),n=Number(s),d=Number(l);if(r.length===4&&o>1900)return`${r}-${s.padStart(2,"0")}-${l.padStart(2,"0")}`;if(l.length===4&&d>1900)return o>12?`${l}-${s.padStart(2,"0")}-${r.padStart(2,"0")}`:n>12?`${l}-${r.padStart(2,"0")}-${s.padStart(2,"0")}`:`${l}-${s.padStart(2,"0")}-${r.padStart(2,"0")}`;if(l.length===2&&!isNaN(d)){let u=d>=50?`19${l}`:`20${l}`;return o>12?`${u}-${s.padStart(2,"0")}-${r.padStart(2,"0")}`:`${u}-${s.padStart(2,"0")}-${r.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function ta(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Ja={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||t.Divisi||"FACILITY CARE",phone:t["No. Hp"]||t["No. HP"],join_date:X(t["Tanggal Masuk"]||t["Tgl Masuk"]),status:t.Status||"",target_pindah_os:X(t["Target Pindah OS"]),target_selesai_os:X(t["Target Selesai OS"]),notes:t.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:X(t["Tanggal Mulai"]),end_date:X(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:X(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:X(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:X(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:X(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:X(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:X(t["Tanggal Target"]||t["Tgl Target"]),completion_date:X(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:X(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:X(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:X(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:X(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:X(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:X(t["Tanggal Back Up"]),completion_date:X(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:X(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:X(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Ua(t,e){let i=ot[t];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Ja[i.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let r=[],s=[],l=[];return e.filter(n=>!ta(n)).forEach((n,d)=>{let u=e.indexOf(n)+2,g=[];a.required.forEach(({key:c,label:f})=>{let h=n[c];if(h==null||String(h).trim()===""){let m=Object.keys(n).filter(y=>y.trim()).join(", ");g.push({column:f,originalValue:h||"",reason:`Kolom "${f}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${m.slice(0,120)}`})}});let p=a.map(n);g.length>0?s.push({row:u,data:p,raw:n,errors:g}):(r.push(n),l.push(p))}),{valid:r,errors:s,mapped:l}}function aa(t){let e=[];return t.SheetNames.forEach(i=>{let a=ot[i];if(!a)return;let r=t.Sheets[i],s=window.XLSX.utils.sheet_to_json(r,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),l=Ua(i,s),o=s.filter(n=>!ta(n));e.push({sheetName:i,module:a.module,label:a.label,total:o.length,valid:l.mapped.length,errorCount:l.errors.length,errors:l.errors,mapped:l.mapped,skipped:!1})}),e}function na(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,r])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(r),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function ia(t){let e=window.XLSX,i=e.utils.book_new(),a=!1;return t.forEach(r=>{if(!r.errors||r.errors.length===0)return;a=!0;let s=r.errors.map(o=>({"No. Baris":o.row,"Kolom Gagal":(o.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(o.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(o.data||{}).map(([n,d])=>[n,d??""]))})),l=e.utils.json_to_sheet(s);e.utils.book_append_sheet(i,l,r.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Ga=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function ra(t){t.innerHTML=`
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
              ${Object.entries(ot).map(([h,{label:m}])=>`<span class="import-sheet-tag">\u{1F4C4} ${h} \u2192 ${m}</span>`).join("")}
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
  `;let e=null,i=null,a=0,r={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function s(h){Object.entries(r).forEach(([m,y])=>{y.style.display=m===h?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let h=document.getElementById("btn-backup-db");h.disabled=!0,h.textContent="\u23F3 Memproses Backup...";try{let m=await x("/api/import/backup");if(m.ok){if(!window.XLSX){V("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let y=window.XLSX,S=y.utils.book_new();Object.entries(m.data.database).forEach(([D,T])=>{let w=T.length>0?T:[{}],$=y.utils.json_to_sheet(w);y.utils.book_append_sheet(S,$,D.substring(0,31))}),y.writeFile(S,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),W("Backup berhasil diunduh!")}else V("Gagal memproses backup: "+(m.data?.error||"Unknown error"))}catch(m){V("Gagal memproses backup: "+m.message)}finally{h.disabled=!1,h.textContent="\u{1F4E6} Backup Database"}});let l=document.getElementById("btn-sync-google");l&&l.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let h=l.innerHTML;l.innerHTML='<span class="spinner"></span> Menyinkronkan...',l.disabled=!0;try{let m=await x("/api/sync/google-sheets",{method:"POST"});m.ok?alert("Sinkronisasi Berhasil: "+(m.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(m.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{l.innerHTML=h,l.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{na(),W("Template Excel berhasil didownload!")});let o=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",h=>{h.stopPropagation(),o.click()}),o.addEventListener("change",h=>{h.target.files[0]&&d(h.target.files[0])}),n.addEventListener("dragover",h=>{h.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",h=>{h.preventDefault(),n.classList.remove("drag-over");let m=h.dataTransfer.files[0];m&&m.name.match(/\.xlsx?$/i)?d(m):V("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,o.value="",document.getElementById("file-info").style.display="none",n.style.display="",s("upload")});async function d(h){e=h,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${h.name} (${(h.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await u(h)}async function u(h){s("validating");let m=document.getElementById("validation-status"),y=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");m.textContent="Membaca file Excel...",y.style.width="20%",await ze(200);let S=await h.arrayBuffer(),D=window.XLSX.read(S,{type:"array",cellDates:!0});m.textContent=`Memvalidasi ${D.SheetNames.length} sheet...`,y.style.width="50%",await ze(100),i=aa(D),y.style.width="100%",m.textContent="Validasi selesai!",await ze(300),g()}catch(S){s("upload"),V("Gagal memproses file: "+S.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function g(){s("preview");let h=i.filter(C=>!C.skipped).length,m=i.reduce((C,B)=>C+B.total,0),y=i.reduce((C,B)=>C+B.valid,0),S=i.reduce((C,B)=>C+B.errorCount,0),D=m>0?Math.round(y/m*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${h} sheet</span>
      <span class="badge badge-secondary">${m} baris</span>
      <span class="badge badge-success">${y} valid (${D}%)</span>
      ${S>0?`<span class="badge badge-danger">${S} error</span>`:""}
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
          ${i.map((C,B)=>`
            <tr class="${C.errorCount>0?"row-error":C.skipped?"row-skipped":"row-ok"}">
              <td><strong>${C.sheetName}</strong></td>
              <td>${C.label}</td>
              <td style="text-align:center">${C.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${C.valid}</span></td>
              <td style="text-align:center">${C.errorCount>0?`<span class="badge badge-danger">${C.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${C.skipped?'<span class="badge badge-neutral">Dilewati</span>':C.errorCount>0&&C.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':C.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':C.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${C.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${B}">\u{1F50D} ${C.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,T.querySelectorAll(".btn-detail-error").forEach(C=>{C.addEventListener("click",()=>{let B=i[Number(C.dataset.idx)];p(B)})});let w=document.getElementById("error-detail-section"),$=document.getElementById("error-detail-container");$.innerHTML="",w.style.display="none";let _=document.getElementById("btn-start-import");y===0?(_.disabled=!0,_.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(_.disabled=!1,S>0?(_.innerHTML=`\u{1F680} Import ${y} Data Valid (${S} dilewati)`,_.title="Baris error akan dilewati, baris valid tetap diimport"):_.innerHTML=`\u{1F680} Mulai Import ${y} Data`)}function p(h){let m=document.getElementById("error-detail-section"),y=document.getElementById("error-detail-container");m.style.display="";let S=h.errors.slice(0,100).map(D=>(Array.isArray(D.errors)?D.errors:[]).map(w=>{let $=typeof w=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${D.row}</span></td>
            <td><strong>${$?w.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${$&&w.originalValue!==void 0?w.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${$?w.reason:w}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${$&&w.aliases?`Gunakan salah satu nama kolom:<br><em>${w.aliases}</em>`:$&&w.hint?w.hint:""}
            </td>
          </tr>
        `}).join("")).join("");y.innerHTML=`
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
            <tbody>${S||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${h.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,m.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{s("upload"),document.getElementById("file-info").style.display="none",n.style.display="",e=null,o.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;ia(i)?W("Log error berhasil didownload."):W("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let h=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";c(h)});async function c(h){s("importing"),a=Date.now();let m=[];Ga.forEach(w=>{let $=i?.find(_=>_.module===w&&_.mapped?.length>0);$&&m.push($)});let y=document.getElementById("import-steps-list");y.innerHTML=m.map(w=>`
      <div class="import-step-item" id="step-item-${w.module}">
        <span class="step-item-icon" id="step-icon-${w.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${w.label} <span class="step-item-count">(${w.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${w.module}"></span>
      </div>
    `).join("");let S=document.getElementById("import-bar"),D=document.getElementById("import-current-status"),T={totalSheets:m.length,totalRows:m.reduce((w,$)=>w+$.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let w=0;w<m.length;w++){let $=m[w],_=document.getElementById(`step-icon-${$.module}`),C=document.getElementById(`step-status-${$.module}`);_.textContent="\u{1F504}",C.textContent="Mengimport...",D.textContent=`Mengimport ${$.label}...`,S.style.width=`${Math.round(w/m.length*100)}%`;try{let B=await x(`/api/import/${$.module}`,{method:"POST",body:JSON.stringify({rows:$.mapped,onDuplicate:h})});if(B.ok){let E=B.data;T.inserted+=E.inserted||0,T.skipped+=E.skipped||0,T.moduleResults.push({label:$.label,inserted:E.inserted||0,skipped:E.skipped||0,status:"ok"}),_.textContent="\u2705",C.innerHTML=`<span class="badge badge-success">${E.inserted||0} berhasil</span>${E.skipped>0?` <span class="badge badge-neutral">${E.skipped} skip</span>`:""}`}else T.failed++,T.moduleResults.push({label:$.label,inserted:0,skipped:0,status:"error",error:B.data?.error}),_.textContent="\u274C",C.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(B){T.failed++,T.moduleResults.push({label:$.label,inserted:0,skipped:0,status:"error",error:B.message}),_.textContent="\u274C",C.innerHTML='<span class="badge badge-danger">Gagal</span>'}await ze(150)}S.style.width="100%",D.textContent="Selesai!",await ze(400),f(T)}function f(h){s("summary");let m=((Date.now()-a)/1e3).toFixed(1),y=h.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${y?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${y?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
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
          ${h.moduleResults.map(S=>`
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,i=null,o.value="",document.getElementById("file-info").style.display="none",n.style.display="",s("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function ze(t){return new Promise(e=>setTimeout(e,t))}A();var st=[],la=[];async function oa(t){st=await K(),la=await Z(),P({container:t,title:"Data SP (Surat Peringatan)",icon:"\u{1F4DC}",apiPath:"/api/sp",enableMobileFilterSheet:!0,itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:st}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await x(`/api/sp?limit=10000&${i}`);if(a.ok){let r=a.data.data.map(l=>({"Nama Karyawan":l.employee_name||"",Divisi:l.division||"",Cabang:l.branch_name||"","Tanggal Sp":l.tanggal||"","Akhir Sp":l.akhir_sp||"","Jenis Sp":l.sp_type||"","Link Document / Foto":l.document_link||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(j(),re));s(r,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(j(),re));i(e,"Template_Import_SP")},onImport:async e=>{let i=l=>{if(!l)return null;let o=String(l||"").toLowerCase(),n=st.find(d=>String(d.label||"").toLowerCase()===o);return n?n.value:null},a=l=>{if(!l)return"";if(l instanceof Date&&!isNaN(l.getTime()))return l.toISOString().slice(0,10);let o=String(l).trim();if(/^\d{4,5}$/.test(o)){let d=Number(o);if(d>2e4&&d<99999){let u=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(o))return o.slice(0,10);let n=o.split(/[\/\-\.]/);if(n.length===3){let[d,u,g]=n.map(p=>p.trim());if(d.length===4&&u.length<=2&&g.length<=2)return`${d}-${u.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&u.length<=2&&d.length<=2)return`${g}-${u.padStart(2,"0")}-${d.padStart(2,"0")}`}return o},r=e.map(l=>({employee_name:String(l["Nama Karyawan"]||"").trim(),division:String(l.Divisi||"").trim(),branch_id:i(String(l.Cabang||"").trim()),tanggal:a(l["Tanggal Sp"]),akhir_sp:a(l["Akhir Sp"]),sp_type:String(l["Jenis Sp"]||"").trim(),document_link:String(l["Link Document / Foto"]||"").trim()})).filter(l=>l.employee_name&&l.branch_id),s=await x("/api/import/sp",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:[{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:la},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:st,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}A();var Ae=[],sa=[];async function da(t){Ae=await K(),sa=await Z(),P({container:t,title:"Data Mutasi",icon:"\u{1F501}",apiPath:"/api/mutasi",enableMobileFilterSheet:!0,itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:Ae},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:Ae}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await x(`/api/mutasi?limit=10000&${i}`);if(a.ok){let r=a.data.data.map(l=>({Tanggal:l.tanggal||"","Nama Karyawan":l.employee_name||"","Cabang Asal":l.from_branch_name||"","Cabang Tujuan":l.to_branch_name||"",Status:l.status||"",Dokumen:l.document_link||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(j(),re));s(r,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(j(),re));i(e,"Template_Import_Mutasi")},onImport:async e=>{let i=l=>{if(!l)return null;let o=String(l||"").toLowerCase(),n=Ae.find(d=>String(d.label||"").toLowerCase()===o);return n?n.value:null},a=l=>{if(!l)return"";if(l instanceof Date&&!isNaN(l.getTime()))return l.toISOString().slice(0,10);let o=String(l).trim();if(/^\d{4,5}$/.test(o)){let d=Number(o);if(d>2e4&&d<99999){let u=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(o))return o.slice(0,10);let n=o.split(/[\/\-\.]/);if(n.length===3){let[d,u,g]=n.map(p=>p.trim());if(d.length===4&&u.length<=2&&g.length<=2)return`${d}-${u.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&u.length<=2&&d.length<=2)return`${g}-${u.padStart(2,"0")}-${d.padStart(2,"0")}`}return o},r=e.map(l=>({tanggal:a(l.Tanggal),employee_name:String(l["Nama Karyawan"]||"").trim(),from_branch_id:i(String(l["Cabang Asal"]||"").trim()),to_branch_id:i(String(l["Cabang Tujuan"]||"").trim()),status:String(l.Status||"").trim(),document_link:String(l.Dokumen||"").trim()})).filter(l=>l.tanggal&&l.employee_name&&l.from_branch_id&&l.to_branch_id),s=await x("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:sa},{type:"select",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Ae,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Ae,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}A();var _t=[],ca=[];async function pa(t){_t=await K(),ca=await Ce(),P({container:t,title:"Data Lembur",icon:"\u23F1\uFE0F",apiPath:"/api/overtime",enableMobileFilterSheet:!0,itemLabel:"Lembur",bulkDelete:!0,columns:[{key:"date",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"branch_name",label:"Cabang Lembur"},{key:"employee_name",label:"Nama Karyawan"},{key:"start_time",label:"Jam Mulai"},{key:"end_time",label:"Jam Selesai"},{key:"break_hours",label:"Istirahat",render:e=>e!=null?`${e} Jam`:"0 Jam"},{key:"total_hours",label:"Total Jam",render:e=>e!=null?`${e} Jam`:"0 Jam"},{key:"reason",label:"Alasan / Ket Lembur"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:_t}],exportOptions:{moduleName:"overtime_records",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await x(`/api/overtime?limit=10000&${i}`);if(a.ok){let r=a.data.data.map(l=>({Tanggal:l.date||"","Cabang Lembur":l.branch_name||"","Nama Karyawan":l.employee_name||"","Jam Mulai":l.start_time||"","Jam Selesai":l.end_time||"","Istirahat (Jam)":l.break_hours||0,"Total Jam":l.total_hours||0,"Alasan / Ket Lembur":l.reason||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(j(),re));s(r,`Data_Lembur_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")}},formFields:[{type:"date",name:"date",label:"Tanggal",required:!0},{type:"select",name:"branch_id",label:"Cabang Lembur",required:!0,options:_t},{type:"select",name:"employee_id",label:"Nama Karyawan",required:!0,options:ca},{type:"time",name:"start_time",label:"Jam Mulai",required:!0},{type:"time",name:"end_time",label:"Jam Selesai",required:!0},{type:"number",name:"break_hours",label:"Istirahat (Jam)",required:!0},{type:"number",name:"total_hours",label:"Total Jam (Jam)",required:!0},{type:"select",name:"reason",label:"Alasan / Ket Lembur",required:!0,options:["Cover manpower","Longshift","Pekerjaan urgent","General Cleaning","Deep Cleaning"]}]})}window._overtimeListenerAttached||(window._overtimeListenerAttached=!0,document.body.addEventListener("input",t=>{if(t.target.name==="start_time"||t.target.name==="end_time"||t.target.name==="break_hours"){let e=t.target.closest("form");if(!e)return;let i=e.querySelector('[name="start_time"]'),a=e.querySelector('[name="end_time"]'),r=e.querySelector('[name="break_hours"]'),s=e.querySelector('[name="total_hours"]');if(!i||!a||!r||!s)return;let l=i.value,o=a.value,n=parseFloat(r.value)||0;if(l&&o){let[d,u]=l.split(":").map(Number),[g,p]=o.split(":").map(Number),c=g*60+p-(d*60+u);c<0&&(c+=24*60);let f=c/60-n;f<0&&(f=0),s.value=Number.isInteger(f)?f:f.toFixed(2)}}}));A();async function ma(t){let e=ie();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}P({container:t,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:i=>new Date(i).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(i,a)=>`<strong>${i||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${a.user_role||""}</small>`},{key:"action",label:"Aksi",render:i=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[i]||"badge-neutral"}">${i}</span>`},{key:"module",label:"Modul",render:i=>`<span style="text-transform:capitalize">${(i||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(i,a)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${i}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async i=>{try{let s=((await(await fetch(`/api/audit-logs?search=${i}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(d=>String(d.id)===String(i));if(!s)return alert("Data tidak ditemukan");let l=d=>{if(!d)return"Tidak ada data";try{return JSON.stringify(JSON.parse(d),null,2)}catch{return d}},o=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${l(s.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${l(s.new_data)}</pre>
           </div>
         </div>
       `,{createModal:n}=await Promise.resolve().then(()=>(_e(),gt));n({title:`Detail Audit Log #${i}`,content:o,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let i=Math.floor(Number(t)-25569);return new Date(i*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let i=t.split(/[\/\-]/);return`${i[2]}-${i[1]}-${i[0]}`}let e=t.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(e){let i=e[1],a=parseInt(e[2],10),r=parseInt(e[3],10);if(a>12&&r<=12)return`${i}-${e[3]}-${e[2]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let i=e.split("-");if(i.length===3&&i[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],r=parseInt(i[2],10),s=a[parseInt(i[1],10)-1];return`${r} ${s} ${i[0]}`}return e};function G(t){return async e=>{if(!Oe()){xe("/login");return}return t(e)}}var Qe=null;function za(){Qe&&clearInterval(Qe);let t=()=>{let e=new Date,i=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),r=document.getElementById("header-clock-time"),s=document.getElementById("header-clock-date");r&&(r.textContent=i),s&&(s.textContent=a)};t(),Qe=setInterval(t,1e3)}async function Qa(){try{let t=await x("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},i=(a,r)=>{let s=document.getElementById(a);s&&(s.textContent=r>0?r:"",s.style.display=r>0?"inline-flex":"none")};i("badge-issues",e.issues?.current||0),i("badge-contracts",e.expiring30?.current||0),i("badge-oo1",e.one_on_one?.current||0),i("badge-schedule",e.schedule?.current||0),i("badge-supply",e.supply?.current||0)}catch{}}var Ne=[];async function Va(){try{let t=await x("/api/dashboard/notifications");if(!t.ok)return;Ne=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Ne.length>0?"block":"none",e.textContent=Ne.length)}catch{}}function Ya(){if(!Ne.length){de({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,i)=>i()});return}let t=`
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
  `;de({title:`Notifikasi (${Ne.length})`,content:t,confirmText:"Tutup",onConfirm:(e,i)=>i()})}function ua(){let t=ie(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
  `;let i=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),r=document.getElementById("topbar-menu-btn"),s=document.getElementById("sidebar-close"),l=()=>{i.classList.add("open"),a.classList.add("show")},o=()=>{i.classList.remove("open"),a.classList.remove("show")};r?.addEventListener("click",l),s?.addEventListener("click",o),a?.addEventListener("click",o),document.querySelectorAll(".nav-item").forEach(d=>d.addEventListener("click",o)),t&&t.role==="input_lembur"&&(document.querySelectorAll(".nav-item").forEach(d=>{d.dataset.route!=="/overtime"&&(d.style.display="none")}),document.querySelectorAll(".nav-section").forEach(d=>{Array.from(d.querySelectorAll(".nav-item")).some(g=>g.style.display!=="none")||(d.style.display="none")}));function n(){let d=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(p=>{let c=p.dataset.route;p.classList.toggle("active",d===c||c!=="/dashboard"&&d.startsWith(c))});let u=document.getElementById("topbar-title"),g=document.querySelector(".nav-item.active .nav-label");u&&g&&(u.textContent=g.textContent)}window.addEventListener("hashchange",n),n(),za(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await x("/api/auth/logout",{method:"POST"}),Ke(),Qe&&clearInterval(Qe),xe("/login")}),Qa(),Va(),document.getElementById("btn-notif")?.addEventListener("click",d=>{d.preventDefault(),Ya()})}async function Wa(){J("/login",({main:e})=>Nt(e)),J("/dashboard",G(({main:e})=>Bt(e))),J("/calendar",G(({main:e})=>Zt(e))),J("/employees",G(({main:e,params:i})=>Mt(e,i))),J("/contracts",G(({main:e,params:i})=>it(e,i))),J("/sp",G(({main:e})=>oa(e))),J("/mutasi",G(({main:e})=>da(e))),J("/overtime",G(({main:e})=>pa(e))),J("/timeline",G(({main:e,params:i})=>Kt(e,i))),J("/issues",G(({main:e,params:i})=>Rt(e,i))),J("/one-on-one",G(({main:e,params:i})=>jt(e,i))),J("/training",G(({main:e})=>qt(e))),J("/relievers",G(({main:e,params:i})=>Jt(e,i))),J("/reports/inspection",G(({main:e})=>Ut(e))),J("/reports/cleaning",G(({main:e})=>Gt(e))),J("/reports/fogging",G(({main:e})=>zt(e))),J("/reports/basecamp",G(({main:e})=>Qt(e))),J("/reports/supply",G(({main:e})=>wt(e,"supply"))),J("/sop",G(({main:e})=>Vt(e))),J("/checklist",G(({main:e})=>Yt(e))),J("/forms",G(({main:e})=>wt(e))),J("/users",G(({main:e})=>Wt(e))),J("/branches",G(({main:e})=>Xt(e))),J("/profile",G(({main:e})=>ea(e))),J("/settings/import",G(({main:e})=>ra(e))),J("/audit-logs",G(({main:e})=>ma(e)));let t=Oe();if(!t&&window.location.hash!=="#/login"&&xe("/login"),t){let e=await x("/api/auth/me");if(e.ok){Re(e.data.data),ua();let i=ie(),a=window.location.hash;i&&i.role==="input_lembur"&&(a==="#/dashboard"||a===""||a==="#/")&&xe("/overtime")}else Ke(),xe("/login")}window.addEventListener("fm:login",()=>{ua();let e=ie();e&&e.role==="input_lembur"?xe("/overtime"):xe("/dashboard")}),$t()}Wa();
