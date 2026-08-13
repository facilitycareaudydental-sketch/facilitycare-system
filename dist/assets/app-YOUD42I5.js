var ma=Object.defineProperty;var ot=(t,e)=>()=>(t&&(e=t(t=0)),e);var st=(t,e)=>{for(var i in e)ma(t,i,{get:e[i],enumerable:!0})};var _e={};st(_e,{API:()=>wt,CLIENT_SIDE_MAX_ROWS:()=>Se,IS_DEVELOPMENT:()=>Ne,apiFetch:()=>x,clearToken:()=>Oe,getToken:()=>Me,getUser:()=>we,setToken:()=>ct,setUser:()=>Re});function Me(){return localStorage.getItem("fm_token")}function ct(t){localStorage.setItem("fm_token",t)}function Oe(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function we(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Re(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function x(t,e={}){let i=Me(),a={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...e.headers||{}};try{let r=`cb=${Date.now()}`,s=t.includes("?")?"&":"?",o=`${wt}${t}${s}${r}`,l=await fetch(o,{...e,headers:a}),n;try{let d=await l.text();try{n=JSON.parse(d)}catch{n={error:`Server Error (${l.status}): ${d.substring(0,80)}...`}}}catch{n={error:"Gagal membaca respon dari server"}}return l.status===401&&(Oe(),window.location.hash="#/login"),{ok:l.ok,status:l.status,data:n}}catch(r){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${r.message})`}}}}var Ne,ua,wt,Se,N=ot(()=>{Ne=!1,ua="https://fm-operations-api.facilitycare-audydental.workers.dev",wt=ua,Se=1e4});var Ct={};st(Ct,{confirmDialog:()=>je,createModal:()=>se});function se({title:t,content:e,onConfirm:i,onCancel:a,confirmText:r="Simpan",cancelText:s="Batal",size:o="md",confirmClass:l="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},d=document.createElement("div");d.className="modal-overlay",d.innerHTML=`
    <div class="modal" style="max-width:${n[o]||n.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${s}</button>
        ${i?`<button class="btn ${l} modal-confirm">${r}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&d.querySelector(".modal-body").appendChild(e);let h=()=>{d.classList.remove("show"),setTimeout(()=>d.remove(),250)};return d.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),h()}),d.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),h()}),i&&d.querySelector(".modal-confirm").addEventListener("click",()=>i(d,h)),d.addEventListener("click",p=>{p.target===d&&(a&&a(),h())}),document.body.appendChild(d),requestAnimationFrame(()=>d.classList.add("show")),{overlay:d,close:h}}function je(t,e,i="Konfirmasi"){return se({title:i,content:`<p>${t}</p>`,onConfirm:(a,r)=>{e(),r()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var xe=ot(()=>{});var le={};st(le,{downloadExcel:()=>P,parseExcel:()=>qe,renderExcelButtons:()=>ba});function qe(t){return new Promise((e,i)=>{let a=new FileReader;a.onload=r=>{try{let s=new Uint8Array(r.target.result),o=XLSX.read(s,{type:"array"}),l=o.SheetNames[0],n=o.Sheets[l];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${o.SheetNames.join(", ")}`),console.log(`Sheet Used: ${l}`);let d=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),h=d.e.r-d.s.r+1,p=d.e.c-d.s.c+1;console.log(`Total Rows (including empty): ${h}`),console.log(`Total Columns: ${p}`);let c=[];for(let y=d.s.c;y<=d.e.c;++y){let b=n[XLSX.utils.encode_cell({c:y,r:d.s.r})];b&&b.v&&c.push(b.v)}console.log(`Headers Found: ${c.join(", ")}`),console.log("---------------------------");let u=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(u,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(u,"__headers",{value:c,enumerable:!1}),e(u)}catch(s){i(s)}},a.onerror=r=>i(r),a.readAsArrayBuffer(t)})}function P(t,e){try{let i=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,i,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function ba(t){return`
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
  `}var K=ot(()=>{});N();var dt={},Qe=null;function U(t,e){dt[t]=e}function Ce(t){window.location.hash=t}function xt(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[i,...a]=e.split("?"),r=dt[i];if(!r){for(let[o,l]of Object.entries(dt))if(o.endsWith("/*")&&i.startsWith(o.slice(0,-2))){r=l;break}}Qe&&(Qe(),Qe=null);let s=document.getElementById("main-content");if(s&&(s.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),r){let o=new URLSearchParams(a.join("?")),l=i.split("/").filter(Boolean),n=await r({path:i,params:o,segments:l,main:s});n&&(Qe=n)}else{let o=s||document.getElementById("app");o&&(o.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var Ke;function ga(){return Ke||(Ke=document.createElement("div"),Ke.id="toast-container",document.body.appendChild(Ke)),Ke}function _t(t,e="info",i=3500){let a=ga(),r=document.createElement("div");r.className=`toast toast-${e}`;let s={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};r.innerHTML=`<span class="toast-icon">${s[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(r),requestAnimationFrame(()=>r.classList.add("show")),setTimeout(()=>{r.classList.remove("show"),setTimeout(()=>r.remove(),350)},i)}var Y=t=>_t(t,"success"),Q=t=>_t(t,"error");xe();N();N();xe();function Ve({columns:t,data:e,onEdit:i,onDelete:a,onView:r,actions:s=[],emptyText:o="Tidak ada data",bulkSelect:l=null}){let n=document.createElement("div");if(n.className="table-wrapper",!e||e.length===0)return n.innerHTML=`<div class="empty-state"><p>${o}</p></div>`,n;let d=document.createElement("table");d.className="data-table";let h=document.createElement("thead"),p=document.createElement("tr");if(l){let u=document.createElement("th");u.style.width="40px",u.style.textAlign="center";let y=document.createElement("input");y.type="checkbox",y.id="select-all-checkbox",y.title="Pilih semua",y.addEventListener("change",()=>{e.forEach(b=>{y.checked?l.selectedIds.add(b.id):l.selectedIds.delete(b.id)}),n.querySelectorAll(".row-checkbox").forEach(b=>b.checked=y.checked),l.onToggle()}),u.appendChild(y),p.appendChild(u)}if(t.forEach(u=>{let y=document.createElement("th");y.textContent=u.label,u.width&&(y.style.width=u.width),p.appendChild(y)}),i||a||r||s.length>0){let u=document.createElement("th");u.textContent="Aksi",u.style.width="120px",p.appendChild(u)}h.appendChild(p),d.appendChild(h);let c=document.createElement("tbody");return e.forEach(u=>{let y=document.createElement("tr");if(l){let b=document.createElement("td");b.style.textAlign="center",b.style.width="40px";let m=document.createElement("input");m.type="checkbox",m.className="row-checkbox",m.checked=l.selectedIds.has(u.id),m.addEventListener("change",()=>{if(m.checked)l.selectedIds.add(u.id);else{l.selectedIds.delete(u.id);let f=document.getElementById("select-all-checkbox");f&&(f.checked=!1)}l.onToggle()}),b.appendChild(m),y.appendChild(b)}if(t.forEach(b=>{let m=document.createElement("td");if(b.render){let f=b.render(u[b.key],u);f instanceof HTMLElement?m.appendChild(f):m.innerHTML=f||""}else m.textContent=u[b.key]!==null&&u[b.key]!==void 0&&u[b.key]!==""?u[b.key]:"";b.nowrap&&(m.style.whiteSpace="nowrap"),y.appendChild(m)}),i||a||r||s.length>0){let b=document.createElement("td");b.className="actions-cell";let m=document.createElement("div");if(m.className="btn-group",r){let f=document.createElement("button");f.className="btn btn-xs btn-ghost",f.innerHTML="\u{1F441}",f.title="Lihat",f.addEventListener("click",()=>r(u)),m.appendChild(f)}if(i){let f=document.createElement("button");f.className="btn btn-xs btn-secondary",f.innerHTML="\u270F\uFE0F",f.title="Edit",f.addEventListener("click",()=>i(u)),m.appendChild(f)}s.forEach(f=>{let w=document.createElement("button");w.className=`btn btn-xs ${f.class||"btn-ghost"}`,w.innerHTML=f.icon||f.label,w.title=f.label,w.addEventListener("click",()=>f.handler(u)),m.appendChild(w)}),b.appendChild(m),y.appendChild(b)}c.appendChild(y)}),d.appendChild(c),n.appendChild(d),n}function We({page:t,pages:e,total:i,limit:a,onPage:r}){if(e<=1)return null;let s=document.createElement("div");s.className="pagination";let o=document.createElement("span");o.className="pagination-info",o.textContent=`Total: ${i} data`,s.appendChild(o);let l=document.createElement("div");l.className="pagination-btns";let n=(p,c,u=!1,y=!1)=>{let b=document.createElement("button");b.className=`btn btn-sm ${y?"btn-primary":"btn-ghost"} pagination-btn`,b.textContent=p,b.disabled=u,b.addEventListener("click",()=>r(c)),l.appendChild(b)};n("\xAB",1,t===1),n("\u2039",t-1,t===1);let d=Math.max(1,t-2),h=Math.min(e,t+2);for(let p=d;p<=h;p++)n(p,p,!1,p===t);return n("\u203A",t+1,t===e),n("\xBB",e,t===e),s.appendChild(l),s}xe();function He(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${He(e.fields)}</div>`;let i=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",r="";switch(e.type){case"textarea":r=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${i} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let o=(e.options||[]).map(p=>{let c=typeof p=="object"?p.value:p,u=typeof p=="object"?p.label:p,y=e.value==c?"selected":"";return`<option value="${c}" ${y}>${u}</option>`}).join("");r=`<select name="${e.name}" class="form-control" ${i}><option value="">-- Pilih ${e.label||""} --</option>${o}</select>`;break;case"combobox":let l=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,n=(e.options||[]).map(p=>{let c=typeof p=="object"?p.value:p,u=typeof p=="object"?p.label||p.value||"":p||"";return(u==="undefined"||u==="[object Object]"||u==="null")&&(u=""),u?`<option value="${u}"></option>`:""}).join(""),d=e.value||"";if(e.value){let p=(e.options||[]).find(c=>(typeof c=="object"?c.value:c)==e.value);if(p){let c=typeof p=="object"?p.label||p.value||"":p||"";c&&c!=="undefined"&&c!=="[object Object]"&&c!=="null"&&(d=c)}}r=`
          <input type="text" name="${e.name}" list="${l}" class="form-control" value="${d}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${l}">${n}</datalist>
        `;break;case"checkbox":r=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let h=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";r=`<input type="date" name="${e.name}" class="form-control" value="${h}" ${i}>`;break;case"number":r=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${i}>`;break;case"email":r=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i}>`;break;case"url":r=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${i}>`;break;default:r=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i} autocomplete="off">`}let s=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${r}${s}</div>`}).join("")}function Ye(t){let e={},i=new FormData(t);for(let[a,r]of i.entries())e[a]=r===""?null:r;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function Xe(t,e){e&&Object.entries(e).forEach(([i,a])=>{let r=t.querySelector(`[name="${i}"]`);r&&(r.hasAttribute("list")||(r.type==="checkbox"?r.checked=!!a:r.type==="date"&&a&&window.parseFlexibleDate?r.value=window.parseFlexibleDate(a):r.value=a??""))})}K();var fe={},$t={on(t,e){fe[t]||(fe[t]=new Set),fe[t].add(e)},off(t,e){fe[t]&&fe[t].delete(e)},emit(t,e){fe[t]&&fe[t].forEach(i=>{try{i(e)}catch(a){console.warn("[calendarBus] Handler error:",a)}})},clear(){Object.keys(fe).forEach(t=>delete fe[t])}},ha=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function me(t){if(!t){$t.emit("data:changed",{module:"unknown"});return}let e=String(t).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");$t.emit("data:changed",{module:e,relevant:ha.has(e)})}function F({container:t,title:e,icon:i,apiPath:a,columns:r,formFields:s,filterFields:o,defaultFilters:l={},itemLabel:n="Data",canCreate:d=!0,canEdit:h=!0,canDelete:p=!0,onBeforeSubmit:c,onAfterLoad:u,onDataLoaded:y,extraActions:b=[],initialSearch:m="",exportOptions:f=null,bulkDelete:w=!1,paginationMode:C="server"}){let S=1,_={...l};m&&(_.search=m);let E=new Set;t.innerHTML=`
    ${w?`
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
        ${f?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${f.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${f.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${f.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${f.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
        ${o.filter(g=>g.type==="search").map(g=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${g.placeholder||"Cari..."}" id="filter-search" value="${_.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${o.filter(g=>g.type!=="search").map(g=>{if(g.type==="select"||g.type==="combobox"){let k=(g.label||"").startsWith("Pilih")?g.label:`Pilih ${g.label||""}`;return`<select class="filter-select" name="${g.name}" id="filter-${g.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${k}</option>${(g.options||[]).map(v=>`<option value="${typeof v=="object"?v.value:v}" ${_[g.name]===(typeof v=="object"?v.value:v)?"selected":""}>${typeof v=="object"?v.label:v}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(g=>g.type!=="search").map(g=>{if(g.type==="select"||g.type==="combobox"){let k=(g.label||"").startsWith("Pilih")?g.label:`Pilih ${g.label||""}`;return`<select class="filter-select filter-select-sheet" name="${g.name}-sheet" id="filter-sheet-${g.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${k}</option>${(g.options||[]).map(v=>`<option value="${typeof v=="object"?v.value:v}" ${_[g.name]===(typeof v=="object"?v.value:v)?"selected":""}>${typeof v=="object"?v.label:v}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function $(){let g=document.getElementById("bulk-toolbar");if(!g)return;let k=document.getElementById("bulk-count"),v=document.getElementById("btn-bulk-delete"),O=document.getElementById("btn-bulk-cancel");k.textContent=`${E.size} item dipilih`,E.size>0?(g.style.display="flex",v.disabled=!1,O.disabled=!1):(g.style.display="none",v.disabled=!0,O.disabled=!0);let A=document.getElementById("select-all-checkbox");if(A){let q=document.querySelectorAll(".row-checkbox");if(q.length>0){let D=[...q].every(ie=>ie.checked),B=[...q].some(ie=>ie.checked);A.checked=D,A.indeterminate=B&&!D}else A.checked=!1,A.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{E.clear(),document.querySelectorAll(".row-checkbox").forEach(k=>k.checked=!1);let g=document.getElementById("select-all-checkbox");g&&(g.checked=!1),$()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(E.size===0)return;let g=[...E],k=document.createElement("div");k.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",k.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${g.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${g.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(k),k.querySelector("#bulk-cancel-btn").addEventListener("click",()=>k.remove()),k.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let v=k.querySelector("#bulk-confirm-btn");v.disabled=!0,v.textContent="Menghapus...";let O=await x(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:g})});k.remove(),O.ok?(Y(`${g.length} ${n} berhasil dihapus.`),E.clear(),$(),me(a),M()):Q(O.data?.error||"Gagal menghapus data.")})});let T=document.getElementById("filter-search"),L;T?.addEventListener("input",g=>{clearTimeout(L),L=setTimeout(()=>{_.search=g.target.value,S=1,E.clear(),$(),M()},400)}),o?.forEach(g=>{(g.type==="select"||g.type==="combobox")&&(document.getElementById(`filter-${g.name}`)?.addEventListener("change",k=>{_[g.name]=k.target.value;let v=document.getElementById(`filter-sheet-${g.name}`);v&&(v.value=k.target.value),S=1,E.clear(),$(),M()}),document.getElementById(`filter-sheet-${g.name}`)?.addEventListener("change",k=>{_[g.name]=k.target.value;let v=document.getElementById(`filter-${g.name}`);v&&(v.value=k.target.value),S=1,E.clear(),$(),M(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{_={},T&&(T.value=""),o?.forEach(g=>{let k=document.getElementById(`filter-${g.name}`);k&&(k.value="");let v=document.getElementById(`filter-sheet-${g.name}`);v&&(v.value="")}),S=1,E.clear(),$(),M()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{_={},T&&(T.value=""),o?.forEach(g=>{let k=document.getElementById(`filter-${g.name}`);k&&(k.value="");let v=document.getElementById(`filter-sheet-${g.name}`);v&&(v.value="")}),S=1,E.clear(),$(),M(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>ke(null)),f&&document.addEventListener("click",function(g){let k=document.getElementById("aksi-menu-main"),v=document.getElementById("btn-aksi-main");k&&v&&!v.contains(g.target)&&!k.contains(g.target)&&k.classList.remove("show-aksi-menu")});let I=document.getElementById("btn-mobile-filter"),G=document.getElementById("filter-options-wrapper"),ne=document.getElementById("btn-close-filter-sheet");if(I&&G&&(I.addEventListener("click",g=>{g.preventDefault(),G.classList.add("sheet-open")}),ne&&ne.addEventListener("click",g=>{g.preventDefault(),G.classList.remove("sheet-open")})),f){document.getElementById(`btn-export-${f.moduleName}`)?.addEventListener("click",async k=>{let v=k.target,O=v.innerHTML;v.innerHTML="\u23F3 Loading...",v.disabled=!0;try{await f.onExport()}catch{Q("Gagal export data")}finally{v.innerHTML=O,v.disabled=!1}}),document.getElementById(`btn-template-${f.moduleName}`)?.addEventListener("click",()=>{f.onTemplate()});let g=document.getElementById(`input-import-${f.moduleName}`);g?.addEventListener("change",async k=>{let v=k.target.files[0];if(!v)return;let O=document.getElementById(`label-import-${f.moduleName}`),A=O?O.querySelector(".import-text"):null,q=A?A.innerText:"";A&&(A.innerText="\u231B Memproses..."),O&&(O.style.pointerEvents="none"),g.disabled=!0;try{let D=await qe(v);if(D.length===0)throw new Error("File kosong atau format salah");await f.onImport(D),Y("Import berhasil!"),me(a),M()}catch(D){Q(D.message||"Gagal import data")}finally{A&&(A.innerText=q),O&&(O.style.pointerEvents="auto"),g.disabled=!1,g.value=""}})}async function M(){$();let g=document.getElementById("table-container");if(!g)return;g.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let k=C==="client",v=k?1:S,O=k?Se:20,A=new URLSearchParams({page:v,limit:O,...Object.fromEntries(Object.entries(_).filter(([,H])=>H))}),q=await x(`${a}?${A}`);if(!q.ok){g.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${q.data?.error||"Error"}</p></div>`;return}let D=q.data?.data||q.data||[],B=q.data?.pagination,ie=D.length;if(k){D=y(D);let H=D.length,z=20,ee=Math.ceil(H/z);S>ee&&ee>0&&(S=ee);let R=(S-1)*z,re=S*z;D=D.slice(R,re),B={page:S,limit:z,total:H,pages:ee}}!1,u&&u(D);let be=Ve({columns:r,data:D,onEdit:h?H=>ke(H):null,actions:b.map(H=>({...H,handler:z=>H.handler(z,M)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:w?{selectedIds:E,onToggle:$}:null});g.innerHTML="",g.appendChild(be);let oe=document.getElementById("pagination-container");if(oe&&(oe.innerHTML="",B&&B.pages>1)){let H=We({page:B.page,pages:B.pages,total:B.total,limit:B.limit,onPage:z=>{S=z,M()}});H&&oe.appendChild(H)}}function ve(g){let k=typeof s=="function"?s(g):s;return He(k)}function ke(g){let k=!!g,v=document.createElement("form");if(v.noValidate=!0,v.innerHTML=ve(g),k){let A=typeof s=="function"?s(g):s;Xe(v,g)}let{close:O}=se({title:k?`Edit ${n}`:`Tambah ${n}`,content:v,size:"lg",confirmText:k?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(A,q)=>{if(!v.reportValidity())return;let D=A.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let B=Ye(v),ie=typeof s=="function"?s(g):s,be=async ee=>{for(let R of ee)if(R.type==="row")await be(R.fields);else if(R.type==="combobox"&&B[R.name]){let re=B[R.name],he=(R.options||[]).find(W=>{let te=String(typeof W=="object"?W.value:W),lt=String(typeof W=="object"?W.label:W);return te===re||lt===re});if(he)B[R.name]=typeof he=="object"?he.value:he;else if(R.createApi){let W={};W[R.createApi.field]=re,R.createApi.extra&&Object.assign(W,R.createApi.extra);let te=await x(R.createApi.path,{method:"POST",body:JSON.stringify(W)});if(te.ok&&te.data?.id)B[R.name]=te.data.id;else if(te.ok&&!te.data?.id)B[R.name]=re;else throw new Error(`Gagal membuat master data: ${te.data?.error||"Unknown error"}`)}}};try{await be(ie)}catch(ee){Q(ee.message),D.disabled=!1,D.textContent=k?"Simpan Perubahan":`Tambah ${n}`;return}c&&(B=await c(B,g));let oe=k?"PUT":"POST",H=k?`${a}/${g.id}`:a,z=await x(H,{method:oe,body:JSON.stringify(B)});z.ok?(Y(k?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),q(),me(a),M()):(Q(z.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=k?"Simpan Perubahan":`Tambah ${n}`)}})}function Ae(g){je(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let k=await x(`${a}/${g.id}`,{method:"DELETE"});k.ok?(Y(`${n} berhasil dihapus.`),me(a),M()):Q(k.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return M(),M}N();N();var $e=null,Ze=null;async function Ee(t=!1){if($e&&!t)return console.log("Employees Raw (Cache Hit)",$e.slice(0,5)),$e;let e=await x(`/api/employees?limit=${Se}&status=Aktif`);return $e=(e.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",$e.slice(0,5)),$e}async function X(t=!1){let i=(await Ee(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",i.slice(0,5)),i}async function j(t=!1){return Ze&&!t||(Ze=((await x("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),Ze}function J(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function pt(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function Te(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function mt(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function ce(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}K();function ut(t,e){let i=new Date,a=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}`;if(!(t.completion_date||t.target_date||t.opening_date||"").startsWith(a))return!1;let s=String(t.status||"").toLowerCase();if(s!=="selesai"&&s!=="completed"&&s!=="done")return!1;let o=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?o.includes("inspeksi"):e==="gcdc"?o.includes("general cleaning")||o.includes("deep cleaning"):!1}N();K();function Et(t,e){let i=String(t.status||"").toLowerCase();return e==="active"?i==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&i==="aktif":!1}N();K();function gt(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let r=new Date(a);r.setDate(a.getDate()+30);let s=new Date(t.end_date);return s.setHours(0,0,0,0),s>=a&&s<=r}return!1}N();K();function Tt(t,e){let i=String(t.status||"").toLowerCase();return e==="open"?i==="open":!1}N();function Dt(t,e){let i=String(t.status||"").toLowerCase();return e==="pending"?i==="pending":!1}var ye={};function Be(t){if(ye[t]){try{ye[t].destroy()}catch{}delete ye[t]}}function fa(){Object.keys(ye).forEach(Be)}var ue=(t,e=0)=>{let i=Number(t);return isNaN(i)||t===null||t===void 0?e:i},De=(t,e="\u2014")=>{if(t==null||t==="")return e;let i=String(t).trim();return i===""||i==="[object Object]"?e:i};function Bt(t,e,i=900){if(!t)return;let a=Math.max(0,Math.round(ue(e)));if(a===0){t.textContent="0";return}let r=Date.now(),s=()=>{let o=Math.min((Date.now()-r)/i,1),l=1-Math.pow(1-o,3);t.textContent=Math.round(l*a).toLocaleString("id-ID"),o<1?requestAnimationFrame(s):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(s)}var ya={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},va=t=>{let e=De(t,"\u2014");return`<span class="status-pill ${ya[e]||"pill-neutral"}">${e}</span>`};var de={family:"Inter",size:11},ge="#94A3B8",Ie="#F1F5F9",bt=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],ka=()=>window.innerWidth<768;function et(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:ka()?"bottom":"top",labels:{font:de,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:de,titleFont:{...de,weight:"700"}}},scales:{x:{grid:{color:Ie},ticks:{font:de,color:ge,maxRotation:0}},y:{grid:{color:Ie},ticks:{font:de,color:ge},beginAtZero:!0}},...t}}var Sa=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),wa=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function It(t=3){return Array(t).fill(0).map((e,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function ae(t,e,i=8e3){try{let a=new AbortController,r=setTimeout(()=>a.abort(),i),s=await x(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(r),!s||!s.ok)return e;let o=s.data;return o?o.data!==void 0?o.data??e:o:e}catch{return e}}function xa(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(a=>{let r=document.getElementById(a);r&&(r.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(a=>{let r=document.getElementById(a);if(r&&r.style.display==="none"){r.style.display="block";let s=r.parentElement;if(s&&!s.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent="Belum ada data",r.style.display="none",s.appendChild(o)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Lt({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Ft({}),["table-contracts","table-issues"].forEach(a=>{let r=document.getElementById(a);r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Pt(t){fa(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">


      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${Sa()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${wa()}</div>

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
          <div class="chart-canvas-wrap" style="height:200px;position:relative;margin-top:10px">
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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${It(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${It(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>ht(t)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async i=>{let a=i.target.value,r=document.getElementById("jadwal-year-label");r&&(r.textContent=a);let s=document.getElementById("skel-jadwal"),o=document.getElementById("chart-jadwal");s&&(s.style.display="block",s.style.position="absolute"),o&&(o.style.display="none");let l=await ae(`/api/dashboard/schedule-chart?year=${a}`,{},8e3);try{At(l)}catch(n){console.warn("ScheduleChart render:",n),pe("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let a=i.target.value,r=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",s=document.getElementById("skel-insp"),o=document.getElementById("chart-insp");s&&(s.style.display="block",s.style.position="absolute"),o&&(o.style.display="none");let l=await ae(r,{},8e3);try{Nt(l)}catch(n){console.warn("InspBar render:",n),pe("skel-insp","chart-insp")}}),t._skelTimeout=setTimeout(()=>xa(),5e3),await ht(t),t._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?ht(t):clearInterval(t._dashRefresh)},6e4)}async function ht(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,i,a,r,s,o,l,n,d,h,p,c,u,y]=await Promise.all([ae("/api/dashboard/kpi",{},8e3),ae("/api/dashboard/issues-trend",{},8e3),ae("/api/dashboard/issues-summary",{},8e3),ae("/api/dashboard/stats",{},8e3),ae("/api/dashboard/calendar",[],8e3),ae("/api/schedule?limit=10000",{data:[]},8e3),ae("/api/employees?limit=10000",{data:[]},8e3),ae("/api/contracts?limit=10000",{data:[]},8e3),ae("/api/issues?limit=10000",{data:[]},8e3),ae("/api/one-on-one?limit=10000",{data:[]},8e3),ae("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),ae(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),ae("/api/relievers?limit=10000",{data:[]},8e3),ae("/api/reports/fogging?limit=10000",{data:[]},8e3)]),b=document.getElementById("filter-insp-month"),m=b?b.value:"",f=m?`/api/dashboard/inspection-bar?month=${m}`:"/api/dashboard/inspection-bar",w=await ae(f,{},8e3);if(e){let C=Array.isArray(o?.data)?o.data:Array.isArray(o)?o:[];window.dashboardSchedules=C;let S=Array.isArray(l?.data)?l.data:Array.isArray(l)?l:[],_=Array.isArray(n?.data)?n.data:Array.isArray(n)?n:[],E=Array.isArray(d?.data)?d.data:Array.isArray(d)?d:[],$=Array.isArray(h?.data)?h.data:Array.isArray(h)?h:[],T=Array.isArray(u?.data)?u.data:Array.isArray(u)?u:[];window.dashboardRelievers=T;let L=Array.isArray(y?.data)?y.data:Array.isArray(y)?y:[];if(window.dashboardFogging=L,e.employees&&(e.employees.current=S.filter(I=>Et(I,"active")).length),e.contracts&&(e.contracts.current=_.filter(I=>gt(I,"active")).length),e.expiring30&&(e.expiring30={current:_.filter(I=>gt(I,"expiring30")).length}),e.issues&&(e.issues.current=E.filter(I=>Tt(I,"open")).length),e.one_on_one&&(e.one_on_one.current=$.filter(I=>Dt(I,"pending")).length),e.schedule){let I=`Q${Math.ceil((new Date().getMonth()+1)/3)}`;e.schedule.current=C.filter(G=>{if(G.period===I)return!0;if(G.target_date){let ne=G.target_date.split("-");if(ne.length>=2){let M=parseInt(ne[1],10);return M&&`Q${Math.ceil(M/3)}`===I}}return!1}).length}e.inspection_month&&(e.inspection_month.current=C.filter(I=>ut(I,"inspeksi")).length),e.cleaning_month&&(e.cleaning_month.current=C.filter(I=>ut(I,"gcdc")).length)}try{Lt(e)}catch(C){console.warn("KPI render:",C)}try{Ft(e)}catch(C){console.warn("MiniStats render:",C)}try{At(c)}catch(C){console.warn("ScheduleChart render:",C),pe("skel-jadwal","chart-jadwal")}try{_a(Array.isArray(a?.by_category)?a.by_category:[])}catch(C){console.warn("Donut render:",C),pe("skel-donut","chart-donut")}try{Ca(i)}catch(C){console.warn("Trend render:",C),pe("skel-trend","chart-trend")}try{Nt(w)}catch(C){console.warn("InspBar render:",C),pe("skel-insp","chart-insp")}try{let C=Array.isArray(r)?r:Array.isArray(r?.recent_issues)?r.recent_issues:[];Ea(C)}catch(C){console.warn("IssuesTable render:",C)}try{let C=Array.isArray(r?.expiring_contracts)?r.expiring_contracts:[];$a(p)}catch(C){console.warn("ContractsTable render:",C)}try{Ta(Array.isArray(s)?s:[])}catch(C){console.warn("Agenda render:",C)}try{Da()}catch(C){console.warn("Quick Actions render:",C)}}function Lt(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let i=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=i.map(a=>{let r=ue(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${r}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${r}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{Bt(a,parseInt(a.dataset.target)||0)})}function Ft(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let i=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,a=new Date().getFullYear(),r=String(new Date().getMonth()+1).padStart(2,"0"),s=`${a}-${r}`,o=p=>`
    <select id="${p}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.preventDefault(); event.stopPropagation();">
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
          <option value="Q1" ${i==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${i==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${i==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${i==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:t.schedule?.current,href:`#/timeline?dash_filter=period_${i.toLowerCase()}`,color:"mini-blue"},{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:o("dash-inspeksi-month"),val:t.inspection_month?.current,href:`#/timeline?dash_filter=inspeksi&month=${s}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:o("dash-gcdc-month"),val:t.cleaning_month?.current,href:`#/timeline?dash_filter=gcdc&month=${s}`,color:"mini-green"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:o("dash-reliefer-month"),val:t.reliever_completed?.current,href:`#/relievers?dash_filter=reliever&month=${s}`,color:"mini-teal"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:o("dash-fogging-month"),val:t.fogging_month?.current,href:`#/reports/fogging?dash_filter=fogging&month=${s}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=l.map(p=>`
    <a href="${p.href}" class="mini-stat ${p.color}" style="text-decoration:none" id="${p.id||""}">
      <div class="mini-stat-icon">${p.icon}</div>
      <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
        <div style="display:flex; align-items:baseline; gap:3px;">
          <div class="mini-stat-value" data-target="${ue(p.val)}">0</div>
          ${p.dropdown?p.dropdown:""}
        </div>
        <div class="mini-stat-text">${p.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(p=>Bt(p,parseInt(p.dataset.target)||0,700));let n=document.getElementById("dash-jadwal-period");n&&n.addEventListener("change",p=>{let c=p.target.value,u=(window.dashboardSchedules||[]).filter(m=>{if(m.period===c)return!0;if(m.target_date){let f=m.target_date.split("-");if(f.length>=2){let w=parseInt(f[1],10);return w&&`Q${Math.ceil(w/3)}`===c}}return!1}).length,y=document.querySelector("#mini-jadwal .mini-stat-value");y&&(y.dataset.target=u,y.textContent=u);let b=document.getElementById("mini-jadwal");b&&(b.href=`#/timeline?dash_filter=period_${c.toLowerCase()}`)});let d=(p,c,u,y,b)=>{let m=document.getElementById(p);if(m){let f=w=>{let C=(u||[]).filter(E=>y(E,w)).length,S=document.querySelector(`#${c} .mini-stat-value`);S&&(S.dataset.target=C,S.textContent=C);let _=document.getElementById(c);_&&(_.href=`${b}&month=${w}`)};f(m.value),m.addEventListener("change",w=>f(w.target.value))}},h=p=>{let c=String(p.status||"").toLowerCase();return c==="done"||c==="selesai"||c==="completed"};d("dash-reliefer-month","mini-reliefer",window.dashboardRelievers,(p,c)=>window.parseFlexibleDate(p.backup_date).startsWith(c)&&h(p),"#/relievers?dash_filter=reliever"),d("dash-inspeksi-month","mini-inspeksi",window.dashboardSchedules,(p,c)=>p.activity_type==="Inspeksi Hygiene"&&h(p)&&window.parseFlexibleDate(p.completion_date||p.target_date).startsWith(c),"#/timeline?dash_filter=inspeksi"),d("dash-gcdc-month","mini-gcdc",window.dashboardSchedules,(p,c)=>(p.activity_type==="General Cleaning"||p.activity_type==="Deep Cleaning")&&h(p)&&window.parseFlexibleDate(p.completion_date||p.target_date).startsWith(c),"#/timeline?dash_filter=gcdc"),d("dash-fogging-month","mini-fogging",window.dashboardFogging,(p,c)=>h(p)&&window.parseFlexibleDate(p.activity_date).startsWith(c),"#/reports/fogging?dash_filter=fogging")}function _a(t){pe("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!e||!i)return;Be("donut");let a=(t||[]).filter(n=>ue(n.count)>0);if(!a.length){Je(e,"Belum ada data permasalahan");return}let r=a.map(n=>`${De(n.category,"Lainnya")}`),s=a.map(n=>ue(n.count)),o=s.reduce((n,d)=>n+d,0);i.innerHTML=a.map((n,d)=>{let h=bt[d%bt.length],p=o>0?Math.round(n.count/o*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${h}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${p}%)</span></div>
          <div class="donut-legend-label">${r[d]}</div>
        </div>
      </div>
    `}).join("");let l={id:"centerText",beforeDraw:function(n){let d=n.width,h=n.height,p=n.ctx;p.restore();let c=(h/80).toFixed(2);p.font="bold "+c+"em Inter",p.textBaseline="middle",p.fillStyle="#1E293B";let u=o.toString(),y=Math.round((d-p.measureText(u).width)/2),b=h/2;p.fillText(u,y,b-4),p.font="600 "+(c*.35).toFixed(2)+"em Inter",p.fillStyle="#64748B";let m="Total",f=Math.round((d-p.measureText(m).width)/2);p.fillText(m,f,b+10),p.save()}};ye.donut=new Chart(e,{type:"doughnut",data:{labels:r,datasets:[{data:s,backgroundColor:bt,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:de,titleFont:{...de,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutout:"75%"},plugins:[l]})}function Ca(t){pe("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;Be("trend"),t=t||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(t.labels||[]).map(o=>{if(!o||typeof o!="string")return"";try{let[l,n]=o.split("-");return(i[Number(n)-1]||n)+" "+String(l).slice(-2)}catch{return o}}),r=(t.open||[]).map(o=>ue(o)),s=(t.closed||[]).map(o=>ue(o));if(!a.length){Je(e,"Belum ada data trend");return}ye.trend=new Chart(e,{type:"line",data:{labels:a,datasets:[{label:"Open",data:r,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:s,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:et({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:ge,maxRotation:45,autoSkip:!0}},y:{grid:{color:Ie},ticks:{font:{family:"Inter",size:9},color:ge},beginAtZero:!0}}})})}function At(t){pe("skel-jadwal","chart-jadwal");let e=document.getElementById("chart-jadwal");if(!e)return;Be("jadwal"),t=t||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(t).some(n=>Array.isArray(n)&&n.some(d=>d>0))){Je(e,"Belum ada data jadwal");return}let r=t["Inspeksi Hygiene"]||Array(12).fill(0),s=t["General Cleaning"]||Array(12).fill(0),o=t["Deep Cleaning"]||Array(12).fill(0),l=t.Fogging||Array(12).fill(0);ye.jadwal=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Inspeksi",data:r,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:s,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:o,backgroundColor:"#F59E0B"},{label:"Fogging",data:l,backgroundColor:"#EF4444"}]},options:et({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:ge,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Ie},ticks:{font:{family:"Inter",size:9},color:ge},min:0}}})})}function Nt(t){pe("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;Be("inspBar"),t=t||{};let i=t.labels||[],a=(t.fc||[]).map(s=>ue(s)),r=(t.spv||[]).map(s=>ue(s));if(!i.length){Je(e,"Belum ada data inspeksi");return}ye.inspBar=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:r,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:et({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:de,color:ge,maxRotation:45,minRotation:30}},y:{grid:{color:Ie},ticks:{font:de,color:ge},min:0,max:100}}})})}function $a(t){pe("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;Be("contractMiniBar"),t=t||{};let i={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(o=>{let l=o.split("-")[1];return i[l]||o}),r=(t.data||[]).map(o=>ue(o));if(!a.length){Je(e,"Belum ada data");return}let s=e.getContext("2d");ye.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:r,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:et({onClick:(o,l)=>{if(l&&l.length>0){let n=l[0].index,d=(t.labels||[])[n];d&&(window.location.hash="#/contracts?month_expiry="+d)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:de,color:ge,maxRotation:0,autoSkip:!1}},y:{grid:{color:Ie,borderDash:[4,4],drawBorder:!1},ticks:{font:de,color:ge,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function Ea(t){let e=document.getElementById("table-issues");if(!e)return;let i=(t||[]).slice(0,8);if(!i.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${i.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${va(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${De(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${De(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function Ta(t){let e=document.getElementById("widget-agenda");if(!e)return;let i=new Date,a=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,s=(t||[]).filter(o=>(o.event_date||"").startsWith(a)).slice(0,10);if(!s.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${s.map(o=>{let l="#3B82F6",n="#EFF6FF",d="Agenda",h=(o.title||"").toLowerCase();return h.includes("inspeksi")?(l="#10B981",n="#ECFDF5",d="Inspeksi"):h.includes("cleaning")||h.includes("gcdc")?(l="#3B82F6",n="#EFF6FF",d="Cleaning"):h.includes("reliefer")?(l="#F59E0B",n="#FFFBEB",d="Reliefer"):h.includes("fogging")&&(l="#8B5CF6",n="#F5F3FF",d="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(o.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${l};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${De(o.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${De(o.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${l}">${d}</div>
        </div>
      `}).join("")}
    </div>
  `}function Da(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function pe(t,e){let i=document.getElementById(t),a=document.getElementById(e);if(i&&(i.style.display="none",i.style.position=""),a){a.style.display="block";let r=a.parentElement;if(r){let s=r.querySelector(".chart-empty");s&&s.remove()}}}function Je(t,e="Belum ada data"){if(!t)return;t.style.display="none";let i=t.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent=e,i.appendChild(r)}}N();async function Mt(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),i=document.getElementById("login-error"),a=document.getElementById("login-btn"),r=document.getElementById("toggle-password"),s=document.getElementById("login-password");r?.addEventListener("click",()=>{let o=s.type==="text";s.type=o?"password":"text",r.style.color=o?"":"var(--primary)"}),e?.addEventListener("submit",async o=>{o.preventDefault(),i.style.display="none";let l=e.username.value.trim(),n=e.password.value;if(!l||!n){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let d=await x("/api/auth/login",{method:"POST",body:JSON.stringify({username:l,password:n})});d.ok&&d.data.success?(ct(d.data.data.token),Re(d.data.data.user),Y("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=d.data.error||"Username atau password salah.",i.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}N();K();async function Ia(){return await j()}function Ba(t,e){let i=String(t.status||"").toLowerCase();return e==="active"?i==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&i==="aktif":!1}async function Ot(t,e){let i=await Ia(),a=e?e.get("dash_filter"):null;F({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",enableMobileFilterSheet:!0,itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:r=>a?r.filter(s=>Ba(s,a)):r,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:r=>Te(r)},{key:"phone",label:"No. HP",render:r=>r?`<a href="tel:${r}">${r}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:r=>window.formatDate(r)},{key:"status",label:"Status",render:r=>J(r)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:r=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:r?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:r?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:i,value:r?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:r?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:r?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:r?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let r=await x(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let s=r.data.data.map(o=>({"Nama Lengkap":o.full_name,Cabang:o.branch_name||"",Divisi:o.division||"","No. HP":o.phone||"","Tgl Masuk":o.join_date||"",Status:o.status||""}));P(s,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{P([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async r=>{let s=n=>{if(!n)return null;let d=String(n||"").toLowerCase(),h=i.find(p=>String(p.label||"").toLowerCase()===d);return h?h.value:null},o=r.map(n=>({full_name:String(n["Nama Lengkap"]||"").trim(),branch_id:s(String(n.Cabang||"").trim()),division:String(n.Divisi||"").trim()||"FACILITY CARE",phone:String(n["No. HP"]||"").trim(),join_date:String(n["Tgl Masuk"]||"").trim(),status:String(n.Status||"").trim(),notes:String(n.Catatan||"").trim()})).filter(n=>n.full_name),l=await x("/api/import/employees",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}N();K();var yt=[],Rt=[];async function Pa(){yt=await j(),Rt=await Ee()}var ft=async t=>{let e=[],i=1;for(;;){let r=await(await Promise.resolve().then(()=>(N(),_e))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${i}`);if(!r.ok)break;let s=r.data?.data||r.data||[],o=Array.isArray(s)?s:[];if(e=e.concat(o),o.length<100||r.data?.pagination&&i>=r.data.pagination.pages)break;i++}return e};async function tt(t,e){await Pa(),F({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",defaultFilters:{},onDataLoaded:a=>a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>Te(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,r)=>r.end_date&&String(r.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':pt(a)},{key:"status",label:"Status",render:a=>J(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:yt},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[s,o]=await Promise.all([ft("/api/employees?status=Aktif"),ft("/api/contracts")]);if(s.length>0){let l=o.filter(p=>p.status==="Aktif"),n=new Set(l.map(p=>p.employee_id)),d=s.filter(p=>!n.has(p.id)),h=`<p style="margin-bottom:12px">Data yang terbaca: <b>${s.length}</b> Karyawan Aktif, dan <b>${l.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${d.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;d.forEach(p=>{let c=o.filter(y=>y.employee_id===p.id),u='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(c.length>0){let y=c[0];u=`Pernah ada kontrak (Status: <b style="color:#EF4444">${y.status}</b>, Selesai: ${window.formatDate(y.end_date)})`}h+=`<li style="margin-bottom:8px"><b>${p.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${p.branch_name||"-"} | ${u}</span></li>`}),h+="</ul>",Promise.resolve().then(()=>(xe(),Ct)).then(p=>p.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:h,cancelText:"Tutup"}))}}catch(s){console.error(s)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let r=document.querySelector(".page-actions");r&&r.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:Rt,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:yt,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await x(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let r=a.data.data.map(s=>({"Nama Lengkap":s.employee_name,Cabang:s.branch_name||"","Div / Bagian":s.division||"","Tanggal Mulai":s.start_date||"","Tanggal Selesai":s.end_date&&String(s.end_date).startsWith("2099")?"":s.end_date||"","Sisa Kontrak":s.end_date&&String(s.end_date).startsWith("2099")?"Tetap":s.days_remaining!==null&&s.days_remaining!==void 0?`${s.days_remaining} Hari`:"",Status:s.status||""}));P(r,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{P([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[r,s]=await Promise.all([x("/api/branches?limit=10000"),ft("/api/employees")]),o=r.data?.data||[],l=s||[];console.log(`Total employee yang berhasil dimuat dari database : ${l.length}`),l.length>0&&(console.log("Contoh 5 employee pertama:"),l.slice(0,5).forEach((m,f)=>{console.log(`${f+1}. ID: ${m.id}, Name: ${m.full_name}, Status: ${m.status}`)}));let n=m=>{if(!m)return null;let f=String(m||"").replace(/\s+/g," ").toLowerCase().trim(),w=o.find(C=>String(C.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===f||String(C.code||"").replace(/\s+/g," ").toLowerCase().trim()===f||String(C.name||"").replace(/\s+/g," ").toLowerCase().trim()===f);return w?w.id:null},d=(m,f)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${f}`),console.log(`Nama dari Excel : "${m}"`),!m)return console.log("Alasan gagal mapping : Nama kosong"),null;let w=String(m||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${w}"`),console.log(`Jumlah employee di database : ${l.length}`);let C=l.find(S=>String(S.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===w);return C?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${C.id}`),C.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},h=m=>{if(!m)return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let f=String(m).trim();if(/^\d{4,5}(\.\d+)?$/.test(f)){let C=Math.floor(Number(f));if(C>2e4&&C<99999){let S=new Date(Date.UTC(1899,11,30)+C*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(f))return f.slice(0,10);let w=f.split(/[\/\-\.]/);if(w.length===3){let[C,S,_]=w.map(E=>E.trim());if(C.length===4&&S.length<=2&&_.length<=2)return`${C}-${S.padStart(2,"0")}-${_.padStart(2,"0")}`;if(_.length===4&&S.length<=2&&C.length<=2)return`${_}-${S.padStart(2,"0")}-${C.padStart(2,"0")}`}return f},p=a.map((m,f)=>{let w=f+2,C=String(m["Nama Lengkap"]||"").trim(),S=m["Tanggal Mulai"],_=h(S);if(!_){let T=a.__worksheet,L=a.__headers||[],I=L.indexOf("Tanggal Mulai"),G="N/A",ne="N/A",M="N/A";if(I!==-1&&T&&window.XLSX){let ke=window.XLSX.utils.encode_cell({c:I,r:w-1});M=ke;let Ae=T[ke];Ae?(G=Ae.t||"undefined",ne=Ae.w||"undefined"):G="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let ve="Unknown";S==null||S===""?ve="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":S instanceof Date&&isNaN(S.getTime())?ve="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":ve="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${w}`),console.log(`Employee Name : ${C}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${I})`),console.log(`Raw Cell Value : "${S}"`),console.log(`JavaScript Type : ${typeof S}`),console.log(`SheetJS Cell Type : ${G}`),console.log(`SheetJS Formatted Value : "${ne}"`),console.log(`Value After Trim : "${String(S||"").trim()}"`),console.log(`Value After Date Parser : "${_}"`),console.log(`Is Empty : ${!S}`),console.log(`Is Invalid Date : ${S instanceof Date?isNaN(S.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${ve}`),console.log(`Workbook Sheet : ${T?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${M}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(m,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(L)),console.log(`==========================
`)}let E=d(C,w),$=null;return E?_||($="Tanggal Mulai kosong atau tidak berformat tanggal"):$="Karyawan tidak ditemukan di Database",{isValid:!!(E&&_),invalidReason:$,rowNum:w,data:{employee_id:E,branch_id:n(String(m.Cabang||"").trim()),division:String(m["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:_,end_date:h(m["Tanggal Selesai"])||"2099-12-31",status:String(m.Status||"").trim(),_rawName:C}}}),c=[],u=[];if(p.forEach(m=>{m.isValid?c.push(m.data):u.push({rowNum:m.rowNum,name:m.data._rawName,reason:m.invalidReason})}),console.log(`Split Validation - Valid: ${c.length}, Invalid: ${u.length}`),c.length===0){let m=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${a.length}
Valid: 0
Invalid: ${u.length}

Daftar Kegagalan (Contoh):
`;u.slice(0,10).forEach(f=>{m+=`- Row ${f.rowNum} | Nama: ${f.name} | Alasan: ${f.reason}
`}),u.length>10&&(m+=`- ... dan ${u.length-10} lainnya.
`),alert(m);return}let y=await x("/api/contracts/import",{method:"POST",body:JSON.stringify(c)}),b=`IMPORT SUMMARY
======================
`;b+=`Total Baris Excel : ${a.length}
`,b+=`Baris Valid       : ${c.length}
`,b+=`Baris Invalid     : ${u.length}

`,y&&y.data&&y.data.metrics?(b+=`Berhasil INSERT   : ${y.data.metrics.inserted}
`,b+=`Berhasil UPDATE   : ${y.data.metrics.updated}
`):b+=`Berhasil diproses : ${c.length}
`,u.length>0&&(b+=`
DAFTAR DATA DILEWATI:
`,u.forEach(m=>{b+=`- Row ${m.rowNum} | ${m.name} | ${m.reason}
`})),alert(b),typeof tt=="function"&&tt()}}})}N();K();var vt=[],Ue=[];function La(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let i of e)if(t.some(a=>a.period===i))return i;return"Q3"}async function Kt(t,e){vt=await j();let i=await X();Ue=["Berlin Ariansyah","Ade Surahman"];let a=m=>m&&!Ue.find(f=>String(typeof f=="object"?f.value:f).toLowerCase()===String(m).toLowerCase())?[...Ue,m]:Ue,r=await x(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),s=m=>{if(!m||m==="-"||String(m).trim()==="")return"";let f=String(m).split("-");return f.length===3&&f[0].length===4?`${f[2]}-${f[1]}-${f[0]}`:m},o=r.data?.data||[],l=La(o),n=e?e.get("dash_filter"):null,d=new Date,h=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`,p={},c=e&&e.get("month")?e.get("month"):null;n==="inspeksi"?p={status:"Done",activity_type:"Inspeksi Hygiene",month:c}:n==="gcdc"?p={status:"Done",activity_type:"GCDC",month:c}:n&&n.startsWith("period_")&&(p={period:n.replace("period_","").toUpperCase()});let u=new Date().getFullYear(),b=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((m,f)=>{let w=String(f+1).padStart(2,"0");return{value:`${u}-${w}`,label:`${m} ${u}`}});F({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:p,onDataLoaded:m=>m.sort((f,w)=>{let C=f.opening_date?new Date(f.opening_date).getTime():0;return(w.opening_date?new Date(w.opening_date).getTime():0)-C}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:m=>mt(m)},{key:"period",label:"Periode",render:m=>ce(m)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:m=>s(m)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:m=>s(m)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:m=>s(m)},{key:"status",label:"Status",render:m=>J(m)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:vt},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:b},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Ue}],formFields:m=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:vt,value:m?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:m?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:m?.period},{name:"pic",label:"PIC",type:"combobox",options:a(m?.pic),value:m?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:m?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:m?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:m?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:m?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:m?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let m=await x(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(m.ok){let f=m.data.data.map(w=>({Cabang:w.branch_name||"",Kegiatan:w.activity_type||"",Periode:w.period||"",PIC:w.pic||"","Tgl Opening":w.opening_date||"","Tgl Target":w.target_date||"","Tgl Selesai":w.completion_date||"",Status:w.status||""}));P(f,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{P([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async m=>{let w=(await x("/api/branches?all=1")).data?.data||[],C=$=>{if(!$)return null;let T=String($||"").toLowerCase(),L=w.find(I=>String(I.full_name||"").toLowerCase()===T||String(I.code||"").toLowerCase()===T||String(I.name||"").toLowerCase()===T);return L?L.id:null},S=$=>{if($==null||$==="")return"";if($ instanceof Date&&!isNaN($.getTime()))return $.toISOString().slice(0,10);let T=String($).trim();if(T===""||T==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(T))return T.slice(0,10);if(/^\d{4,5}$/.test(T)){let I=Number(T);if(I>2e4&&I<99999){let G=new Date(Date.UTC(1899,11,30)+I*864e5);return isNaN(G.getTime())?"":G.toISOString().slice(0,10)}}let L=T.split(/[\/\-\.]/);if(L.length===3){let[I,G,ne]=L.map(M=>M.trim());if(I.length===4&&G.length<=2&&ne.length<=2)return`${I}-${G.padStart(2,"0")}-${ne.padStart(2,"0")}`;if(ne.length===4&&G.length<=2&&I.length<=2)return`${ne}-${G.padStart(2,"0")}-${I.padStart(2,"0")}`}return T},_=m.map($=>({branch_id:C(String($.Cabang||"").trim()),activity_type:String($.Kegiatan||"").trim(),period:String($.Periode||"").trim(),pic:String($.PIC||$.Pic||"").trim(),opening_date:S($["Tgl Opening"]||$["Tanggal Opening"]||$["Tgl Openir"]),target_date:S($["Tgl Target"]||$["Tanggal Target"]),completion_date:S($["Tgl Selesai"]||$["Tanggal Selesai"]),status:String($.Status||"").trim(),notes:String($.Catatan||$.Keterangan||"").trim()})).filter($=>$.activity_type&&$.period),E=await x("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:_,onDuplicate:"update"})});if(!E.ok)throw new Error(E.data?.error||"Import gagal");return E.data}}})}N();K();var kt=[],at=[];function Fa(t,e){let i=String(t.status||"").toLowerCase();return e==="open"?i==="open":!1}async function jt(t,e){let i=e?e.get("dash_filter"):null;kt=await j(),at=await X();let a=n=>n&&!at.find(d=>d.value===n)?[...at,{value:n,label:n}]:at,r=new Date().getFullYear(),s=["2025","2026","2027","2028","2029","2030"],l=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((n,d)=>{let h=String(d+1).padStart(2,"0");return{value:`${r}-${h}`,label:`${n} ${r}`}});F({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:n=>i?n.filter(d=>Fa(d,i)):n,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:n=>`<span class="badge badge-secondary">${n}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:n=>`<span title="${n}">${n?.length>50?n.slice(0,50)+"\u2026":n}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:n=>`<span title="${n||""}">${n?.length>40?n.slice(0,40)+"\u2026":n||"-"}</span>`},{key:"status",label:"Status",render:n=>J(n)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"day_count",label:"Hari",render:n=>n??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:kt},{type:"select",name:"month",label:"Bulan",options:l},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:s}],formFields:n=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:n?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:kt,value:n?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:n?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:n?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:n?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(n?.employee_name),value:n?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(n?.fc_specialist),value:n?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:n?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:n?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let n=await x(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let d=n.data.data.map(h=>({Tanggal:h.report_date||"",Cabang:h.branch_name||"",Kategori:h.category||"",Sumber:h.source||"",Keluhan:h.complaint||"","Nama FC":h.employee_name||"","FC Spesialis":h.fc_specialist||"",Solusi:h.solution||"","Tgl Selesai":h.completion_date||"",Status:h.status||""}));P(d,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{P([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async n=>{let h=(await x("/api/branches?all=1")).data?.data||[],p=y=>{if(!y)return null;let b=String(y||"").toLowerCase(),m=h.find(f=>String(f.full_name||"").toLowerCase()===b||String(f.code||"").toLowerCase()===b||String(f.name||"").toLowerCase()===b);return m?m.id:null},c=n.map(y=>({branch_id:p(String(y.Cabang||"").trim()),report_date:String(y.Tanggal||"").trim(),category:String(y.Kategori||"").trim(),source:String(y.Sumber||"").trim(),complaint:String(y.Keluhan||"").trim(),employee_name:String(y["Nama FC"]||"").trim(),fc_specialist:String(y["FC Spesialis"]||"").trim(),solution:String(y.Solusi||"").trim(),completion_date:String(y["Tgl Selesai"]||"").trim(),status:String(y.Status||"").trim()})).filter(y=>y.report_date&&y.complaint&&y.category),u=await x("/api/import/issues",{method:"POST",body:JSON.stringify({rows:c,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}N();var Pe=[];function Aa(t,e){let i=String(t.status||"").toLowerCase();return e==="pending"?i==="pending":!1}async function Ht(t,e){let i=e?e.get("dash_filter"):null;Pe=await j();let a=await X(),r=["Ade","Berlin"],s=l=>l&&!a.find(n=>n.value===l)?[...a,{value:l,label:l}]:a,o=l=>l&&!r.find(n=>(typeof n=="object"?n.value:n)===l)?[...r,l]:r;F({container:t,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:l=>i?l.filter(n=>Aa(n,i)):l,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:l=>`<span title="${l||""}">${l?.length>50?l.slice(0,50)+"\u2026":l||"-"}</span>`},{key:"solution",label:"Solusi",render:l=>`<span title="${l||""}">${l?.length>40?l.slice(0,40)+"\u2026":l||"-"}</span>`},{key:"status",label:"Status",render:l=>J(l)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:l=>window.formatDate(l)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:l=>l?`<a href="${l}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:Pe},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async l=>{let n=new URLSearchParams(l||{}).toString(),d=await x(`/api/one-on-one?limit=10000&${n}`);if(d.ok){let h=d.data.data.map(c=>({Tanggal:c.meeting_date||"",Cabang:c.branch_name||"","Nama Karyawan":c.employee_name||"",PIC:c.pic||"",Masalah:c.problem||"",Solusi:c.solution||"",Status:c.status||"","Tgl Selesai":c.completion_date||"",Dokumen:c.document_link||""})),{downloadExcel:p}=await Promise.resolve().then(()=>(K(),le));p(h,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let l=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(K(),le));n(l,"Template_Import_OneOnOne")},onImport:async l=>{let n=c=>{if(!c)return null;let u=String(c||"").toLowerCase(),y=Pe.find(b=>String(b.label||"").toLowerCase()===u);return y?y.value:null},d=c=>{if(!c)return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let u=String(c).trim();if(/^\d{4,5}$/.test(u)){let b=Number(u);if(b>2e4&&b<99999){let m=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);let y=u.split(/[\/\-\.]/);if(y.length===3){let[b,m,f]=y.map(w=>w.trim());if(b.length===4&&m.length<=2&&f.length<=2)return`${b}-${m.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&m.length<=2&&b.length<=2)return`${f}-${m.padStart(2,"0")}-${b.padStart(2,"0")}`}return u},h=l.map(c=>({meeting_date:d(c.Tanggal),employee_name:String(c["Nama Karyawan"]||"").trim(),branch_id:n(String(c.Cabang||"").trim()),pic:String(c.PIC||"").trim(),problem:String(c.Masalah||"").trim(),solution:String(c.Solusi||"").trim(),status:String(c.Status||"").trim(),completion_date:d(c["Tgl Selesai"]),document_link:String(c.Dokumen||"").trim()})).filter(c=>c.meeting_date&&c.employee_name&&c.branch_id),p=await x("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:h,onDuplicate:"update"})});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}},formFields:l=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:l?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:l?.branch_id&&!Pe.find(n=>n.value==l.branch_id)?[...Pe,{value:l.branch_id,label:l.branch_name||l.branch_id}]:Pe,createApi:{path:"/api/branches",field:"full_name"},value:l?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:s(l?.employee_name),value:l?.employee_name},{name:"pic",label:"PIC",type:"select",options:o(l?.pic),createApi:{path:"/api/pic",field:"name"},value:l?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:l?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:l?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:l?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:l?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:l?.document_link}]})}N();async function qt(t){let e=await j(),i=await X(),a=["Ade Surahman","Berlin Ariansyah","Mizwar","Fajar","Ade","Berlin"],r=n=>n&&!i.find(d=>d.value===n)?[...i,{value:n,label:n}]:i,s=n=>n&&!a.find(d=>(typeof d=="object"?d.value:d)===n)?[...a,n]:a,o=Array.from({length:5},(n,d)=>String(new Date().getFullYear()-d));F({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:n=>{try{let d=JSON.parse(n);return Array.isArray(d)?d.join(", "):n||"-"}catch{return n||"-"}}},{key:"score",label:"Nilai",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer / peserta..."},{type:"select",name:"batch",label:"Batch",options:["Batch 1","Batch 2","Batch 3","Batch 4","Batch 5"]},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"trainer",label:"Trainer",options:["Berlin Ariansyah"]},{type:"select",name:"year",label:"Tahun",options:o}],exportOptions:{moduleName:"training",onExport:async n=>{let d=new URLSearchParams(n||{}).toString(),h=await x(`/api/training?limit=10000&${d}`);if(h.ok){let p=h.data.data.map(u=>{let y=u.participants||"";try{let b=JSON.parse(y);y=Array.isArray(b)?b.join(", "):y}catch{}return{Tanggal:u.training_date||"",Batch:u.batch||"",Materi:u.subject||"",Cabang:u.branch_name||"",Trainer:u.trainer||"",Peserta:y,Nilai:u.score!==null&&u.score!==void 0?u.score:"",Dokumen:u.document_link||""}}),{downloadExcel:c}=await Promise.resolve().then(()=>(K(),le));c(p,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:d}=await Promise.resolve().then(()=>(K(),le));d(n,"Template_Import_Training")},onImport:async n=>{let d=u=>{if(!u)return null;let y=String(u||"").toLowerCase(),b=e.find(m=>String(m.label||"").toLowerCase()===y);return b?b.value:null},h=u=>{if(!u)return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let y=String(u).trim();if(/^\d{4,5}$/.test(y)){let m=Number(y);if(m>2e4&&m<99999){let f=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);let b=y.split(/[\/\-\.]/);if(b.length===3){let[m,f,w]=b.map(C=>C.trim());if(m.length===4&&f.length<=2&&w.length<=2)return`${m}-${f.padStart(2,"0")}-${w.padStart(2,"0")}`;if(w.length===4&&f.length<=2&&m.length<=2)return`${w}-${f.padStart(2,"0")}-${m.padStart(2,"0")}`}return y},p=n.map(u=>({training_date:h(u.Tanggal),batch:String(u.Batch||"").trim(),subject:String(u.Materi||"").trim(),branch_id:d(String(u.Cabang||"").trim()),trainer:String(u.Trainer||"").trim(),participants:String(u.Peserta||"").trim(),score:u.Nilai?Number(u.Nilai):null,document_link:String(u.Dokumen||"").trim()})).filter(u=>u.training_date&&u.subject&&u.branch_id),c=await x("/api/training/import",{method:"POST",body:JSON.stringify(p)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}},formFields:n=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:n?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:n?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:n?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:n?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:s(n?.trainer),value:n?.trainer}]},{name:"participants",label:"Peserta",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let d=JSON.parse(n?.participants);return Array.isArray(d)?d.join(", "):n?.participants||""}catch{return n?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:n?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:n?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],onBeforeSubmit:async n=>n})}N();xe();K();function Jt({container:t,title:e,icon:i,apiPath:a,columns:r,formFields:s,filterFields:o,defaultFilters:l={},itemLabel:n="Data",canCreate:d=!0,canEdit:h=!0,canDelete:p=!0,onBeforeSubmit:c,onAfterLoad:u,onDataLoaded:y,extraActions:b=[],initialSearch:m="",exportOptions:f=null,bulkDelete:w=!1,paginationMode:C="server"}){let S=1,_={...l};m&&(_.search=m);let E=new Set;t.innerHTML=`
    ${w?`
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
        ${f?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${f.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${f.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${f.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${f.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
        ${o.filter(g=>g.type==="search").map(g=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${g.placeholder||"Cari..."}" id="filter-search" value="${_.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${o.filter(g=>g.type!=="search").map(g=>{if(g.type==="select"||g.type==="combobox"){let k=(g.label||"").startsWith("Pilih")?g.label:`Pilih ${g.label||""}`;return`<select class="filter-select" name="${g.name}" id="filter-${g.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${k}</option>${(g.options||[]).map(v=>`<option value="${typeof v=="object"?v.value:v}" ${_[g.name]===(typeof v=="object"?v.value:v)?"selected":""}>${typeof v=="object"?v.label:v}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(g=>g.type!=="search").map(g=>{if(g.type==="select"||g.type==="combobox"){let k=(g.label||"").startsWith("Pilih")?g.label:`Pilih ${g.label||""}`;return`<select class="filter-select filter-select-sheet" name="${g.name}-sheet" id="filter-sheet-${g.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${k}</option>${(g.options||[]).map(v=>`<option value="${typeof v=="object"?v.value:v}" ${_[g.name]===(typeof v=="object"?v.value:v)?"selected":""}>${typeof v=="object"?v.label:v}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function $(){let g=document.getElementById("bulk-toolbar");if(!g)return;let k=document.getElementById("bulk-count"),v=document.getElementById("btn-bulk-delete"),O=document.getElementById("btn-bulk-cancel");k.textContent=`${E.size} item dipilih`,E.size>0?(g.style.display="flex",v.disabled=!1,O.disabled=!1):(g.style.display="none",v.disabled=!0,O.disabled=!0);let A=document.getElementById("select-all-checkbox");if(A){let q=document.querySelectorAll(".row-checkbox");if(q.length>0){let D=[...q].every(ie=>ie.checked),B=[...q].some(ie=>ie.checked);A.checked=D,A.indeterminate=B&&!D}else A.checked=!1,A.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{E.clear(),document.querySelectorAll(".row-checkbox").forEach(k=>k.checked=!1);let g=document.getElementById("select-all-checkbox");g&&(g.checked=!1),$()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(E.size===0)return;let g=[...E],k=document.createElement("div");k.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",k.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${g.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${g.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(k),k.querySelector("#bulk-cancel-btn").addEventListener("click",()=>k.remove()),k.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let v=k.querySelector("#bulk-confirm-btn");v.disabled=!0,v.textContent="Menghapus...";let O=await x(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:g})});k.remove(),O.ok?(Y(`${g.length} ${n} berhasil dihapus.`),E.clear(),$(),me(a),M()):Q(O.data?.error||"Gagal menghapus data.")})});let T=document.getElementById("filter-search"),L;T?.addEventListener("input",g=>{clearTimeout(L),L=setTimeout(()=>{_.search=g.target.value,S=1,E.clear(),$(),M()},400)}),o?.forEach(g=>{(g.type==="select"||g.type==="combobox")&&(document.getElementById(`filter-${g.name}`)?.addEventListener("change",k=>{_[g.name]=k.target.value;let v=document.getElementById(`filter-sheet-${g.name}`);v&&(v.value=k.target.value),S=1,E.clear(),$(),M()}),document.getElementById(`filter-sheet-${g.name}`)?.addEventListener("change",k=>{_[g.name]=k.target.value;let v=document.getElementById(`filter-${g.name}`);v&&(v.value=k.target.value),S=1,E.clear(),$(),M(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{_={},T&&(T.value=""),o?.forEach(g=>{let k=document.getElementById(`filter-${g.name}`);k&&(k.value="");let v=document.getElementById(`filter-sheet-${g.name}`);v&&(v.value="")}),S=1,E.clear(),$(),M()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{_={},T&&(T.value=""),o?.forEach(g=>{let k=document.getElementById(`filter-${g.name}`);k&&(k.value="");let v=document.getElementById(`filter-sheet-${g.name}`);v&&(v.value="")}),S=1,E.clear(),$(),M(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>ke(null)),f&&document.addEventListener("click",function(g){let k=document.getElementById("aksi-menu-main"),v=document.getElementById("btn-aksi-main");k&&v&&!v.contains(g.target)&&!k.contains(g.target)&&k.classList.remove("show-aksi-menu")});let I=document.getElementById("btn-mobile-filter"),G=document.getElementById("filter-options-wrapper"),ne=document.getElementById("btn-close-filter-sheet");if(I&&G&&(I.addEventListener("click",g=>{g.preventDefault(),G.classList.add("sheet-open")}),ne&&ne.addEventListener("click",g=>{g.preventDefault(),G.classList.remove("sheet-open")})),f){document.getElementById(`btn-export-${f.moduleName}`)?.addEventListener("click",async k=>{let v=k.target,O=v.innerHTML;v.innerHTML="\u23F3 Loading...",v.disabled=!0;try{await f.onExport()}catch{Q("Gagal export data")}finally{v.innerHTML=O,v.disabled=!1}}),document.getElementById(`btn-template-${f.moduleName}`)?.addEventListener("click",()=>{f.onTemplate()});let g=document.getElementById(`input-import-${f.moduleName}`);g?.addEventListener("change",async k=>{let v=k.target.files[0];if(!v)return;let O=document.getElementById(`label-import-${f.moduleName}`),A=O?O.querySelector(".import-text"):null,q=A?A.innerText:"";A&&(A.innerText="\u231B Memproses..."),O&&(O.style.pointerEvents="none"),g.disabled=!0;try{let D=await qe(v);if(D.length===0)throw new Error("File kosong atau format salah");await f.onImport(D),Y("Import berhasil!"),me(a),M()}catch(D){Q(D.message||"Gagal import data")}finally{A&&(A.innerText=q),O&&(O.style.pointerEvents="auto"),g.disabled=!1,g.value=""}})}async function M(){$();let g=document.getElementById("table-container");if(!g)return;g.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let k=C==="client",v=k?1:S,O=k?Se:20,A=new URLSearchParams({page:v,limit:O,...Object.fromEntries(Object.entries(_).filter(([,H])=>H))}),q=await x(`${a}?${A}`);if(!q.ok){g.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${q.data?.error||"Error"}</p></div>`;return}let D=q.data?.data||q.data||[],B=q.data?.pagination,ie=D.length;if(k){D=y(D);let H=D.length,z=20,ee=Math.ceil(H/z);S>ee&&ee>0&&(S=ee);let R=(S-1)*z,re=S*z;D=D.slice(R,re),B={page:S,limit:z,total:H,pages:ee}}!1,u&&u(D);let be=Ve({columns:r,data:D,onEdit:h?H=>ke(H):null,actions:b.map(H=>({...H,handler:z=>H.handler(z,M)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:w?{selectedIds:E,onToggle:$}:null});g.innerHTML="",g.appendChild(be);let oe=document.getElementById("pagination-container");if(oe&&(oe.innerHTML="",B&&B.pages>1)){let H=We({page:B.page,pages:B.pages,total:B.total,limit:B.limit,onPage:z=>{S=z,M()}});H&&oe.appendChild(H)}}function ve(g){let k=typeof s=="function"?s(g):s;return He(k)}function ke(g){let k=!!g,v=document.createElement("form");if(v.noValidate=!0,v.innerHTML=ve(g),k){let A=typeof s=="function"?s(g):s;Xe(v,g)}let{close:O}=se({title:k?`Edit ${n}`:`Tambah ${n}`,content:v,size:"lg",confirmText:k?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(A,q)=>{if(!v.reportValidity())return;let D=A.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let B=Ye(v),ie=typeof s=="function"?s(g):s,be=async ee=>{for(let R of ee)if(R.type==="row")await be(R.fields);else if(R.type==="combobox"&&B[R.name]){let re=B[R.name],he=(R.options||[]).find(W=>{let te=String(typeof W=="object"?W.value:W),lt=String(typeof W=="object"?W.label:W);return te===re||lt===re});if(he)B[R.name]=typeof he=="object"?he.value:he;else if(R.createApi){let W={};W[R.createApi.field]=re,R.createApi.extra&&Object.assign(W,R.createApi.extra);let te=await x(R.createApi.path,{method:"POST",body:JSON.stringify(W)});if(te.ok&&te.data?.id)B[R.name]=te.data.id;else if(te.ok&&!te.data?.id)B[R.name]=re;else throw new Error(`Gagal membuat master data: ${te.data?.error||"Unknown error"}`)}}};try{await be(ie)}catch(ee){Q(ee.message),D.disabled=!1,D.textContent=k?"Simpan Perubahan":`Tambah ${n}`;return}c&&(B=await c(B,g));let oe=k?"PUT":"POST",H=k?`${a}/${g.id}`:a,z=await x(H,{method:oe,body:JSON.stringify(B)});z.ok?(Y(k?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),q(),me(a),M()):(Q(z.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=k?"Simpan Perubahan":`Tambah ${n}`)}})}function Ae(g){je(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let k=await x(`${a}/${g.id}`,{method:"DELETE"});k.ok?(Y(`${n} berhasil dihapus.`),me(a),M()):Q(k.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return M(),M}N();K();async function Ut(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await j(),a=await X(),r=e?e.get("dash_filter"):null,s={};if(r==="reliever"){let c=new Date,u=`${c.getFullYear()}-${String(c.getMonth()+1).padStart(2,"0")}`;s={status:"Done",month:e&&e.get("month")?e.get("month"):u}}console.log("RAW",await Ee()),console.log("OPTIONS",a);let o=c=>c&&!a.find(u=>u.value===c)?[...a,{value:c,label:c}]:a,l=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],n=c=>c&&!l.includes(c)?[...l,c]:l,d=new Date().getFullYear(),p=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((c,u)=>{let y=String(u+1).padStart(2,"0");return{value:`${d}-${y}`,label:`${c} ${d}`}});Jt({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:s,onDataLoaded:c=>c.sort((u,y)=>{let b=u.backup_date?new Date(u.backup_date).getTime():0;return(y.backup_date?new Date(y.backup_date).getTime():0)-b}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:c=>ce(c)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:c=>window.formatDate(c)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:c=>window.formatDate(c)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:c=>c?`<span class="badge badge-info">${c}</span>`:"-"},{key:"status",label:"Status",render:c=>J(c)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:l},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"month",label:"Bulan",options:p},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:c=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:c?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:c?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:o(c?.original_fc_name),value:c?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:n(c?.reliever_name),value:c?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:c?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:c?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:c?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:c?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:c?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let c=await x(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(c.ok){let u=c.data.data.map(y=>({Cabang:y.branch_name||"","Nama Facility care":y.original_fc_name||"",Periode:y.period||"",Relifer:y.reliever_name||"","Tanggal Back Up":y.backup_date||"","Tanggal Selesai":y.completion_date||"",Keterangan:y.reason||"",Shift:y.shift||"",Status:y.status||""}));u.length===0&&u.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),P(u,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{P([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async c=>{let y=(await x("/api/branches?all=1")).data?.data||[],b=w=>{if(!w)return null;let C=String(w||"").toLowerCase(),S=y.find(_=>String(_.full_name||"").toLowerCase()===C||String(_.code||"").toLowerCase()===C||String(_.name||"").toLowerCase()===C);return S?S.id:null},m=c.map(w=>({branch_name:String(w.Cabang||"").trim(),backup_date:String(w["Tanggal Back Up"]||w["Tanggal Backup"]||"").trim(),original_fc_name:String(w["Nama Facility care"]||w["FC Digantikan"]||"").trim(),reliever_name:String(w.Relifer||w.Reliefer||"").trim(),period:String(w.Periode||"").trim(),reason:String(w.Keterangan||"").trim(),shift:String(w.Shift||"").trim(),completion_date:String(w["Tanggal Selesai"]||"").trim(),status:String(w.Status||"").trim()})).filter(w=>w.reliever_name&&w.backup_date),f=await x("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!f.ok)throw new Error(f.data?.error||"Import gagal");return f.data}}})}N();K();async function Gt(t){let e=await j(),i=Array.from({length:4},(a,r)=>String(new Date().getFullYear()-r));F({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>ce(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>J(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),s=await x(`/api/reports/inspection?limit=10000&${r}`);if(s.ok){let o=s.data.data.map(l=>({Cabang:l.branch_name||"",Periode:l.period||"",Tanggal:l.inspection_date||"","Point FC":l.fc_score!==null&&l.fc_score!==void 0?l.fc_score:"","Point SPV":l.spv_score!==null&&l.spv_score!==void 0?l.spv_score:"",Status:l.status||"","Link Dokumen":l.document_link||""}));P(o,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{P([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let r=n=>{if(!n)return null;let d=String(n||"").toLowerCase(),h=e.find(p=>String(p.label||"").toLowerCase()===d);return h?h.value:null},s=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let d=String(n).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let p=Number(d);if(p>2e4&&p<99999){let c=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(c.getTime())?"":c.toISOString().slice(0,10)}}let h=d.split(/[\/\-\.]/);if(h.length===3){let[p,c,u]=h.map(y=>y.trim());if(p.length===4&&c.length<=2&&u.length<=2)return`${p}-${c.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&c.length<=2&&p.length<=2)return`${u}-${c.padStart(2,"0")}-${p.padStart(2,"0")}`}return d},o=a.map(n=>({branch_id:r(String(n.Cabang||"").trim()),period:String(n.Periode||"").trim(),inspection_date:s(n.Tanggal),fc_score:n["Point FC"]!==void 0&&n["Point FC"]!==""?Number(n["Point FC"]):null,spv_score:n["Point SPV"]!==void 0&&n["Point SPV"]!==""?Number(n["Point SPV"]):null,status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.inspection_date),l=await x("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}N();K();async function zt(t){let e=await j(),i=Array.from({length:4},(a,r)=>String(new Date().getFullYear()-r));F({container:t,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>ce(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>J(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),s=await x(`/api/reports/cleaning?limit=10000&${r}`);if(s.ok){let o=s.data.data.map(l=>({Cabang:l.branch_name||"",Jenis:l.activity_type||"",Periode:l.period||"",Tanggal:l.activity_date||"",Status:l.status||"","Link Dokumen":l.document_link||""}));P(o,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{P([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let r=n=>{if(!n)return null;let d=String(n||"").toLowerCase(),h=e.find(p=>String(p.label||"").toLowerCase()===d);return h?h.value:null},s=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let d=String(n).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let p=Number(d);if(p>2e4&&p<99999){let c=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(c.getTime())?"":c.toISOString().slice(0,10)}}let h=d.split(/[\/\-\.]/);if(h.length===3){let[p,c,u]=h.map(y=>y.trim());if(p.length===4&&c.length<=2&&u.length<=2)return`${p}-${c.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&c.length<=2&&p.length<=2)return`${u}-${c.padStart(2,"0")}-${p.padStart(2,"0")}`}return d},o=a.map(n=>({branch_id:r(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:s(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),l=await x("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}N();K();async function Qt(t,e){let i=await j(),a=Array.from({length:4},(o,l)=>String(new Date().getFullYear()-l)),r=e?e.get("dash_filter"):null,s={};if(r==="fogging"){let o=new Date,l=String(o.getMonth()+1).padStart(2,"0"),n=String(o.getFullYear()),d=e?e.get("month"):null;d&&d.length===7&&(n=d.split("-")[0],l=d.split("-")[1]),s={status:"Done",month:l,year:n}}F({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:s,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:o=>`<span class="badge badge-warning">${o}</span>`},{key:"period",label:"Periode",render:o=>ce(o)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>J(o)},{key:"document_link",label:"Dokumen",render:o=>o?`<a href="${o}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:o=>o||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:i,value:o?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:o?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:o?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:o?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:o?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async o=>{let l=new URLSearchParams(o||{}).toString(),n=await x(`/api/reports/fogging?limit=10000&${l}`);if(n.ok){let d=n.data.data.map(h=>({Cabang:h.branch_name||"",Jenis:h.activity_type||"Fogging",Periode:h.period||"",Tanggal:h.activity_date||"",Status:h.status||"","Link Dokumen":h.document_link||""}));P(d,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{P([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async o=>{let l=p=>{if(!p)return null;let c=String(p||"").toLowerCase(),u=i.find(y=>String(y.label||"").toLowerCase()===c);return u?u.value:null},n=p=>{if(p==null||p==="")return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let c=String(p).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let y=Number(c);if(y>2e4&&y<99999){let b=new Date(Date.UTC(1899,11,30)+y*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}let u=c.split(/[\/\-\.]/);if(u.length===3){let[y,b,m]=u.map(f=>f.trim());if(y.length===4&&b.length<=2&&m.length<=2)return`${y}-${b.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&b.length<=2&&y.length<=2)return`${m}-${b.padStart(2,"0")}-${y.padStart(2,"0")}`}return c},d=o.map(p=>({branch_id:l(String(p.Cabang||"").trim()),activity_type:String(p.Jenis||p.Kegiatan||"Fogging").trim(),period:String(p.Periode||"").trim(),activity_date:n(p.Tanggal),status:String(p.Status||"").trim(),document_link:String(p["Link Dokumen"]||"").trim(),notes:String(p.Catatan||p.Keterangan||"").trim()})).filter(p=>p.branch_id&&p.period&&p.activity_date),h=await x("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(d)});if(!h.ok)throw new Error(h.data?.error||"Import gagal");return h.data}}})}N();K();async function Vt(t){let e=await j(),i=await X(),a=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],r=Array.from({length:4},(l,n)=>String(new Date().getFullYear()-n)),s=l=>l&&!i.find(n=>n.value===l)?[...i,{value:l,label:l}]:i,o=l=>l&&!a.find(n=>n.value===l)?[...a,{value:l,label:l}]:a;F({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:l=>`<span title="${l||""}">${l?.length>60?l.slice(0,60)+"\u2026":l||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>J(l)},{key:"notes",label:"Keterangan",render:l=>l?.length>40?l.slice(0,40)+"\u2026":l||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:l=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:l?.branch_id},{name:"pic",label:"PIC",type:"select",options:o(l?.pic),value:l?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:l?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:l?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:l?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:l?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async l=>{let n=new URLSearchParams(l||{}).toString(),d=await x(`/api/reports/basecamp?limit=10000&${n}`);if(d.ok){let h=d.data.data.map(p=>({"Tgl Info":p.info_date||"",Cabang:p.branch_name||"",Permasalahan:p.problem||"",PIC:p.pic||"","Tgl Done":p.done_date||"",Status:p.status||"",Keterangan:p.notes||""}));P(h,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{P([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async l=>{let n=c=>{if(!c)return null;let u=String(c||"").toLowerCase(),y=e.find(b=>String(b.label||"").toLowerCase()===u);return y?y.value:null},d=c=>{if(c==null||c==="")return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let u=String(c).trim();if(u===""||u==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);if(/^\d{4,5}$/.test(u)){let b=Number(u);if(b>2e4&&b<99999){let m=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}let y=u.split(/[\/\-\.]/);if(y.length===3){let[b,m,f]=y.map(w=>w.trim());if(b.length===4&&m.length<=2&&f.length<=2)return`${b}-${m.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&m.length<=2&&b.length<=2)return`${f}-${m.padStart(2,"0")}-${b.padStart(2,"0")}`}return u},h=l.map(c=>({info_date:d(c["Tgl Info"]||c["Tanggal Info"]),branch_id:n(String(c.Cabang||"").trim()),problem:String(c.Permasalahan||"").trim(),pic:String(c.PIC||"").trim(),done_date:d(c["Tgl Done"]||c["Tanggal Done"]),status:String(c.Status||"").trim(),notes:String(c.Keterangan||c.Catatan||"").trim()})).filter(c=>c.info_date&&c.branch_id&&c.problem),p=await x("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(h)});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}}})}async function Wt(t){F({container:t,title:"SOP",icon:"\u{1F4DA}",apiPath:"/api/sop",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(N(),_e)),r=await a(`/api/sop?limit=10000&${i}`);if(r.ok){let s=r.data.data.map(l=>({"Nama SOP":l.name||"",Kategori:l.category||"",Dokumen:l.document_link||"",Catatan:l.notes||l.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(K(),le));o(s,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(K(),le));i(e,"Template_Import_SOP")},onImport:async e=>{let i=e.map(s=>({name:String(s["Nama SOP"]||"").trim(),category:String(s.Kategori||"").trim(),document_link:String(s.Dokumen||"").trim(),description:String(s.Catatan||"").trim()})).filter(s=>s.name),{apiFetch:a}=await Promise.resolve().then(()=>(N(),_e)),r=await a("/api/sop/import",{method:"POST",body:JSON.stringify(i)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function Yt(t){F({container:t,title:"Master Checklist",icon:"\u{1F4CB}",apiPath:"/api/checklist",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(N(),_e)),r=await a(`/api/checklist?limit=10000&${i}`);if(r.ok){let s=r.data.data.map(l=>({"Nama Checklist":l.name||"",Kategori:l.category||"",Dokumen:l.document_link||"",Deskripsi:l.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(K(),le));o(s,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(K(),le));i(e,"Template_Import_Checklist")},onImport:async e=>{let i=e.map(s=>({name:String(s["Nama Checklist"]||"").trim(),category:String(s.Kategori||"").trim(),document_link:String(s.Dokumen||"").trim(),description:String(s.Deskripsi||"").trim()})).filter(s=>s.name),{apiFetch:a}=await Promise.resolve().then(()=>(N(),_e)),r=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(i)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}N();xe();K();async function St(t,e="forms"){if(e==="supply")return Ma(t);Na(t)}function Na(t){F({container:t,title:"Master Form",icon:"\u{1F4D1}",apiPath:"/api/forms",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await x(`/api/forms?limit=10000&${i}`);a.data?.data?P(a.data.data,"Data_Master_Form"):Q("Gagal export data master form")},onImport:async e=>{let i=await x("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!i.ok)throw new Error(i.data?.error||"Import failed");return i.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function Ma(t){let i=((await x("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));F({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,r)=>r.branch_name_ref||r.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let r=JSON.parse(a);return Array.isArray(r)?r.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let r=JSON.parse(a);return Array.isArray(r)?r.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>J(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let r=a?.tools_items;try{r=Array.isArray(JSON.parse(r))?JSON.parse(r).join(", "):r}catch{}let s=a?.chemical_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:a?.branch_id&&!i.find(o=>o.value==a.branch_id)?[...i,{value:a.branch_id,label:a.branch_name||a.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:r},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:s},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),s=await x(`/api/reports/supply?limit=10000&${r}`);if(s.ok){let o=s.data.data.map(l=>{let n=l.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let d=l.chemical_items;try{d=Array.isArray(JSON.parse(d))?JSON.parse(d).join(", "):d}catch{}return{Waktu:l.submitted_at||"",Pengirim:l.submitter_name||"",Cabang:l.branch_name_ref||l.branch_name||"","Alat/Barang":n||"",Chemical:d||"",Catatan:l.additional_notes||"",Status:l.status||"","Diproses Oleh":l.processed_by||""}});P(o,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{P([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let s=(await x("/api/branches?all=1")).data?.data||[],o=h=>{if(!h)return null;let p=String(h||"").toLowerCase(),c=s.find(u=>String(u.full_name||"").toLowerCase()===p||String(u.code||"").toLowerCase()===p||String(u.name||"").toLowerCase()===p);return c?c.id:null},l=h=>{if(h==null||h==="")return"";if(h instanceof Date&&!isNaN(h.getTime()))return h.toISOString().slice(0,10);let p=String(h).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let u=Number(p);if(u>2e4&&u<99999){let y=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}let c=p.split(/[\/\-\.]/);if(c.length===3){let[u,y,b]=c.map(m=>m.trim());if(u.length===4&&y.length<=2&&b.length<=2)return`${u}-${y.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&y.length<=2&&u.length<=2)return`${b}-${y.padStart(2,"0")}-${u.padStart(2,"0")}`}return p},n=a.map(h=>({submitted_at:l(h.Waktu||h.Tanggal),submitter_name:String(h.Pengirim||"").trim(),branch_id:o(String(h.Cabang||"").trim()),tools_items:String(h["Alat/Barang"]||h.Alat||"").trim(),chemical_items:String(h.Chemical||"").trim(),additional_notes:String(h.Catatan||h.Keterangan||"").trim(),status:String(h.Status||"").trim(),processed_by:String(h["Diproses Oleh"]||h.PIC||"").trim()})).filter(h=>h.submitted_at&&h.submitter_name&&h.branch_id),d=await x("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,r)=>{let s=se({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(o,l)=>{let n=o.querySelector("#supply-status").value,d=o.querySelector("#supply-processed-by").value;(await x(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:d})})).ok?(Y("Status diperbarui."),l(),r()):Q("Gagal update status.")}})}}]})}N();K();async function Xt(t){let e=we();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}F({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let a=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await x(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let a=i.data.data.map(r=>({"Nama Lengkap":r.full_name||"",Username:r.username||"",Email:r.email||"",Role:r.role||"",Status:r.is_active?"Aktif":"Nonaktif"}));P(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{P([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let a=i.map(s=>({full_name:String(s["Nama Lengkap"]||"").trim(),username:String(s.Username||"").trim(),email:String(s.Email||"").trim(),role:String(s.Role||"").trim()||"viewer",password:String(s.Password||"").trim()})).filter(s=>s.username&&s.password&&s.email&&s.full_name),r=await x("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}N();K();async function Zt(t){F({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await x(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)P(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{P([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let i=e.map(r=>({code:String(r["Kode Cabang"]||"").trim(),name:String(r["Nama Pendek"]||"").trim(),full_name:String(r["Nama Lengkap"]||"").trim(),city:String(r.Kota||"").trim()})).filter(r=>r.code&&r.name),a=await x("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}N();async function ea(t){let e=new Date,i=[];t.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),r()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),r()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(s=>s.addEventListener("change",r));async function a(){try{let s=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;i=(await x(`/api/dashboard/calendar?month=${s}`)).data?.data||[]}catch(s){console.warn("[Calendar] Failed to load events, rendering empty grid:",s),i=[]}}async function r(){let s=document.getElementById("calendar-grid");if(s){s.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let o=e.getFullYear(),l=e.getMonth(),n=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),d=document.getElementById("cal-month-label");d&&(d.textContent=n);let h=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(S=>S.value)),p=i.filter(S=>h.has(S.type)),c={};p.forEach(S=>{let _=(S.event_date||"").slice(0,10);c[_]||(c[_]=[]),c[_].push(S)});let u=new Date(o,l,1).getDay(),y=new Date(o,l+1,0).getDate(),b=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],m=new Date().toISOString().slice(0,10),f='<div class="calendar-grid">';b.forEach(S=>{f+=`<div class="cal-day-header">${S}</div>`});for(let S=0;S<u;S++)f+='<div class="cal-cell cal-cell-empty"></div>';for(let S=1;S<=y;S++){let _=`${o}-${String(l+1).padStart(2,"0")}-${String(S).padStart(2,"0")}`,E=c[_]||[],$=_===m;f+=`
          <div class="cal-cell ${$?"cal-today":""} ${E.length?"cal-has-events":""}"
               data-date="${_}" tabindex="0" role="button" aria-label="${_}">
            <div class="cal-day-num ${$?"today-num":""}">${S}</div>
            <div class="cal-events-preview">
              ${E.slice(0,3).map(T=>`
                <div class="cal-event-dot cal-color-${T.color||"gray"}" title="${nt(T.title||T.type)}">
                  <span class="cal-event-dot-label">${Oa(T.title||T.branch_name||T.type,18)}</span>
                </div>
              `).join("")}
              ${E.length>3?`<div class="cal-more">+${E.length-3} lagi</div>`:""}
            </div>
          </div>`}let C=(u+y)%7;if(C!==0)for(let S=0;S<7-C;S++)f+='<div class="cal-cell cal-cell-empty"></div>';f+="</div>",s.innerHTML=f,s.querySelectorAll(".cal-cell[data-date]").forEach(S=>{S.addEventListener("click",()=>{let _=S.dataset.date,E=c[_]||[];if(!E.length)return;let $=document.getElementById("cal-event-list"),T=new Date(_+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=T,document.getElementById("cal-event-items").innerHTML=E.map(L=>`
            <div class="cal-event-item cal-color-border-${L.color||"gray"}">
              <div class="cal-event-type">${Ra(L.type)}</div>
              <div class="cal-event-title">${nt(L.title||"-")}</div>
              <div class="cal-event-branch">${nt(L.branch_name||"")}</div>
              ${L.status?`<div class="cal-event-status">${nt(L.status)}</div>`:""}
              ${L.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${L.days_remaining} hari</div>`:""}
            </div>
          `).join(""),$.style.display="block"})})}catch(o){console.error("[Calendar] Render error:",o),s&&(s.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}r()}function Oa(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function nt(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Ra(t){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[t]||t}N();async function ta(t){let e=we(),i=(e?.full_name||e?.username||"U")[0].toUpperCase(),r={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
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
  `;let s=localStorage.getItem("fm_token"),o=document.getElementById("session-info");if(s&&o)try{let l=JSON.parse(atob(s.split(".")[1])),n=new Date(l.exp*1e3);o.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{o.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async l=>{l.preventDefault();let n=document.getElementById("pwd-error"),d=document.getElementById("pwd-success"),h=document.getElementById("btn-save-pwd");n.style.display="none",d.style.display="none";let p=l.target,c=p.current_password.value,u=p.new_password.value,y=p.confirm_password.value;if(u!==y){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(u.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}h.disabled=!0,h.textContent="\u23F3 Menyimpan...";let b=await x("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:c,new_password:u})});h.disabled=!1,h.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',b.ok?(d.textContent="\u2705 Password berhasil diubah.",d.style.display="block",p.reset(),Y("Password berhasil diubah.")):(n.textContent=b.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}N();var it={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function Z(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let r=Number(e);if(r>2e4&&r<99999){let s=new Date(Date.UTC(1899,11,30)+r*864e5);return isNaN(s.getTime())?null:s.toISOString().slice(0,10)}}let i=e.split(/[\/\-\.]/);if(i.length===3){let[r,s,o]=i.map(h=>h.trim()),l=Number(r),n=Number(s),d=Number(o);if(r.length===4&&l>1900)return`${r}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`;if(o.length===4&&d>1900)return l>12?`${o}-${s.padStart(2,"0")}-${r.padStart(2,"0")}`:n>12?`${o}-${r.padStart(2,"0")}-${s.padStart(2,"0")}`:`${o}-${s.padStart(2,"0")}-${r.padStart(2,"0")}`;if(o.length===2&&!isNaN(d)){let h=d>=50?`19${o}`:`20${o}`;return l>12?`${h}-${s.padStart(2,"0")}-${r.padStart(2,"0")}`:`${h}-${s.padStart(2,"0")}-${r.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function aa(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Ka={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",phone:t["No. Hp"],join_date:Z(t["Tanggal Masuk"]),status:t.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:Z(t["Tanggal Mulai"]),end_date:Z(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:Z(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:Z(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:Z(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:Z(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:Z(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:Z(t["Tanggal Target"]||t["Tgl Target"]),completion_date:Z(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:Z(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:Z(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:Z(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:Z(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:Z(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:Z(t["Tanggal Back Up"]),completion_date:Z(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:Z(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:Z(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function ja(t,e){let i=it[t];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Ka[i.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let r=[],s=[],o=[];return e.filter(n=>!aa(n)).forEach((n,d)=>{let h=e.indexOf(n)+2,p=[];a.required.forEach(({key:u,label:y})=>{let b=n[u];if(b==null||String(b).trim()===""){let m=Object.keys(n).filter(f=>f.trim()).join(", ");p.push({column:y,originalValue:b||"",reason:`Kolom "${y}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${m.slice(0,120)}`})}});let c=a.map(n);p.length>0?s.push({row:h,data:c,raw:n,errors:p}):(r.push(n),o.push(c))}),{valid:r,errors:s,mapped:o}}function na(t){let e=[];return t.SheetNames.forEach(i=>{let a=it[i];if(!a)return;let r=t.Sheets[i],s=window.XLSX.utils.sheet_to_json(r,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),o=ja(i,s),l=s.filter(n=>!aa(n));e.push({sheetName:i,module:a.module,label:a.label,total:l.length,valid:o.mapped.length,errorCount:o.errors.length,errors:o.errors,mapped:o.mapped,skipped:!1})}),e}function ia(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,r])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(r),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function ra(t){let e=window.XLSX,i=e.utils.book_new(),a=!1;return t.forEach(r=>{if(!r.errors||r.errors.length===0)return;a=!0;let s=r.errors.map(l=>({"No. Baris":l.row,"Kolom Gagal":(l.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(l.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(l.data||{}).map(([n,d])=>[n,d??""]))})),o=e.utils.json_to_sheet(s);e.utils.book_append_sheet(i,o,r.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Ha=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function la(t){t.innerHTML=`
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
              ${Object.entries(it).map(([b,{label:m}])=>`<span class="import-sheet-tag">\u{1F4C4} ${b} \u2192 ${m}</span>`).join("")}
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
  `;let e=null,i=null,a=0,r={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function s(b){Object.entries(r).forEach(([m,f])=>{f.style.display=m===b?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let b=document.getElementById("btn-backup-db");b.disabled=!0,b.textContent="\u23F3 Memproses Backup...";try{let m=await x("/api/import/backup");if(m.ok){if(!window.XLSX){Q("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let f=window.XLSX,w=f.utils.book_new();Object.entries(m.data.database).forEach(([C,S])=>{let _=S.length>0?S:[{}],E=f.utils.json_to_sheet(_);f.utils.book_append_sheet(w,E,C.substring(0,31))}),f.writeFile(w,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),Y("Backup berhasil diunduh!")}else Q("Gagal memproses backup: "+(m.data?.error||"Unknown error"))}catch(m){Q("Gagal memproses backup: "+m.message)}finally{b.disabled=!1,b.textContent="\u{1F4E6} Backup Database"}});let o=document.getElementById("btn-sync-google");o&&o.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let b=o.innerHTML;o.innerHTML='<span class="spinner"></span> Menyinkronkan...',o.disabled=!0;try{let m=await x("/api/sync/google-sheets",{method:"POST"});m.ok?alert("Sinkronisasi Berhasil: "+(m.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(m.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{o.innerHTML=b,o.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{ia(),Y("Template Excel berhasil didownload!")});let l=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",b=>{b.stopPropagation(),l.click()}),l.addEventListener("change",b=>{b.target.files[0]&&d(b.target.files[0])}),n.addEventListener("dragover",b=>{b.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",b=>{b.preventDefault(),n.classList.remove("drag-over");let m=b.dataTransfer.files[0];m&&m.name.match(/\.xlsx?$/i)?d(m):Q("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,l.value="",document.getElementById("file-info").style.display="none",n.style.display="",s("upload")});async function d(b){e=b,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${b.name} (${(b.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await h(b)}async function h(b){s("validating");let m=document.getElementById("validation-status"),f=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");m.textContent="Membaca file Excel...",f.style.width="20%",await Ge(200);let w=await b.arrayBuffer(),C=window.XLSX.read(w,{type:"array",cellDates:!0});m.textContent=`Memvalidasi ${C.SheetNames.length} sheet...`,f.style.width="50%",await Ge(100),i=na(C),f.style.width="100%",m.textContent="Validasi selesai!",await Ge(300),p()}catch(w){s("upload"),Q("Gagal memproses file: "+w.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function p(){s("preview");let b=i.filter(T=>!T.skipped).length,m=i.reduce((T,L)=>T+L.total,0),f=i.reduce((T,L)=>T+L.valid,0),w=i.reduce((T,L)=>T+L.errorCount,0),C=m>0?Math.round(f/m*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${b} sheet</span>
      <span class="badge badge-secondary">${m} baris</span>
      <span class="badge badge-success">${f} valid (${C}%)</span>
      ${w>0?`<span class="badge badge-danger">${w} error</span>`:""}
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
          ${i.map((T,L)=>`
            <tr class="${T.errorCount>0?"row-error":T.skipped?"row-skipped":"row-ok"}">
              <td><strong>${T.sheetName}</strong></td>
              <td>${T.label}</td>
              <td style="text-align:center">${T.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${T.valid}</span></td>
              <td style="text-align:center">${T.errorCount>0?`<span class="badge badge-danger">${T.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${T.skipped?'<span class="badge badge-neutral">Dilewati</span>':T.errorCount>0&&T.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':T.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':T.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${T.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${L}">\u{1F50D} ${T.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,S.querySelectorAll(".btn-detail-error").forEach(T=>{T.addEventListener("click",()=>{let L=i[Number(T.dataset.idx)];c(L)})});let _=document.getElementById("error-detail-section"),E=document.getElementById("error-detail-container");E.innerHTML="",_.style.display="none";let $=document.getElementById("btn-start-import");f===0?($.disabled=!0,$.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):($.disabled=!1,w>0?($.innerHTML=`\u{1F680} Import ${f} Data Valid (${w} dilewati)`,$.title="Baris error akan dilewati, baris valid tetap diimport"):$.innerHTML=`\u{1F680} Mulai Import ${f} Data`)}function c(b){let m=document.getElementById("error-detail-section"),f=document.getElementById("error-detail-container");m.style.display="";let w=b.errors.slice(0,100).map(C=>(Array.isArray(C.errors)?C.errors:[]).map(_=>{let E=typeof _=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${C.row}</span></td>
            <td><strong>${E?_.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${E&&_.originalValue!==void 0?_.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${E?_.reason:_}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${E&&_.aliases?`Gunakan salah satu nama kolom:<br><em>${_.aliases}</em>`:E&&_.hint?_.hint:""}
            </td>
          </tr>
        `}).join("")).join("");f.innerHTML=`
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
            <tbody>${w||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${b.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,m.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{s("upload"),document.getElementById("file-info").style.display="none",n.style.display="",e=null,l.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;ra(i)?Y("Log error berhasil didownload."):Y("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let b=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";u(b)});async function u(b){s("importing"),a=Date.now();let m=[];Ha.forEach(_=>{let E=i?.find($=>$.module===_&&$.mapped?.length>0);E&&m.push(E)});let f=document.getElementById("import-steps-list");f.innerHTML=m.map(_=>`
      <div class="import-step-item" id="step-item-${_.module}">
        <span class="step-item-icon" id="step-icon-${_.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${_.label} <span class="step-item-count">(${_.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${_.module}"></span>
      </div>
    `).join("");let w=document.getElementById("import-bar"),C=document.getElementById("import-current-status"),S={totalSheets:m.length,totalRows:m.reduce((_,E)=>_+E.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let _=0;_<m.length;_++){let E=m[_],$=document.getElementById(`step-icon-${E.module}`),T=document.getElementById(`step-status-${E.module}`);$.textContent="\u{1F504}",T.textContent="Mengimport...",C.textContent=`Mengimport ${E.label}...`,w.style.width=`${Math.round(_/m.length*100)}%`;try{let L=await x(`/api/import/${E.module}`,{method:"POST",body:JSON.stringify({rows:E.mapped,onDuplicate:b})});if(L.ok){let I=L.data;S.inserted+=I.inserted||0,S.skipped+=I.skipped||0,S.moduleResults.push({label:E.label,inserted:I.inserted||0,skipped:I.skipped||0,status:"ok"}),$.textContent="\u2705",T.innerHTML=`<span class="badge badge-success">${I.inserted||0} berhasil</span>${I.skipped>0?` <span class="badge badge-neutral">${I.skipped} skip</span>`:""}`}else S.failed++,S.moduleResults.push({label:E.label,inserted:0,skipped:0,status:"error",error:L.data?.error}),$.textContent="\u274C",T.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(L){S.failed++,S.moduleResults.push({label:E.label,inserted:0,skipped:0,status:"error",error:L.message}),$.textContent="\u274C",T.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Ge(150)}w.style.width="100%",C.textContent="Selesai!",await Ge(400),y(S)}function y(b){s("summary");let m=((Date.now()-a)/1e3).toFixed(1),f=b.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${f?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${f?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
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
          ${b.moduleResults.map(w=>`
            <tr>
              <td>${w.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${w.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${w.skipped}</span></td>
              <td style="text-align:center">
                ${w.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${w.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,i=null,l.value="",document.getElementById("file-info").style.display="none",n.style.display="",s("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Ge(t){return new Promise(e=>setTimeout(e,t))}N();var rt=[],oa=[];async function sa(t){rt=await j(),oa=await X(),F({container:t,title:"Data SP (Surat Peringatan)",icon:"\u{1F4DC}",apiPath:"/api/sp",enableMobileFilterSheet:!0,itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:rt}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await x(`/api/sp?limit=10000&${i}`);if(a.ok){let r=a.data.data.map(o=>({"Nama Karyawan":o.employee_name||"",Divisi:o.division||"",Cabang:o.branch_name||"","Tanggal Sp":o.tanggal||"","Akhir Sp":o.akhir_sp||"","Jenis Sp":o.sp_type||"","Link Document / Foto":o.document_link||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(K(),le));s(r,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(K(),le));i(e,"Template_Import_SP")},onImport:async e=>{let i=o=>{if(!o)return null;let l=String(o||"").toLowerCase(),n=rt.find(d=>String(d.label||"").toLowerCase()===l);return n?n.value:null},a=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let l=String(o).trim();if(/^\d{4,5}$/.test(l)){let d=Number(l);if(d>2e4&&d<99999){let h=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);let n=l.split(/[\/\-\.]/);if(n.length===3){let[d,h,p]=n.map(c=>c.trim());if(d.length===4&&h.length<=2&&p.length<=2)return`${d}-${h.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&h.length<=2&&d.length<=2)return`${p}-${h.padStart(2,"0")}-${d.padStart(2,"0")}`}return l},r=e.map(o=>({employee_name:String(o["Nama Karyawan"]||"").trim(),division:String(o.Divisi||"").trim(),branch_id:i(String(o.Cabang||"").trim()),tanggal:a(o["Tanggal Sp"]),akhir_sp:a(o["Akhir Sp"]),sp_type:String(o["Jenis Sp"]||"").trim(),document_link:String(o["Link Document / Foto"]||"").trim()})).filter(o=>o.employee_name&&o.branch_id),s=await x("/api/import/sp",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:[{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:oa},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:rt,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}N();var Le=[],ca=[];async function da(t){Le=await j(),ca=await X(),F({container:t,title:"Data Mutasi",icon:"\u{1F501}",apiPath:"/api/mutasi",enableMobileFilterSheet:!0,itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:Le},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:Le}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await x(`/api/mutasi?limit=10000&${i}`);if(a.ok){let r=a.data.data.map(o=>({Tanggal:o.tanggal||"","Nama Karyawan":o.employee_name||"","Cabang Asal":o.from_branch_name||"","Cabang Tujuan":o.to_branch_name||"",Status:o.status||"",Dokumen:o.document_link||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(K(),le));s(r,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(K(),le));i(e,"Template_Import_Mutasi")},onImport:async e=>{let i=o=>{if(!o)return null;let l=String(o||"").toLowerCase(),n=Le.find(d=>String(d.label||"").toLowerCase()===l);return n?n.value:null},a=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let l=String(o).trim();if(/^\d{4,5}$/.test(l)){let d=Number(l);if(d>2e4&&d<99999){let h=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);let n=l.split(/[\/\-\.]/);if(n.length===3){let[d,h,p]=n.map(c=>c.trim());if(d.length===4&&h.length<=2&&p.length<=2)return`${d}-${h.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&h.length<=2&&d.length<=2)return`${p}-${h.padStart(2,"0")}-${d.padStart(2,"0")}`}return l},r=e.map(o=>({tanggal:a(o.Tanggal),employee_name:String(o["Nama Karyawan"]||"").trim(),from_branch_id:i(String(o["Cabang Asal"]||"").trim()),to_branch_id:i(String(o["Cabang Tujuan"]||"").trim()),status:String(o.Status||"").trim(),document_link:String(o.Dokumen||"").trim()})).filter(o=>o.tanggal&&o.employee_name&&o.from_branch_id&&o.to_branch_id),s=await x("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:ca},{type:"select",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Le,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Le,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let i=Math.floor(Number(t)-25569);return new Date(i*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let i=t.split(/[\/\-]/);return`${i[2]}-${i[1]}-${i[0]}`}let e=t.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(e){let i=e[1],a=parseInt(e[2],10),r=parseInt(e[3],10);if(a>12&&r<=12)return`${i}-${e[3]}-${e[2]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let i=e.split("-");if(i.length===3&&i[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],r=parseInt(i[2],10),s=a[parseInt(i[1],10)-1];return`${r} ${s} ${i[0]}`}return e};function V(t){return async e=>{if(!Me()){Ce("/login");return}return t(e)}}var ze=null;function qa(){ze&&clearInterval(ze);let t=()=>{let e=new Date,i=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),r=document.getElementById("header-clock-time"),s=document.getElementById("header-clock-date");r&&(r.textContent=i),s&&(s.textContent=a)};t(),ze=setInterval(t,1e3)}async function Ja(){try{let t=await x("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},i=(a,r)=>{let s=document.getElementById(a);s&&(s.textContent=r>0?r:"",s.style.display=r>0?"inline-flex":"none")};i("badge-issues",e.issues?.current||0),i("badge-contracts",e.expiring30?.current||0),i("badge-oo1",e.one_on_one?.current||0),i("badge-schedule",e.schedule?.current||0),i("badge-supply",e.supply?.current||0)}catch{}}var Fe=[];async function Ua(){try{let t=await x("/api/dashboard/notifications");if(!t.ok)return;Fe=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Fe.length>0?"block":"none",e.textContent=Fe.length)}catch{}}function Ga(){if(!Fe.length){se({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,i)=>i()});return}let t=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${Fe.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;se({title:`Notifikasi (${Fe.length})`,content:t,confirmText:"Tutup",onConfirm:(e,i)=>i()})}function pa(){let t=we(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
  `;let i=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),r=document.getElementById("topbar-menu-btn"),s=document.getElementById("sidebar-close"),o=()=>{i.classList.add("open"),a.classList.add("show")},l=()=>{i.classList.remove("open"),a.classList.remove("show")};r?.addEventListener("click",o),s?.addEventListener("click",l),a?.addEventListener("click",l),document.querySelectorAll(".nav-item").forEach(d=>d.addEventListener("click",l));function n(){let d=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(c=>{let u=c.dataset.route;c.classList.toggle("active",d===u||u!=="/dashboard"&&d.startsWith(u))});let h=document.getElementById("topbar-title"),p=document.querySelector(".nav-item.active .nav-label");h&&p&&(h.textContent=p.textContent)}window.addEventListener("hashchange",n),n(),qa(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await x("/api/auth/logout",{method:"POST"}),Oe(),ze&&clearInterval(ze),Ce("/login")}),Ja(),Ua(),document.getElementById("btn-notif")?.addEventListener("click",d=>{d.preventDefault(),Ga()})}async function za(){U("/login",({main:e})=>Mt(e)),U("/dashboard",V(({main:e})=>Pt(e))),U("/calendar",V(({main:e})=>ea(e))),U("/employees",V(({main:e,params:i})=>Ot(e,i))),U("/contracts",V(({main:e,params:i})=>tt(e,i))),U("/sp",V(({main:e})=>sa(e))),U("/mutasi",V(({main:e})=>da(e))),U("/timeline",V(({main:e,params:i})=>Kt(e,i))),U("/issues",V(({main:e,params:i})=>jt(e,i))),U("/one-on-one",V(({main:e,params:i})=>Ht(e,i))),U("/training",V(({main:e})=>qt(e))),U("/relievers",V(({main:e,params:i})=>Ut(e,i))),U("/reports/inspection",V(({main:e})=>Gt(e))),U("/reports/cleaning",V(({main:e})=>zt(e))),U("/reports/fogging",V(({main:e})=>Qt(e))),U("/reports/basecamp",V(({main:e})=>Vt(e))),U("/reports/supply",V(({main:e})=>St(e,"supply"))),U("/sop",V(({main:e})=>Wt(e))),U("/checklist",V(({main:e})=>Yt(e))),U("/forms",V(({main:e})=>St(e))),U("/users",V(({main:e})=>Xt(e))),U("/branches",V(({main:e})=>Zt(e))),U("/profile",V(({main:e})=>ta(e))),U("/settings/import",V(({main:e})=>la(e)));let t=Me();if(!t&&window.location.hash!=="#/login"&&Ce("/login"),t){let e=await x("/api/auth/me");e.ok?(Re(e.data.data),pa()):(Oe(),Ce("/login"))}window.addEventListener("fm:login",()=>{pa(),Ce("/dashboard")}),xt()}za();
