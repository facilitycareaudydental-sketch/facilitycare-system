var wa=Object.defineProperty;var gt=(t,e)=>()=>(t&&(e=t(t=0)),e);var bt=(t,e)=>{for(var i in e)wa(t,i,{get:e[i],enumerable:!0})};var Ie={};bt(Ie,{API:()=>It,CLIENT_SIDE_MAX_ROWS:()=>De,IS_DEVELOPMENT:()=>Ke,apiFetch:()=>C,clearToken:()=>ze,getToken:()=>je,getUser:()=>te,setToken:()=>ht,setUser:()=>He});function je(){return localStorage.getItem("fm_token")}function ht(t){localStorage.setItem("fm_token",t)}function ze(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function te(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function He(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function C(t,e={}){let i=je(),a={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...e.headers||{}};try{let r=`cb=${Date.now()}`,s=t.includes("?")?"&":"?",o=`${It}${t}${s}${r}`,l=await fetch(o,{...e,headers:a}),n;try{let c=await l.text();try{n=JSON.parse(c)}catch{n={error:`Server Error (${l.status}): ${c.substring(0,80)}...`}}}catch{n={error:"Gagal membaca respon dari server"}}return l.status===401&&(ze(),window.location.hash="#/login"),{ok:l.ok,status:l.status,data:n}}catch(r){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${r.message})`}}}}var Ke,Sa,It,De,N=gt(()=>{Ke=!1,Sa="https://fm-operations-api.facilitycare-audydental.workers.dev",It=Sa,De=1e4});var yt={};bt(yt,{confirmDialog:()=>Se,createModal:()=>se});function se({title:t,content:e,onConfirm:i,onCancel:a,confirmText:r="Simpan",cancelText:s="Batal",size:o="md",confirmClass:l="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},c=document.createElement("div");c.className="modal-overlay",c.innerHTML=`
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
  `,e instanceof HTMLElement&&c.querySelector(".modal-body").appendChild(e);let v=()=>{c.classList.remove("show"),setTimeout(()=>c.remove(),250)};return c.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),v()}),c.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),v()}),i&&c.querySelector(".modal-confirm").addEventListener("click",()=>i(c,v)),c.addEventListener("click",b=>{b.target===c&&(a&&a(),v())}),document.body.appendChild(c),requestAnimationFrame(()=>c.classList.add("show")),{overlay:c,close:v}}function Se(t,e,i="Konfirmasi"){return se({title:i,content:`<p>${t}</p>`,onConfirm:(a,r)=>{e(),r()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var ye=gt(()=>{});var le={};bt(le,{downloadExcel:()=>L,parseExcel:()=>Ue,renderExcelButtons:()=>Ca});function Ue(t){return new Promise((e,i)=>{let a=new FileReader;a.onload=r=>{try{let s=new Uint8Array(r.target.result),o=XLSX.read(s,{type:"array"}),l=o.SheetNames[0],n=o.Sheets[l];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${o.SheetNames.join(", ")}`),console.log(`Sheet Used: ${l}`);let c=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),v=c.e.r-c.s.r+1,b=c.e.c-c.s.c+1;console.log(`Total Rows (including empty): ${v}`),console.log(`Total Columns: ${b}`);let u=[];for(let h=c.s.c;h<=c.e.c;++h){let f=n[XLSX.utils.encode_cell({c:h,r:c.s.r})];f&&f.v&&u.push(f.v)}console.log(`Headers Found: ${u.join(", ")}`),console.log("---------------------------");let p=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(p,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(p,"__headers",{value:u,enumerable:!1}),e(p)}catch(s){i(s)}},a.onerror=r=>i(r),a.readAsArrayBuffer(t)})}function L(t,e){try{let i=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,i,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function Ca(t){return`
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
  `}var K=gt(()=>{});N();var ft={},Xe=null;function q(t,e){ft[t]=e}function $e(t){window.location.hash=t}function Bt(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[i,...a]=e.split("?"),r=ft[i];if(!r){for(let[o,l]of Object.entries(ft))if(o.endsWith("/*")&&i.startsWith(o.slice(0,-2))){r=l;break}}Xe&&(Xe(),Xe=null);let s=document.getElementById("main-content");if(s&&(s.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),r){let o=new URLSearchParams(a.join("?")),l=i.split("/").filter(Boolean),n=await r({path:i,params:o,segments:l,main:s});n&&(Xe=n)}else{let o=s||document.getElementById("app");o&&(o.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var qe;function _a(){return qe||(qe=document.createElement("div"),qe.id="toast-container",document.body.appendChild(qe)),qe}function Lt(t,e="info",i=3500){let a=_a(),r=document.createElement("div");r.className=`toast toast-${e}`;let s={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};r.innerHTML=`<span class="toast-icon">${s[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(r),requestAnimationFrame(()=>r.classList.add("show")),setTimeout(()=>{r.classList.remove("show"),setTimeout(()=>r.remove(),350)},i)}var J=t=>Lt(t,"success"),j=t=>Lt(t,"error");ye();N();N();ye();function Ze({columns:t,data:e,onEdit:i,onDelete:a,onView:r,actions:s=[],emptyText:o="Tidak ada data",bulkSelect:l=null}){let n=document.createElement("div");if(n.className="table-wrapper",!e||e.length===0)return n.innerHTML=`<div class="empty-state"><p>${o}</p></div>`,n;let c=document.createElement("table");c.className="data-table";let v=document.createElement("thead"),b=document.createElement("tr");if(l){let p=document.createElement("th");p.style.width="40px",p.style.textAlign="center";let h=document.createElement("input");h.type="checkbox",h.id="select-all-checkbox",h.title="Pilih semua",h.addEventListener("change",()=>{e.forEach(f=>{h.checked?l.selectedIds.add(f.id):l.selectedIds.delete(f.id)}),n.querySelectorAll(".row-checkbox").forEach(f=>f.checked=h.checked),l.onToggle()}),p.appendChild(h),b.appendChild(p)}if(t.forEach(p=>{let h=document.createElement("th");h.textContent=p.label,p.width&&(h.style.width=p.width),b.appendChild(h)}),i||a||r||s.length>0){let p=document.createElement("th");p.textContent="Aksi",p.style.width="120px",b.appendChild(p)}v.appendChild(b),c.appendChild(v);let u=document.createElement("tbody");return e.forEach(p=>{let h=document.createElement("tr");if(l){let f=document.createElement("td");f.style.textAlign="center",f.style.width="40px";let m=document.createElement("input");m.type="checkbox",m.className="row-checkbox",m.checked=l.selectedIds.has(p.id),m.addEventListener("change",()=>{if(m.checked)l.selectedIds.add(p.id);else{l.selectedIds.delete(p.id);let d=document.getElementById("select-all-checkbox");d&&(d.checked=!1)}l.onToggle()}),f.appendChild(m),h.appendChild(f)}if(t.forEach(f=>{let m=document.createElement("td");if(f.render){let d=f.render(p[f.key],p);d instanceof HTMLElement?m.appendChild(d):m.innerHTML=d||""}else m.textContent=p[f.key]!==null&&p[f.key]!==void 0&&p[f.key]!==""?p[f.key]:"";f.nowrap&&(m.style.whiteSpace="nowrap"),h.appendChild(m)}),i||a||r||s.length>0){let f=document.createElement("td");f.className="actions-cell";let m=document.createElement("div");if(m.className="btn-group",r){let d=document.createElement("button");d.className="btn btn-xs btn-ghost",d.innerHTML="\u{1F441}",d.title="Lihat",d.addEventListener("click",()=>r(p)),m.appendChild(d)}if(i){let d=document.createElement("button");d.className="btn btn-xs btn-secondary",d.innerHTML="\u270F\uFE0F",d.title="Edit",d.addEventListener("click",()=>i(p)),m.appendChild(d)}s.forEach(d=>{let g=document.createElement("button");g.className=`btn btn-xs ${d.class||"btn-ghost"}`,g.innerHTML=d.icon||d.label,g.title=d.label,g.addEventListener("click",()=>d.handler(p)),m.appendChild(g)}),f.appendChild(m),h.appendChild(f)}u.appendChild(h)}),c.appendChild(u),n.appendChild(c),n}function et({page:t,pages:e,total:i,limit:a,onPage:r}){if(e<=1)return null;let s=document.createElement("div");s.className="pagination";let o=document.createElement("span");o.className="pagination-info",o.textContent=`Total: ${i} data`,s.appendChild(o);let l=document.createElement("div");l.className="pagination-btns";let n=(b,u,p=!1,h=!1)=>{let f=document.createElement("button");f.className=`btn btn-sm ${h?"btn-primary":"btn-ghost"} pagination-btn`,f.textContent=b,f.disabled=p,f.addEventListener("click",()=>r(u)),l.appendChild(f)};n("\xAB",1,t===1),n("\u2039",t-1,t===1);let c=Math.max(1,t-2),v=Math.min(e,t+2);for(let b=c;b<=v;b++)n(b,b,!1,b===t);return n("\u203A",t+1,t===e),n("\xBB",e,t===e),s.appendChild(l),s}ye();function Je(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${Je(e.fields)}</div>`;let i=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",r="";switch(e.type){case"textarea":r=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${i} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let o=(e.options||[]).map(p=>{let h=typeof p=="object"?p.value:p,f=typeof p=="object"?p.label:p,m=e.value==h?"selected":"";return`<option value="${h}" ${m}>${f}</option>`}).join("");r=`<select name="${e.name}" class="form-control" ${i}><option value="">-- Pilih ${e.label||""} --</option>${o}</select>`;break;case"combobox":let l=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,n=(e.options||[]).map(p=>{let h=typeof p=="object"?p.value:p,f=typeof p=="object"?p.label||p.value||"":p||"";return(f==="undefined"||f==="[object Object]"||f==="null")&&(f=""),f?`<option value="${f}"></option>`:""}).join(""),c=e.value||"";if(e.value){let p=(e.options||[]).find(h=>(typeof h=="object"?h.value:h)==e.value);if(p){let h=typeof p=="object"?p.label||p.value||"":p||"";h&&h!=="undefined"&&h!=="[object Object]"&&h!=="null"&&(c=h)}}r=`
          <input type="text" name="${e.name}" list="${l}" class="form-control" value="${c}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${l}">${n}</datalist>
        `;break;case"checkbox":r=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let v=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";r=`<input type="date" name="${e.name}" class="form-control" value="${v}" ${i}>`;break;case"month":let b=e.value?String(e.value).slice(0,7):"";r=`<input type="month" name="${e.name}" class="form-control" value="${b}" ${i}>`;break;case"number":let u=e.readonly?"readonly":"";r=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${i} ${u}>`;break;case"email":r=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i}>`;break;case"url":r=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${i}>`;break;default:r=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i} autocomplete="off">`}let s=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${r}${s}</div>`}).join("")}function tt(t){let e={},i=new FormData(t);for(let[a,r]of i.entries())e[a]=r===""?null:r;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function at(t,e){e&&Object.entries(e).forEach(([i,a])=>{let r=t.querySelector(`[name="${i}"]`);r&&(r.hasAttribute("list")||(r.type==="checkbox"?r.checked=!!a:r.type==="date"&&a&&window.parseFlexibleDate?r.value=window.parseFlexibleDate(a):r.type==="month"&&a?r.value=String(a).slice(0,7):r.value=a??""))})}K();var _e={},nt={on(t,e){_e[t]||(_e[t]=new Set),_e[t].add(e)},off(t,e){_e[t]&&_e[t].delete(e)},emit(t,e){_e[t]&&_e[t].forEach(i=>{try{i(e)}catch(a){console.warn("[calendarBus] Handler error:",a)}})},clear(){Object.keys(_e).forEach(t=>delete _e[t])}},$a=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function ce(t){if(!t){nt.emit("data:changed",{module:"unknown"});return}let e=String(t).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");nt.emit("data:changed",{module:e,relevant:$a.has(e)})}function F({container:t,title:e,icon:i,apiPath:a,columns:r,formFields:s,filterFields:o,defaultFilters:l={},itemLabel:n="Data",canCreate:c=!0,canEdit:v=!0,canDelete:b=!0,onBeforeSubmit:u,onAfterLoad:p,onDataLoaded:h,extraActions:f=[],initialSearch:m="",exportOptions:d=null,bulkDelete:g=!1,paginationMode:x="server"}){let w=te();w&&w.role==="viewer"&&(c=!1,v=!1,b=!1,g=!1,d=null),w&&w.role==="editor_khusus"&&a!=="/api/issues"&&a!=="/api/overtime"&&(c=!1,v=!1,b=!1,g=!1,d=null),w&&w.role==="input_lembur"&&(a==="/api/overtime"?(b=!1,g=!1,d=null):(c=!1,v=!1,b=!1,g=!1,d=null));let k=1,$={...l};m&&($.search=m);let E=new Set;t.innerHTML=`
    ${g?`
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
        ${c?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
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
    

    ${o&&o.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${o.filter(y=>y.type==="search").map(y=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${y.placeholder||"Cari..."}" id="filter-search" value="${$.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${o.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let _=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select" name="${y.name}" id="filter-${y.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${_}</option>${(y.options||[]).map(S=>`<option value="${typeof S=="object"?S.value:S}" ${$[y.name]===(typeof S=="object"?S.value:S)?"selected":""}>${typeof S=="object"?S.label:S}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let _=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select filter-select-sheet" name="${y.name}-sheet" id="filter-sheet-${y.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${_}</option>${(y.options||[]).map(S=>`<option value="${typeof S=="object"?S.value:S}" ${$[y.name]===(typeof S=="object"?S.value:S)?"selected":""}>${typeof S=="object"?S.label:S}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function T(){let y=document.getElementById("bulk-toolbar");if(!y)return;let _=document.getElementById("bulk-count"),S=document.getElementById("btn-bulk-delete"),R=document.getElementById("btn-bulk-cancel");_.textContent=`${E.size} item dipilih`,E.size>0?(y.style.display="flex",S.disabled=!1,R.disabled=!1):(y.style.display="none",S.disabled=!0,R.disabled=!0);let M=document.getElementById("select-all-checkbox");if(M){let V=document.querySelectorAll(".row-checkbox");if(V.length>0){let B=[...V].every(oe=>oe.checked),A=[...V].some(oe=>oe.checked);M.checked=B,M.indeterminate=A&&!B}else M.checked=!1,M.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{E.clear(),document.querySelectorAll(".row-checkbox").forEach(_=>_.checked=!1);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1),T()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(E.size===0)return;let y=[...E],_=document.createElement("div");_.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",_.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${y.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${y.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(_),_.querySelector("#bulk-cancel-btn").addEventListener("click",()=>_.remove()),_.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let S=_.querySelector("#bulk-confirm-btn");S.disabled=!0,S.textContent="Menghapus...";let R=await C(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:y})});_.remove(),R.ok?(J(`${y.length} ${n} berhasil dihapus.`),E.clear(),T(),ce(a),P()):j(R.data?.error||"Gagal menghapus data.")})});let D=document.getElementById("filter-search"),I;D?.addEventListener("input",y=>{clearTimeout(I),I=setTimeout(()=>{$.search=y.target.value,k=1,E.clear(),T(),P()},400)}),o?.forEach(y=>{(y.type==="select"||y.type==="combobox")&&(document.getElementById(`filter-${y.name}`)?.addEventListener("change",_=>{$[y.name]=_.target.value;let S=document.getElementById(`filter-sheet-${y.name}`);S&&(S.value=_.target.value),k=1,E.clear(),T(),P()}),document.getElementById(`filter-sheet-${y.name}`)?.addEventListener("change",_=>{$[y.name]=_.target.value;let S=document.getElementById(`filter-${y.name}`);S&&(S.value=_.target.value),k=1,E.clear(),T(),P(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{$={},D&&(D.value=""),o?.forEach(y=>{let _=document.getElementById(`filter-${y.name}`);_&&(_.value="");let S=document.getElementById(`filter-sheet-${y.name}`);S&&(S.value="")}),k=1,E.clear(),T(),P()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{$={},D&&(D.value=""),o?.forEach(y=>{let _=document.getElementById(`filter-${y.name}`);_&&(_.value="");let S=document.getElementById(`filter-sheet-${y.name}`);S&&(S.value="")}),k=1,E.clear(),T(),P(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>fe(null)),d&&document.addEventListener("click",function(y){let _=document.getElementById("aksi-menu-main"),S=document.getElementById("btn-aksi-main");_&&S&&!S.contains(y.target)&&!_.contains(y.target)&&_.classList.remove("show-aksi-menu")});let O=document.getElementById("btn-mobile-filter"),X=document.getElementById("filter-options-wrapper"),Y=document.getElementById("btn-close-filter-sheet");if(O&&X&&(O.addEventListener("click",y=>{y.preventDefault(),X.classList.add("sheet-open")}),Y&&Y.addEventListener("click",y=>{y.preventDefault(),X.classList.remove("sheet-open")})),d){document.getElementById(`btn-export-${d.moduleName}`)?.addEventListener("click",async _=>{let S=_.target,R=S.innerHTML;S.innerHTML="\u23F3 Loading...",S.disabled=!0;try{await d.onExport()}catch{j("Gagal export data")}finally{S.innerHTML=R,S.disabled=!1}}),document.getElementById(`btn-template-${d.moduleName}`)?.addEventListener("click",()=>{d.onTemplate()});let y=document.getElementById(`input-import-${d.moduleName}`);y?.addEventListener("change",async _=>{let S=_.target.files[0];if(!S)return;let R=document.getElementById(`label-import-${d.moduleName}`),M=R?R.querySelector(".import-text"):null,V=M?M.innerText:"";M&&(M.innerText="\u231B Memproses..."),R&&(R.style.pointerEvents="none"),y.disabled=!0;try{let B=await Ue(S);if(B.length===0)throw new Error("File kosong atau format salah");await d.onImport(B),J("Import berhasil!"),ce(a),P()}catch(B){j(B.message||"Gagal import data")}finally{M&&(M.innerText=V),R&&(R.style.pointerEvents="auto"),y.disabled=!1,y.value=""}})}async function P(){T();let y=document.getElementById("table-container");if(!y)return;y.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let _=x==="client",S=_?1:k,R=_?De:20,M=new URLSearchParams({page:S,limit:R,...Object.fromEntries(Object.entries($).filter(([,U])=>U))}),V=await C(`${a}?${M}`);if(!V.ok){y.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${V.data?.error||"Error"}</p></div>`;return}let B=V.data?.data||V.data||[],A=V.data?.pagination,oe=B.length;if(_){B=h(B);let U=B.length,W=20,ie=Math.ceil(U/W);k>ie&&ie>0&&(k=ie);let z=(k-1)*W,de=k*W;B=B.slice(z,de),A={page:k,limit:W,total:U,pages:ie}}!1,p&&p(B);let ke=Ze({columns:r,data:B,onEdit:v?U=>fe(U):null,actions:f.map(U=>({...U,handler:W=>U.handler(W,P)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:g?{selectedIds:E,onToggle:T}:null});y.innerHTML="",y.appendChild(ke);let pe=document.getElementById("pagination-container");if(pe&&(pe.innerHTML="",A&&A.pages>1)){let U=et({page:A.page,pages:A.pages,total:A.total,limit:A.limit,onPage:W=>{k=W,P()}});U&&pe.appendChild(U)}}function Te(y){let _=typeof s=="function"?s(y):s;return Je(_)}function fe(y){let _=!!y,S=document.createElement("form");if(S.noValidate=!0,S.innerHTML=Te(y),_){let M=typeof s=="function"?s(y):s;at(S,y)}let{close:R}=se({title:_?`Edit ${n}`:`Tambah ${n}`,content:S,size:"lg",confirmText:_?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(M,V)=>{if(!S.reportValidity())return;let B=M.querySelector(".modal-confirm");B.disabled=!0,B.textContent="Menyimpan...";let A=tt(S),oe=typeof s=="function"?s(y):s,ke=async ie=>{for(let z of ie)if(z.type==="row")await ke(z.fields);else if(z.type==="combobox"&&A[z.name]){let de=A[z.name],we=(z.options||[]).find(Z=>{let re=String(typeof Z=="object"?Z.value:Z),ut=String(typeof Z=="object"?Z.label:Z);return re===de||ut===de});if(we)A[z.name]=typeof we=="object"?we.value:we;else if(z.createApi){let Z={};Z[z.createApi.field]=de,z.createApi.extra&&Object.assign(Z,z.createApi.extra);let re=await C(z.createApi.path,{method:"POST",body:JSON.stringify(Z)});if(re.ok&&re.data?.id)A[z.name]=re.data.id;else if(re.ok&&!re.data?.id)A[z.name]=de;else throw new Error(`Gagal membuat master data: ${re.data?.error||"Unknown error"}`)}}};try{await ke(oe)}catch(ie){j(ie.message),B.disabled=!1,B.textContent=_?"Simpan Perubahan":`Tambah ${n}`;return}u&&(A=await u(A,y));let pe=_?"PUT":"POST",U=_?`${a}/${y.id}`:a,W=await C(U,{method:pe,body:JSON.stringify(A)});W.ok?(J(_?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),V(),ce(a),P()):(j(W.data?.error||"Gagal menyimpan data."),B.disabled=!1,B.textContent=_?"Simpan Perubahan":`Tambah ${n}`)}})}function ne(y){Se(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let _=await C(`${a}/${y.id}`,{method:"DELETE"});_.ok?(J(`${n} berhasil dihapus.`),ce(a),P()):j(_.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return P(),P}N();N();var Be=null,it=null;async function Ee(t=!1){if(Be&&!t)return console.log("Employees Raw (Cache Hit)",Be.slice(0,5)),Be;let e=await C(`/api/employees?limit=${De}&status=Aktif`);return Be=(e.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",Be.slice(0,5)),Be}async function ae(t=!1){let i=(await Ee(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",i.slice(0,5)),i}async function H(t=!1){return it&&!t||(it=((await C("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),it}function Q(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function vt(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function Le(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function xt(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function me(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}K();N();K();N();K();N();K();function rt(t){if(t.target.name==="report_date"||t.target.name==="completion_date"){let e=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),a=document.querySelector('input[name="day_count"]');if(e&&i&&a)if(e.value&&i.value){let r=new Date(e.value),s=new Date(i.value),o=Math.floor((s-r)/864e5);a.value=isNaN(o)?"":o}else a.value=""}}document.body.removeEventListener("input",rt);document.body.addEventListener("input",rt);document.body.removeEventListener("change",rt);document.body.addEventListener("change",rt);N();var Ce={};function Pe(t){if(Ce[t]){try{Ce[t].destroy()}catch{}delete Ce[t]}}function Ea(){Object.keys(Ce).forEach(Pe)}var ve=(t,e=0)=>{let i=Number(t);return isNaN(i)||t===null||t===void 0?e:i},Fe=(t,e="\u2014")=>{if(t==null||t==="")return e;let i=String(t).trim();return i===""||i==="[object Object]"?e:i};function ot(t,e,i=900){if(!t)return;t._animFrame&&cancelAnimationFrame(t._animFrame);let a=Math.max(0,Math.round(ve(e))),r=t.textContent?t.textContent.replace(/[^\d]/g,""):"",s=r!==""&&parseInt(r)||0;if(a===s){t.textContent=a.toLocaleString("id-ID"),t._animFrame=null;return}let o=Date.now(),l=()=>{let n=Math.min((Date.now()-o)/i,1),c=1-Math.pow(1-n,3);t.textContent=Math.round(s+c*(a-s)).toLocaleString("id-ID"),n<1?t._animFrame=requestAnimationFrame(l):(t.textContent=a.toLocaleString("id-ID"),t._animFrame=null)};t._animFrame=requestAnimationFrame(l)}var Ta={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},Da=t=>{let e=Fe(t,"\u2014");return`<span class="status-pill ${Ta[e]||"pill-neutral"}">${e}</span>`};var ge={family:"Inter",size:11},xe="#94A3B8",Ae="#F1F5F9",kt=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],Ia=()=>window.innerWidth<768;function st(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:Ia()?"bottom":"top",labels:{font:ge,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ge,titleFont:{...ge,weight:"700"}}},scales:{x:{grid:{color:Ae},ticks:{font:ge,color:xe,maxRotation:0}},y:{grid:{color:Ae},ticks:{font:ge,color:xe},beginAtZero:!0}},...t}}var Ba=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),La=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Ft(t=3){return Array(t).fill(0).map((e,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function ue(t,e,i=8e3){try{let a=new AbortController,r=setTimeout(()=>a.abort(),i),s=await C(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(r),!s||!s.ok)return e;let o=s.data;return o?o.data!==void 0?o.data??e:o:e}catch{return e}}function Fa(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(a=>{let r=document.getElementById(a);r&&(r.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(a=>{let r=document.getElementById(a);if(r&&r.style.display==="none"){r.style.display="block";let s=r.parentElement;if(s&&!s.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent="Belum ada data",r.style.display="none",s.appendChild(o)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Nt({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Mt({}),["table-contracts","table-issues"].forEach(a=>{let r=document.getElementById(a);r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Pt(t){Ea(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">
      
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h2 style="margin:0; font-size:1.2rem; color:var(--text-1);">Dashboard Overview</h2>
        <button id="btn-dash-refresh" class="btn-primary" style="padding:6px 12px; font-size:0.8rem; display:flex; align-items:center; gap:6px; cursor:pointer; background:var(--primary); color:#fff; border:none; border-radius:6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
          Refresh Data
        </button>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${Ba()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${La()}</div>

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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${Ft(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${Ft(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>At(t)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async i=>{let a=i.target.value,r=document.getElementById("jadwal-year-label");r&&(r.textContent=a);let s=document.getElementById("skel-jadwal"),o=document.getElementById("chart-jadwal");s&&(s.style.display="block",s.style.position="absolute"),o&&(o.style.display="none");let l=await ue(`/api/dashboard/schedule-chart?year=${a}`,{},8e3);try{Ot(l)}catch(n){console.warn("ScheduleChart render:",n),be("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let a=i.target.value,r=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",s=document.getElementById("skel-insp"),o=document.getElementById("chart-insp");s&&(s.style.display="block",s.style.position="absolute"),o&&(o.style.display="none");let l=await ue(r,{},8e3);try{Rt(l)}catch(n){console.warn("InspBar render:",n),be("skel-insp","chart-insp")}}),t._skelTimeout=setTimeout(()=>Fa(),5e3),await At(t)}async function At(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,i,a,r,s,o,l,n]=await Promise.all([ue("/api/dashboard/kpi",{},8e3),ue("/api/dashboard/issues-trend",{},8e3),ue("/api/dashboard/issues-summary",{},8e3),ue("/api/dashboard/stats",{},8e3),ue("/api/dashboard/calendar",[],8e3),ue("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),ue(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),ue("/api/dashboard/dropdown-aggregates",{},8e3)]),c=document.getElementById("filter-insp-month"),v=c?c.value:"",b=v?`/api/dashboard/inspection-bar?month=${v}`:"/api/dashboard/inspection-bar",u=await ue(b,{},8e3);window.dashboardAggregates=n||{};try{Nt(e)}catch(p){console.warn("KPI render:",p)}try{Mt(e)}catch(p){console.warn("MiniStats render:",p)}try{Ot(l)}catch(p){console.warn("ScheduleChart render:",p),be("skel-jadwal","chart-jadwal")}try{Aa(Array.isArray(a?.by_category)?a.by_category:[])}catch(p){console.warn("Donut render:",p),be("skel-donut","chart-donut")}try{Pa(i)}catch(p){console.warn("Trend render:",p),be("skel-trend","chart-trend")}try{Rt(u)}catch(p){console.warn("InspBar render:",p),be("skel-insp","chart-insp")}try{let p=Array.isArray(r)?r:Array.isArray(r?.recent_issues)?r.recent_issues:[];Ma(p)}catch(p){console.warn("IssuesTable render:",p)}try{let p=Array.isArray(r?.expiring_contracts)?r.expiring_contracts:[];Na(o)}catch(p){console.warn("ContractsTable render:",p)}try{Oa(Array.isArray(s)?s:[])}catch(p){console.warn("Agenda render:",p)}try{Ra()}catch(p){console.warn("Quick Actions render:",p)}}function Nt(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let i=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=i.map(a=>{let r=ve(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${r}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${r}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{ot(a,parseInt(a.dataset.target)||0)})}function Mt(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let i=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,a=new Date().getFullYear(),r=String(new Date().getMonth()+1).padStart(2,"0"),s=`${a}-${r}`,o=p=>`
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
  `,l=window.dashboardAggregates||{},n=p=>{let h=window.dashboardAggregates?.schedule_quarter_details?.[p];return h?`Insp: ${h.insp} \u2022 GC: ${h.gc} \u2022 DC: ${h.dc} \u2022 Fog: ${h.fog}`:""},c=p=>{let h=window.dashboardAggregates?.schedule_quarter_details?.[p];return h?`Rincian ${p}: Inspeksi: ${h.insp} | General Cleaning: ${h.gc} | Deep Cleaning: ${h.dc} | Fogging: ${h.fog}`:`Total Kegiatan ${p}`},v=[{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.preventDefault(); event.stopPropagation();">
          <option value="Q1" ${i==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${i==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${i==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${i==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:l.schedule_by_quarter?.[i]??t.schedule?.current??0,href:`#/timeline?dash_filter=period_${i.toLowerCase()}`,color:"mini-blue",extra:`<div id="mini-jadwal-detail" style="font-size:0.6rem; color:var(--text-3); font-weight:500; margin-top:2px; line-height:1.2; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${c(i)}">${n(i)}</div>`},{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:o("dash-inspeksi-month"),val:l.inspeksi_by_month?.[s]??t.inspection_month?.current??0,href:`#/timeline?dash_filter=inspeksi&month=${s}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:o("dash-gcdc-month"),val:l.gcdc_by_month?.[s]??t.cleaning_month?.current??0,href:`#/timeline?dash_filter=gcdc&month=${s}`,color:"mini-green"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:o("dash-reliefer-month"),val:l.relievers_by_month?.[s]??t.reliever_completed?.current??0,href:`#/relievers?dash_filter=reliever&month=${s}`,color:"mini-teal"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:o("dash-fogging-month"),val:l.fogging_by_month?.[s]??t.fogging_month?.current??0,href:`#/reports/fogging?dash_filter=fogging&month=${s}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=v.map(p=>`
    <a href="${p.href}" class="mini-stat ${p.color}" style="text-decoration:none" id="${p.id||""}">
      <div class="mini-stat-icon">${p.icon}</div>
      <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
        <div style="display:flex; align-items:baseline; gap:3px;">
          <div class="mini-stat-value" data-target="${ve(p.val)}">0</div>
          ${p.dropdown?p.dropdown:""}
        </div>
        <div class="mini-stat-text">${p.label}</div>
        ${p.extra||""}
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(p=>ot(p,parseInt(p.dataset.target)||0,700));let b=document.getElementById("dash-jadwal-period");if(b){let p=h=>{let m=(window.dashboardAggregates?.schedule_by_quarter||{})[h]||0,d=document.querySelector("#mini-jadwal .mini-stat-value");d&&(d.dataset.target=m,ot(d,m,400));let g=document.getElementById("mini-jadwal-detail");g&&(g.textContent=n(h),g.title=c(h));let x=document.getElementById("mini-jadwal");x&&(x.href=`#/timeline?dash_filter=period_${h.toLowerCase()}`,x.title=c(h))};b.addEventListener("change",h=>p(h.target.value))}let u=(p,h,f,m)=>{let d=document.getElementById(p);if(d){let g=x=>{let k=(window.dashboardAggregates?.[f]||{})[x]||0,$=document.querySelector(`#${h} .mini-stat-value`);$&&($.dataset.target=k,ot($,k,400));let E=document.getElementById(h);E&&(E.href=`${m}&month=${x}`)};d.addEventListener("change",x=>g(x.target.value))}};u("dash-reliefer-month","mini-reliefer","relievers_by_month","#/relievers?dash_filter=reliever"),u("dash-inspeksi-month","mini-inspeksi","inspeksi_by_month","#/timeline?dash_filter=inspeksi"),u("dash-gcdc-month","mini-gcdc","gcdc_by_month","#/timeline?dash_filter=gcdc"),u("dash-fogging-month","mini-fogging","fogging_by_month","#/reports/fogging?dash_filter=fogging")}function Aa(t){be("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!e||!i)return;Pe("donut");let a=(t||[]).filter(n=>ve(n.count)>0);if(!a.length){Ge(e,"Belum ada data permasalahan");return}let r=a.map(n=>`${Fe(n.category,"Lainnya")}`),s=a.map(n=>ve(n.count)),o=s.reduce((n,c)=>n+c,0);i.innerHTML=a.map((n,c)=>{let v=kt[c%kt.length],b=o>0?Math.round(n.count/o*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${v}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${b}%)</span></div>
          <div class="donut-legend-label">${r[c]}</div>
        </div>
      </div>
    `}).join("");let l={id:"centerText",beforeDraw:function(n){let c=n.width,v=n.height,b=n.ctx;b.restore();let u=(v/80).toFixed(2);b.font="bold "+u+"em Inter",b.textBaseline="middle",b.fillStyle="#1E293B";let p=o.toString(),h=Math.round((c-b.measureText(p).width)/2),f=v/2;b.fillText(p,h,f-4),b.font="600 "+(u*.35).toFixed(2)+"em Inter",b.fillStyle="#64748B";let m="Total",d=Math.round((c-b.measureText(m).width)/2);b.fillText(m,d,f+10),b.save()}};Ce.donut=new Chart(e,{type:"doughnut",data:{labels:r,datasets:[{data:s,backgroundColor:kt,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ge,titleFont:{...ge,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutout:"75%"},plugins:[l]})}function Pa(t){be("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;Pe("trend"),t=t||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(t.labels||[]).map(o=>{if(!o||typeof o!="string")return"";try{let[l,n]=o.split("-");return(i[Number(n)-1]||n)+" "+String(l).slice(-2)}catch{return o}}),r=(t.open||[]).map(o=>ve(o)),s=(t.closed||[]).map(o=>ve(o));if(!a.length){Ge(e,"Belum ada data trend");return}Ce.trend=new Chart(e,{type:"line",data:{labels:a,datasets:[{label:"Open",data:r,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:s,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:st({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:xe,maxRotation:45,autoSkip:!0}},y:{grid:{color:Ae},ticks:{font:{family:"Inter",size:9},color:xe},beginAtZero:!0}}})})}function Ot(t){be("skel-jadwal","chart-jadwal");let e=document.getElementById("chart-jadwal");if(!e)return;Pe("jadwal"),t=t||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(t).some(n=>Array.isArray(n)&&n.some(c=>c>0))){Ge(e,"Belum ada data jadwal");return}let r=t["Inspeksi Hygiene"]||Array(12).fill(0),s=t["General Cleaning"]||Array(12).fill(0),o=t["Deep Cleaning"]||Array(12).fill(0),l=t.Fogging||Array(12).fill(0);Ce.jadwal=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Inspeksi",data:r,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:s,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:o,backgroundColor:"#F59E0B"},{label:"Fogging",data:l,backgroundColor:"#EF4444"}]},options:st({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:xe,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Ae},ticks:{font:{family:"Inter",size:9},color:xe},min:0}}})})}function Rt(t){be("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;Pe("inspBar"),t=t||{};let i=t.labels||[],a=(t.fc||[]).map(s=>ve(s)),r=(t.spv||[]).map(s=>ve(s));if(!i.length){Ge(e,"Belum ada data inspeksi");return}Ce.inspBar=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:r,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:st({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ge,color:xe,maxRotation:45,minRotation:45}},y:{grid:{color:Ae},ticks:{font:ge,color:xe},min:0,max:100}}})})}function Na(t){be("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;Pe("contractMiniBar"),t=t||{};let i={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(o=>{let l=o.split("-")[1];return i[l]||o}),r=(t.data||[]).map(o=>ve(o));if(!a.length){Ge(e,"Belum ada data");return}let s=e.getContext("2d");Ce.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:r,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:st({onClick:(o,l)=>{if(l&&l.length>0){let n=l[0].index,c=(t.labels||[])[n];c&&(window.location.hash="#/contracts?month_expiry="+c)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ge,color:xe,maxRotation:0,autoSkip:!1}},y:{grid:{color:Ae,borderDash:[4,4],drawBorder:!1},ticks:{font:ge,color:xe,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function Ma(t){let e=document.getElementById("table-issues");if(!e)return;let i=(t||[]).slice(0,8);if(!i.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${i.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${Da(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Fe(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Fe(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function Oa(t){let e=document.getElementById("widget-agenda");if(!e)return;let i=new Date,a=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,s=(t||[]).filter(o=>(o.event_date||"").startsWith(a)).slice(0,10);if(!s.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${s.map(o=>{let l="#3B82F6",n="#EFF6FF",c="Agenda",v=(o.title||"").toLowerCase();return v.includes("inspeksi")?(l="#10B981",n="#ECFDF5",c="Inspeksi"):v.includes("cleaning")||v.includes("gcdc")?(l="#3B82F6",n="#EFF6FF",c="Cleaning"):v.includes("reliefer")?(l="#F59E0B",n="#FFFBEB",c="Reliefer"):v.includes("fogging")&&(l="#8B5CF6",n="#F5F3FF",c="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(o.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${l};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Fe(o.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Fe(o.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${l}">${c}</div>
        </div>
      `}).join("")}
    </div>
  `}function Ra(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function be(t,e){let i=document.getElementById(t),a=document.getElementById(e);if(i&&(i.style.display="none",i.style.position=""),a){a.style.display="block";let r=a.parentElement;if(r){let s=r.querySelector(".chart-empty");s&&s.remove()}}}function Ge(t,e="Belum ada data"){if(!t)return;t.style.display="none";let i=t.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent=e,i.appendChild(r)}}N();async function Kt(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),i=document.getElementById("login-error"),a=document.getElementById("login-btn"),r=document.getElementById("toggle-password"),s=document.getElementById("login-password");r?.addEventListener("click",()=>{let o=s.type==="text";s.type=o?"password":"text",r.style.color=o?"":"var(--primary)"}),e?.addEventListener("submit",async o=>{o.preventDefault(),i.style.display="none";let l=e.username.value.trim(),n=e.password.value;if(!l||!n){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let c=await C("/api/auth/login",{method:"POST",body:JSON.stringify({username:l,password:n})});c.ok&&c.data.success?(ht(c.data.data.token),He(c.data.data.user),J("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=c.data.error||"Username atau password salah.",i.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}N();K();async function Ka(){return await H()}function Ne(t){if(!t||t==="-"||String(t).trim()==="")return"-";let e=String(t).trim(),i=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],a=e.match(/^(\d{4})-(\d{1,2})/);if(a){let s=a[1],o=parseInt(a[2],10)-1;if(o>=0&&o<12)return`${i[o]} ${s}`}let r=e.split(/[\/\-\.]/);if(r.length===2&&r[1].length===4){let s=parseInt(r[0],10)-1;if(s>=0&&s<12)return`${i[s]} ${r[1]}`}if(r.length===3&&r[2].length===4){let s=parseInt(r[1],10)-1;if(s>=0&&s<12)return`${i[s]} ${r[2]}`}for(let s=0;s<12;s++)if(e.toLowerCase().includes(i[s].toLowerCase()))return e;return e}function ja(t,e){let i=String(t.status||"").toLowerCase();return e==="active"?i==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&i==="aktif":!1}async function jt(t,e){let i=await Ka(),a=e?e.get("dash_filter"):null;F({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",enableMobileFilterSheet:!0,itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:r=>a?r.filter(s=>ja(s,a)):r,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:r=>Le(r)},{key:"phone",label:"No. HP",render:r=>r?`<a href="tel:${r}">${r}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:r=>window.formatDate(r)},{key:"target_pindah_os",label:"Target Pindah OS",render:r=>Ne(r)},{key:"target_selesai_os",label:"Target Selesai OS",render:r=>Ne(r)},{key:"status",label:"Status",render:r=>Q(r)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:r=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:r?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:r?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:i,value:r?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:r?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:r?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:r?.status||""}]},{type:"row",fields:[{name:"target_pindah_os",label:"Target Pindah OS (Bulan)",type:"month",value:r?.target_pindah_os?String(r.target_pindah_os).slice(0,7):""},{name:"target_selesai_os",label:"Target Selesai OS (Bulan)",type:"month",value:r?.target_selesai_os?String(r.target_selesai_os).slice(0,7):""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let r=await C(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let s=r.data.data.map(o=>({"Nama Lengkap":o.full_name,Cabang:o.branch_name||"",Divisi:o.division||"","No. HP":o.phone||"","Tgl Masuk":o.join_date||"",Status:o.status||"","Target Pindah OS":Ne(o.target_pindah_os)==="-"?"":Ne(o.target_pindah_os),"Target Selesai OS":Ne(o.target_selesai_os)==="-"?"":Ne(o.target_selesai_os)}));L(s,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif","Target Pindah OS":"September 2026","Target Selesai OS":"Oktober 2026"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif","Target Pindah OS":"November 2026","Target Selesai OS":"Desember 2026"}],"Template_Import_Karyawan")},onImport:async r=>{let s=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),v=i.find(b=>String(b.label||"").toLowerCase()===c);return v?v.value:null},o=r.map(n=>({full_name:String(n["Nama Lengkap"]||"").trim(),branch_id:s(String(n.Cabang||"").trim()),division:String(n.Divisi||n["Div / Bagian"]||"").trim()||"FACILITY CARE",phone:String(n["No. HP"]||n["No. Hp"]||"").trim(),join_date:String(n["Tgl Masuk"]||n["Tanggal Masuk"]||"").trim(),status:String(n.Status||"").trim(),target_pindah_os:String(n["Target Pindah OS"]||"").trim(),target_selesai_os:String(n["Target Selesai OS"]||"").trim(),notes:String(n.Catatan||"").trim()})).filter(n=>n.full_name),l=await C("/api/import/employees",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}N();K();var St=[],zt=[];async function za(){St=await H(),zt=await Ee()}var wt=async t=>{let e=[],i=1;for(;;){let r=await(await Promise.resolve().then(()=>(N(),Ie))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${i}`);if(!r.ok)break;let s=r.data?.data||r.data||[],o=Array.isArray(s)?s:[];if(e=e.concat(o),o.length<100||r.data?.pagination&&i>=r.data.pagination.pages)break;i++}return e};async function lt(t,e){await za(),F({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",defaultFilters:{},onDataLoaded:a=>a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>Le(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,r)=>r.end_date&&String(r.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':vt(a)},{key:"status",label:"Status",render:a=>Q(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:St},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[s,o]=await Promise.all([wt("/api/employees?status=Aktif"),wt("/api/contracts")]);if(s.length>0){let l=o.filter(b=>b.status==="Aktif"),n=new Set(l.map(b=>b.employee_id)),c=s.filter(b=>!n.has(b.id)),v=`<p style="margin-bottom:12px">Data yang terbaca: <b>${s.length}</b> Karyawan Aktif, dan <b>${l.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${c.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;c.forEach(b=>{let u=o.filter(h=>h.employee_id===b.id),p='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(u.length>0){let h=u[0];p=`Pernah ada kontrak (Status: <b style="color:#EF4444">${h.status}</b>, Selesai: ${window.formatDate(h.end_date)})`}v+=`<li style="margin-bottom:8px"><b>${b.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${b.branch_name||"-"} | ${p}</span></li>`}),v+="</ul>",Promise.resolve().then(()=>(ye(),yt)).then(b=>b.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:v,cancelText:"Tutup"}))}}catch(s){console.error(s)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let r=document.querySelector(".page-actions");r&&r.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:zt,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:St,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await C(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let r=a.data.data.map(s=>({"Nama Lengkap":s.employee_name,Cabang:s.branch_name||"","Div / Bagian":s.division||"","Tanggal Mulai":s.start_date||"","Tanggal Selesai":s.end_date&&String(s.end_date).startsWith("2099")?"":s.end_date||"","Sisa Kontrak":s.end_date&&String(s.end_date).startsWith("2099")?"Tetap":s.days_remaining!==null&&s.days_remaining!==void 0?`${s.days_remaining} Hari`:"",Status:s.status||""}));L(r,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[r,s]=await Promise.all([C("/api/branches?limit=10000"),wt("/api/employees")]),o=r.data?.data||[],l=s||[];console.log(`Total employee yang berhasil dimuat dari database : ${l.length}`),l.length>0&&(console.log("Contoh 5 employee pertama:"),l.slice(0,5).forEach((m,d)=>{console.log(`${d+1}. ID: ${m.id}, Name: ${m.full_name}, Status: ${m.status}`)}));let n=m=>{if(!m)return null;let d=String(m||"").replace(/\s+/g," ").toLowerCase().trim(),g=o.find(x=>String(x.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===d||String(x.code||"").replace(/\s+/g," ").toLowerCase().trim()===d||String(x.name||"").replace(/\s+/g," ").toLowerCase().trim()===d);return g?g.id:null},c=(m,d)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${d}`),console.log(`Nama dari Excel : "${m}"`),!m)return console.log("Alasan gagal mapping : Nama kosong"),null;let g=String(m||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${g}"`),console.log(`Jumlah employee di database : ${l.length}`);let x=l.find(w=>String(w.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===g);return x?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${x.id}`),x.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},v=m=>{if(!m)return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let d=String(m).trim();if(/^\d{4,5}(\.\d+)?$/.test(d)){let x=Math.floor(Number(d));if(x>2e4&&x<99999){let w=new Date(Date.UTC(1899,11,30)+x*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);let g=d.split(/[\/\-\.]/);if(g.length===3){let[x,w,k]=g.map($=>$.trim());if(x.length===4&&w.length<=2&&k.length<=2)return`${x}-${w.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&w.length<=2&&x.length<=2)return`${k}-${w.padStart(2,"0")}-${x.padStart(2,"0")}`}return d},b=a.map((m,d)=>{let g=d+2,x=String(m["Nama Lengkap"]||"").trim(),w=m["Tanggal Mulai"],k=v(w);if(!k){let T=a.__worksheet,D=a.__headers||[],I=D.indexOf("Tanggal Mulai"),O="N/A",X="N/A",Y="N/A";if(I!==-1&&T&&window.XLSX){let Te=window.XLSX.utils.encode_cell({c:I,r:g-1});Y=Te;let fe=T[Te];fe?(O=fe.t||"undefined",X=fe.w||"undefined"):O="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let P="Unknown";w==null||w===""?P="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":w instanceof Date&&isNaN(w.getTime())?P="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":P="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${g}`),console.log(`Employee Name : ${x}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${I})`),console.log(`Raw Cell Value : "${w}"`),console.log(`JavaScript Type : ${typeof w}`),console.log(`SheetJS Cell Type : ${O}`),console.log(`SheetJS Formatted Value : "${X}"`),console.log(`Value After Trim : "${String(w||"").trim()}"`),console.log(`Value After Date Parser : "${k}"`),console.log(`Is Empty : ${!w}`),console.log(`Is Invalid Date : ${w instanceof Date?isNaN(w.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${P}`),console.log(`Workbook Sheet : ${T?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${Y}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(m,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(D)),console.log(`==========================
`)}let $=c(x,g),E=null;return $?k||(E="Tanggal Mulai kosong atau tidak berformat tanggal"):E="Karyawan tidak ditemukan di Database",{isValid:!!($&&k),invalidReason:E,rowNum:g,data:{employee_id:$,branch_id:n(String(m.Cabang||"").trim()),division:String(m["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:k,end_date:v(m["Tanggal Selesai"])||"2099-12-31",status:String(m.Status||"").trim(),_rawName:x}}}),u=[],p=[];if(b.forEach(m=>{m.isValid?u.push(m.data):p.push({rowNum:m.rowNum,name:m.data._rawName,reason:m.invalidReason})}),console.log(`Split Validation - Valid: ${u.length}, Invalid: ${p.length}`),u.length===0){let m=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${a.length}
Valid: 0
Invalid: ${p.length}

Daftar Kegagalan (Contoh):
`;p.slice(0,10).forEach(d=>{m+=`- Row ${d.rowNum} | Nama: ${d.name} | Alasan: ${d.reason}
`}),p.length>10&&(m+=`- ... dan ${p.length-10} lainnya.
`),alert(m);return}let h=await C("/api/contracts/import",{method:"POST",body:JSON.stringify(u)}),f=`IMPORT SUMMARY
======================
`;f+=`Total Baris Excel : ${a.length}
`,f+=`Baris Valid       : ${u.length}
`,f+=`Baris Invalid     : ${p.length}

`,h&&h.data&&h.data.metrics?(f+=`Berhasil INSERT   : ${h.data.metrics.inserted}
`,f+=`Berhasil UPDATE   : ${h.data.metrics.updated}
`):f+=`Berhasil diproses : ${u.length}
`,p.length>0&&(f+=`
DAFTAR DATA DILEWATI:
`,p.forEach(m=>{f+=`- Row ${m.rowNum} | ${m.name} | ${m.reason}
`})),alert(f),typeof lt=="function"&&lt()}}})}N();K();var _t=[],Ve=[];function Ha(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let i of e)if(t.some(a=>a.period===i))return i;return"Q3"}async function Ht(t,e){_t=await H();let i=await ae();Ve=["Berlin Ariansyah","Ade Surahman"];let a=m=>m&&!Ve.find(d=>String(typeof d=="object"?d.value:d).toLowerCase()===String(m).toLowerCase())?[...Ve,m]:Ve,r=await C(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),s=m=>{if(!m||m==="-"||String(m).trim()==="")return"";let d=String(m).split("-");return d.length===3&&d[0].length===4?`${d[2]}-${d[1]}-${d[0]}`:m},o=r.data?.data||[],l=Ha(o),n=e?e.get("dash_filter"):null,c=new Date,v=`${c.getFullYear()}-${String(c.getMonth()+1).padStart(2,"0")}`,b={},u=e&&e.get("month")?e.get("month"):null;n==="inspeksi"?b={status:"Done",activity_type:"Inspeksi Hygiene",month:u}:n==="gcdc"?b={status:"Done",activity_type:"GCDC",month:u}:n&&n.startsWith("period_")&&(b={period:n.replace("period_","").toUpperCase()});let p=new Date().getFullYear(),f=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((m,d)=>{let g=String(d+1).padStart(2,"0");return{value:`${p}-${g}`,label:`${m} ${p}`}});F({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:b,onDataLoaded:m=>m.sort((d,g)=>{let x=d.opening_date?new Date(d.opening_date).getTime():0;return(g.opening_date?new Date(g.opening_date).getTime():0)-x}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:m=>xt(m)},{key:"period",label:"Periode",render:m=>me(m)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:m=>s(m)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:m=>s(m)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:m=>s(m)},{key:"status",label:"Status",render:m=>Q(m)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:_t},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:f},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Ve}],formFields:m=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:_t,value:m?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:m?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:m?.period},{name:"pic",label:"PIC",type:"combobox",options:a(m?.pic),value:m?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:m?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:m?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:m?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:m?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:m?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let m=await C(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(m.ok){let d=m.data.data.map(g=>({Cabang:g.branch_name||"",Kegiatan:g.activity_type||"",Periode:g.period||"",PIC:g.pic||"","Tgl Opening":g.opening_date||"","Tgl Target":g.target_date||"","Tgl Selesai":g.completion_date||"",Status:g.status||""}));L(d,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async m=>{let g=(await C("/api/branches?all=1")).data?.data||[],x=E=>{if(!E)return null;let T=String(E||"").toLowerCase().trim(),D=T.match(/^(\d{3})/);if(D){let O=g.find(X=>String(X.code||"")===D[1]);if(O)return O.id}let I=g.find(O=>String(O.full_name||"").toLowerCase()===T||String(O.code||"").toLowerCase()===T||String(O.name||"").toLowerCase()===T);return I?I.id:null},w=E=>{if(E==null||E==="")return"";if(E instanceof Date&&!isNaN(E.getTime()))return E.toISOString().slice(0,10);let T=String(E).trim();if(T===""||T==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(T))return T.slice(0,10);if(/^\d{4,5}$/.test(T)){let I=Number(T);if(I>2e4&&I<99999){let O=new Date(Date.UTC(1899,11,30)+I*864e5);return isNaN(O.getTime())?"":O.toISOString().slice(0,10)}}let D=T.split(/[\/\-\.]/);if(D.length===3){let[I,O,X]=D.map(Y=>Y.trim());if(I.length===4&&O.length<=2&&X.length<=2)return`${I}-${O.padStart(2,"0")}-${X.padStart(2,"0")}`;if(X.length===4&&O.length<=2&&I.length<=2)return`${X}-${O.padStart(2,"0")}-${I.padStart(2,"0")}`}return T},k=m.map(E=>({branch_id:x(String(E.Cabang||"").trim()),branch_name:String(E.Cabang||"").trim(),activity_type:String(E.Kegiatan||"").trim(),period:String(E.Periode||"").trim(),pic:String(E.PIC||E.Pic||"").trim(),opening_date:w(E["Tgl Opening"]||E["Tanggal Opening"]||E["Tgl Openir"]),target_date:w(E["Tgl Target"]||E["Tanggal Target"]),completion_date:w(E["Tgl Selesai"]||E["Tanggal Selesai"]),status:String(E.Status||"").trim(),notes:String(E.Catatan||E.Keterangan||"").trim()})).filter(E=>E.activity_type&&E.period),$=await C("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:k,onDuplicate:"update"})});if(!$.ok)throw new Error($.data?.error||"Import gagal");return ce("schedule"),$.data}}})}N();K();var Ct=[],dt=[];function ct(t){if(t.target.name==="report_date"||t.target.name==="completion_date"){let e=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),a=document.querySelector('input[name="day_count"]');if(e&&i&&a)if(e.value&&i.value){let r=new Date(e.value),s=new Date(i.value),o=Math.floor((s-r)/864e5);a.value=isNaN(o)?"":o}else a.value=""}}document.body.removeEventListener("input",ct);document.body.addEventListener("input",ct);document.body.removeEventListener("change",ct);document.body.addEventListener("change",ct);function qa(t,e){let i=String(t.status||"").toLowerCase();return e==="open"?i==="open":!1}async function qt(t,e){let i=e?e.get("dash_filter"):null;Ct=await H(),dt=await ae();let a=n=>n&&!dt.find(c=>c.value===n)?[...dt,{value:n,label:n}]:dt,r=new Date().getFullYear(),s=["2025","2026","2027","2028","2029","2030"],l=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((n,c)=>{let v=String(c+1).padStart(2,"0");return{value:`${r}-${v}`,label:`${n} ${r}`}});F({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:n=>i?n.filter(c=>qa(c,i)):n,columns:[{key:"report_date",label:"Tanggal Info",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:n=>`<span class="badge badge-secondary">${n}</span>`},{key:"source",label:"Sumber Laporan"},{key:"complaint",label:"Keluhan",render:n=>`<span title="${n}">${n?.length>50?n.slice(0,50)+"...":n}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:n=>`<span title="${n||""}">${n?.length>40?n.slice(0,40)+"...":n||"-"}</span>`},{key:"status",label:"Status",render:n=>Q(n)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"day_count",label:"Day",render:n=>n??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:Ct},{type:"select",name:"month",label:"Bulan",options:l},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:s}],formFields:n=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:n?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Ct,value:n?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:n?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:n?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:n?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(n?.employee_name),value:n?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(n?.fc_specialist),value:n?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:n?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:n?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date},{name:"day_count",label:"Day",type:"number",value:n?.day_count,readonly:!0}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let n=await C(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let c=n.data.data.map(v=>({"Tanggal Info":v.report_date||"",Cabang:v.branch_name||"",Kategori:v.category||"","Sumber Laporan":v.source||"",Keluhan:v.complaint||"","Nama FC":v.employee_name||"","FC Spesialis":v.fc_specialist||"",Solusi:v.solution||"",Status:v.status||"","Tanggal Selesai":v.completion_date||"",Day:v.day_count!==null?v.day_count:""}));L(c,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan",Status:"Done","Tanggal Selesai":"2024-03-02",Day:1}],"Template_Import_Permasalahan")},onImport:async n=>{let v=(await C("/api/branches?all=1")).data?.data||[],b=h=>{if(!h)return null;let f=String(h||"").toLowerCase(),m=v.find(d=>String(d.full_name||"").toLowerCase()===f||String(d.code||"").toLowerCase()===f||String(d.name||"").toLowerCase()===f);return m?m.id:null},u=n.map(h=>({branch_id:b(String(h.Cabang||"").trim()),report_date:String(h["Tanggal Info"]||h.Tanggal||"").trim(),category:String(h.Kategori||"").trim(),source:String(h["Sumber Laporan"]||h.Sumber||"").trim(),complaint:String(h.Keluhan||"").trim(),employee_name:String(h["Nama FC"]||"").trim(),fc_specialist:String(h["FC Spesialis"]||"").trim(),solution:String(h.Solusi||"").trim(),completion_date:String(h["Tanggal Selesai"]||h["Tgl Selesai"]||"").trim(),status:String(h.Status||"").trim(),day_count:h.Day||h.Hari||null})).filter(h=>h.report_date&&h.complaint&&h.category),p=await C("/api/import/issues",{method:"POST",body:JSON.stringify({rows:u,onDuplicate:"update"})});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}}})}N();var Me=[];function Ja(t,e){let i=String(t.status||"").toLowerCase();return e==="pending"?i==="pending":!1}async function Jt(t,e){let i=e?e.get("dash_filter"):null;Me=await H();let a=await ae(),r=["Ade","Berlin"],s=l=>l&&!a.find(n=>n.value===l)?[...a,{value:l,label:l}]:a,o=l=>l&&!r.find(n=>(typeof n=="object"?n.value:n)===l)?[...r,l]:r;F({container:t,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:l=>i?l.filter(n=>Ja(n,i)):l,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:l=>`<span title="${l||""}">${l?.length>50?l.slice(0,50)+"\u2026":l||"-"}</span>`},{key:"solution",label:"Solusi",render:l=>`<span title="${l||""}">${l?.length>40?l.slice(0,40)+"\u2026":l||"-"}</span>`},{key:"status",label:"Status",render:l=>Q(l)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:l=>window.formatDate(l)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:l=>l?`<a href="${l}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:Me},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async l=>{let n=new URLSearchParams(l||{}).toString(),c=await C(`/api/one-on-one?limit=10000&${n}`);if(c.ok){let v=c.data.data.map(u=>({Tanggal:u.meeting_date||"",Cabang:u.branch_name||"","Nama Karyawan":u.employee_name||"",PIC:u.pic||"",Masalah:u.problem||"",Solusi:u.solution||"",Status:u.status||"","Tgl Selesai":u.completion_date||"",Dokumen:u.document_link||""})),{downloadExcel:b}=await Promise.resolve().then(()=>(K(),le));b(v,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let l=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(K(),le));n(l,"Template_Import_OneOnOne")},onImport:async l=>{let n=u=>{if(!u)return null;let p=String(u||"").toLowerCase(),h=Me.find(f=>String(f.label||"").toLowerCase()===p);return h?h.value:null},c=u=>{if(!u)return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let p=String(u).trim();if(/^\d{4,5}$/.test(p)){let f=Number(p);if(f>2e4&&f<99999){let m=new Date(Date.UTC(1899,11,30)+f*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);let h=p.split(/[\/\-\.]/);if(h.length===3){let[f,m,d]=h.map(g=>g.trim());if(f.length===4&&m.length<=2&&d.length<=2)return`${f}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&m.length<=2&&f.length<=2)return`${d}-${m.padStart(2,"0")}-${f.padStart(2,"0")}`}return p},v=l.map(u=>({meeting_date:c(u.Tanggal),employee_name:String(u["Nama Karyawan"]||"").trim(),branch_id:n(String(u.Cabang||"").trim()),pic:String(u.PIC||"").trim(),problem:String(u.Masalah||"").trim(),solution:String(u.Solusi||"").trim(),status:String(u.Status||"").trim(),completion_date:c(u["Tgl Selesai"]),document_link:String(u.Dokumen||"").trim()})).filter(u=>u.meeting_date&&u.employee_name&&u.branch_id),b=await C("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:v,onDuplicate:"update"})});if(!b.ok)throw new Error(b.data?.error||"Import gagal");return b.data}},formFields:l=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:l?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:l?.branch_id&&!Me.find(n=>n.value==l.branch_id)?[...Me,{value:l.branch_id,label:l.branch_name||l.branch_id}]:Me,createApi:{path:"/api/branches",field:"full_name"},value:l?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:s(l?.employee_name),value:l?.employee_name},{name:"pic",label:"PIC",type:"select",options:o(l?.pic),createApi:{path:"/api/pic",field:"name"},value:l?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:l?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:l?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:l?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:l?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:l?.document_link}]})}N();async function Ut(t){let e=await H(),i=await ae(),a=["Ade Surahman","Berlin Ariansyah","Mizwar","Fajar","Ade","Berlin"],r=n=>n&&!i.find(c=>c.value===n)?[...i,{value:n,label:n}]:i,s=n=>n&&!a.find(c=>(typeof c=="object"?c.value:c)===n)?[...a,n]:a,o=Array.from({length:5},(n,c)=>String(new Date().getFullYear()-c));F({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:n=>{try{let c=JSON.parse(n);return Array.isArray(c)?c.join(", "):n||"-"}catch{return n||"-"}}},{key:"score",label:"Nilai",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer / peserta..."},{type:"select",name:"batch",label:"Batch",options:["Batch 1","Batch 2","Batch 3","Batch 4","Batch 5"]},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"trainer",label:"Trainer",options:["Berlin Ariansyah"]},{type:"select",name:"year",label:"Tahun",options:o}],exportOptions:{moduleName:"training",onExport:async n=>{let c=new URLSearchParams(n||{}).toString(),v=await C(`/api/training?limit=10000&${c}`);if(v.ok){let b=v.data.data.map(p=>{let h=p.participants||"";try{let f=JSON.parse(h);h=Array.isArray(f)?f.join(", "):h}catch{}return{Tanggal:p.training_date||"",Batch:p.batch||"",Materi:p.subject||"",Cabang:p.branch_name||"",Trainer:p.trainer||"",Peserta:h,Nilai:p.score!==null&&p.score!==void 0?p.score:"",Dokumen:p.document_link||""}}),{downloadExcel:u}=await Promise.resolve().then(()=>(K(),le));u(b,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:c}=await Promise.resolve().then(()=>(K(),le));c(n,"Template_Import_Training")},onImport:async n=>{let c=p=>{if(!p)return null;let h=String(p||"").toLowerCase(),f=e.find(m=>String(m.label||"").toLowerCase()===h);return f?f.value:null},v=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let h=String(p).trim();if(/^\d{4,5}$/.test(h)){let m=Number(h);if(m>2e4&&m<99999){let d=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);let f=h.split(/[\/\-\.]/);if(f.length===3){let[m,d,g]=f.map(x=>x.trim());if(m.length===4&&d.length<=2&&g.length<=2)return`${m}-${d.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&d.length<=2&&m.length<=2)return`${g}-${d.padStart(2,"0")}-${m.padStart(2,"0")}`}return h},b=n.map(p=>({training_date:v(p.Tanggal),batch:String(p.Batch||"").trim(),subject:String(p.Materi||"").trim(),branch_id:c(String(p.Cabang||"").trim()),trainer:String(p.Trainer||"").trim(),participants:String(p.Peserta||"").trim(),score:p.Nilai?Number(p.Nilai):null,document_link:String(p.Dokumen||"").trim()})).filter(p=>p.training_date&&p.subject&&p.branch_id),u=await C("/api/training/import",{method:"POST",body:JSON.stringify(b)});if(!u.ok)throw new Error(u.data?.error||"Import gagal")}},formFields:n=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:n?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:n?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:n?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:n?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:s(n?.trainer),value:n?.trainer}]},{name:"participants",label:"Peserta",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let c=JSON.parse(n?.participants);return Array.isArray(c)?c.join(", "):n?.participants||""}catch{return n?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:n?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:n?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],onBeforeSubmit:async n=>n})}N();ye();K();function Gt({container:t,title:e,icon:i,apiPath:a,columns:r,formFields:s,filterFields:o,defaultFilters:l={},itemLabel:n="Data",canCreate:c=!0,canEdit:v=!0,canDelete:b=!0,onBeforeSubmit:u,onAfterLoad:p,onDataLoaded:h,extraActions:f=[],initialSearch:m="",exportOptions:d=null,bulkDelete:g=!1,paginationMode:x="server"}){let w=te();w&&w.role==="viewer"&&(c=!1,v=!1,b=!1,g=!1,d=null),w&&w.role==="editor_khusus"&&a!=="/api/issues"&&a!=="/api/overtime"&&(c=!1,v=!1,b=!1,g=!1,d=null),w&&w.role==="input_lembur"&&(a==="/api/overtime"?(b=!1,g=!1,d=null):(c=!1,v=!1,b=!1,g=!1,d=null));let k=1,$={...l};m&&($.search=m);let E=new Set;t.innerHTML=`
    ${g?`
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
        ${c?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
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
    

    ${o&&o.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${o.filter(y=>y.type==="search").map(y=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${y.placeholder||"Cari..."}" id="filter-search" value="${$.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${o.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let _=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select" name="${y.name}" id="filter-${y.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${_}</option>${(y.options||[]).map(S=>`<option value="${typeof S=="object"?S.value:S}" ${$[y.name]===(typeof S=="object"?S.value:S)?"selected":""}>${typeof S=="object"?S.label:S}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let _=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select filter-select-sheet" name="${y.name}-sheet" id="filter-sheet-${y.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${_}</option>${(y.options||[]).map(S=>`<option value="${typeof S=="object"?S.value:S}" ${$[y.name]===(typeof S=="object"?S.value:S)?"selected":""}>${typeof S=="object"?S.label:S}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function T(){let y=document.getElementById("bulk-toolbar");if(!y)return;let _=document.getElementById("bulk-count"),S=document.getElementById("btn-bulk-delete"),R=document.getElementById("btn-bulk-cancel");_.textContent=`${E.size} item dipilih`,E.size>0?(y.style.display="flex",S.disabled=!1,R.disabled=!1):(y.style.display="none",S.disabled=!0,R.disabled=!0);let M=document.getElementById("select-all-checkbox");if(M){let V=document.querySelectorAll(".row-checkbox");if(V.length>0){let B=[...V].every(oe=>oe.checked),A=[...V].some(oe=>oe.checked);M.checked=B,M.indeterminate=A&&!B}else M.checked=!1,M.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{E.clear(),document.querySelectorAll(".row-checkbox").forEach(_=>_.checked=!1);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1),T()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(E.size===0)return;let y=[...E],_=document.createElement("div");_.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",_.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${y.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${y.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(_),_.querySelector("#bulk-cancel-btn").addEventListener("click",()=>_.remove()),_.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let S=_.querySelector("#bulk-confirm-btn");S.disabled=!0,S.textContent="Menghapus...";let R=await C(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:y})});_.remove(),R.ok?(J(`${y.length} ${n} berhasil dihapus.`),E.clear(),T(),ce(a),P()):j(R.data?.error||"Gagal menghapus data.")})});let D=document.getElementById("filter-search"),I;D?.addEventListener("input",y=>{clearTimeout(I),I=setTimeout(()=>{$.search=y.target.value,k=1,E.clear(),T(),P()},400)}),o?.forEach(y=>{(y.type==="select"||y.type==="combobox")&&(document.getElementById(`filter-${y.name}`)?.addEventListener("change",_=>{$[y.name]=_.target.value;let S=document.getElementById(`filter-sheet-${y.name}`);S&&(S.value=_.target.value),k=1,E.clear(),T(),P()}),document.getElementById(`filter-sheet-${y.name}`)?.addEventListener("change",_=>{$[y.name]=_.target.value;let S=document.getElementById(`filter-${y.name}`);S&&(S.value=_.target.value),k=1,E.clear(),T(),P(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{$={},D&&(D.value=""),o?.forEach(y=>{let _=document.getElementById(`filter-${y.name}`);_&&(_.value="");let S=document.getElementById(`filter-sheet-${y.name}`);S&&(S.value="")}),k=1,E.clear(),T(),P()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{$={},D&&(D.value=""),o?.forEach(y=>{let _=document.getElementById(`filter-${y.name}`);_&&(_.value="");let S=document.getElementById(`filter-sheet-${y.name}`);S&&(S.value="")}),k=1,E.clear(),T(),P(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>fe(null)),d&&document.addEventListener("click",function(y){let _=document.getElementById("aksi-menu-main"),S=document.getElementById("btn-aksi-main");_&&S&&!S.contains(y.target)&&!_.contains(y.target)&&_.classList.remove("show-aksi-menu")});let O=document.getElementById("btn-mobile-filter"),X=document.getElementById("filter-options-wrapper"),Y=document.getElementById("btn-close-filter-sheet");if(O&&X&&(O.addEventListener("click",y=>{y.preventDefault(),X.classList.add("sheet-open")}),Y&&Y.addEventListener("click",y=>{y.preventDefault(),X.classList.remove("sheet-open")})),d){document.getElementById(`btn-export-${d.moduleName}`)?.addEventListener("click",async _=>{let S=_.target,R=S.innerHTML;S.innerHTML="\u23F3 Loading...",S.disabled=!0;try{await d.onExport()}catch{j("Gagal export data")}finally{S.innerHTML=R,S.disabled=!1}}),document.getElementById(`btn-template-${d.moduleName}`)?.addEventListener("click",()=>{d.onTemplate()});let y=document.getElementById(`input-import-${d.moduleName}`);y?.addEventListener("change",async _=>{let S=_.target.files[0];if(!S)return;let R=document.getElementById(`label-import-${d.moduleName}`),M=R?R.querySelector(".import-text"):null,V=M?M.innerText:"";M&&(M.innerText="\u231B Memproses..."),R&&(R.style.pointerEvents="none"),y.disabled=!0;try{let B=await Ue(S);if(B.length===0)throw new Error("File kosong atau format salah");await d.onImport(B),J("Import berhasil!"),ce(a),P()}catch(B){j(B.message||"Gagal import data")}finally{M&&(M.innerText=V),R&&(R.style.pointerEvents="auto"),y.disabled=!1,y.value=""}})}async function P(){T();let y=document.getElementById("table-container");if(!y)return;y.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let _=x==="client",S=_?1:k,R=_?De:20,M=new URLSearchParams({page:S,limit:R,...Object.fromEntries(Object.entries($).filter(([,U])=>U))}),V=await C(`${a}?${M}`);if(!V.ok){y.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${V.data?.error||"Error"}</p></div>`;return}let B=V.data?.data||V.data||[],A=V.data?.pagination,oe=B.length;if(_){B=h(B);let U=B.length,W=20,ie=Math.ceil(U/W);k>ie&&ie>0&&(k=ie);let z=(k-1)*W,de=k*W;B=B.slice(z,de),A={page:k,limit:W,total:U,pages:ie}}!1,p&&p(B);let ke=Ze({columns:r,data:B,onEdit:v?U=>fe(U):null,actions:f.map(U=>({...U,handler:W=>U.handler(W,P)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:g?{selectedIds:E,onToggle:T}:null});y.innerHTML="",y.appendChild(ke);let pe=document.getElementById("pagination-container");if(pe&&(pe.innerHTML="",A&&A.pages>1)){let U=et({page:A.page,pages:A.pages,total:A.total,limit:A.limit,onPage:W=>{k=W,P()}});U&&pe.appendChild(U)}}function Te(y){let _=typeof s=="function"?s(y):s;return Je(_)}function fe(y){let _=!!y,S=document.createElement("form");if(S.noValidate=!0,S.innerHTML=Te(y),_){let M=typeof s=="function"?s(y):s;at(S,y)}let{close:R}=se({title:_?`Edit ${n}`:`Tambah ${n}`,content:S,size:"lg",confirmText:_?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(M,V)=>{if(!S.reportValidity())return;let B=M.querySelector(".modal-confirm");B.disabled=!0,B.textContent="Menyimpan...";let A=tt(S),oe=typeof s=="function"?s(y):s,ke=async ie=>{for(let z of ie)if(z.type==="row")await ke(z.fields);else if(z.type==="combobox"&&A[z.name]){let de=A[z.name],we=(z.options||[]).find(Z=>{let re=String(typeof Z=="object"?Z.value:Z),ut=String(typeof Z=="object"?Z.label:Z);return re===de||ut===de});if(we)A[z.name]=typeof we=="object"?we.value:we;else if(z.createApi){let Z={};Z[z.createApi.field]=de,z.createApi.extra&&Object.assign(Z,z.createApi.extra);let re=await C(z.createApi.path,{method:"POST",body:JSON.stringify(Z)});if(re.ok&&re.data?.id)A[z.name]=re.data.id;else if(re.ok&&!re.data?.id)A[z.name]=de;else throw new Error(`Gagal membuat master data: ${re.data?.error||"Unknown error"}`)}}};try{await ke(oe)}catch(ie){j(ie.message),B.disabled=!1,B.textContent=_?"Simpan Perubahan":`Tambah ${n}`;return}u&&(A=await u(A,y));let pe=_?"PUT":"POST",U=_?`${a}/${y.id}`:a,W=await C(U,{method:pe,body:JSON.stringify(A)});W.ok?(J(_?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),V(),ce(a),P()):(j(W.data?.error||"Gagal menyimpan data."),B.disabled=!1,B.textContent=_?"Simpan Perubahan":`Tambah ${n}`)}})}function ne(y){Se(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let _=await C(`${a}/${y.id}`,{method:"DELETE"});_.ok?(J(`${n} berhasil dihapus.`),ce(a),P()):j(_.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return P(),P}N();K();async function Vt(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await H(),a=await ae(),r=e?e.get("dash_filter"):null,s={};if(r==="reliever"){let u=new Date,p=`${u.getFullYear()}-${String(u.getMonth()+1).padStart(2,"0")}`;s={status:"Done",month:e&&e.get("month")?e.get("month"):p}}console.log("RAW",await Ee()),console.log("OPTIONS",a);let o=u=>u&&!a.find(p=>p.value===u)?[...a,{value:u,label:u}]:a,l=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],n=u=>u&&!l.includes(u)?[...l,u]:l,c=new Date().getFullYear(),b=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((u,p)=>{let h=String(p+1).padStart(2,"0");return{value:`${c}-${h}`,label:`${u} ${c}`}});Gt({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:s,onDataLoaded:u=>u.sort((p,h)=>{let f=p.backup_date?new Date(p.backup_date).getTime():0;return(h.backup_date?new Date(h.backup_date).getTime():0)-f}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:u=>me(u)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:u=>window.formatDate(u)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:u=>window.formatDate(u)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:u=>u?`<span class="badge badge-info">${u}</span>`:"-"},{key:"status",label:"Status",render:u=>Q(u)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:l},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"month",label:"Bulan",options:b},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:u=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:u?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:u?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:o(u?.original_fc_name),value:u?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:n(u?.reliever_name),value:u?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:u?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:u?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:u?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:u?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:u?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let u=await C(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(u.ok){let p=u.data.data.map(h=>({Cabang:h.branch_name||"","Nama Facility care":h.original_fc_name||"",Periode:h.period||"",Relifer:h.reliever_name||"","Tanggal Back Up":h.backup_date||"","Tanggal Selesai":h.completion_date||"",Keterangan:h.reason||"",Shift:h.shift||"",Status:h.status||""}));p.length===0&&p.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),L(p,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async u=>{let h=(await C("/api/branches?all=1")).data?.data||[],f=g=>{if(!g)return null;let x=String(g||"").toLowerCase(),w=h.find(k=>String(k.full_name||"").toLowerCase()===x||String(k.code||"").toLowerCase()===x||String(k.name||"").toLowerCase()===x);return w?w.id:null},m=u.map(g=>({branch_name:String(g.Cabang||"").trim(),backup_date:String(g["Tanggal Back Up"]||g["Tanggal Backup"]||"").trim(),original_fc_name:String(g["Nama Facility care"]||g["FC Digantikan"]||"").trim(),reliever_name:String(g.Relifer||g.Reliefer||"").trim(),period:String(g.Periode||"").trim(),reason:String(g.Keterangan||"").trim(),shift:String(g.Shift||"").trim(),completion_date:String(g["Tanggal Selesai"]||"").trim(),status:String(g.Status||"").trim()})).filter(g=>g.reliever_name&&g.backup_date),d=await C("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}}})}N();K();async function Qt(t){let e=await H(),i=Array.from({length:4},(a,r)=>String(new Date().getFullYear()-r));F({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>me(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>Q(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),s=await C(`/api/reports/inspection?limit=10000&${r}`);if(s.ok){let o=s.data.data.map(l=>({Cabang:l.branch_name||"",Periode:l.period||"",Tanggal:l.inspection_date||"","Point FC":l.fc_score!==null&&l.fc_score!==void 0?l.fc_score:"","Point SPV":l.spv_score!==null&&l.spv_score!==void 0?l.spv_score:"",Status:l.status||"","Link Dokumen":l.document_link||""}));L(o,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let r=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),v=e.find(b=>String(b.label||"").toLowerCase()===c);return v?v.value:null},s=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let c=String(n).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let b=Number(c);if(b>2e4&&b<99999){let u=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}let v=c.split(/[\/\-\.]/);if(v.length===3){let[b,u,p]=v.map(h=>h.trim());if(b.length===4&&u.length<=2&&p.length<=2)return`${b}-${u.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&u.length<=2&&b.length<=2)return`${p}-${u.padStart(2,"0")}-${b.padStart(2,"0")}`}return c},o=a.map(n=>({branch_id:r(String(n.Cabang||"").trim()),period:String(n.Periode||"").trim(),inspection_date:s(n.Tanggal),fc_score:n["Point FC"]!==void 0&&n["Point FC"]!==""?Number(n["Point FC"]):null,spv_score:n["Point SPV"]!==void 0&&n["Point SPV"]!==""?Number(n["Point SPV"]):null,status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.inspection_date),l=await C("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}N();K();async function Yt(t){let e=await H(),i=Array.from({length:4},(a,r)=>String(new Date().getFullYear()-r));F({container:t,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>me(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>Q(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),s=await C(`/api/reports/cleaning?limit=10000&${r}`);if(s.ok){let o=s.data.data.map(l=>({Cabang:l.branch_name||"",Jenis:l.activity_type||"",Periode:l.period||"",Tanggal:l.activity_date||"",Status:l.status||"","Link Dokumen":l.document_link||""}));L(o,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let r=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),v=e.find(b=>String(b.label||"").toLowerCase()===c);return v?v.value:null},s=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let c=String(n).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let b=Number(c);if(b>2e4&&b<99999){let u=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}let v=c.split(/[\/\-\.]/);if(v.length===3){let[b,u,p]=v.map(h=>h.trim());if(b.length===4&&u.length<=2&&p.length<=2)return`${b}-${u.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&u.length<=2&&b.length<=2)return`${p}-${u.padStart(2,"0")}-${b.padStart(2,"0")}`}return c},o=a.map(n=>({branch_id:r(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:s(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),l=await C("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}N();K();async function Wt(t,e){let i=await H(),a=Array.from({length:4},(o,l)=>String(new Date().getFullYear()-l)),r=e?e.get("dash_filter"):null,s={};if(r==="fogging"){let o=new Date,l=String(o.getMonth()+1).padStart(2,"0"),n=String(o.getFullYear()),c=e?e.get("month"):null;c&&c.length===7&&(n=c.split("-")[0],l=c.split("-")[1]),s={status:"Done",month:l,year:n}}F({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:s,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:o=>`<span class="badge badge-warning">${o}</span>`},{key:"period",label:"Periode",render:o=>me(o)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>Q(o)},{key:"document_link",label:"Dokumen",render:o=>o?`<a href="${o}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:o=>o||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:i,value:o?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:o?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:o?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:o?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:o?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async o=>{let l=new URLSearchParams(o||{}).toString(),n=await C(`/api/reports/fogging?limit=10000&${l}`);if(n.ok){let c=n.data.data.map(v=>({Cabang:v.branch_name||"",Jenis:v.activity_type||"Fogging",Periode:v.period||"",Tanggal:v.activity_date||"",Status:v.status||"","Link Dokumen":v.document_link||""}));L(c,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async o=>{let l=b=>{if(!b)return null;let u=String(b||"").toLowerCase(),p=i.find(h=>String(h.label||"").toLowerCase()===u);return p?p.value:null},n=b=>{if(b==null||b==="")return"";if(b instanceof Date&&!isNaN(b.getTime()))return b.toISOString().slice(0,10);let u=String(b).trim();if(u===""||u==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);if(/^\d{4,5}$/.test(u)){let h=Number(u);if(h>2e4&&h<99999){let f=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}let p=u.split(/[\/\-\.]/);if(p.length===3){let[h,f,m]=p.map(d=>d.trim());if(h.length===4&&f.length<=2&&m.length<=2)return`${h}-${f.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&f.length<=2&&h.length<=2)return`${m}-${f.padStart(2,"0")}-${h.padStart(2,"0")}`}return u},c=o.map(b=>({branch_id:l(String(b.Cabang||"").trim()),activity_type:String(b.Jenis||b.Kegiatan||"Fogging").trim(),period:String(b.Periode||"").trim(),activity_date:n(b.Tanggal),status:String(b.Status||"").trim(),document_link:String(b["Link Dokumen"]||"").trim(),notes:String(b.Catatan||b.Keterangan||"").trim()})).filter(b=>b.branch_id&&b.period&&b.activity_date),v=await C("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(c)});if(!v.ok)throw new Error(v.data?.error||"Import gagal");return v.data}}})}N();K();async function Xt(t){let e=await H(),i=await ae(),a=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],r=Array.from({length:4},(l,n)=>String(new Date().getFullYear()-n)),s=l=>l&&!i.find(n=>n.value===l)?[...i,{value:l,label:l}]:i,o=l=>l&&!a.find(n=>n.value===l)?[...a,{value:l,label:l}]:a;F({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:l=>`<span title="${l||""}">${l?.length>60?l.slice(0,60)+"\u2026":l||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>Q(l)},{key:"notes",label:"Keterangan",render:l=>l?.length>40?l.slice(0,40)+"\u2026":l||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:l=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:l?.branch_id},{name:"pic",label:"PIC",type:"select",options:o(l?.pic),value:l?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:l?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:l?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:l?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:l?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async l=>{let n=new URLSearchParams(l||{}).toString(),c=await C(`/api/reports/basecamp?limit=10000&${n}`);if(c.ok){let v=c.data.data.map(b=>({"Tgl Info":b.info_date||"",Cabang:b.branch_name||"",Permasalahan:b.problem||"",PIC:b.pic||"","Tgl Done":b.done_date||"",Status:b.status||"",Keterangan:b.notes||""}));L(v,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async l=>{let n=u=>{if(!u)return null;let p=String(u||"").toLowerCase(),h=e.find(f=>String(f.label||"").toLowerCase()===p);return h?h.value:null},c=u=>{if(u==null||u==="")return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let p=String(u).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let f=Number(p);if(f>2e4&&f<99999){let m=new Date(Date.UTC(1899,11,30)+f*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}let h=p.split(/[\/\-\.]/);if(h.length===3){let[f,m,d]=h.map(g=>g.trim());if(f.length===4&&m.length<=2&&d.length<=2)return`${f}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&m.length<=2&&f.length<=2)return`${d}-${m.padStart(2,"0")}-${f.padStart(2,"0")}`}return p},v=l.map(u=>({info_date:c(u["Tgl Info"]||u["Tanggal Info"]),branch_id:n(String(u.Cabang||"").trim()),problem:String(u.Permasalahan||"").trim(),pic:String(u.PIC||"").trim(),done_date:c(u["Tgl Done"]||u["Tanggal Done"]),status:String(u.Status||"").trim(),notes:String(u.Keterangan||u.Catatan||"").trim()})).filter(u=>u.info_date&&u.branch_id&&u.problem),b=await C("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(v)});if(!b.ok)throw new Error(b.data?.error||"Import gagal");return b.data}}})}async function Zt(t){F({container:t,title:"SOP",icon:"\u{1F4DA}",apiPath:"/api/sop",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(N(),Ie)),r=await a(`/api/sop?limit=10000&${i}`);if(r.ok){let s=r.data.data.map(l=>({"Nama SOP":l.name||"",Kategori:l.category||"",Dokumen:l.document_link||"",Catatan:l.notes||l.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(K(),le));o(s,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(K(),le));i(e,"Template_Import_SOP")},onImport:async e=>{let i=e.map(s=>({name:String(s["Nama SOP"]||"").trim(),category:String(s.Kategori||"").trim(),document_link:String(s.Dokumen||"").trim(),description:String(s.Catatan||"").trim()})).filter(s=>s.name),{apiFetch:a}=await Promise.resolve().then(()=>(N(),Ie)),r=await a("/api/sop/import",{method:"POST",body:JSON.stringify(i)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function ea(t){F({container:t,title:"Master Checklist",icon:"\u{1F4CB}",apiPath:"/api/checklist",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(N(),Ie)),r=await a(`/api/checklist?limit=10000&${i}`);if(r.ok){let s=r.data.data.map(l=>({"Nama Checklist":l.name||"",Kategori:l.category||"",Dokumen:l.document_link||"",Deskripsi:l.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(K(),le));o(s,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(K(),le));i(e,"Template_Import_Checklist")},onImport:async e=>{let i=e.map(s=>({name:String(s["Nama Checklist"]||"").trim(),category:String(s.Kategori||"").trim(),document_link:String(s.Dokumen||"").trim(),description:String(s.Deskripsi||"").trim()})).filter(s=>s.name),{apiFetch:a}=await Promise.resolve().then(()=>(N(),Ie)),r=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(i)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}N();K();var Ua=["Lantai & Sudut Ruangan","Dinding, Partisi & Plint","Kaca, Cermin & Partisi Kaca","Sanitair & Area Basah Toilet","Plafon, Lampu & Kisi Ventilasi","Mebel, Meja & Kursi Kerja/Pasien","Peralatan Khusus & Medis Non-Steril","Wadah Sampah & Utilitas","Teras","Lobby 1","Receptionist","Snack Corner","DU 101","Ronsen","Toilet 1","Janitor","Ruang Limbah","Ruang Steril","Area Tangga","Lobby 2","Playground","Musholla","Tempat Wudhu","Ruang Dokter","Ruang Team","Gudang Farmasi & Bahan","RoofTop","Ruang Kompressor","Lift","Area Genset","Area Tangga Exit","Taman","Pos Security","Toilet Security","Halaman Parkir"];async function ta(t){let e=Ua;try{let i=await C("/api/hygiene-standards/rooms");i.ok&&Array.isArray(i.data)&&i.data.length>0&&(e=i.data)}catch{}F({container:t,title:"Master Hygiene",icon:"\u2728",apiPath:"/api/hygiene-standards",bulkDelete:!1,enableMobileFilterSheet:!0,itemLabel:"Parameter Kebersihan",columns:[{key:"room_name",label:"Ruangan",render:i=>`<span class="badge badge-info" style="font-weight:600; white-space:nowrap;">${i||"-"}</span>`},{key:"item_name",label:"Item / Objek",render:i=>`<strong style="color:var(--text-1);">${i||"-"}</strong>`},{key:"cleanliness_standard",label:"Standar Kebersihan",render:i=>`<span style="line-height:1.5; color:var(--text-2);">${i||"-"}</span>`}],filterFields:[{type:"search",placeholder:"Cari ruangan, item, atau standar kebersihan..."},{type:"select",name:"room_name",label:"Ruangan",options:e}],formFields:i=>[{type:"row",fields:[{name:"room_name",label:"Ruangan / Kategori",type:"select",required:!0,options:e,value:i?.room_name||""},{name:"item_name",label:"Item / Objek",required:!0,placeholder:"Nama item/objek pembersihan",value:i?.item_name||""}]},{name:"cleanliness_standard",label:"Standar Kebersihan Fisik",type:"textarea",required:!0,rows:3,placeholder:"Kondisi fisik yang dipersyaratkan (misal: Bebas debu, tidak lengket, kering...)",value:i?.cleanliness_standard||""}],exportOptions:{moduleName:"hygiene_standards",onExport:async i=>{let a=new URLSearchParams(i||{}).toString(),r=await C(`/api/hygiene-standards?all=1&${a}`);if(r.ok){let o=(r.data?.data||r.data||[]).map(l=>({Ruangan:l.room_name||"","Item/Objek":l.item_name||"","Standar Kebersihan":l.cleanliness_standard||""}));L(o,`Master_Hygiene_FCMS_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data master hygiene")},onTemplate:async()=>{L([{Ruangan:"Teras","Item/Objek":"Sudut/Pojok","Standar Kebersihan":"Bebas noda, kerak lumut, puntung rokok, dan tumpukan kotoran kering"},{Ruangan:"Lobby 1","Item/Objek":"Lantai Lobby","Standar Kebersihan":"Mengkilap bersih, bebas minyak, tidak licin, nat ubin cerah"}],"Template_Master_Hygiene_FCMS")},onImport:async i=>{let a=i.map(s=>({room_name:String(s.Ruangan||s.room_name||"").trim(),item_name:String(s["Item/Objek"]||s.Item||s.item_name||"").trim(),cleanliness_standard:String(s["Standar Kebersihan"]||s.Standar||s.cleanliness_standard||"").trim()})).filter(s=>s.room_name&&s.item_name&&s.cleanliness_standard);if(a.length===0)throw new Error("Tidak ada baris valid untuk diimpor. Pastikan header: Ruangan, Item/Objek, Standar Kebersihan");let r=await C("/api/hygiene-standards/import",{method:"POST",body:JSON.stringify(a)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}N();ye();K();async function $t(t,e="forms"){if(e==="supply")return Va(t);Ga(t)}function Ga(t){F({container:t,title:"Master Form",icon:"\u{1F4D1}",apiPath:"/api/forms",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await C(`/api/forms?limit=10000&${i}`);a.data?.data?L(a.data.data,"Data_Master_Form"):j("Gagal export data master form")},onImport:async e=>{let i=await C("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!i.ok)throw new Error(i.data?.error||"Import failed");return i.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function Va(t){let i=((await C("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));F({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,r)=>r.branch_name_ref||r.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let r=JSON.parse(a);return Array.isArray(r)?r.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let r=JSON.parse(a);return Array.isArray(r)?r.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>Q(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let r=a?.tools_items;try{r=Array.isArray(JSON.parse(r))?JSON.parse(r).join(", "):r}catch{}let s=a?.chemical_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:a?.branch_id&&!i.find(o=>o.value==a.branch_id)?[...i,{value:a.branch_id,label:a.branch_name||a.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:r},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:s},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),s=await C(`/api/reports/supply?limit=10000&${r}`);if(s.ok){let o=s.data.data.map(l=>{let n=l.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let c=l.chemical_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}return{Waktu:l.submitted_at||"",Pengirim:l.submitter_name||"",Cabang:l.branch_name_ref||l.branch_name||"","Alat/Barang":n||"",Chemical:c||"",Catatan:l.additional_notes||"",Status:l.status||"","Diproses Oleh":l.processed_by||""}});L(o,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let s=(await C("/api/branches?all=1")).data?.data||[],o=v=>{if(!v)return null;let b=String(v||"").toLowerCase(),u=s.find(p=>String(p.full_name||"").toLowerCase()===b||String(p.code||"").toLowerCase()===b||String(p.name||"").toLowerCase()===b);return u?u.id:null},l=v=>{if(v==null||v==="")return"";if(v instanceof Date&&!isNaN(v.getTime()))return v.toISOString().slice(0,10);let b=String(v).trim();if(b===""||b==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(b))return b.slice(0,10);if(/^\d{4,5}$/.test(b)){let p=Number(b);if(p>2e4&&p<99999){let h=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}let u=b.split(/[\/\-\.]/);if(u.length===3){let[p,h,f]=u.map(m=>m.trim());if(p.length===4&&h.length<=2&&f.length<=2)return`${p}-${h.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&h.length<=2&&p.length<=2)return`${f}-${h.padStart(2,"0")}-${p.padStart(2,"0")}`}return b},n=a.map(v=>({submitted_at:l(v.Waktu||v.Tanggal),submitter_name:String(v.Pengirim||"").trim(),branch_id:o(String(v.Cabang||"").trim()),tools_items:String(v["Alat/Barang"]||v.Alat||"").trim(),chemical_items:String(v.Chemical||"").trim(),additional_notes:String(v.Catatan||v.Keterangan||"").trim(),status:String(v.Status||"").trim(),processed_by:String(v["Diproses Oleh"]||v.PIC||"").trim()})).filter(v=>v.submitted_at&&v.submitter_name&&v.branch_id),c=await C("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,r)=>{let s=se({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(o,l)=>{let n=o.querySelector("#supply-status").value,c=o.querySelector("#supply-processed-by").value;(await C(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:c})})).ok?(J("Status diperbarui."),l(),r()):j("Gagal update status.")}})}}]})}N();K();async function aa(t){let e=te();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}F({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",editor_khusus:"badge-warning",input_lembur:"badge-info",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let a=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"editor_khusus",label:"Editor (Hanya Lembur & Masalah)"},{value:"input_lembur",label:"Input Lembur (Hanya Isi Lembur)"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await C(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let a=i.data.data.map(r=>({"Nama Lengkap":r.full_name||"",Username:r.username||"",Email:r.email||"",Role:r.role||"",Status:r.is_active?"Aktif":"Nonaktif"}));L(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let a=i.map(s=>({full_name:String(s["Nama Lengkap"]||"").trim(),username:String(s.Username||"").trim(),email:String(s.Email||"").trim(),role:String(s.Role||"").trim()||"viewer",password:String(s.Password||"").trim()})).filter(s=>s.username&&s.password&&s.email&&s.full_name),r=await C("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}N();K();async function na(t){F({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await C(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)L(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let i=e.map(r=>({code:String(r["Kode Cabang"]||"").trim(),name:String(r["Nama Pendek"]||"").trim(),full_name:String(r["Nama Lengkap"]||"").trim(),city:String(r.Kota||"").trim()})).filter(r=>r.code&&r.name),a=await C("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}N();async function ia(t){let e=new Date,i=[];t.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),r()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),r()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(o=>o.addEventListener("change",r));async function a(){try{let o=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;i=(await C(`/api/dashboard/calendar?month=${o}`)).data?.data||[]}catch(o){console.warn("[Calendar] Failed to load events, rendering empty grid:",o),i=[]}}async function r(){let o=document.getElementById("calendar-grid");if(o){o.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let l=e.getFullYear(),n=e.getMonth(),c=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),v=document.getElementById("cal-month-label");v&&(v.textContent=c);let b=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(k=>k.value)),u=i.filter(k=>b.has(k.type)),p={};u.forEach(k=>{let $=Qa(k.event_date);$&&(p[$]||(p[$]=[]),p[$].push(k))});let h=new Date(l,n,1).getDay(),f=new Date(l,n+1,0).getDate(),m=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],d=new Date().toISOString().slice(0,10),g='<div class="calendar-grid">';m.forEach(k=>{g+=`<div class="cal-day-header">${k}</div>`});for(let k=0;k<h;k++)g+='<div class="cal-cell cal-cell-empty"></div>';for(let k=1;k<=f;k++){let $=`${l}-${String(n+1).padStart(2,"0")}-${String(k).padStart(2,"0")}`,E=p[$]||[],T=$===d;g+=`
          <div class="cal-cell ${T?"cal-today":""} ${E.length?"cal-has-events":""}"
               data-date="${$}" tabindex="0" role="button" aria-label="${$}">
            <div class="cal-day-num ${T?"today-num":""}">${k}</div>
            <div class="cal-events-preview">
              ${E.slice(0,3).map(D=>`
                <div class="cal-event-dot cal-color-${D.color||"gray"}" title="${he(D.title||D.type)}">
                  <span class="cal-event-dot-label">${Ya(D.title||D.branch_name||D.type,18)}</span>
                </div>
              `).join("")}
              ${E.length>3?`<div class="cal-more">+${E.length-3} lagi</div>`:""}
            </div>
          </div>`}let w=(h+f)%7;if(w!==0)for(let k=0;k<7-w;k++)g+='<div class="cal-cell cal-cell-empty"></div>';g+="</div>",o.innerHTML=g,o.querySelectorAll(".cal-cell[data-date]").forEach(k=>{k.addEventListener("click",()=>{let $=k.dataset.date,E=p[$]||[];if(!E.length)return;let T=document.getElementById("cal-event-list"),D=new Date($+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=D,document.getElementById("cal-event-items").innerHTML=E.map(I=>`
            <div class="cal-event-item cal-color-border-${I.color||"gray"}">
              <div class="cal-event-type">
                ${Wa(I.type)}
                ${I.date_type?` <span class="badge" style="font-size:0.75rem;padding:2px 6px;background:rgba(0,0,0,0.06);margin-left:4px;">${he(I.date_type)}</span>`:""}
              </div>
              <div class="cal-event-title" style="font-weight:600;font-size:0.95rem;margin-top:2px;">${he(I.title||"-")}</div>
              ${I.branch_name?`<div class="cal-event-branch" style="color:var(--text-2);font-size:0.85rem;margin-top:2px;">\u{1F4CD} ${he(I.branch_name)}</div>`:""}
              ${I.pic||I.period||I.opening_date||I.target_date||I.completion_date||I.notes?`
                <div class="cal-event-meta" style="margin-top:6px;padding-top:6px;border-top:1px dashed var(--border);font-size:0.8rem;color:var(--text-2);display:flex;flex-direction:column;gap:3px;">
                  ${I.pic?`<div>\u{1F464} <b>PIC:</b> ${he(I.pic)}</div>`:""}
                  ${I.period?`<div>\u{1F4C5} <b>Periode:</b> ${he(I.period)}</div>`:""}
                  ${I.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${he(Et(I.opening_date))}</div>`:""}
                  ${I.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${he(Et(I.target_date))}</div>`:""}
                  ${I.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${he(Et(I.completion_date))}</div>`:""}
                  ${I.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${he(I.notes)}</div>`:""}
                </div>
              `:""}
              ${I.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${I.status==="Done"?"badge-success":"badge-warning"}">${he(I.status)}</span></div>`:""}
              ${I.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${I.days_remaining} hari</div>`:""}
            </div>
          `).join(""),T.style.display="block"})})}catch(l){console.error("[Calendar] Render error:",l),o&&(o.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let s=()=>{a().then(()=>r())};nt.on("data:changed",s),r()}function Qa(t){if(!t||t==="-"||String(t).trim()==="")return"";let e=String(t).trim();if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);let i=e.split(/[\/\-\.]/);if(i.length===3){let[a,r,s]=i.map(o=>o.trim());if(s.length===4)return`${s}-${r.padStart(2,"0")}-${a.padStart(2,"0")}`;if(a.length===4)return`${a}-${r.padStart(2,"0")}-${s.padStart(2,"0")}`}return e.slice(0,10)}function Et(t){if(!t||t==="-"||String(t).trim()==="")return"";let i=String(t).trim().split(/[\/\-\.]/);if(i.length===3){let[a,r,s]=i.map(o=>o.trim());if(a.length===4)return`${s.padStart(2,"0")}-${r.padStart(2,"0")}-${a}`;if(s.length===4)return`${a.padStart(2,"0")}-${r.padStart(2,"0")}-${s}`}return t}function Ya(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function he(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Wa(t){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging"}[t]||t}N();async function ra(t){let e=te(),i=(e?.full_name||e?.username||"U")[0].toUpperCase(),r={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
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
  `;let s=localStorage.getItem("fm_token"),o=document.getElementById("session-info");if(s&&o)try{let l=JSON.parse(atob(s.split(".")[1])),n=new Date(l.exp*1e3);o.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{o.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async l=>{l.preventDefault();let n=document.getElementById("pwd-error"),c=document.getElementById("pwd-success"),v=document.getElementById("btn-save-pwd");n.style.display="none",c.style.display="none";let b=l.target,u=b.current_password.value,p=b.new_password.value,h=b.confirm_password.value;if(p!==h){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(p.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}v.disabled=!0,v.textContent="\u23F3 Menyimpan...";let f=await C("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:u,new_password:p})});v.disabled=!1,v.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',f.ok?(c.textContent="\u2705 Password berhasil diubah.",c.style.display="block",b.reset(),J("Password berhasil diubah.")):(n.textContent=f.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}N();var pt={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function ee(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let r=Number(e);if(r>2e4&&r<99999){let s=new Date(Date.UTC(1899,11,30)+r*864e5);return isNaN(s.getTime())?null:s.toISOString().slice(0,10)}}let i=e.split(/[\/\-\.]/);if(i.length===3){let[r,s,o]=i.map(v=>v.trim()),l=Number(r),n=Number(s),c=Number(o);if(r.length===4&&l>1900)return`${r}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`;if(o.length===4&&c>1900)return l>12?`${o}-${s.padStart(2,"0")}-${r.padStart(2,"0")}`:n>12?`${o}-${r.padStart(2,"0")}-${s.padStart(2,"0")}`:`${o}-${s.padStart(2,"0")}-${r.padStart(2,"0")}`;if(o.length===2&&!isNaN(c)){let v=c>=50?`19${o}`:`20${o}`;return l>12?`${v}-${s.padStart(2,"0")}-${r.padStart(2,"0")}`:`${v}-${s.padStart(2,"0")}-${r.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function oa(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Xa={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||t.Divisi||"FACILITY CARE",phone:t["No. Hp"]||t["No. HP"],join_date:ee(t["Tanggal Masuk"]||t["Tgl Masuk"]),status:t.Status||"",target_pindah_os:ee(t["Target Pindah OS"]),target_selesai_os:ee(t["Target Selesai OS"]),notes:t.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:ee(t["Tanggal Mulai"]),end_date:ee(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:ee(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:ee(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:ee(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:ee(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:ee(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:ee(t["Tanggal Target"]||t["Tgl Target"]),completion_date:ee(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:ee(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:ee(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:ee(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:ee(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:ee(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:ee(t["Tanggal Back Up"]),completion_date:ee(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:ee(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:ee(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Za(t,e){let i=pt[t];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Xa[i.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let r=[],s=[],o=[];return e.filter(n=>!oa(n)).forEach((n,c)=>{let v=e.indexOf(n)+2,b=[];a.required.forEach(({key:p,label:h})=>{let f=n[p];if(f==null||String(f).trim()===""){let m=Object.keys(n).filter(d=>d.trim()).join(", ");b.push({column:h,originalValue:f||"",reason:`Kolom "${h}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${m.slice(0,120)}`})}});let u=a.map(n);b.length>0?s.push({row:v,data:u,raw:n,errors:b}):(r.push(n),o.push(u))}),{valid:r,errors:s,mapped:o}}function sa(t){let e=[];return t.SheetNames.forEach(i=>{let a=pt[i];if(!a)return;let r=t.Sheets[i],s=window.XLSX.utils.sheet_to_json(r,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),o=Za(i,s),l=s.filter(n=>!oa(n));e.push({sheetName:i,module:a.module,label:a.label,total:l.length,valid:o.mapped.length,errorCount:o.errors.length,errors:o.errors,mapped:o.mapped,skipped:!1})}),e}function la(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,r])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(r),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function da(t){let e=window.XLSX,i=e.utils.book_new(),a=!1;return t.forEach(r=>{if(!r.errors||r.errors.length===0)return;a=!0;let s=r.errors.map(l=>({"No. Baris":l.row,"Kolom Gagal":(l.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(l.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(l.data||{}).map(([n,c])=>[n,c??""]))})),o=e.utils.json_to_sheet(s);e.utils.book_append_sheet(i,o,r.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var en=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function ca(t){t.innerHTML=`
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
              ${Object.entries(pt).map(([f,{label:m}])=>`<span class="import-sheet-tag">\u{1F4C4} ${f} \u2192 ${m}</span>`).join("")}
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
  `;let e=null,i=null,a=0,r={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function s(f){Object.entries(r).forEach(([m,d])=>{d.style.display=m===f?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let f=document.getElementById("btn-backup-db");f.disabled=!0,f.textContent="\u23F3 Memproses Backup...";try{let m=await C("/api/import/backup");if(m.ok){if(!window.XLSX){j("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let d=window.XLSX,g=d.utils.book_new();Object.entries(m.data.database).forEach(([x,w])=>{let k=w.length>0?w:[{}],$=d.utils.json_to_sheet(k);d.utils.book_append_sheet(g,$,x.substring(0,31))}),d.writeFile(g,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),J("Backup berhasil diunduh!")}else j("Gagal memproses backup: "+(m.data?.error||"Unknown error"))}catch(m){j("Gagal memproses backup: "+m.message)}finally{f.disabled=!1,f.textContent="\u{1F4E6} Backup Database"}});let o=document.getElementById("btn-sync-google");o&&o.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let f=o.innerHTML;o.innerHTML='<span class="spinner"></span> Menyinkronkan...',o.disabled=!0;try{let m=await C("/api/sync/google-sheets",{method:"POST"});m.ok?alert("Sinkronisasi Berhasil: "+(m.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(m.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{o.innerHTML=f,o.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{la(),J("Template Excel berhasil didownload!")});let l=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",f=>{f.stopPropagation(),l.click()}),l.addEventListener("change",f=>{f.target.files[0]&&c(f.target.files[0])}),n.addEventListener("dragover",f=>{f.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",f=>{f.preventDefault(),n.classList.remove("drag-over");let m=f.dataTransfer.files[0];m&&m.name.match(/\.xlsx?$/i)?c(m):j("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,l.value="",document.getElementById("file-info").style.display="none",n.style.display="",s("upload")});async function c(f){e=f,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${f.name} (${(f.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await v(f)}async function v(f){s("validating");let m=document.getElementById("validation-status"),d=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");m.textContent="Membaca file Excel...",d.style.width="20%",await Qe(200);let g=await f.arrayBuffer(),x=window.XLSX.read(g,{type:"array",cellDates:!0});m.textContent=`Memvalidasi ${x.SheetNames.length} sheet...`,d.style.width="50%",await Qe(100),i=sa(x),d.style.width="100%",m.textContent="Validasi selesai!",await Qe(300),b()}catch(g){s("upload"),j("Gagal memproses file: "+g.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function b(){s("preview");let f=i.filter(T=>!T.skipped).length,m=i.reduce((T,D)=>T+D.total,0),d=i.reduce((T,D)=>T+D.valid,0),g=i.reduce((T,D)=>T+D.errorCount,0),x=m>0?Math.round(d/m*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${f} sheet</span>
      <span class="badge badge-secondary">${m} baris</span>
      <span class="badge badge-success">${d} valid (${x}%)</span>
      ${g>0?`<span class="badge badge-danger">${g} error</span>`:""}
    `;let w=document.getElementById("preview-table-container");w.innerHTML=`
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
          ${i.map((T,D)=>`
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
                ${T.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${D}">\u{1F50D} ${T.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,w.querySelectorAll(".btn-detail-error").forEach(T=>{T.addEventListener("click",()=>{let D=i[Number(T.dataset.idx)];u(D)})});let k=document.getElementById("error-detail-section"),$=document.getElementById("error-detail-container");$.innerHTML="",k.style.display="none";let E=document.getElementById("btn-start-import");d===0?(E.disabled=!0,E.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(E.disabled=!1,g>0?(E.innerHTML=`\u{1F680} Import ${d} Data Valid (${g} dilewati)`,E.title="Baris error akan dilewati, baris valid tetap diimport"):E.innerHTML=`\u{1F680} Mulai Import ${d} Data`)}function u(f){let m=document.getElementById("error-detail-section"),d=document.getElementById("error-detail-container");m.style.display="";let g=f.errors.slice(0,100).map(x=>(Array.isArray(x.errors)?x.errors:[]).map(k=>{let $=typeof k=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${x.row}</span></td>
            <td><strong>${$?k.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${$&&k.originalValue!==void 0?k.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${$?k.reason:k}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${$&&k.aliases?`Gunakan salah satu nama kolom:<br><em>${k.aliases}</em>`:$&&k.hint?k.hint:""}
            </td>
          </tr>
        `}).join("")).join("");d.innerHTML=`
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
            <tbody>${g||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${f.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,m.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{s("upload"),document.getElementById("file-info").style.display="none",n.style.display="",e=null,l.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;da(i)?J("Log error berhasil didownload."):J("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let f=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";p(f)});async function p(f){s("importing"),a=Date.now();let m=[];en.forEach(k=>{let $=i?.find(E=>E.module===k&&E.mapped?.length>0);$&&m.push($)});let d=document.getElementById("import-steps-list");d.innerHTML=m.map(k=>`
      <div class="import-step-item" id="step-item-${k.module}">
        <span class="step-item-icon" id="step-icon-${k.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${k.label} <span class="step-item-count">(${k.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${k.module}"></span>
      </div>
    `).join("");let g=document.getElementById("import-bar"),x=document.getElementById("import-current-status"),w={totalSheets:m.length,totalRows:m.reduce((k,$)=>k+$.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let k=0;k<m.length;k++){let $=m[k],E=document.getElementById(`step-icon-${$.module}`),T=document.getElementById(`step-status-${$.module}`);E.textContent="\u{1F504}",T.textContent="Mengimport...",x.textContent=`Mengimport ${$.label}...`,g.style.width=`${Math.round(k/m.length*100)}%`;try{let D=await C(`/api/import/${$.module}`,{method:"POST",body:JSON.stringify({rows:$.mapped,onDuplicate:f})});if(D.ok){let I=D.data;w.inserted+=I.inserted||0,w.skipped+=I.skipped||0,w.moduleResults.push({label:$.label,inserted:I.inserted||0,skipped:I.skipped||0,status:"ok"}),E.textContent="\u2705",T.innerHTML=`<span class="badge badge-success">${I.inserted||0} berhasil</span>${I.skipped>0?` <span class="badge badge-neutral">${I.skipped} skip</span>`:""}`}else w.failed++,w.moduleResults.push({label:$.label,inserted:0,skipped:0,status:"error",error:D.data?.error}),E.textContent="\u274C",T.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(D){w.failed++,w.moduleResults.push({label:$.label,inserted:0,skipped:0,status:"error",error:D.message}),E.textContent="\u274C",T.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Qe(150)}g.style.width="100%",x.textContent="Selesai!",await Qe(400),h(w)}function h(f){s("summary");let m=((Date.now()-a)/1e3).toFixed(1),d=f.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${d?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${d?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
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
        <div class="stat-value">${m}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${f.moduleResults.map(g=>`
            <tr>
              <td>${g.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${g.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${g.skipped}</span></td>
              <td style="text-align:center">
                ${g.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${g.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,i=null,l.value="",document.getElementById("file-info").style.display="none",n.style.display="",s("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Qe(t){return new Promise(e=>setTimeout(e,t))}N();var mt=[],pa=[];async function ma(t){mt=await H(),pa=await ae(),F({container:t,title:"Data SP (Surat Peringatan)",icon:"\u{1F4DC}",apiPath:"/api/sp",enableMobileFilterSheet:!0,itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:mt}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await C(`/api/sp?limit=10000&${i}`);if(a.ok){let r=a.data.data.map(o=>({"Nama Karyawan":o.employee_name||"",Divisi:o.division||"",Cabang:o.branch_name||"","Tanggal Sp":o.tanggal||"","Akhir Sp":o.akhir_sp||"","Jenis Sp":o.sp_type||"","Link Document / Foto":o.document_link||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(K(),le));s(r,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(K(),le));i(e,"Template_Import_SP")},onImport:async e=>{let i=o=>{if(!o)return null;let l=String(o||"").toLowerCase(),n=mt.find(c=>String(c.label||"").toLowerCase()===l);return n?n.value:null},a=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let l=String(o).trim();if(/^\d{4,5}$/.test(l)){let c=Number(l);if(c>2e4&&c<99999){let v=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);let n=l.split(/[\/\-\.]/);if(n.length===3){let[c,v,b]=n.map(u=>u.trim());if(c.length===4&&v.length<=2&&b.length<=2)return`${c}-${v.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&v.length<=2&&c.length<=2)return`${b}-${v.padStart(2,"0")}-${c.padStart(2,"0")}`}return l},r=e.map(o=>({employee_name:String(o["Nama Karyawan"]||"").trim(),division:String(o.Divisi||"").trim(),branch_id:i(String(o.Cabang||"").trim()),tanggal:a(o["Tanggal Sp"]),akhir_sp:a(o["Akhir Sp"]),sp_type:String(o["Jenis Sp"]||"").trim(),document_link:String(o["Link Document / Foto"]||"").trim()})).filter(o=>o.employee_name&&o.branch_id),s=await C("/api/import/sp",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:[{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:pa},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:mt,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}N();var Oe=[],ua=[];async function ga(t){Oe=await H(),ua=await ae(),F({container:t,title:"Data Mutasi",icon:"\u{1F501}",apiPath:"/api/mutasi",enableMobileFilterSheet:!0,itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:Oe},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:Oe}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await C(`/api/mutasi?limit=10000&${i}`);if(a.ok){let r=a.data.data.map(o=>({Tanggal:o.tanggal||"","Nama Karyawan":o.employee_name||"","Cabang Asal":o.from_branch_name||"","Cabang Tujuan":o.to_branch_name||"",Status:o.status||"",Dokumen:o.document_link||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(K(),le));s(r,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(K(),le));i(e,"Template_Import_Mutasi")},onImport:async e=>{let i=o=>{if(!o)return null;let l=String(o||"").toLowerCase(),n=Oe.find(c=>String(c.label||"").toLowerCase()===l);return n?n.value:null},a=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let l=String(o).trim();if(/^\d{4,5}$/.test(l)){let c=Number(l);if(c>2e4&&c<99999){let v=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);let n=l.split(/[\/\-\.]/);if(n.length===3){let[c,v,b]=n.map(u=>u.trim());if(c.length===4&&v.length<=2&&b.length<=2)return`${c}-${v.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&v.length<=2&&c.length<=2)return`${b}-${v.padStart(2,"0")}-${c.padStart(2,"0")}`}return l},r=e.map(o=>({tanggal:a(o.Tanggal),employee_name:String(o["Nama Karyawan"]||"").trim(),from_branch_id:i(String(o["Cabang Asal"]||"").trim()),to_branch_id:i(String(o["Cabang Tujuan"]||"").trim()),status:String(o.Status||"").trim(),document_link:String(o.Dokumen||"").trim()})).filter(o=>o.tanggal&&o.employee_name&&o.from_branch_id&&o.to_branch_id),s=await C("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:ua},{type:"select",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Oe,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Oe,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}N();var Tt=[],ba=[];async function ha(t){Tt=await H(),ba=await Ee(),F({container:t,title:"Data Lembur",icon:"\u23F1\uFE0F",apiPath:"/api/overtime",enableMobileFilterSheet:!0,itemLabel:"Lembur",bulkDelete:!0,columns:[{key:"date",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"branch_name",label:"Cabang Lembur"},{key:"employee_name",label:"Nama Karyawan"},{key:"start_time",label:"Jam Mulai"},{key:"end_time",label:"Jam Selesai"},{key:"break_hours",label:"Istirahat",render:e=>e!=null?`${e} Jam`:"0 Jam"},{key:"total_hours",label:"Total Jam",render:e=>e!=null?`${e} Jam`:"0 Jam"},{key:"reason",label:"Alasan / Ket Lembur"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:Tt}],exportOptions:{moduleName:"overtime_records",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await C(`/api/overtime?limit=10000&${i}`);if(a.ok){let r=a.data.data.map(o=>({Tanggal:o.date||"","Cabang Lembur":o.branch_name||"","Nama Karyawan":o.employee_name||"","Jam Mulai":o.start_time||"","Jam Selesai":o.end_time||"","Istirahat (Jam)":o.break_hours||0,"Total Jam":o.total_hours||0,"Alasan / Ket Lembur":o.reason||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(K(),le));s(r,`Data_Lembur_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")}},formFields:[{type:"date",name:"date",label:"Tanggal",required:!0},{type:"select",name:"branch_id",label:"Cabang Lembur",required:!0,options:Tt},{type:"select",name:"employee_id",label:"Nama Karyawan",required:!0,options:ba},{type:"time",name:"start_time",label:"Jam Mulai",required:!0},{type:"time",name:"end_time",label:"Jam Selesai",required:!0},{type:"number",name:"break_hours",label:"Istirahat (Jam)",required:!0},{type:"number",name:"total_hours",label:"Total Jam (Jam)",required:!0},{type:"select",name:"reason",label:"Alasan / Ket Lembur",required:!0,options:["Cover manpower","Longshift","Pekerjaan urgent","General Cleaning","Deep Cleaning"]}]})}window._overtimeListenerAttached||(window._overtimeListenerAttached=!0,document.body.addEventListener("input",t=>{if(t.target.name==="start_time"||t.target.name==="end_time"||t.target.name==="break_hours"){let e=t.target.closest("form");if(!e)return;let i=e.querySelector('[name="start_time"]'),a=e.querySelector('[name="end_time"]'),r=e.querySelector('[name="break_hours"]'),s=e.querySelector('[name="total_hours"]');if(!i||!a||!r||!s)return;let o=i.value,l=a.value,n=parseFloat(r.value)||0;if(o&&l){let[c,v]=o.split(":").map(Number),[b,u]=l.split(":").map(Number),p=b*60+u-(c*60+v);p<0&&(p+=24*60);let h=p/60-n;h<0&&(h=0),s.value=Number.isInteger(h)?h:h.toFixed(2)}}}));N();async function fa(t){let e=te();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}F({container:t,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:i=>new Date(i).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(i,a)=>`<strong>${i||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${a.user_role||""}</small>`},{key:"action",label:"Aksi",render:i=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[i]||"badge-neutral"}">${i}</span>`},{key:"module",label:"Modul",render:i=>`<span style="text-transform:capitalize">${(i||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(i,a)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${i}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async i=>{try{let s=((await(await fetch(`/api/audit-logs?search=${i}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(c=>String(c.id)===String(i));if(!s)return alert("Data tidak ditemukan");let o=c=>{if(!c)return"Tidak ada data";try{return JSON.stringify(JSON.parse(c),null,2)}catch{return c}},l=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${o(s.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${o(s.new_data)}</pre>
           </div>
         </div>
       `,{createModal:n}=await Promise.resolve().then(()=>(ye(),yt));n({title:`Detail Audit Log #${i}`,content:l,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}N();ye();K();var Dt={"2026-09":"September 2026","2026-10":"Oktober 2026","2026-11":"November 2026","2026-12":"Desember 2026"},Ye={scheduled:{label:"Terjadwal",badgeClass:"badge-warning",icon:"\u{1F7E1}",bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},in_progress:{label:"Sedang Proses",badgeClass:"badge-info",icon:"\u{1F535}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},completed:{label:"Selesai",badgeClass:"badge-success",icon:"\u{1F7E2}",bg:"#DCFCE7",color:"#166534",border:"#86EFAC"}};async function ya(t){t.innerHTML=`
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
  `;let e="board",i=null,a=[],r=te(),s=r&&["superadmin","admin","manager","spv"].includes(r.role);C("/api/branches?all=1").then(d=>{d.ok&&d.data?.data&&(a=d.data.data)}).catch(()=>{});async function o(){try{let d=await C("/api/os-migrations/summary");if(!d.ok)throw new Error("Gagal memuat data migrasi");i=d.data?.data||d.data,l(i),c()}catch(d){document.getElementById("migration-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${d.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function l(d){let g=document.getElementById("migration-summary-cards");if(!g||!d)return;let x=d.total_branches||0,w=d.vendors?.AEON?.total||0,k=d.vendors?.["Triple S"]?.total||0,$=d.by_status?.completed||0,E=d.by_status?.in_progress||0,T=d.by_status?.scheduled||0,D=x>0?Math.round($/x*100):0;g.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${x} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.8rem; font-weight:800; color:#1E40AF; line-height:1.2;">${w} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
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
          <span style="font-size:0.85rem; font-weight:700; color:#15803D;">${D}%</span>
        </div>
        <div style="width:100%; height:8px; background:#E2E8F0; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${D}%; height:100%; background:#16A34A; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.75rem; color:var(--text-2); font-weight:500;">
          <span>\u{1F7E2} Selesai: <strong>${$}</strong></span>
          <span>\u{1F535} Proses: <strong>${E}</strong></span>
          <span>\u{1F7E1} Terjadwal: <strong>${T}</strong></span>
        </div>
      </div>
    `}function n(){if(!i||!i.matrix)return[];let d=(document.getElementById("migration-search")?.value||"").toLowerCase().trim(),g=document.getElementById("filter-vendor")?.value||"",x=document.getElementById("filter-month")?.value||"",w=document.getElementById("filter-status")?.value||"";return i.matrix.filter(k=>!(g&&k.vendor!==g||x&&k.target_month!==x||w&&k.status!==w||d&&!`${k.branch_code} ${k.branch_name} ${k.branch_full_name} ${k.vendor} ${k.notes||""} ${k.pic_vendor||""}`.toLowerCase().includes(d)))}function c(){let d=n(),g=document.getElementById("filter-stats-text");g&&(g.innerHTML=`Menampilkan <strong>${d.length}</strong> dari <strong>${i?.total_branches||0}</strong> cabang`);let x=document.getElementById("migration-view-container");e==="board"?v(x,d):b(x,d)}function v(d,g){if(!g.length){d.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;return}let x=["AEON","Triple S"],w=["2026-09","2026-10","2026-11","2026-12"],k="";x.forEach($=>{let E=g.filter(Y=>Y.vendor===$);if(document.getElementById("filter-vendor")?.value&&document.getElementById("filter-vendor").value!==$||E.length===0)return;let T=E.length,D=$==="AEON"?"#1E40AF":"#0F766E",I=$==="AEON"?"#EFF6FF":"#F0FDFA",O=$==="AEON"?"#BFDBFE":"#99F6E4",X=w.filter(Y=>E.some(P=>P.target_month===Y));k+=`
        <div class="vendor-board-section" style="margin-bottom:32px; background:#fff; border:1px solid var(--border); border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <!-- Vendor Section Header -->
          <div style="padding:14px 20px; background:${I}; border-bottom:1px solid ${O}; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.1rem; font-weight:800; color:${D};">${$==="AEON"?"\u{1F3E2} AEON":"\u{1F6E1}\uFE0F Triple S"}</span>
              <span class="badge" style="background:#fff; color:${D}; border:1px solid ${O}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px;">
                ${T} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
              Format Timeline Berdasarkan Dokumen Migrasi
            </div>
          </div>

          <!-- Months Grid (Hanya menampilkan bulan yang memiliki jadwal) -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px; padding:18px;">
            ${X.map(Y=>{let P=E.filter(ne=>ne.target_month===Y);return`
                <div class="month-column" style="background:#F8FAFC; border:1px solid var(--border); border-radius:10px; display:flex; flex-direction:column; min-height:220px;">
                  <!-- Month Header -->
                  <div style="padding:12px 14px; background:#F1F5F9; border-bottom:1px solid var(--border); border-radius:9px 9px 0 0; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">
                      \u{1F4C5} ${(Dt[Y]||Y).split(" ")[0]}
                    </div>
                    <span class="badge" style="background:#E2E8F0; color:var(--text-2); font-weight:700; font-size:0.75rem; padding:2px 7px; border-radius:4px;">
                      ${P.length}
                    </span>
                  </div>

                  <!-- Branch Cards Container -->
                  <div style="padding:10px; display:flex; flex-direction:column; gap:8px; flex:1; max-height:480px; overflow-y:auto;">
                    ${P.map(ne=>{let y=Ye[ne.status]||Ye.scheduled;return`
                        <div class="branch-card" style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 12px; box-shadow:0 1px 2px rgba(0,0,0,0.03); transition:transform 0.15s, box-shadow 0.15s;">
                          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px; margin-bottom:6px;">
                            <div style="font-weight:700; font-size:0.9rem; color:var(--text-1); line-height:1.3;">
                              <span style="color:#2563EB; font-family:monospace; margin-right:4px;">${ne.branch_code}.</span>${ne.branch_name}
                            </div>
                            ${s?`
                              <div style="display:flex; gap:2px;">
                                <button class="btn-icon-mini edit-card-btn" data-id="${ne.id}" title="Edit" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--text-3);">\u270F\uFE0F</button>
                                <button class="btn-icon-mini delete-card-btn" data-id="${ne.id}" data-name="${ne.branch_name}" title="Hapus" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--danger);">\u{1F5D1}\uFE0F</button>
                              </div>
                            `:""}
                          </div>

                          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px; margin-top:8px;">
                            <!-- Status Selector Dropdown -->
                            <div style="position:relative;">
                              <button class="status-toggle-btn" data-id="${ne.id}" data-status="${ne.status}" style="background:${y.bg}; color:${y.color}; border:1px solid ${y.border}; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; display:inline-flex; align-items:center; gap:4px; cursor:pointer;" title="Klik untuk ubah status">
                                ${y.icon} ${y.label} \u25BE
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
      `}),d.innerHTML=k,u(d)}function b(d,g){if(!g.length){d.innerHTML=`
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
                ${s?'<th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${g.map((x,w)=>{let k=Ye[x.status]||Ye.scheduled;return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3);">${w+1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">${x.branch_full_name}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${x.vendor==="AEON"?"#DBEAFE":"#CCFBF1"}; color:${x.vendor==="AEON"?"#1E40AF":"#0F766E"}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${x.vendor}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:var(--text-2);">
                      ${Dt[x.target_month]||x.target_month}
                    </td>
                    <td style="padding:12px 16px;">
                      <button class="status-toggle-btn" data-id="${x.id}" data-status="${x.status}" style="background:${k.bg}; color:${k.color}; border:1px solid ${k.border}; padding:4px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; cursor:pointer;" title="Klik untuk ubah status">
                        ${k.icon} ${k.label} \u25BE
                      </button>
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-2);">
                      ${x.actual_date||"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; font-weight:600; color:var(--text-2);">
                      ${x.employee_count>0?`\u{1F464} ${x.employee_count}`:"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3); max-width:200px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${x.notes||"-"}
                    </td>
                    ${s?`
                      <td style="padding:12px 16px; text-align:right;">
                        <button class="btn btn-outline btn-sm edit-card-btn" data-id="${x.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-card-btn" data-id="${x.id}" data-name="${x.branch_name}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,u(d)}function u(d){d.querySelectorAll(".status-toggle-btn").forEach(g=>{g.addEventListener("click",x=>{x.stopPropagation();let w=g.dataset.id,k=g.dataset.status;p(g,w,k)})}),d.querySelectorAll(".edit-card-btn").forEach(g=>{g.addEventListener("click",()=>{let x=g.dataset.id,w=i?.matrix?.find(k=>String(k.id)===String(x));w&&h(w)})}),d.querySelectorAll(".delete-card-btn").forEach(g=>{g.addEventListener("click",()=>{let x=g.dataset.id,w=g.dataset.name||"Cabang";Se(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${w}</strong>?`,async()=>{try{if(!(await C(`/api/os-migrations/${x}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus");J(`Jadwal migrasi ${w} berhasil dihapus`),o()}catch(k){j(k.message)}})})})}function p(d,g,x){document.querySelectorAll(".status-popup-menu").forEach(T=>T.remove());let w=d.getBoundingClientRect(),k=document.createElement("div");k.className="status-popup-menu",k.style.cssText=`
      position: fixed;
      top: ${w.bottom+window.scrollY+4}px;
      left: ${w.left+window.scrollX}px;
      background: #fff;
      border: 1px solid var(--border);
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 9999;
      padding: 6px 0;
      min-width: 160px;
    `;let $=[{key:"scheduled",label:"\u{1F7E1} Terjadwal"},{key:"in_progress",label:"\u{1F535} Sedang Proses"},{key:"completed",label:"\u{1F7E2} Selesai"}];k.innerHTML=$.map(T=>`
      <button class="status-option-item" data-key="${T.key}" style="
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
        color: ${T.key===x?"#2563EB":"var(--text-1)"};
        font-weight: ${T.key===x?"700":"500"};
      ">
        ${T.label} ${T.key===x?"\u2713":""}
      </button>
    `).join("");let E=T=>{!k.contains(T.target)&&T.target!==d&&(k.remove(),document.removeEventListener("click",E))};setTimeout(()=>document.addEventListener("click",E),50),k.querySelectorAll(".status-option-item").forEach(T=>{T.addEventListener("click",async()=>{let D=T.dataset.key;if(k.remove(),document.removeEventListener("click",E),D!==x)try{if(!(await C(`/api/os-migrations/${g}/status`,{method:"PATCH",body:JSON.stringify({status:D})})).ok)throw new Error("Gagal memperbarui status");J("Status berhasil diubah"),o()}catch(I){j(I.message)}})}),document.body.appendChild(k)}function h(d=null){let g=!!d,x=g?"Edit Jadwal Migrasi OS":"Tambah Jadwal Migrasi OS",k=`
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${a.map($=>`
      <option value="${$.id}" ${d&&d.branch_id===$.id?"selected":""}>
        ${$.full_name}
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
    `;se({title:x,content:k,confirmText:g?"Simpan Perubahan":"Tambah Cabang",onConfirm:async($,E)=>{let T=$.querySelector("#migration-form");if(!T.checkValidity()){T.reportValidity();return}let D=new FormData(T),I={branch_id:parseInt(D.get("branch_id"),10),vendor:D.get("vendor"),target_month:D.get("target_month"),status:D.get("status"),actual_date:D.get("actual_date")||null,pic_vendor:D.get("pic_vendor")||null,notes:D.get("notes")||null};try{let O=g?`/api/os-migrations/${d.id}`:"/api/os-migrations",Y=await C(O,{method:g?"PUT":"POST",body:JSON.stringify(I)});if(!Y.ok){let P=await Y.json().catch(()=>({}));throw new Error(P.error||"Gagal menyimpan data")}J(g?"Jadwal migrasi berhasil diupdate":"Jadwal migrasi berhasil ditambahkan"),E(),o()}catch(O){j(O.message)}}})}let f=document.getElementById("view-board-btn"),m=document.getElementById("view-table-btn");f.addEventListener("click",()=>{e="board",f.style.background="#fff",f.classList.remove("btn-ghost"),m.style.background="transparent",m.classList.add("btn-ghost"),c()}),m.addEventListener("click",()=>{e="table",m.style.background="#fff",m.classList.remove("btn-ghost"),f.style.background="transparent",f.classList.add("btn-ghost"),c()}),["migration-search","filter-vendor","filter-month","filter-status"].forEach(d=>{let g=document.getElementById(d);g&&(g.addEventListener("input",c),g.addEventListener("change",c))}),document.getElementById("btn-add-migration").addEventListener("click",()=>{h()}),document.getElementById("btn-export-migrations").addEventListener("click",()=>{let d=n();if(!d.length){j("Tidak ada data untuk diexport");return}let g=d.map((x,w)=>({No:w+1,"Kode Cabang":x.branch_code,"Nama Cabang":x.branch_name,"Cabang Lengkap":x.branch_full_name,"Vendor OS":x.vendor,"Target Bulan":Dt[x.target_month]||x.target_month,Status:Ye[x.status]?.label||x.status,"Tanggal Realisasi":x.actual_date||"","Jml Personil FC":x.employee_count||0,"PIC Vendor":x.pic_vendor||"",Catatan:x.notes||""}));L(g,`Timeline_Migrasi_OS_${new Date().toISOString().slice(0,10)}`),J("File Excel berhasil diunduh")}),o()}N();ye();K();var va={Persiapan:{bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE"},"Area Umum":{bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD"},"Ruang Tindakan":{bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},"Ruang Penunjang":{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},Toilet:{bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4"},"Area Luar":{bg:"#FEF3C7",color:"#92400E",border:"#FDE68A"},"Kontrol & Check":{bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA"},Istirahat:{bg:"#E2E8F0",color:"#475569",border:"#CBD5E1"},"Layanan Pasien":{bg:"#DCFCE7",color:"#15803D",border:"#86EFAC"},"Sampah Umum":{bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},"Sampah Medis":{bg:"#FEE2E2",color:"#991B1B",border:"#FCA5A5"},Penutup:{bg:"#E0E7FF",color:"#3730A3",border:"#C7D2FE"}};async function xa(t){t.innerHTML=`
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
  `;let e="pagi",i="timeline",a=[],r=te(),s=r&&["superadmin","admin","manager","spv"].includes(r.role);async function o(){try{let d=await C("/api/dac?all=1");if(!d.ok)throw new Error("Gagal memuat standar DAC");a=d.data?.data||d.data||[];let g=a.filter($=>$.shift==="pagi").length,x=a.filter($=>$.shift==="siang").length;document.getElementById("badge-count-pagi").textContent=g,document.getElementById("badge-count-siang").textContent=x,document.getElementById("badge-count-all").textContent=a.length;let w=document.getElementById("dac-category-filter"),k=Array.from(new Set(a.map($=>$.category))).sort();w.innerHTML='<option value="">Semua Kategori</option>'+k.map($=>`<option value="${$}">${$}</option>`).join(""),n(),c()}catch(d){document.getElementById("dac-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${d.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function l(){let d=(document.getElementById("dac-search")?.value||"").toLowerCase().trim(),g=document.getElementById("dac-category-filter")?.value||"",x=document.getElementById("dac-period-filter")?.value||"";return a.filter(w=>{if(e!=="all"&&w.shift!==e||g&&w.category!==g)return!1;if(x==="morning"){let k=parseInt(w.start_time.split(":")[0],10);if(w.shift==="pagi"&&k>=12||w.shift==="siang"&&k>=17)return!1}else if(x==="afternoon"){let k=parseInt(w.start_time.split(":")[0],10);if(w.shift==="pagi"&&k<12||w.shift==="siang"&&k<17)return!1}return!(d&&!`${w.start_time} ${w.end_time} ${w.category} ${w.location} ${w.activity}`.toLowerCase().includes(d))})}function n(){let d=document.getElementById("dac-summary-cards");if(!d)return;let g=a.filter(E=>E.shift==="pagi"),x=a.filter(E=>E.shift==="siang"),w=e==="pagi"?g:e==="siang"?x:a,k=e==="pagi"?"08:00 \u2013 17:00":e==="siang"?"13:00 \u2013 22:00":"08:00 \u2013 22:00",$=e==="pagi"?"Shift Pagi":e==="siang"?"Shift Siang":"Semua Shift";d.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Aktivitas (${$})</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${w.length} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Aktivitas</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Urutan Terjadwal dari Awal s/d Penutup</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Jam Kerja Standar</div>
        <div style="font-size:1.5rem; font-weight:800; color:#1E40AF; line-height:1.2;">${k}</div>
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
    `}function c(){let d=l(),g=document.getElementById("dac-stats-text");g&&(g.innerHTML=`Menampilkan <strong>${d.length}</strong> aktivitas`);let x=document.getElementById("dac-view-container");i==="timeline"?v(x,d):b(x,d)}function v(d,g){if(!g.length){d.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;return}d.innerHTML=`
      <div class="timeline-container" style="display:flex; flex-direction:column; gap:12px; max-width:960px; margin:0 auto;">
        ${g.map(x=>{let w=va[x.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},k=x.shift==="pagi"?"\u{1F305} Pagi":"\u{1F307} Siang",$=x.shift==="pagi"?"#FEF3C7":"#E0E7FF",E=x.shift==="pagi"?"#92400E":"#3730A3";return`
            <div class="dac-timeline-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:16px 20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; gap:16px; align-items:flex-start; position:relative;">
              <!-- Time Column -->
              <div style="min-width:130px; text-align:center; background:#F8FAFC; border:1px solid var(--border); border-radius:8px; padding:10px 8px;">
                <div style="font-size:1.1rem; font-weight:800; color:var(--text-1); font-family:monospace;">
                  ${x.start_time} - ${x.end_time}
                </div>
                <div style="font-size:0.75rem; color:var(--text-3); font-weight:600; margin-top:4px;">
                  Item #${x.item_no}
                </div>
                ${e==="all"?`
                  <div style="margin-top:6px;">
                    <span class="badge" style="background:${$}; color:${E}; font-size:0.7rem; font-weight:700; padding:2px 6px; border-radius:4px;">
                      ${k}
                    </span>
                  </div>
                `:""}
              </div>

              <!-- Content Column -->
              <div style="flex:1;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                  <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span class="badge" style="background:${w.bg}; color:${w.color}; border:1px solid ${w.border}; font-weight:700; font-size:0.75rem; padding:3px 10px; border-radius:6px;">
                      ${x.category}
                    </span>
                    <span style="font-weight:700; font-size:0.95rem; color:var(--text-1);">
                      \u{1F4CD} ${x.location}
                    </span>
                  </div>
                  ${s?`
                    <div style="display:flex; gap:4px;">
                      <button class="btn btn-ghost btn-sm edit-dac-btn" data-id="${x.id}" title="Edit" style="padding:4px 8px; font-size:0.8rem; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-ghost btn-sm delete-dac-btn" data-id="${x.id}" data-name="${x.location}" title="Hapus" style="padding:4px 8px; font-size:0.8rem; color:var(--danger); cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </div>
                  `:""}
                </div>

                <div style="font-size:0.9rem; color:var(--text-2); line-height:1.5; background:#FAFBFD; padding:10px 14px; border-radius:8px; border-left:3px solid ${w.color}; margin-top:8px;">
                  ${x.activity}
                </div>
              </div>
            </div>
          `}).join("")}
      </div>
    `,u(d)}function b(d,g){if(!g.length){d.innerHTML=`
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
                ${e==="all"?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">SHIFT</th>':""}
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:110px;">JAM</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:140px;">KATEGORI</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:200px;">LOKASI / AREA</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">AKTIVITAS</th>
                ${s?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right; width:100px;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${g.map(x=>{let w=va[x.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"};return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 14px; font-size:0.85rem; color:var(--text-3); font-weight:600;">${x.item_no}</td>
                    ${e==="all"?`
                      <td style="padding:12px 14px;">
                        <span class="badge" style="background:${x.shift==="pagi"?"#FEF3C7":"#E0E7FF"}; color:${x.shift==="pagi"?"#92400E":"#3730A3"}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:4px;">
                          ${x.shift==="pagi"?"Pagi":"Siang"}
                        </span>
                      </td>
                    `:""}
                    <td style="padding:12px 14px; font-family:monospace; font-weight:700; font-size:0.85rem; color:var(--text-1); white-space:nowrap;">
                      ${x.start_time} - ${x.end_time}
                    </td>
                    <td style="padding:12px 14px;">
                      <span class="badge" style="background:${w.bg}; color:${w.color}; border:1px solid ${w.border}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px; white-space:nowrap;">
                        ${x.category}
                      </span>
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; font-weight:600; color:var(--text-1);">
                      ${x.location}
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; color:var(--text-2); line-height:1.4;">
                      ${x.activity}
                    </td>
                    ${s?`
                      <td style="padding:12px 14px; text-align:right; white-space:nowrap;">
                        <button class="btn btn-outline btn-sm edit-dac-btn" data-id="${x.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-dac-btn" data-id="${x.id}" data-name="${x.location}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,u(d)}function u(d){d.querySelectorAll(".edit-dac-btn").forEach(g=>{g.addEventListener("click",()=>{let x=g.dataset.id,w=a.find(k=>String(k.id)===String(x));w&&p(w)})}),d.querySelectorAll(".delete-dac-btn").forEach(g=>{g.addEventListener("click",()=>{let x=g.dataset.id,w=g.dataset.name||"item";Se(`Apakah Anda yakin ingin menghapus aktivitas pada lokasi <strong>${w}</strong>?`,async()=>{try{if(!(await C(`/api/dac/${x}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus item");J("Aktivitas DAC berhasil dihapus"),o()}catch(k){j(k.message)}})})})}function p(d=null){let g=!!d,x=g?"Edit Aktivitas DAC":"Tambah Aktivitas DAC",w=["Persiapan","Area Umum","Ruang Tindakan","Ruang Penunjang","Toilet","Area Luar","Kontrol & Check","Istirahat","Layanan Pasien","Sampah Umum","Sampah Medis","Penutup"],k=`
      <form id="dac-form" style="display:flex; flex-direction:column; gap:14px;">
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Shift *</label>
            <select name="shift" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="pagi" ${d?.shift==="pagi"||!d&&e==="pagi"?"selected":""}>\u{1F305} Shift Pagi (08:00 - 17:00)</option>
              <option value="siang" ${d?.shift==="siang"||!d&&e==="siang"?"selected":""}>\u{1F307} Shift Siang (13:00 - 22:00)</option>
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
            ${w.map($=>`<option value="${$}">`).join("")}
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
    `;se({title:x,content:k,confirmText:g?"Simpan Perubahan":"Tambah Aktivitas",onConfirm:async($,E)=>{let T=$.querySelector("#dac-form");if(!T.checkValidity()){T.reportValidity();return}let D=new FormData(T),I={shift:D.get("shift"),item_no:D.get("item_no")?parseInt(D.get("item_no"),10):void 0,start_time:D.get("start_time"),end_time:D.get("end_time"),category:D.get("category"),location:D.get("location"),activity:D.get("activity")};try{let O=g?`/api/dac/${d.id}`:"/api/dac",Y=await C(O,{method:g?"PUT":"POST",body:JSON.stringify(I)});if(!Y.ok){let P=await Y.json().catch(()=>({}));throw new Error(P.error||"Gagal menyimpan aktivitas")}J(g?"Aktivitas berhasil diperbarui":"Aktivitas berhasil ditambahkan"),E(),o()}catch(O){j(O.message)}}})}function h(){let d=l();if(!d.length){j("Tidak ada data untuk dicetak");return}let g=e==="pagi"?"SHIFT PAGI (08:00 - 17:00)":e==="siang"?"SHIFT SIANG (13:00 - 22:00)":"SELURUH SHIFT (PAGI & SIANG)",x=window.open("","_blank");if(!x){j("Popup diblokir browser, mohon izinkan popup untuk mencetak");return}x.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>STANDAR ACUAN DAC - ${g}</title>
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
          <div class="subtitle">STANDAR RUTINITAS DAILY ACTIVITY CLEANING (DAC) \u2014 ${g}</div>
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
            ${d.map(w=>`
              <tr>
                <td style="text-align:center; font-weight:bold;">${w.item_no}</td>
                ${e==="all"?`<td>${w.shift==="pagi"?"Pagi":"Siang"}</td>`:""}
                <td class="jam">${w.start_time} - ${w.end_time}</td>
                <td><strong>${w.category}</strong></td>
                <td><strong>${w.location}</strong></td>
                <td>${w.activity}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
        <div style="margin-top:20px; text-align:right; font-size:9pt; color:#64748b;">
          Dicetak pada: ${new Date().toLocaleString("id-ID")} | Standar Resmi Operasional Facility Care Audy Dental
        </div>
      </body>
      </html>
    `),x.document.close(),x.focus(),setTimeout(()=>{x.print()},400)}document.querySelectorAll(".shift-tab-btn").forEach(d=>{d.addEventListener("click",()=>{document.querySelectorAll(".shift-tab-btn").forEach(g=>{g.style.background="#F1F5F9",g.style.color="var(--text-2)"}),d.style.background="#2563EB",d.style.color="#fff",e=d.dataset.shift,n(),c()})});let f=document.getElementById("view-timeline-btn"),m=document.getElementById("view-table-btn");f.addEventListener("click",()=>{i="timeline",f.style.background="#fff",f.classList.remove("btn-ghost"),m.style.background="transparent",m.classList.add("btn-ghost"),c()}),m.addEventListener("click",()=>{i="table",m.style.background="#fff",m.classList.remove("btn-ghost"),f.style.background="transparent",f.classList.add("btn-ghost"),c()}),["dac-search","dac-category-filter","dac-period-filter"].forEach(d=>{let g=document.getElementById(d);g&&(g.addEventListener("input",c),g.addEventListener("change",c))}),document.getElementById("btn-add-dac").addEventListener("click",()=>{p()}),document.getElementById("btn-print-dac").addEventListener("click",()=>{h()}),document.getElementById("btn-export-dac").addEventListener("click",()=>{let d=l();if(!d.length){j("Tidak ada data untuk diexport");return}let g=d.map(x=>({No:x.item_no,Shift:x.shift==="pagi"?"Shift Pagi":"Shift Siang","Jam Mulai":x.start_time,"Jam Selesai":x.end_time,Kategori:x.category,"Lokasi / Area":x.location,"Rincian Aktivitas":x.activity}));L(g,`DAC_Daily_Activity_Cleaning_${e}_${new Date().toISOString().slice(0,10)}`),J("File Excel berhasil diunduh")}),o()}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let i=Math.floor(Number(t)-25569);return new Date(i*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let i=t.split(/[\/\-]/);return`${i[2]}-${i[1]}-${i[0]}`}let e=t.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(e){let i=e[1],a=parseInt(e[2],10),r=parseInt(e[3],10);if(a>12&&r<=12)return`${i}-${e[3]}-${e[2]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let i=e.split("-");if(i.length===3&&i[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],r=parseInt(i[2],10),s=a[parseInt(i[1],10)-1];return`${r} ${s} ${i[0]}`}return e};function G(t){return async e=>{if(!je()){$e("/login");return}return t(e)}}var We=null;function tn(){We&&clearInterval(We);let t=()=>{let e=new Date,i=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),r=document.getElementById("header-clock-time"),s=document.getElementById("header-clock-date");r&&(r.textContent=i),s&&(s.textContent=a)};t(),We=setInterval(t,1e3)}async function an(){try{let t=await C("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},i=(a,r)=>{let s=document.getElementById(a);s&&(s.textContent=r>0?r:"",s.style.display=r>0?"inline-flex":"none")};i("badge-issues",e.issues?.current||0),i("badge-contracts",e.expiring30?.current||0),i("badge-oo1",e.one_on_one?.current||0),i("badge-schedule",e.schedule?.current||0),i("badge-supply",e.supply?.current||0)}catch{}}var Re=[];async function nn(){try{let t=await C("/api/dashboard/notifications");if(!t.ok)return;Re=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Re.length>0?"block":"none",e.textContent=Re.length)}catch{}}function rn(){if(!Re.length){se({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,i)=>i()});return}let t=`
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
  `;se({title:`Notifikasi (${Re.length})`,content:t,confirmText:"Tutup",onConfirm:(e,i)=>i()})}function ka(){let t=te(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
                <span class="topbar-greeting-time">${(()=>{let c=new Date().getHours();return c>=4&&c<11?"Selamat Pagi":c>=11&&c<15?"Selamat Siang":c>=15&&c<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${t?.full_name||t?.username||"Guest"}</span> \u{1F44B}
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
  `;let i=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),r=document.getElementById("topbar-menu-btn"),s=document.getElementById("sidebar-close"),o=()=>{i.classList.add("open"),a.classList.add("show")},l=()=>{i.classList.remove("open"),a.classList.remove("show")};r?.addEventListener("click",o),s?.addEventListener("click",l),a?.addEventListener("click",l),document.querySelectorAll(".nav-item").forEach(c=>c.addEventListener("click",l)),t&&t.role==="input_lembur"&&(document.querySelectorAll(".nav-item").forEach(c=>{c.dataset.route!=="/overtime"&&(c.style.display="none")}),document.querySelectorAll(".nav-section").forEach(c=>{Array.from(c.querySelectorAll(".nav-item")).some(b=>b.style.display!=="none")||(c.style.display="none")}));function n(){let c=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(u=>{let p=u.dataset.route;u.classList.toggle("active",c===p||p!=="/dashboard"&&c.startsWith(p))});let v=document.getElementById("topbar-title"),b=document.querySelector(".nav-item.active .nav-label");v&&b&&(v.textContent=b.textContent)}window.addEventListener("hashchange",n),n(),tn(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await C("/api/auth/logout",{method:"POST"}),ze(),We&&clearInterval(We),$e("/login")}),an(),nn(),document.getElementById("btn-notif")?.addEventListener("click",c=>{c.preventDefault(),rn()})}async function on(){q("/login",({main:e})=>Kt(e)),q("/dashboard",G(({main:e})=>Pt(e))),q("/calendar",G(({main:e})=>ia(e))),q("/employees",G(({main:e,params:i})=>jt(e,i))),q("/contracts",G(({main:e,params:i})=>lt(e,i))),q("/sp",G(({main:e})=>ma(e))),q("/mutasi",G(({main:e})=>ga(e))),q("/overtime",G(({main:e})=>ha(e))),q("/timeline",G(({main:e,params:i})=>Ht(e,i))),q("/os-migrations",G(({main:e})=>ya(e))),q("/issues",G(({main:e,params:i})=>qt(e,i))),q("/one-on-one",G(({main:e,params:i})=>Jt(e,i))),q("/training",G(({main:e})=>Ut(e))),q("/relievers",G(({main:e,params:i})=>Vt(e,i))),q("/reports/inspection",G(({main:e})=>Qt(e))),q("/reports/cleaning",G(({main:e})=>Yt(e))),q("/reports/fogging",G(({main:e})=>Wt(e))),q("/reports/basecamp",G(({main:e})=>Xt(e))),q("/reports/supply",G(({main:e})=>$t(e,"supply"))),q("/sop",G(({main:e})=>Zt(e))),q("/checklist",G(({main:e})=>ea(e))),q("/hygiene-standards",G(({main:e})=>ta(e))),q("/dac",G(({main:e})=>xa(e))),q("/forms",G(({main:e})=>$t(e))),q("/users",G(({main:e})=>aa(e))),q("/branches",G(({main:e})=>na(e))),q("/profile",G(({main:e})=>ra(e))),q("/settings/import",G(({main:e})=>ca(e))),q("/audit-logs",G(({main:e})=>fa(e)));let t=je();if(!t&&window.location.hash!=="#/login"&&$e("/login"),t){let e=await C("/api/auth/me");if(e.ok){He(e.data.data),ka();let i=te(),a=window.location.hash;i&&i.role==="input_lembur"&&(a==="#/dashboard"||a===""||a==="#/")&&$e("/overtime")}else ze(),$e("/login")}window.addEventListener("fm:login",()=>{ka();let e=te();e&&e.role==="input_lembur"?$e("/overtime"):$e("/dashboard")}),Bt()}on();
