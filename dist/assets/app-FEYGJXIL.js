var mn=Object.defineProperty;var Ut=(e,t)=>()=>(e&&(t=e(e=0)),t);var Gt=(e,t)=>{for(var n in t)mn(e,n,{get:t[n],enumerable:!0})};var Vt={};Gt(Vt,{confirmDialog:()=>re,createModal:()=>Z});function Z({title:e,content:t,onConfirm:n,onCancel:r,confirmText:i="Simpan",cancelText:m="Batal",size:f="md",confirmClass:u="btn-primary"}){let o={sm:"400px",md:"560px",lg:"720px",xl:"900px"},p=document.createElement("div");p.className="modal-overlay",p.innerHTML=`
    <div class="modal" style="max-width:${o[f]||o.md}">
      <div class="modal-header">
        <h3 class="modal-title">${e}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof t=="string"?t:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${m}</button>
        ${n?`<button class="btn ${u} modal-confirm">${i}</button>`:""}
      </div>
    </div>
  `,t instanceof HTMLElement&&p.querySelector(".modal-body").appendChild(t);let $=()=>{p.classList.remove("show"),setTimeout(()=>p.remove(),250)};return p.querySelector(".modal-close").addEventListener("click",()=>{r&&r(),$()}),p.querySelector(".modal-cancel").addEventListener("click",()=>{r&&r(),$()}),n&&p.querySelector(".modal-confirm").addEventListener("click",()=>n(p,$)),p.addEventListener("click",F=>{F.target===p&&(r&&r(),$())}),document.body.appendChild(p),requestAnimationFrame(()=>p.classList.add("show")),{overlay:p,close:$}}function re(e,t,n="Konfirmasi"){if(typeof e=="object"&&e!==null){let r=e;return Z({title:r.title||"Konfirmasi",content:typeof r.message=="string"?`<p>${r.message}</p>`:r.message,onConfirm:async(i,m)=>{r.onConfirm&&await r.onConfirm(),m()},onCancel:r.onCancel,confirmText:r.confirmText||"Ya, Lanjutkan",cancelText:r.cancelText||"Batal",confirmClass:r.confirmClass||"btn-danger"})}return Z({title:n,content:typeof e=="string"?`<p>${e}</p>`:e,onConfirm:async(r,i)=>{t&&await t(),i()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var ge=Ut(()=>{});var fe={};Gt(fe,{downloadExcel:()=>ne,parseExcel:()=>Ge,renderExcelButtons:()=>xn});function Ge(e){return new Promise((t,n)=>{let r=new FileReader;r.onload=i=>{try{let m=new Uint8Array(i.target.result),f=XLSX.read(m,{type:"array"}),u=f.SheetNames[0],o=f.Sheets[u];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${e.name}`),console.log(`File Size: ${(e.size/1024).toFixed(2)} KB`),console.log(`File Type: ${e.type||"unknown"}`),console.log(`Sheets Found: ${f.SheetNames.join(", ")}`),console.log(`Sheet Used: ${u}`);let p=XLSX.utils.decode_range(o["!ref"]||"A1:A1"),$=p.e.r-p.s.r+1,F=p.e.c-p.s.c+1;console.log(`Total Rows (including empty): ${$}`),console.log(`Total Columns: ${F}`);let E=[];for(let j=p.s.c;j<=p.e.c;++j){let T=o[XLSX.utils.encode_cell({c:j,r:p.s.r})];T&&T.v&&E.push(T.v)}console.log(`Headers Found: ${E.join(", ")}`),console.log("---------------------------");let C=XLSX.utils.sheet_to_json(o,{defval:""});Object.defineProperty(C,"__worksheet",{value:o,enumerable:!1}),Object.defineProperty(C,"__headers",{value:E,enumerable:!1}),t(C)}catch(m){n(m)}},r.onerror=i=>n(i),r.readAsArrayBuffer(e)})}function ne(e,t){try{let n=XLSX.utils.json_to_sheet(e),r=XLSX.utils.book_new();XLSX.utils.book_append_sheet(r,n,"Data"),XLSX.writeFile(r,`${t}.xlsx`)}catch(n){throw console.error("Error generating Excel file:",n),n}}function xn(e){return`
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
  `}var pe=Ut(()=>{});var bn="https://fm-operations-api.facilitycare-audydental.workers.dev",hn=bn,Le=1e4;function Ze(){return localStorage.getItem("fm_token")}function Jt(e){localStorage.setItem("fm_token",e)}function et(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ae(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function tt(e){localStorage.setItem("fm_user",JSON.stringify(e))}function X(e,t="read",n=null){if(!e||!e.role)return!1;let r=e.role;if(r==="superadmin"||r==="admin")return!0;let i=null,m="read";return n!==null?(i=t,m=n):m=t||"read",m==="export"||m==="import"||m==="admin"?!1:r==="editor_khusus"?i==="issues"||i==="overtime"?m==="read"||m==="write"||m==="delete":m==="read":r==="input_lembur"&&i==="overtime"?m==="read"||m==="write":m==="read"}async function q(e,t={}){let n=Ze(),r={"Content-Type":"application/json",...n?{Authorization:`Bearer ${n}`}:{},...t.headers||{}};try{let i=`cb=${Date.now()}`,m=e.includes("?")?"&":"?",f=`${hn}${e}${m}${i}`,u=await fetch(f,{...t,headers:r}),o;try{let p=await u.text();try{o=JSON.parse(p)}catch{o={error:`Server Error (${u.status}): ${p.substring(0,80)}...`}}}catch{o={error:"Gagal membaca respon dari server"}}return u.status===401&&(et(),window.location.hash="#/login"),{ok:u.ok,status:u.status,data:o}}catch(i){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${i.message})`}}}}var ht={},at=null;function oe(e,t){ht[e]=t}function $e(e){window.location.hash=e}function Yt(){async function e(){let t=window.location.hash.replace("#","")||"/dashboard",[n,...r]=t.split("?"),i=ht[n];if(!i){for(let[f,u]of Object.entries(ht))if(f.endsWith("/*")&&n.startsWith(f.slice(0,-2))){i=u;break}}at&&(at(),at=null);let m=document.getElementById("main-content");if(m&&(m.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),i){let f=new URLSearchParams(r.join("?")),u=n.split("/").filter(Boolean),o=await i({path:n,params:f,segments:u,main:m});o&&(at=o)}else{let f=m||document.getElementById("app");f&&(f.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",e),e()}var He;function fn(){return He||(He=document.createElement("div"),He.id="toast-container",document.body.appendChild(He)),He}function ft(e,t="info",n=3500){let r=fn(),i=document.createElement("div");i.className=`toast toast-${t}`;let m={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};i.innerHTML=`<span class="toast-icon">${m[t]||"\u2139"}</span><span class="toast-msg">${e}</span>`,r.appendChild(i),requestAnimationFrame(()=>i.classList.add("show")),setTimeout(()=>{i.classList.remove("show"),setTimeout(()=>i.remove(),350)},n)}var H=e=>ft(e,"success"),N=e=>ft(e,"error");var xt=e=>ft(e,"info");ge();ge();function nt({columns:e,data:t,onEdit:n,onDelete:r,onView:i,actions:m=[],emptyText:f="Tidak ada data",bulkSelect:u=null}){let o=document.createElement("div");if(o.className="table-wrapper",!t||t.length===0)return o.innerHTML=`<div class="empty-state"><p>${f}</p></div>`,o;let p=document.createElement("table");p.className="data-table";let $=document.createElement("thead"),F=document.createElement("tr");if(u){let C=document.createElement("th");C.style.width="40px",C.style.textAlign="center";let j=document.createElement("input");j.type="checkbox",j.id="select-all-checkbox",j.title="Pilih semua",j.addEventListener("change",()=>{t.forEach(T=>{j.checked?u.selectedIds.add(T.id):u.selectedIds.delete(T.id)}),o.querySelectorAll(".row-checkbox").forEach(T=>T.checked=j.checked),u.onToggle()}),C.appendChild(j),F.appendChild(C)}if(e.forEach(C=>{let j=document.createElement("th");j.textContent=C.label,C.width&&(j.style.width=C.width),F.appendChild(j)}),n||r||i||m.length>0){let C=document.createElement("th");C.textContent="Aksi",C.style.width="120px",F.appendChild(C)}$.appendChild(F),p.appendChild($);let E=document.createElement("tbody");return t.forEach(C=>{let j=document.createElement("tr");if(u){let T=document.createElement("td");T.style.textAlign="center",T.style.width="40px";let D=document.createElement("input");D.type="checkbox",D.className="row-checkbox",D.checked=u.selectedIds.has(C.id),D.addEventListener("change",()=>{if(D.checked)u.selectedIds.add(C.id);else{u.selectedIds.delete(C.id);let _=document.getElementById("select-all-checkbox");_&&(_.checked=!1)}u.onToggle()}),T.appendChild(D),j.appendChild(T)}if(e.forEach(T=>{let D=document.createElement("td");if(T.render){let _=T.render(C[T.key],C);_ instanceof HTMLElement?D.appendChild(_):D.innerHTML=_||""}else D.textContent=C[T.key]!==null&&C[T.key]!==void 0&&C[T.key]!==""?C[T.key]:"";T.nowrap&&(D.style.whiteSpace="nowrap"),j.appendChild(D)}),n||r||i||m.length>0){let T=document.createElement("td");T.className="actions-cell";let D=document.createElement("div");if(D.className="btn-group",i){let _=document.createElement("button");_.className="btn btn-xs btn-ghost",_.innerHTML="\u{1F441}",_.title="Lihat",_.addEventListener("click",()=>i(C)),D.appendChild(_)}if(n){let _=document.createElement("button");_.className="btn btn-xs btn-secondary",_.innerHTML="\u270F\uFE0F",_.title="Edit",_.addEventListener("click",()=>n(C)),D.appendChild(_)}m.forEach(_=>{let g=document.createElement("button");g.className=`btn btn-xs ${_.class||"btn-ghost"}`,g.innerHTML=_.icon||_.label,g.title=_.label,g.addEventListener("click",()=>_.handler(C)),D.appendChild(g)}),T.appendChild(D),j.appendChild(T)}E.appendChild(j)}),p.appendChild(E),o.appendChild(p),o}function it({page:e,pages:t,total:n,limit:r,onPage:i}){if(t<=1)return null;let m=document.createElement("div");m.className="pagination";let f=document.createElement("span");f.className="pagination-info",f.textContent=`Total: ${n} data`,m.appendChild(f);let u=document.createElement("div");u.className="pagination-btns";let o=(F,E,C=!1,j=!1)=>{let T=document.createElement("button");T.className=`btn btn-sm ${j?"btn-primary":"btn-ghost"} pagination-btn`,T.textContent=F,T.disabled=C,T.addEventListener("click",()=>i(E)),u.appendChild(T)};o("\xAB",1,e===1),o("\u2039",e-1,e===1);let p=Math.max(1,e-2),$=Math.min(t,e+2);for(let F=p;F<=$;F++)o(F,F,!1,F===e);return o("\u203A",e+1,e===t),o("\xBB",t,e===t),m.appendChild(u),m}ge();function Ue(e){return e.map(t=>{if(t.type==="hidden")return`<input type="hidden" name="${t.name}" value="${t.value||""}">`;if(t.type==="html")return t.html||"";if(t.type==="row")return`<div class="form-row">${Ue(t.fields)}</div>`;let n=t.required?"required":"",r=t.label?`<label class="form-label">${t.label}${t.required?' <span class="required">*</span>':""}</label>`:"",i="";switch(t.type){case"textarea":i=`<textarea name="${t.name}" class="form-control" placeholder="${t.placeholder||""}" ${n} rows="${t.rows||3}">${t.value||""}</textarea>`;break;case"select":let f=(t.options||[]).map(C=>{let j=typeof C=="object"?C.value:C,T=typeof C=="object"?C.label:C,D=t.value==j?"selected":"";return`<option value="${j}" ${D}>${T}</option>`}).join("");i=`<select name="${t.name}" class="form-control" ${n}><option value="">-- Pilih ${t.label||""} --</option>${f}</select>`;break;case"combobox":let u=`dl-${t.name}-${Math.random().toString(36).substring(7)}`,o=(t.options||[]).map(C=>{let j=typeof C=="object"?C.value:C,T=typeof C=="object"?C.label||C.value||"":C||"";return(T==="undefined"||T==="[object Object]"||T==="null")&&(T=""),T?`<option value="${T}"></option>`:""}).join(""),p=t.value||"";if(t.value){let C=(t.options||[]).find(j=>(typeof j=="object"?j.value:j)==t.value);if(C){let j=typeof C=="object"?C.label||C.value||"":C||"";j&&j!=="undefined"&&j!=="[object Object]"&&j!=="null"&&(p=j)}}i=`
          <input type="text" name="${t.name}" list="${u}" class="form-control" value="${p}" placeholder="Pilih atau ketik baru..." ${n} autocomplete="off">
          <datalist id="${u}">${o}</datalist>
        `;break;case"checkbox":i=`<label class="checkbox-label"><input type="checkbox" name="${t.name}" value="1" ${t.value?"checked":""}> ${t.checkLabel||t.label}</label>`;break;case"date":let $=window.parseFlexibleDate&&t.value?window.parseFlexibleDate(t.value):t.value||"";i=`<input type="date" name="${t.name}" class="form-control" value="${$}" ${n}>`;break;case"month":let F=t.value?String(t.value).slice(0,7):"";i=`<input type="month" name="${t.name}" class="form-control" value="${F}" ${n}>`;break;case"number":let E=t.readonly?"readonly":"";i=`<input type="number" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" min="${t.min||""}" max="${t.max||""}" step="${t.step||"1"}" ${n} ${E}>`;break;case"email":i=`<input type="email" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${n} autocomplete="${t.autocomplete||"off"}">`;break;case"password":i=`<input type="password" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${n} autocomplete="${t.autocomplete||"new-password"}">`;break;case"url":i=`<input type="url" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||"https://..."}" ${n}>`;break;default:i=`<input type="${t.type||"text"}" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${n} autocomplete="${t.autocomplete||"off"}">`}let m=t.hint?`<div class="form-hint">${t.hint}</div>`:"";return`<div class="form-group ${t.class||""}">${r}${i}${m}</div>`}).join("")}function rt(e){let t={},n=new FormData(e);for(let[r,i]of n.entries())t[r]=i===""?null:i;return e.querySelectorAll("input[type=checkbox]").forEach(r=>{r.checked||(t[r.name]=null)}),t}function ot(e,t){t&&Object.entries(t).forEach(([n,r])=>{let i=e.querySelector(`[name="${n}"]`);i&&(i.hasAttribute("list")||(i.type==="checkbox"?i.checked=!!r:i.type==="date"&&r&&window.parseFlexibleDate?i.value=window.parseFlexibleDate(r):i.type==="month"&&r?i.value=String(r).slice(0,7):i.value=r??""))})}pe();var _e={},st={on(e,t){_e[e]||(_e[e]=new Set),_e[e].add(t)},off(e,t){_e[e]&&_e[e].delete(t)},emit(e,t){_e[e]&&_e[e].forEach(n=>{try{n(t)}catch(r){console.warn("[calendarBus] Handler error:",r)}})},clear(){Object.keys(_e).forEach(e=>delete _e[e])}},yn=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function xe(e){if(!e){st.emit("data:changed",{module:"unknown"});return}let t=String(e).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");st.emit("data:changed",{module:t,relevant:yn.has(t)})}function me({container:e,title:t,icon:n,apiPath:r,columns:i,formFields:m,filterFields:f,defaultFilters:u={},itemLabel:o="Data",canCreate:p=!0,canEdit:$=!0,canDelete:F=!0,onBeforeSubmit:E,onAfterLoad:C,onDataLoaded:j,extraActions:T=[],initialSearch:D="",exportOptions:_=null,bulkDelete:g=!1,paginationMode:A="server"}){let s=ae();s&&(s.role==="superadmin"||s.role==="admin")?X(s,"delete")||(F=!1,g=!1):(_=null,g=!1,s&&s.role==="editor_khusus"?r!=="/api/issues"&&r!=="/api/overtime"&&(p=!1,$=!1,F=!1):s&&s.role==="input_lembur"?(r==="/api/overtime"||(p=!1,$=!1),F=!1):(p=!1,$=!1,F=!1));let c=1,v={...u};D&&(v.search=D);let x=new Set;e.innerHTML=`
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
      <h1 class="page-title">${n} ${t}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${p?`<button class="btn btn-primary" id="btn-create">+ Tambah ${o}</button>`:""}
        ${_?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${_.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${_.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${_.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${_.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
    

    ${f&&f.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${f.filter(b=>b.type==="search").map(b=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${b.placeholder||"Cari..."}" id="filter-search" value="${v.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${f.filter(b=>b.type!=="search").map(b=>{if(b.type==="select"||b.type==="combobox"){let h=(b.label||"").startsWith("Pilih")?b.label:`Pilih ${b.label||""}`;return`<select class="filter-select" name="${b.name}" id="filter-${b.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${h}</option>${(b.options||[]).map(S=>`<option value="${typeof S=="object"?S.value:S}" ${v[b.name]===(typeof S=="object"?S.value:S)?"selected":""}>${typeof S=="object"?S.label:S}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${f.filter(b=>b.type!=="search").map(b=>{if(b.type==="select"||b.type==="combobox"){let h=(b.label||"").startsWith("Pilih")?b.label:`Pilih ${b.label||""}`;return`<select class="filter-select filter-select-sheet" name="${b.name}-sheet" id="filter-sheet-${b.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${h}</option>${(b.options||[]).map(S=>`<option value="${typeof S=="object"?S.value:S}" ${v[b.name]===(typeof S=="object"?S.value:S)?"selected":""}>${typeof S=="object"?S.label:S}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function w(){let b=document.getElementById("bulk-toolbar");if(!b)return;let h=document.getElementById("bulk-count"),S=document.getElementById("btn-bulk-delete"),O=document.getElementById("btn-bulk-cancel");h.textContent=`${x.size} item dipilih`,x.size>0?(b.style.display="flex",S.disabled=!1,O.disabled=!1):(b.style.display="none",S.disabled=!0,O.disabled=!0);let z=document.getElementById("select-all-checkbox");if(z){let G=document.querySelectorAll(".row-checkbox");if(G.length>0){let L=[...G].every(V=>V.checked),R=[...G].some(V=>V.checked);z.checked=L,z.indeterminate=R&&!L}else z.checked=!1,z.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{x.clear(),document.querySelectorAll(".row-checkbox").forEach(h=>h.checked=!1);let b=document.getElementById("select-all-checkbox");b&&(b.checked=!1),w()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(x.size===0)return;let b=[...x],h=document.createElement("div");h.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",h.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${b.length} ${o}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${b.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(h),h.querySelector("#bulk-cancel-btn").addEventListener("click",()=>h.remove()),h.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let S=h.querySelector("#bulk-confirm-btn");S.disabled=!0,S.textContent="Menghapus...";let O=await q(`${r}/bulk`,{method:"DELETE",body:JSON.stringify({ids:b})});h.remove(),O.ok?(H(`${b.length} ${o} berhasil dihapus.`),x.clear(),w(),xe(r),I()):N(O.data?.error||"Gagal menghapus data.")})});let a=document.getElementById("filter-search"),d;a?.addEventListener("input",b=>{clearTimeout(d),d=setTimeout(()=>{v.search=b.target.value,c=1,x.clear(),w(),I()},400)}),f?.forEach(b=>{(b.type==="select"||b.type==="combobox")&&(document.getElementById(`filter-${b.name}`)?.addEventListener("change",h=>{v[b.name]=h.target.value;let S=document.getElementById(`filter-sheet-${b.name}`);S&&(S.value=h.target.value),c=1,x.clear(),w(),I()}),document.getElementById(`filter-sheet-${b.name}`)?.addEventListener("change",h=>{v[b.name]=h.target.value;let S=document.getElementById(`filter-${b.name}`);S&&(S.value=h.target.value),c=1,x.clear(),w(),I(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{v={},a&&(a.value=""),f?.forEach(b=>{let h=document.getElementById(`filter-${b.name}`);h&&(h.value="");let S=document.getElementById(`filter-sheet-${b.name}`);S&&(S.value="")}),c=1,x.clear(),w(),I()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{v={},a&&(a.value=""),f?.forEach(b=>{let h=document.getElementById(`filter-${b.name}`);h&&(h.value="");let S=document.getElementById(`filter-sheet-${b.name}`);S&&(S.value="")}),c=1,x.clear(),w(),I(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>B(null)),_&&document.addEventListener("click",function(b){let h=document.getElementById("aksi-menu-main"),S=document.getElementById("btn-aksi-main");h&&S&&!S.contains(b.target)&&!h.contains(b.target)&&h.classList.remove("show-aksi-menu")});let y=document.getElementById("btn-mobile-filter"),k=document.getElementById("filter-options-wrapper"),M=document.getElementById("btn-close-filter-sheet");if(y&&k&&(y.addEventListener("click",b=>{b.preventDefault(),k.classList.add("sheet-open")}),M&&M.addEventListener("click",b=>{b.preventDefault(),k.classList.remove("sheet-open")})),_){document.getElementById(`btn-export-${_.moduleName}`)?.addEventListener("click",async h=>{let S=h.target,O=S.innerHTML;S.innerHTML="\u23F3 Loading...",S.disabled=!0;try{await _.onExport()}catch{N("Gagal export data")}finally{S.innerHTML=O,S.disabled=!1}}),document.getElementById(`btn-template-${_.moduleName}`)?.addEventListener("click",()=>{_.onTemplate()});let b=document.getElementById(`input-import-${_.moduleName}`);b?.addEventListener("change",async h=>{let S=h.target.files[0];if(!S)return;let O=document.getElementById(`label-import-${_.moduleName}`),z=O?O.querySelector(".import-text"):null,G=z?z.innerText:"";z&&(z.innerText="\u231B Memproses..."),O&&(O.style.pointerEvents="none"),b.disabled=!0;try{let L=await Ge(S);if(L.length===0)throw new Error("File kosong atau format salah");await _.onImport(L),H("Import berhasil!"),xe(r),I()}catch(L){N(L.message||"Gagal import data")}finally{z&&(z.innerText=G),O&&(O.style.pointerEvents="auto"),b.disabled=!1,b.value=""}})}async function I(){w();let b=document.getElementById("table-container");if(!b)return;b.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let h=A==="client",S=h?1:c,O=h?Le:20,z=new URLSearchParams({page:S,limit:O,...Object.fromEntries(Object.entries(v).filter(([,U])=>U))}),G=await q(`${r}?${z}`);if(!G.ok){b.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${G.data?.error||"Error"}</p></div>`;return}let L=G.data?.data||G.data||[],R=G.data?.pagination,V=L.length;if(h){L=j(L);let U=L.length,Y=20,ee=Math.ceil(U/Y);c>ee&&ee>0&&(c=ee);let W=(c-1)*Y,te=c*Y;L=L.slice(W,te),R={page:c,limit:Y,total:U,pages:ee}}!1,C&&C(L);let J=nt({columns:i,data:L,onEdit:$?U=>B(U):null,actions:T.map(U=>({...U,handler:Y=>U.handler(Y,I)})),emptyText:`Tidak ada ${String(o||"").toLowerCase()}`,bulkSelect:g?{selectedIds:x,onToggle:w}:null});b.innerHTML="",b.appendChild(J);let Q=document.getElementById("pagination-container");if(Q&&(Q.innerHTML="",R&&R.pages>1)){let U=it({page:R.page,pages:R.pages,total:R.total,limit:R.limit,onPage:Y=>{c=Y,I()}});U&&Q.appendChild(U)}}function K(b){let h=typeof m=="function"?m(b):m;return Ue(h)}function B(b){let h=!!b,S=document.createElement("form");if(S.noValidate=!0,S.innerHTML=K(b),h){let z=typeof m=="function"?m(b):m;ot(S,b)}let{close:O}=Z({title:h?`Edit ${o}`:`Tambah ${o}`,content:S,size:"lg",confirmText:h?"Simpan Perubahan":`Tambah ${o}`,onConfirm:async(z,G)=>{if(!S.reportValidity())return;let L=z.querySelector(".modal-confirm");L.disabled=!0,L.textContent="Menyimpan...";let R=rt(S),V=typeof m=="function"?m(b):m,J=async ee=>{for(let W of ee)if(W.type==="row")await J(W.fields);else if(W.type==="combobox"&&R[W.name]){let te=R[W.name],de=(W.options||[]).find(ie=>{let ce=String(typeof ie=="object"?ie.value:ie),Pe=String(typeof ie=="object"?ie.label:ie);return ce===te||Pe===te});if(de)R[W.name]=typeof de=="object"?de.value:de;else if(W.createApi){let ie={};ie[W.createApi.field]=te,W.createApi.extra&&Object.assign(ie,W.createApi.extra);let ce=await q(W.createApi.path,{method:"POST",body:JSON.stringify(ie)});if(ce.ok&&ce.data?.id)R[W.name]=ce.data.id;else if(ce.ok&&!ce.data?.id)R[W.name]=te;else throw new Error(`Gagal membuat master data: ${ce.data?.error||"Unknown error"}`)}}};try{await J(V)}catch(ee){N(ee.message),L.disabled=!1,L.textContent=h?"Simpan Perubahan":`Tambah ${o}`;return}E&&(R=await E(R,b));let Q=h?"PUT":"POST",U=h?`${r}/${b.id}`:r,Y=await q(U,{method:Q,body:JSON.stringify(R)});Y.ok?(H(h?`${o} berhasil diperbarui.`:`${o} berhasil ditambahkan.`),G(),xe(r),I()):(N(Y.data?.error||"Gagal menyimpan data."),L.disabled=!1,L.textContent=h?"Simpan Perubahan":`Tambah ${o}`)}})}function P(b){re(`Hapus ${o} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let h=await q(`${r}/${b.id}`,{method:"DELETE"});h.ok?(H(`${o} berhasil dihapus.`),xe(r),I()):N(h.data?.error||"Gagal menghapus.")},`Hapus ${o}`)}return I(),I}var Ie=null,lt=null;async function Be(e=!1){if(Ie&&!e)return console.log("Employees Raw (Cache Hit)",Ie.slice(0,5)),Ie;let t=await q(`/api/employees?limit=${Le}&status=Aktif`);return Ie=(t.data?.data||[]).map(n=>({value:n.id,label:n.full_name})),console.log("Employees Raw",t.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",Ie.slice(0,5)),Ie}async function be(e=!1){let n=(await Be(e)).map(r=>({value:r.label,label:r.label}));return console.log("Employee Options",n.slice(0,5)),n}async function se(e=!1){return lt&&!e||(lt=((await q("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name}))),lt}function he(e){let t={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!e||e==="-"||String(e).trim()===""?"":`<span class="badge ${t[e]||"badge-neutral"}">${e}</span>`}function Wt(e){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[e]||"badge-neutral"}">${e||"-"}</span>`}function Se(e){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[e]||"badge-neutral"}">${e||"-"}</span>`}pe();var vt=[],Je=[];function vn(e){if(!Array.isArray(e))return"Q3";let t=["Q4","Q3","Q2","Q1"];for(let n of t)if(e.some(r=>r.period===n))return n;return"Q3"}async function Qt(e,t){vt=await se();let n=await be();Je=["Berlin Ariansyah","Ade Surahman"];let r=D=>D&&!Je.find(_=>String(typeof _=="object"?_.value:_).toLowerCase()===String(D).toLowerCase())?[...Je,D]:Je,i=await q(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),m=D=>{if(!D||D==="-"||String(D).trim()==="")return"";let _=String(D).split("-");return _.length===3&&_[0].length===4?`${_[2]}-${_[1]}-${_[0]}`:D},f=i.data?.data||[],u=vn(f),o=t?t.get("dash_filter"):null,p=new Date,$=`${p.getFullYear()}-${String(p.getMonth()+1).padStart(2,"0")}`,F={},E=t&&t.get("month")?t.get("month"):null;o==="inspeksi"?F={status:"Done",activity_type:"Inspeksi Hygiene",month:E}:o==="gcdc"?F={status:"Done",activity_type:"GCDC",month:E}:o&&o.startsWith("period_")&&(F={period:o.replace("period_","").toUpperCase()});let C=new Date().getFullYear(),T=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((D,_)=>{let g=String(_+1).padStart(2,"0");return{value:`${C}-${g}`,label:`${D} ${C}`}});me({container:e,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!1,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:F,onDataLoaded:D=>D.sort((_,g)=>{let A=_.opening_date?new Date(_.opening_date).getTime():0;return(g.opening_date?new Date(g.opening_date).getTime():0)-A}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:D=>Wt(D)},{key:"period",label:"Periode",render:D=>Se(D)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:D=>m(D)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:D=>m(D)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:D=>m(D)},{key:"status",label:"Status",render:D=>he(D)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:vt},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:T},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Je}],formFields:D=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:vt,value:D?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:D?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:D?.period},{name:"pic",label:"PIC",type:"combobox",options:r(D?.pic),value:D?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:D?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:D?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:D?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:D?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:D?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let D=await q(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(D.ok){let _=D.data.data.map(g=>({Cabang:g.branch_name||"",Kegiatan:g.activity_type||"",Periode:g.period||"",PIC:g.pic||"","Tgl Opening":g.opening_date||"","Tgl Target":g.target_date||"","Tgl Selesai":g.completion_date||"",Status:g.status||""}));ne(_,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async D=>{let g=(await q("/api/branches?all=1")).data?.data||[],A=v=>{if(!v)return null;let x=String(v||"").toLowerCase().trim(),w=x.match(/^(\d{3})/);if(w){let d=g.find(y=>String(y.code||"")===w[1]);if(d)return d.id}let a=g.find(d=>String(d.full_name||"").toLowerCase()===x||String(d.code||"").toLowerCase()===x||String(d.name||"").toLowerCase()===x);return a?a.id:null},s=v=>{if(v==null||v==="")return"";if(v instanceof Date&&!isNaN(v.getTime()))return v.toISOString().slice(0,10);let x=String(v).trim();if(x===""||x==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(x))return x.slice(0,10);if(/^\d{4,5}$/.test(x)){let a=Number(x);if(a>2e4&&a<99999){let d=new Date(Date.UTC(1899,11,30)+a*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let w=x.split(/[\/\-\.]/);if(w.length===3){let[a,d,y]=w.map(k=>k.trim());if(a.length===4&&d.length<=2&&y.length<=2)return`${a}-${d.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&d.length<=2&&a.length<=2)return`${y}-${d.padStart(2,"0")}-${a.padStart(2,"0")}`}return x},l=D.map(v=>({branch_id:A(String(v.Cabang||"").trim()),branch_name:String(v.Cabang||"").trim(),activity_type:String(v.Kegiatan||"").trim(),period:String(v.Periode||"").trim(),pic:String(v.PIC||v.Pic||"").trim(),opening_date:s(v["Tgl Opening"]||v["Tanggal Opening"]||v["Tgl Openir"]),target_date:s(v["Tgl Target"]||v["Tanggal Target"]),completion_date:s(v["Tgl Selesai"]||v["Tanggal Selesai"]),status:String(v.Status||"").trim(),notes:String(v.Catatan||v.Keterangan||"").trim()})).filter(v=>v.activity_type&&v.period),c=await q("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return xe("schedule"),c.data}}})}ge();var kt=[{id:1,full_name:"ACHMAD NURYAINI",branch_id:7,branch_name:"007. Bekasi",division:"FACILITY CARE",phone:"081298765431",join_date:"2023-10-09",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:2,full_name:"ADAM JHORDY",branch_id:32,branch_name:"032. Cawang",division:"FACILITY CARE",phone:"085712345678",join_date:"2024-03-18",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:3,full_name:"ADE MUSTIKA REZDAFITRA",branch_id:16,branch_name:"016. Cibubur",division:"FACILITY CARE",phone:"081345678901",join_date:"2022-06-21",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:4,full_name:"ADE SURAHMAN",branch_id:0,branch_name:"000. Management",division:"FACILITY CARE",phone:"081223344556",join_date:"2021-01-15",status:"Aktif",notes:"Trainer & Supervisor Operasional",target_pindah_os:"",target_selesai_os:""},{id:5,full_name:"ADITYA",branch_id:26,branch_name:"026. Tangerang City",division:"FACILITY CARE",phone:"087811223344",join_date:"2024-10-19",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:6,full_name:"ADITYA PRATAMA",branch_id:54,branch_name:"054. Bogor Gunung Batu",division:"FACILITY CARE",phone:"089699887766",join_date:"2025-09-04",status:"Aktif",notes:"Lulusan Training Batch 5 nilai tertinggi",target_pindah_os:"",target_selesai_os:""},{id:7,full_name:"AGUNG SEPTIADI",branch_id:0,branch_name:"000. Management",division:"FC - RELIEFER",phone:"082155667788",join_date:"2026-02-12",status:"Aktif",notes:"Personil cover cabang mobile",target_pindah_os:"",target_selesai_os:""}],Xt=["FACILITY CARE","SECURITY","FC - RELIEFER"],kn=["Aktif","Tidak Aktif","Resign","Cut"];function ze(e){if(!e)return"-";try{let t=String(e).trim();if(/^\d{5}(\.\d+)?$/.test(t)){let r=new Date(Math.round((Number(t)-25569)*86400*1e3));if(!isNaN(r.getTime()))return r.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}let n=new Date(t);return isNaN(n.getTime())?t:n.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Zt(e){if(!e||e==="-")return"\u{1F464}";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function ea(e,t){let n=[],r=[],i="all",m="all",f="all",u="",o=!1,p=localStorage.getItem("fcms_employees_view_mode")||"card",$=t?t.get("dash_filter"):null;$==="active"?i="Aktif":$==="reliefer"&&(m="FC - RELIEFER");let F=ae(),E=X(F,"employees","write"),C=X(F,"employees","write"),j=X(F,"employees","delete");try{r=await se()}catch(l){console.error("Error fetching branches:",l)}async function T(){o=!0,s();try{let l=await q("/api/employees?limit=10000");if(l.ok){let c=l.data?.data||l.data||[];n=Array.isArray(c)&&c.length>0?c:kt}else n=kt}catch(l){console.warn("Gagal load employees, menggunakan data lokal:",l),n=kt}finally{o=!1,s()}}async function D(){if(!X(F,"export")){N("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:l}=await Promise.resolve().then(()=>(pe(),fe)),c=n.map((v,x)=>({No:x+1,"Nama Lengkap":v.full_name||"",Cabang:v.branch_name||"",Divisi:v.division||"","No. HP":v.phone||"","Tgl Masuk":v.join_date||"",Status:v.status||"","Target Pindah OS":v.target_pindah_os||"","Target Selesai OS":v.target_selesai_os||"",Catatan:v.notes||""}));l(c,`Data_Karyawan_FCMS_${new Date().toISOString().split("T")[0]}`),H("Berhasil mengekspor data karyawan ke Excel")}catch(l){N("Gagal ekspor Excel: "+l.message)}}function _(l=null){let c=!!l;if(c&&!C){N("Anda tidak memiliki izin untuk mengedit data karyawan.");return}if(!c&&!E){N("Anda tidak memiliki izin untuk menambah data karyawan.");return}let v=!!l,x=l?.division||"FACILITY CARE",w=l?.status||"Aktif",a=document.createElement("div");a.className="emp-modal-body",a.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Lengkap Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <input type="text" id="emp-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: Achmad Nuryaini" value="${l?.full_name?String(l.full_name).replace(/"/g,"&quot;"):""}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              No. Handphone / WhatsApp
            </label>
            <input type="tel" id="emp-modal-phone" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="08xx-xxxx-xxxx" value="${l?.phone?String(l.phone).replace(/"/g,"&quot;"):""}" />
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${r.map(d=>`
                <option value="${d.value}" ${String(l?.branch_id)===String(d.value)?"selected":""}>
                  ${d.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi / Bagian <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-division" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${Xt.map(d=>`
                <option value="${d}" ${x===d?"selected":""}>${d}</option>
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
              value="${l?.join_date||""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${kn.map(d=>`
                <option value="${d}" ${w===d?"selected":""}>${d}</option>
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
              value="${l?.target_pindah_os?String(l.target_pindah_os).slice(0,7):""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Target Selesai OS (Opsional)
            </label>
            <input type="month" id="emp-modal-target-selesai" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${l?.target_selesai_os?String(l.target_selesai_os).slice(0,7):""}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan / Keterangan Karyawan
          </label>
          <textarea id="emp-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Catatan keahlian, tugas khusus, riwayat cover...">${l?.notes||""}</textarea>
        </div>
      </div>
    `,Z({title:v?"\u270F\uFE0F Edit Data Karyawan":"\u2795 Tambah Data Karyawan",content:a,size:"md",confirmText:v?"Simpan Perubahan":"Simpan Karyawan",confirmClass:"btn-primary",onConfirm:async(d,y)=>{let k=(a.querySelector("#emp-modal-name").value||"").trim(),M=(a.querySelector("#emp-modal-phone").value||"").trim(),I=a.querySelector("#emp-modal-branch").value,K=a.querySelector("#emp-modal-division").value,B=a.querySelector("#emp-modal-joindate").value||null,P=a.querySelector("#emp-modal-status").value,b=a.querySelector("#emp-modal-target-pindah").value||null,h=a.querySelector("#emp-modal-target-selesai").value||null,S=(a.querySelector("#emp-modal-notes").value||"").trim();if(!k||!I||!K||!P){N("Harap lengkapi Nama Lengkap, Cabang, Divisi, dan Status!");return}let O={full_name:k,phone:M,branch_id:Number(I),division:K,join_date:B,status:P,target_pindah_os:b,target_selesai_os:h,notes:S};try{let z;v?z=await q(`/api/employees/${l.id}`,{method:"PUT",body:JSON.stringify(O)}):z=await q("/api/employees",{method:"POST",body:JSON.stringify(O)}),z.ok?(H(v?"Data karyawan berhasil diperbarui!":"Karyawan baru berhasil ditambahkan!"),y(),await T()):N(z.data?.error||"Gagal menyimpan data karyawan")}catch(z){N("Terjadi kesalahan: "+z.message)}}})}async function g(l){re({title:"Nonaktifkan Karyawan?",message:`Apakah Anda yakin ingin menonaktifkan personil <strong>${l.full_name}</strong> dari cabang <em>${l.branch_name||"cabang"}</em>?`,confirmText:"Ya, Nonaktifkan",confirmClass:"btn-danger",onConfirm:async()=>{try{let c=await q(`/api/employees/${l.id}`,{method:"DELETE"});c.ok?(H("Status karyawan berhasil diubah menjadi Tidak Aktif"),await T()):N(c.data?.error||"Gagal menonaktifkan data")}catch(c){N("Terjadi kesalahan: "+c.message)}}})}function A(){return n.filter(l=>{if(u){let c=u.toLowerCase(),v=(l.full_name||"").toLowerCase().includes(c),x=(l.branch_name||"").toLowerCase().includes(c),w=(l.phone||"").toLowerCase().includes(c),a=(l.division||"").toLowerCase().includes(c),d=(l.notes||"").toLowerCase().includes(c);if(!v&&!x&&!w&&!a&&!d)return!1}return!(i!=="all"&&(i==="Aktif"&&l.status!=="Aktif"||i==="Nonaktif"&&l.status==="Aktif")||m!=="all"&&l.division!==m||f!=="all"&&String(l.branch_id)!==String(f))})}function s(){let l=A(),c=n.length,v=n.filter(h=>h.status==="Aktif").length,x=n.filter(h=>h.division==="FACILITY CARE"&&h.status==="Aktif").length,w=n.filter(h=>(h.division==="SECURITY"||h.division==="FC - RELIEFER")&&h.status==="Aktif").length,a=n.length,d=n.filter(h=>h.status==="Aktif").length,y=n.filter(h=>h.status!=="Aktif").length;e.innerHTML=`
      <style>
        .emp-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
          width: 100%;
          max-width: 100%;
          min-width: 0;
          box-sizing: border-box;
          overflow-x: hidden;
        }
        .emp-hero-card {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
          border-radius: 16px;
          padding: 24px 28px;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
          position: relative;
          overflow: hidden;
          box-sizing: border-box;
          width: 100%;
          max-width: 100%;
          min-width: 0;
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
          box-sizing: border-box;
          width: 100%;
          max-width: 100%;
          min-width: 0;
        }
        .emp-search-row {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
          align-items: center;
        }
        .emp-view-toggle {
          display: inline-flex;
          background: #F1F5F9;
          padding: 4px;
          border-radius: 10px;
          border: 1px solid #E2E8F0;
          gap: 4px;
          flex-shrink: 0;
        }
        .emp-view-btn {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 6px 12px;
          border-radius: 7px;
          border: none;
          background: transparent;
          color: #64748B;
          font-size: 0.82rem;
          font-weight: 700;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .emp-view-btn:hover {
          color: #0F172A;
        }
        .emp-view-btn.active {
          background: white;
          color: #2563EB;
          box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }
        .emp-table-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
          box-shadow: 0 2px 8px rgba(0,0,0,0.04);
          overflow: hidden;
          width: 100%;
          box-sizing: border-box;
        }
        .emp-table-responsive {
          width: 100%;
          overflow-x: auto;
        }
        .emp-table {
          width: 100%;
          border-collapse: collapse;
          text-align: left;
          font-size: 0.85rem;
        }
        .emp-table th {
          background: #F8FAFC;
          color: #475569;
          font-weight: 700;
          font-size: 0.75rem;
          text-transform: uppercase;
          letter-spacing: 0.04em;
          padding: 13px 16px;
          border-bottom: 1px solid #E2E8F0;
          white-space: nowrap;
        }
        .emp-table td {
          padding: 13px 16px;
          border-bottom: 1px solid #F1F5F9;
          color: #1E293B;
          vertical-align: middle;
        }
        .emp-table tbody tr {
          transition: background 0.15s ease;
        }
        .emp-table tbody tr:hover {
          background: #F8FAFC;
        }
        .emp-table tbody tr:last-child td {
          border-bottom: none;
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
          grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
          gap: 16px;
          width: 100%;
          max-width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .emp-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
          padding: 16px 18px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.04);
          display: flex;
          flex-direction: column;
          gap: 13px;
          width: 100% !important;
          max-width: 100% !important;
          min-width: 0 !important;
          box-sizing: border-box !important;
          overflow: hidden;
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
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
          flex-wrap: wrap;
        }
        
        .emp-div-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.74rem;
          font-weight: 700;
          flex-shrink: 0;
          white-space: nowrap;
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
          font-size: 0.74rem;
          font-weight: 700;
          flex-shrink: 0;
          white-space: nowrap;
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
          gap: 12px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .emp-avatar {
          width: 50px;
          height: 50px;
          border-radius: 12px;
          background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
          color: white;
          display: flex;
          align-items: center;
          justify-content: center;
          font-weight: 800;
          font-size: 1.1rem;
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
          font-size: 1.02rem;
          font-weight: 800;
          color: #0F172A;
          line-height: 1.25;
          word-break: break-word;
          overflow-wrap: anywhere;
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
          padding: 11px 12px;
          display: flex;
          flex-direction: column;
          gap: 8px;
          font-size: 0.83rem;
          width: 100%;
          max-width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .emp-info-row {
          display: flex;
          align-items: flex-start;
          justify-content: space-between;
          gap: 8px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .emp-info-label {
          color: #64748B;
          display: inline-flex;
          align-items: center;
          gap: 5px;
          flex-shrink: 0;
          white-space: nowrap;
          font-size: 0.8rem;
        }
        .emp-info-val {
          color: #0F172A;
          font-weight: 600;
          text-align: right;
          min-width: 0;
          flex: 1;
          word-break: break-word;
          overflow-wrap: anywhere;
          font-size: 0.82rem;
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
          font-size: 0.78rem;
          color: #475569;
          background: #FFFBEB;
          border: 1px solid #FEF3C7;
          border-radius: 8px;
          padding: 8px 10px;
          line-height: 1.4;
          word-break: break-word;
          overflow-wrap: anywhere;
          width: 100%;
          box-sizing: border-box;
        }

        /* Card footer */
        .emp-card-footer {
          display: flex;
          justify-content: space-between;
          align-items: center;
          border-top: 1px solid #F1F5F9;
          padding-top: 11px;
          margin-top: auto;
          gap: 8px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
          flex-wrap: wrap;
        }
        .emp-action-btns {
          display: flex;
          align-items: center;
          gap: 6px;
          margin-left: auto;
          flex-shrink: 0;
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
          .emp-page {
            gap: 14px;
            width: 100% !important;
            max-width: 100% !important;
            overflow-x: hidden !important;
          }
          .emp-hero-card {
            padding: 16px 14px !important;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
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
            padding: 14px !important;
            gap: 12px;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
          .emp-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .emp-select {
            width: 100%;
          }
          .emp-view-toggle {
            width: 100%;
            display: flex;
          }
          .emp-view-btn {
            flex: 1;
            justify-content: center;
          }
          .emp-cards-grid {
            grid-template-columns: minmax(0, 1fr) !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .emp-card {
            padding: 14px !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .emp-info-card {
            padding: 10px !important;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
        }
      </style>

      <div class="emp-page">

        <!-- Hero Header Card -->
        <div class="emp-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">\u{1F465}</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Master Karyawan FCMS
                </h1>
              </div>
            </div>

            <div class="emp-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${E?`
                <button id="emp-btn-add" class="btn btn-primary emp-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Karyawan</span>
                </button>
              `:""}
              
              ${X(F,"export")?`<button id="emp-btn-excel" class="btn btn-ghost emp-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="emp-btn-refresh" class="btn btn-ghost emp-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${o?"spin-icon":""}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="emp-stats-row">
            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(37, 99, 235, 0.2); color:#93C5FD;">\u{1F465}</div>
              <div>
                <div class="emp-stat-val">${c}</div>
                <div class="emp-stat-lbl">Total Karyawan</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">\u{1F7E2}</div>
              <div>
                <div class="emp-stat-val">${v}</div>
                <div class="emp-stat-lbl">Personil Aktif</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(14, 165, 233, 0.2); color:#7DD3FC;">\u{1F3E2}</div>
              <div>
                <div class="emp-stat-val">${x}</div>
                <div class="emp-stat-lbl">Facility Care</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(124, 58, 237, 0.2); color:#C4B5FD;">\u{1F6E1}\uFE0F</div>
              <div>
                <div class="emp-stat-val">${w}</div>
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
              <option value="all">Semua Cabang (${r.length})</option>
              ${r.map(h=>`
                <option value="${h.value}" ${String(f)===String(h.value)?"selected":""}>
                  ${h.label}
                </option>
              `).join("")}
            </select>

            <select id="emp-division-filter" class="emp-select">
              <option value="all">Semua Divisi</option>
              ${Xt.map(h=>`
                <option value="${h}" ${m===h?"selected":""}>
                  ${h}
                </option>
              `).join("")}
            </select>

            <div class="emp-view-toggle" title="Pilih Tampilan">
              <button type="button" class="emp-view-btn ${p==="card"?"active":""}" data-view="card">
                <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="emp-view-btn ${p==="table"?"active":""}" data-view="table">
                <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="emp-filter-pills">
            <div class="emp-pill ${i==="all"&&m==="all"?"active":""}" data-status="all" data-div="all">
              <span>Semua Karyawan</span>
              <span class="emp-pill-count">${a}</span>
            </div>
            <div class="emp-pill ${i==="Aktif"?"active":""}" data-status="Aktif" data-div="all">
              <span class="emp-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="emp-pill-count">${d}</span>
            </div>
            <div class="emp-pill ${m==="FC - RELIEFER"?"active":""}" data-status="all" data-div="FC - RELIEFER">
              <span class="emp-pill-dot" style="background:#8B5CF6;"></span>
              <span>FC - Reliefer</span>
              <span class="emp-pill-count">${n.filter(h=>h.division==="FC - RELIEFER").length}</span>
            </div>
            <div class="emp-pill ${i==="Nonaktif"?"active":""}" data-status="Nonaktif" data-div="all">
              <span class="emp-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="emp-pill-count">${y}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive List Container (Card / Table) -->
        <div id="emp-list-container">
          ${l.length===0?`
            <div class="emp-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada karyawan yang cocok</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian, filter cabang, atau filter divisi.</p>
              <button id="emp-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
          `:p==="table"?`
            <div class="emp-table-card">
              <div class="emp-table-responsive">
                <table class="emp-table">
                  <thead>
                    <tr>
                      <th style="width:45px; text-align:center;">No</th>
                      <th>Karyawan</th>
                      <th>Cabang</th>
                      <th>Divisi</th>
                      <th>No. HP</th>
                      <th>Tgl Masuk</th>
                      <th>Target Migrasi OS / Catatan</th>
                      <th style="text-align:center;">Status</th>
                      <th style="text-align:center;">Aksi</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${l.map((h,S)=>{let O=h.division||"FACILITY CARE",z="emp-div-fc",G="";O==="SECURITY"?(z="emp-div-sec",G="sec"):O==="FC - RELIEFER"&&(z="emp-div-rel",G="rel");let L=h.status==="Aktif",R=L?"emp-status-aktif":"emp-status-nonaktif",V=L?"\u{1F7E2} Aktif":`\u26AA ${h.status||"Nonaktif"}`;return`
                        <tr data-id="${h.id}">
                          <td style="text-align:center; font-weight:700; color:#64748B;">${S+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="emp-avatar ${G}" style="width:36px; height:36px; border-radius:9px; font-size:0.85rem;">
                                ${Zt(h.full_name)}
                              </div>
                              <div>
                                <div style="font-weight:700; color:#0F172A; white-space:nowrap;">${h.full_name||"-"}</div>
                                <div style="font-size:0.75rem; color:#64748B;">ID #${h.id}</div>
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#334155;">
                            \u{1F4CD} ${h.branch_name||"Cabang Tidak Ditentukan"}
                          </td>
                          <td style="white-space:nowrap;">
                            <span class="emp-div-badge ${z}">
                              ${O==="SECURITY"?"\u{1F6E1}\uFE0F":O==="FC - RELIEFER"?"\u{1F504}":"\u{1F9F9}"} ${O}
                            </span>
                          </td>
                          <td style="white-space:nowrap;">
                            ${h.phone?`
                              <a href="tel:${h.phone}" class="emp-phone-link">
                                \u{1F4DE} ${h.phone}
                              </a>
                            `:'<span style="color:#94A3B8;">-</span>'}
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            ${ze(h.join_date)}
                          </td>
                          <td style="max-width:260px;">
                            ${h.target_pindah_os||h.target_selesai_os?`
                              <div style="font-size:0.78rem; font-weight:600; color:#4F46E5; margin-bottom:2px;">
                                \u{1F504} ${h.target_pindah_os?"Pindah: "+ze(h.target_pindah_os):""}
                                ${h.target_selesai_os?(h.target_pindah_os?" \u2192 ":"")+"Selesai: "+ze(h.target_selesai_os):""}
                              </div>
                            `:""}
                            ${h.notes?`
                              <div style="font-size:0.78rem; color:#475569; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${String(h.notes).replace(/"/g,"&quot;")}">
                                \u{1F4DD} ${h.notes}
                              </div>
                            `:!h.target_pindah_os&&!h.target_selesai_os?'<span style="color:#94A3B8;">-</span>':""}
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="emp-status-badge ${R}">
                              ${V}
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <div style="display:inline-flex; align-items:center; gap:6px;">
                              ${h.phone?`
                                <a href="https://wa.me/${h.phone.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="emp-icon-btn" title="Chat WhatsApp" style="color:#16A34A;">
                                  <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                                </a>
                              `:""}
                              ${C?`
                                <button class="emp-icon-btn edit" data-action="edit" data-id="${h.id}" title="Edit Data Karyawan">
                                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              `:""}
                              ${j?`
                                <button class="emp-icon-btn delete" data-action="delete" data-id="${h.id}" title="Nonaktifkan Karyawan">
                                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                                </button>
                              `:""}
                            </div>
                          </td>
                        </tr>
                      `}).join("")}
                  </tbody>
                </table>
              </div>
            </div>
          `:`
            <div class="emp-cards-grid">
              ${l.map(h=>{let S=h.division||"FACILITY CARE",O="emp-div-fc",z="";S==="SECURITY"?(O="emp-div-sec",z="sec"):S==="FC - RELIEFER"&&(O="emp-div-rel",z="rel");let G=h.status==="Aktif",L=G?"emp-status-aktif":"emp-status-nonaktif",R=G?"\u{1F7E2} Aktif":`\u26AA ${h.status||"Nonaktif"}`;return`
              <div class="emp-card" data-id="${h.id}">
                <!-- Top Badges Row -->
                <div class="emp-card-top">
                  <span class="emp-div-badge ${O}">
                    ${S==="SECURITY"?"\u{1F6E1}\uFE0F":S==="FC - RELIEFER"?"\u{1F504}":"\u{1F9F9}"} ${S}
                  </span>
                  <div style="display:flex; align-items:center; gap:8px;">
                    <span class="emp-status-badge ${L}">
                      ${R}
                    </span>
                  </div>
                </div>

                <!-- Profile Box -->
                <div class="emp-profile-box">
                  <div class="emp-avatar ${z}">
                    ${Zt(h.full_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="emp-name" style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${h.full_name||"-"}
                    </div>
                    <div class="emp-subtitle">
                      Personil ${S}
                    </div>
                  </div>
                </div>

                <!-- Information Grid -->
                <div class="emp-info-card">
                  <div class="emp-info-row">
                    <span class="emp-info-label">\u{1F4CD} Cabang:</span>
                    <span class="emp-info-val">${h.branch_name||"Cabang Tidak Ditentukan"}</span>
                  </div>
                  <div class="emp-info-row">
                    <span class="emp-info-label">\u{1F4DE} No. HP:</span>
                    <span class="emp-info-val">
                      ${h.phone?`
                        <a href="tel:${h.phone}" class="emp-phone-link">
                          <span>${h.phone}</span>
                        </a>
                      `:'<span style="color:#94A3B8;">-</span>'}
                    </span>
                  </div>
                  <div class="emp-info-row">
                    <span class="emp-info-label">\u{1F5D3}\uFE0F Tgl Masuk:</span>
                    <span class="emp-info-val">${ze(h.join_date)}</span>
                  </div>
                  ${h.target_pindah_os||h.target_selesai_os?`
                    <div class="emp-info-row target-os" style="border-top:1px dashed #E2E8F0; padding-top:6px; margin-top:2px; flex-direction:column; align-items:flex-start; gap:3px; width:100%;">
                      <span class="emp-info-label" style="font-weight:600; color:#475569;">\u{1F504} Target Migrasi OS:</span>
                      <span class="emp-info-val" style="font-size:0.78rem; color:#4F46E5; text-align:left; width:100%; word-break:break-word; overflow-wrap:anywhere;">
                        ${h.target_pindah_os?"Pindah: "+ze(h.target_pindah_os):""}
                        ${h.target_selesai_os?(h.target_pindah_os?" \u2192 ":"")+"Selesai: "+ze(h.target_selesai_os):""}
                      </span>
                    </div>
                  `:""}
                </div>

                <!-- Notes box (if any) -->
                ${h.notes?`
                  <div class="emp-notes-box">
                    \u{1F4DD} <strong>Catatan:</strong> ${h.notes}
                  </div>
                `:""}

                <!-- Footer Actions -->
                <div class="emp-card-footer">
                  <div style="font-size:0.75rem; color:#94A3B8;">
                    ID #${h.id}
                  </div>

                  <div class="emp-action-btns">
                    ${h.phone?`
                      <a href="https://wa.me/${h.phone.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="emp-icon-btn" title="Chat WhatsApp" style="color:#16A34A;">
                        <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                      </a>
                    `:""}

                    ${C?`
                      <button class="emp-icon-btn edit" data-action="edit" data-id="${h.id}" title="Edit Data Karyawan">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                    `:""}
                    ${j?`
                      <button class="emp-icon-btn delete" data-action="delete" data-id="${h.id}" title="Nonaktifkan Karyawan">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                      </button>
                    `:""}
                  </div>
                </div>
              </div>
            `}).join("")}
            </div>
          `}
        </div>

      </div>
    `,e.querySelectorAll(".emp-view-btn").forEach(h=>{h.addEventListener("click",()=>{let S=h.getAttribute("data-view");S&&S!==p&&(p=S,localStorage.setItem("fcms_employees_view_mode",S),s())})});let k=e.querySelector("#emp-search-input");k&&k.addEventListener("input",h=>{u=h.target.value,s();let S=e.querySelector("#emp-search-input");S&&(S.focus(),S.setSelectionRange(S.value.length,S.value.length))});let M=e.querySelector("#emp-branch-filter");M&&M.addEventListener("change",h=>{f=h.target.value,s()});let I=e.querySelector("#emp-division-filter");I&&I.addEventListener("change",h=>{m=h.target.value,s()}),e.querySelectorAll(".emp-pill").forEach(h=>{h.addEventListener("click",()=>{i=h.dataset.status;let S=h.dataset.div;S&&S!=="all"?m=S:h.dataset.status==="all"&&(m="all"),s()})});let K=e.querySelector("#emp-reset-filter");K&&K.addEventListener("click",()=>{u="",f="all",m="all",i="all",s()});let B=e.querySelector("#emp-btn-add");B&&B.addEventListener("click",()=>_());let P=e.querySelector("#emp-btn-excel");P&&P.addEventListener("click",()=>D());let b=e.querySelector("#emp-btn-refresh");b&&b.addEventListener("click",()=>T()),e.querySelectorAll(".emp-icon-btn.edit").forEach(h=>{h.addEventListener("click",S=>{let O=Number(S.currentTarget.getAttribute("data-id")),z=n.find(G=>G.id===O);z&&_(z)})}),e.querySelectorAll(".emp-icon-btn.delete").forEach(h=>{h.addEventListener("click",S=>{let O=Number(S.currentTarget.getAttribute("data-id")),z=n.find(G=>G.id===O);z&&g(z)})})}await T()}ge();pe();function Ye(e){if(!e||e==="-")return"-";if(String(e).startsWith("2099"))return"Tetap / PKWTT";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Fe(e){return e?String(e).startsWith("2099"):!1}function ta(e){if(!e||e==="-")return"\u{1F464}";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function aa(e,t){let n=[],r=[],i=[],m={},f="",u="all",o="all",p="all",$="all",F=!1,E=localStorage.getItem("fcms_contracts_view_mode")||"card",C=t?t.get("dash_filter"):null;C==="expiring30"?($="30",p="Aktif"):C==="active"&&(p="Aktif");let j=ae(),T=X(j,"contracts","write"),D=X(j,"contracts","write"),_=X(j,"contracts","delete");try{let[a,d]=await Promise.all([se(),Be()]);r=Array.isArray(a)?a:[],i=Array.isArray(d)?d:[],i.forEach(y=>{let k=y.value!==void 0?y.value:y.id,M=y.label||y.full_name;k&&y.phone&&(m[k]=y.phone),M&&y.phone&&(m[M.toLowerCase()]=y.phone)})}catch(a){console.error("Error fetching cache options for contracts:",a)}async function g(){F=!0,w();try{let a=await q("/api/contracts?limit=10000");if(a.ok){let d=a.data?.data||a.data||[];n=Array.isArray(d)?d:[]}else n=[],N("Gagal memuat data kontrak dari server.")}catch(a){console.error("Error load contracts:",a),n=[],N("Terjadi kesalahan koneksi saat memuat kontrak.")}finally{F=!1,w()}}function A(){return n.filter(a=>{if(f.trim()){let d=f.toLowerCase().trim(),y=String(a.employee_name||"").toLowerCase(),k=String(a.branch_name||"").toLowerCase(),M=String(a.pkwt_number||"").toLowerCase(),I=String(a.contract_type||"").toLowerCase(),K=String(a.notes||"").toLowerCase();if(!(y.includes(d)||k.includes(d)||M.includes(d)||I.includes(d)||K.includes(d)))return!1}if(u!=="all"&&String(a.branch_id)!==String(u)||o!=="all"&&a.division!==o||p==="Aktif"&&a.status!=="Aktif")return!1;if(p==="Expiring30"){if(!!!(a.end_date&&String(a.end_date).trim()&&a.end_date!=="-"))return!1;let y=Fe(a.end_date),k=a.days_remaining!==null&&a.days_remaining!==void 0&&a.days_remaining!==""?Number(a.days_remaining):null;if(a.status!=="Aktif"||y||k===null||isNaN(k)||k<0||k>30)return!1}if(p==="Tetap"&&(!Fe(a.end_date)||a.status!=="Aktif")||p==="Nonaktif"&&a.status==="Aktif")return!1;if($!=="all"){if(!!!(a.end_date&&String(a.end_date).trim()&&a.end_date!=="-"))return!1;let y=parseInt($,10),k=Fe(a.end_date),M=a.days_remaining!==null&&a.days_remaining!==void 0&&a.days_remaining!==""?Number(a.days_remaining):null;if(k||a.status!=="Aktif"||M===null||isNaN(M)||M<0||M>y)return!1}return!0})}async function s(){Z({title:"\u{1F50D} Analisis Selisih Kontrak Karyawan",content:`
        <div style="text-align:center; padding:24px 10px;">
          <div style="font-size:2rem; margin-bottom:10px;">\u23F3</div>
          <p style="color:#64748B; font-size:0.9rem;">Sedang menyinkronkan data karyawan aktif dengan kontrak aktif di database...</p>
        </div>
      `,cancelText:"Tutup"});try{let[a,d]=await Promise.all([q("/api/employees?limit=10000&status=Aktif"),q("/api/contracts?limit=10000")]),y=a.ok?a.data?.data||a.data||[]:[],k=d.ok?d.data?.data||d.data||[]:[],M=new Set(k.filter(B=>B.status==="Aktif").map(B=>B.employee_id)),I=y.filter(B=>!M.has(B.id)),K="";I.length===0?K=`
          <div style="text-align:center; padding:20px 10px;">
            <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F389}</div>
            <h4 style="font-size:1.1rem; font-weight:800; color:#10B981; margin:0 0 8px 0;">Semua Karyawan Aktif Memiliki Kontrak!</h4>
            <p style="font-size:0.86rem; color:#64748B; margin:0;">
              Total <strong>${y.length}</strong> karyawan aktif tercatat memiliki perikatan kontrak aktif yang sah di sistem.
            </p>
          </div>
        `:K=`
          <div style="display:flex; flex-direction:column; gap:12px;">
            <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:10px; padding:12px; font-size:0.84rem; color:#991B1B;">
              Ditemukan <strong>${I.length}</strong> dari <strong>${y.length}</strong> karyawan aktif yang <strong>belum memiliki Kontrak Aktif</strong>.
            </div>
            <div style="max-height:360px; overflow-y:auto; display:flex; flex-direction:column; gap:8px; padding-right:4px;">
              ${I.map(B=>{let P=k.filter(h=>h.employee_id===B.id),b='<span style="color:#D97706; font-weight:600;">Belum pernah ada data kontrak</span>';if(P.length>0){let h=P[0];b=`Riwayat: Status <strong>${h.status}</strong> (Selesai: ${Ye(h.end_date)})`}return`
                  <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:8px; padding:10px 12px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">
                    <div>
                      <div style="font-weight:700; font-size:0.88rem; color:#0F172A;">${B.full_name}</div>
                      <div style="font-size:0.76rem; color:#64748B;">
                        Cabang: ${B.branch_name||"-"} \u2022 Divisi: ${B.division||"FACILITY CARE"}
                      </div>
                      <div style="font-size:0.75rem; color:#475569; margin-top:2px;">
                        ${b}
                      </div>
                    </div>
                    ${D?`
                      <button class="btn btn-primary btn-add-contract-for-emp" data-empid="${B.id}" data-empname="${String(B.full_name).replace(/"/g,"&quot;")}" data-branchid="${B.branch_id||""}" data-division="${B.division||"FACILITY CARE"}" style="font-size:0.76rem; padding:5px 10px; border-radius:6px; background:#2563EB;">
                        + Buat Kontrak
                      </button>
                    `:""}
                  </div>
                `}).join("")}
            </div>
          </div>
        `,Z({title:"Analisis Selisih Kontrak Karyawan",content:K,cancelText:"Tutup"}),setTimeout(()=>{document.querySelectorAll(".btn-add-contract-for-emp").forEach(B=>{B.addEventListener("click",P=>{let b=P.currentTarget,h=b.getAttribute("data-empid"),S=b.getAttribute("data-empname"),O=b.getAttribute("data-branchid"),z=b.getAttribute("data-division"),G=document.querySelector(".modal-close");G&&G.click(),c({employee_id:h,employee_name:S,branch_id:O,division:z,status:"Aktif"})})})},100)}catch(a){console.error(a),N("Gagal melakukan analisis selisih data kontrak.")}}function l(a,d){if(!a)return"";let y=String(a).split("-").map(Number);if(y.length!==3||!y[0]||!y[1]||!y[2])return"";let k=new Date(y[0],y[1]-1,y[2]);if(isNaN(k.getTime()))return"";if(d==="KONTRAK 6 BULAN")k.setMonth(k.getMonth()+6),k.setDate(k.getDate()-1);else if(d==="KONTRAK 1 TAHUN")k.setFullYear(k.getFullYear()+1),k.setDate(k.getDate()-1);else if(d==="KONTRAK 2 TAHUN")k.setFullYear(k.getFullYear()+2),k.setDate(k.getDate()-1);else return"";let M=k.getFullYear(),I=String(k.getMonth()+1).padStart(2,"0"),K=String(k.getDate()).padStart(2,"0");return`${M}-${I}-${K}`}function c(a=null){let d=!!(a&&a.id);if(d&&!T){N("Anda tidak memiliki izin untuk mengedit kontrak.");return}if(!d&&!D){N("Anda tidak memiliki izin untuk menambah kontrak.");return}let y=!!(a&&a.id),k=y?`\u270F\uFE0F Edit Kontrak: ${a.employee_name||"Karyawan"}`:"\u2795 Tambah Kontrak Karyawan Baru",M=a?.employee_id||"",I=a?.branch_id||"",K=a?.division||"FACILITY CARE",B=a?.status||"Aktif",P=a?.contract_type||"KONTRAK 1 TAHUN",b=a?.pkwt_number||"PKWT 1",h=a?.notes||"",S=!1;a&&(a.contract_type==="PKWTT (TETAP)"||a.pkwt_number==="TETAP"?S=!0:a.contract_type&&(a.contract_type.startsWith("KONTRAK")||a.contract_type.includes("BULAN")||a.contract_type.includes("TAHUN"))||a.pkwt_number&&a.pkwt_number.startsWith("PKWT")?S=!1:S=Fe(a.end_date));let O=a?.start_date?String(a.start_date).slice(0,10):"",z=a?.end_date&&!Fe(a.end_date)?String(a.end_date).slice(0,10):"";!S&&O&&!z&&(z=l(O,P));let G=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-employee" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Karyawan --</option>
              ${i.map(R=>{let V=R.value!==void 0?R.value:R.id,J=R.label||R.full_name||"Karyawan",Q=R.division||R.branch_name?` (${R.division||"FC"} - ${R.branch_name||"Cabang -"})`:"";return`
                <option value="${V}" ${String(M)===String(V)?"selected":""}>
                  ${J}${Q}
                </option>
              `}).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${r.map(R=>`
                <option value="${R.value}" ${String(I)===String(R.value)?"selected":""}>
                  ${R.label}
                </option>
              `).join("")}
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi / Bagian <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-division" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="FACILITY CARE" ${K==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${K==="SECURITY"?"selected":""}>SECURITY</option>
              <option value="FC - RELIEFER" ${K==="FC - RELIEFER"?"selected":""}>FC - RELIEFER</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Kontrak <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Aktif" ${B==="Aktif"?"selected":""}>\u{1F7E2} Aktif</option>
              <option value="Tidak Aktif" ${B==="Tidak Aktif"?"selected":""}>\u26AA Tidak Aktif</option>
              <option value="Resign" ${B==="Resign"?"selected":""}>\u{1F534} Resign</option>
              <option value="Cut" ${B==="Cut"?"selected":""}>\u2702\uFE0F Cut / Pemutusan</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tipe / Durasi Kontrak
            </label>
            <select id="con-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="KONTRAK 6 BULAN" ${P==="KONTRAK 6 BULAN"?"selected":""}>KONTRAK 6 BULAN</option>
              <option value="KONTRAK 1 TAHUN" ${P==="KONTRAK 1 TAHUN"?"selected":""}>KONTRAK 1 TAHUN</option>
              <option value="KONTRAK 2 TAHUN" ${P==="KONTRAK 2 TAHUN"?"selected":""}>KONTRAK 2 TAHUN</option>
              <option value="PKWTT (TETAP)" ${P==="PKWTT (TETAP)"?"selected":""}>PKWTT (TETAP)</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tingkat / No. PKWT
            </label>
            <select id="con-modal-pkwt" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="PKWT 1" ${b==="PKWT 1"?"selected":""}>PKWT 1</option>
              <option value="PKWT 2" ${b==="PKWT 2"?"selected":""}>PKWT 2</option>
              <option value="PKWT 3" ${b==="PKWT 3"?"selected":""}>PKWT 3</option>
              <option value="PKWT 4" ${b==="PKWT 4"?"selected":""}>PKWT 4</option>
              <option value="PKWT 5" ${b==="PKWT 5"?"selected":""}>PKWT 5</option>
              <option value="PKWT 6" ${b==="PKWT 6"?"selected":""}>PKWT 6</option>
              <option value="TETAP" ${b==="TETAP"?"selected":""}>PKWTT / TETAP</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Mulai Kontrak
            </label>
            <input type="date" id="con-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${O}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Selesai Kontrak
            </label>
            <input type="date" id="con-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; ${S?"opacity:0.5; background-color:#F8FAFC; cursor:not-allowed;":"background-color:#FFFFFF;"}" 
              value="${z}" ${S?"disabled":""} />
            <div style="margin-top:6px; display:flex; align-items:center; gap:6px;">
              <input type="checkbox" id="con-modal-is-tetap" ${S?"checked":""} style="cursor:pointer;" />
              <label for="con-modal-is-tetap" style="font-size:0.78rem; color:#475569; cursor:pointer; user-select:none;">
                Karyawan Tetap / PKWTT (Tanpa Batas Tanggal)
              </label>
            </div>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan Tambahan
          </label>
          <textarea id="con-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; resize:vertical;" 
            placeholder="Keterangan perpanjangan, klausul khusus, atau nomor SK...">${h}</textarea>
        </div>
      </div>
    `,L=Z({title:k,content:G,confirmText:y?"\u{1F4BE} Simpan Perubahan":"\u2795 Tambah Kontrak",cancelText:"Batal",onConfirm:async(R,V)=>{let J=R.querySelector("#con-modal-employee"),Q=R.querySelector("#con-modal-branch"),U=R.querySelector("#con-modal-division"),Y=R.querySelector("#con-modal-status"),ee=R.querySelector("#con-modal-start"),W=R.querySelector("#con-modal-end"),te=R.querySelector("#con-modal-is-tetap"),de=R.querySelector("#con-modal-type"),ie=R.querySelector("#con-modal-pkwt"),ce=R.querySelector("#con-modal-notes"),Pe=J?.value,Ot=Q?.value,on=U?.value||"FACILITY CARE",sn=Y?.value||"Aktif",ut=ee?.value?.trim()||null,Ht=te?.checked,ln=W?.value?.trim(),gt=Ht?"2099-12-31":ln||null,dn=de?.value,pn=ie?.value,cn=ce?.value?.trim()||null;if(!Pe){N("Harap pilih nama karyawan!"),J?.focus();return}if(!Ot){N("Harap pilih cabang penempatan!"),Q?.focus();return}if(!Ht&&ut&&gt&&gt<ut){N("Tanggal selesai kontrak tidak boleh lebih awal dari tanggal mulai kontrak!"),W?.focus();return}let un={employee_id:parseInt(Pe,10),branch_id:parseInt(Ot,10),division:on,status:sn,start_date:ut,end_date:gt,contract_type:dn,pkwt_number:pn,notes:cn};try{let mt=y?`/api/contracts/${a.id}`:"/api/contracts",bt=await q(mt,{method:y?"PUT":"POST",body:JSON.stringify(un)});if(bt.ok)H(y?"Data kontrak berhasil diperbarui!":"Kontrak baru berhasil ditambahkan!"),V(),await g();else{let gn=bt.data?.error||bt.data?.message||"Gagal menyimpan data kontrak.";N(gn)}}catch(mt){console.error(mt),N("Terjadi kesalahan koneksi ke server.")}}});setTimeout(()=>{let R=L?.overlay||document.querySelector(".modal-overlay:last-child");if(!R)return;let V=R.querySelector("#con-modal-is-tetap"),J=R.querySelector("#con-modal-end"),Q=R.querySelector("#con-modal-start"),U=R.querySelector("#con-modal-type"),Y=R.querySelector("#con-modal-pkwt");if(!V||!J)return;function ee(W,te=!0){W?(V.checked=!0,J.disabled=!0,J.style.opacity="0.5",J.style.backgroundColor="#F8FAFC",J.style.cursor="not-allowed",J.value="",te&&(U&&(U.value="PKWTT (TETAP)"),Y&&(Y.value="TETAP"))):(V.checked=!1,J.disabled=!1,J.style.opacity="1",J.style.backgroundColor="#FFFFFF",J.style.cursor="default",te&&(U&&U.value==="PKWTT (TETAP)"&&(U.value="KONTRAK 1 TAHUN"),Y&&Y.value==="TETAP"&&(Y.value="PKWT 1")))}V.addEventListener("change",W=>{ee(W.target.checked,!0),W.target.checked||J.focus()}),U&&U.addEventListener("change",W=>{let te=W.target.value;te==="PKWTT (TETAP)"?(ee(!0,!1),Y&&(Y.value="TETAP")):(ee(!1,!1),Y&&Y.value==="TETAP"&&(Y.value="PKWT 1"),Q&&Q.value&&(J.value=l(Q.value,te)))}),Y&&Y.addEventListener("change",W=>{W.target.value==="TETAP"?(U&&(U.value="PKWTT (TETAP)"),ee(!0,!1)):(U&&U.value==="PKWTT (TETAP)"&&(U.value="KONTRAK 1 TAHUN"),ee(!1,!1),Q&&Q.value&&U&&(J.value=l(Q.value,U.value)))}),Q&&Q.addEventListener("change",W=>{!V.checked&&W.target.value&&U&&(J.value=l(W.target.value,U.value))})},20)}function v(a){Z({title:"\u{1F5D1}\uFE0F Hapus Data Kontrak",content:`
        <div style="font-size:0.92rem; color:#334155; line-height:1.5;">
          Apakah Anda yakin ingin menghapus catatan kontrak untuk <strong>${a.employee_name||"karyawan ini"}</strong> (ID #${a.id})?
          <div style="margin-top:10px; padding:10px; background:#FEF2F2; border-radius:8px; border:1px solid #FECACA; font-size:0.8rem; color:#991B1B;">
            \u26A0\uFE0F Tindakan ini akan menghapus riwayat masa berlaku kontrak terkait dari sistem.
          </div>
        </div>
      `,confirmText:"Ya, Hapus",cancelText:"Batal",onConfirm:async(d,y)=>{try{let k=await q(`/api/contracts/${a.id}`,{method:"DELETE"});k.ok?(H("Data kontrak berhasil dihapus."),y(),await g()):N("Gagal menghapus kontrak: "+(k.data?.message||"Terjadi kesalahan"))}catch(k){console.error(k),N("Terjadi kendala koneksi.")}}})}function x(){if(!X(j,"export")){N("Anda tidak memiliki izin untuk mengekspor data.");return}if(!n.length){N("Belum ada data kontrak untuk diekspor.");return}let a=n.map(d=>{let y=Fe(d.end_date),k=!!(d.end_date&&String(d.end_date).trim()&&d.end_date!=="-"),M=d.days_remaining!==null&&d.days_remaining!==void 0&&d.days_remaining!==""?Number(d.days_remaining):null;return{"Nama Lengkap":d.employee_name||"",Cabang:d.branch_name||"",Divisi:d.division||"","No. PKWT":d.pkwt_number||"","Tipe Kontrak":d.contract_type||"","Tanggal Mulai":d.start_date||"","Tanggal Selesai":y?"Tetap / PKWTT":k?d.end_date:"-","Sisa Hari":y?"Tetap":k&&M!==null&&!isNaN(M)?M+" Hari":"-",Status:d.status||"",Catatan:d.notes||""}});ne(a,`Data_Kontrak_Karyawan_${new Date().toISOString().slice(0,10)}`),H("File Excel Data Kontrak berhasil diunduh!")}function w(){let a=A(),d=n.length,y=n.filter(L=>L.status==="Aktif").length,k=n.filter(L=>{if(!!!(L.end_date&&String(L.end_date).trim()&&L.end_date!=="-"))return!1;let V=Fe(L.end_date),J=L.days_remaining!==null&&L.days_remaining!==void 0&&L.days_remaining!==""?Number(L.days_remaining):null;return L.status==="Aktif"&&!V&&J!==null&&!isNaN(J)&&J>=0&&J<=30}).length,M=n.filter(L=>Fe(L.end_date)&&L.status==="Aktif").length,I=n.filter(L=>L.status!=="Aktif").length;e.innerHTML=`
      <style>
        .con-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
          width: 100%;
          max-width: 100%;
          min-width: 0;
          box-sizing: border-box;
          overflow-x: hidden;
        }

        /* Hero Header Card */
        .con-hero-card {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
          border-radius: 16px;
          padding: 24px 28px;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
          position: relative;
          overflow: hidden;
          box-sizing: border-box;
          width: 100%;
          max-width: 100%;
          min-width: 0;
        }
        .con-hero-card::after {
          content: '';
          position: absolute;
          top: -40px;
          right: -40px;
          width: 220px;
          height: 220px;
          background: radial-gradient(circle, rgba(59, 130, 246, 0.2) 0%, transparent 70%);
          border-radius: 50%;
          pointer-events: none;
        }
        .con-hero-actions {
          display: flex;
          align-items: center;
          gap: 8px;
          flex-wrap: wrap;
        }
        .con-hero-actions .btn {
          height: 38px;
          box-sizing: border-box;
          transition: all 0.2s ease;
        }
        .con-btn-add:hover {
          background: #1D4ED8 !important;
          transform: translateY(-1px);
        }
        .con-btn-excel:hover, .con-btn-refresh:hover {
          background: rgba(255, 255, 255, 0.16) !important;
          border-color: rgba(255, 255, 255, 0.3) !important;
          transform: translateY(-1px);
        }
        .con-btn-missing:hover {
          background: rgba(239, 68, 68, 0.25) !important;
          border-color: rgba(239, 68, 68, 0.6) !important;
          transform: translateY(-1px);
        }
        .con-stats-row {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin-top: 20px;
        }
        .con-stat-box {
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
        .con-stat-box:hover {
          background: rgba(255, 255, 255, 0.08);
          transform: translateY(-2px);
        }
        .con-stat-icon {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 1.3rem;
          flex-shrink: 0;
        }
        .con-stat-val {
          font-size: 1.55rem;
          font-weight: 800;
          line-height: 1.1;
          letter-spacing: -0.02em;
        }
        .con-stat-lbl {
          font-size: 0.72rem;
          text-transform: uppercase;
          color: #94A3B8;
          font-weight: 700;
          letter-spacing: 0.05em;
          margin-top: 2px;
        }

        /* Controls / Filter Bar */
        .con-controls-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
          padding: 16px 20px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.03);
          display: flex;
          flex-direction: column;
          gap: 14px;
          box-sizing: border-box;
          width: 100%;
          max-width: 100%;
          min-width: 0;
        }
        .con-search-row {
          display: flex;
          gap: 12px;
          align-items: center;
          flex-wrap: wrap;
        }
        .con-search-wrap {
          position: relative;
          flex: 1;
          min-width: 240px;
        }
        .con-search-icon {
          position: absolute;
          left: 14px;
          top: 50%;
          transform: translateY(-50%);
          color: #94A3B8;
          pointer-events: none;
        }
        .con-search-input {
          width: 100%;
          padding: 10px 14px 10px 40px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          outline: none;
          transition: border-color 0.15s ease;
        }
        .con-search-input:focus {
          border-color: #2563EB;
          box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
        }
        .con-select {
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
        .con-select:focus {
          border-color: #2563EB;
        }

        /* Swipeable Pills */
        .con-filter-pills {
          display: flex;
          align-items: center;
          gap: 8px;
          overflow-x: auto;
          white-space: nowrap;
          scrollbar-width: none;
          padding-bottom: 2px;
        }
        .con-filter-pills::-webkit-scrollbar {
          display: none;
        }
        .con-pill {
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
        .con-pill:hover {
          background: #F1F5F9;
          color: #1E293B;
        }
        .con-pill.active {
          background: #0F172A;
          color: white;
          border-color: #0F172A;
        }
        .con-pill-dot {
          width: 8px;
          height: 8px;
          border-radius: 50%;
        }
        .con-pill-count {
          padding: 1px 6px;
          border-radius: 10px;
          font-size: 0.72rem;
          background: rgba(0,0,0,0.06);
        }
        .con-pill.active .con-pill-count {
          background: rgba(255,255,255,0.2);
          color: white;
        }

        /* View Mode Toggle */
        .con-view-toggle {
          display: inline-flex;
          align-items: center;
          background: #F1F5F9;
          padding: 3px;
          border-radius: 10px;
          border: 1px solid #CBD5E1;
          gap: 3px;
          flex-shrink: 0;
        }
        .con-view-btn {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 6px 12px;
          border-radius: 7px;
          font-size: 0.82rem;
          font-weight: 600;
          border: none;
          background: transparent;
          color: #64748B;
          cursor: pointer;
          transition: all 0.15s ease;
          height: 34px;
          box-sizing: border-box;
          user-select: none;
        }
        .con-view-btn:hover {
          color: #0F172A;
          background: rgba(255, 255, 255, 0.5);
        }
        .con-view-btn.active {
          background: #FFFFFF;
          color: #2563EB;
          font-weight: 700;
          box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
        }

        /* Modern Table View */
        .con-table-card {
          background: #FFFFFF;
          border: 1px solid #E2E8F0;
          border-radius: 14px;
          box-shadow: 0 2px 6px rgba(0, 0, 0, 0.03);
          overflow: hidden;
        }
        .con-table-responsive {
          width: 100%;
          overflow-x: auto;
          -webkit-overflow-scrolling: touch;
        }
        .con-table {
          width: 100%;
          border-collapse: collapse;
          text-align: left;
          font-size: 0.88rem;
          min-width: 960px;
        }
        .con-table th {
          background: #F8FAFC;
          color: #475569;
          font-size: 0.76rem;
          font-weight: 700;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          padding: 13px 16px;
          border-bottom: 1.5px solid #E2E8F0;
          white-space: nowrap;
        }
        .con-table td {
          padding: 13px 16px;
          border-bottom: 1px solid #F1F5F9;
          color: #1E293B;
          vertical-align: middle;
        }
        .con-table tbody tr:hover td {
          background: #F8FAFC;
        }
        .con-table tbody tr:last-child td {
          border-bottom: none;
        }

        /* Cards Grid */
        .con-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
          gap: 16px;
          width: 100%;
          max-width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .con-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
          padding: 16px 18px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.04);
          display: flex;
          flex-direction: column;
          gap: 13px;
          width: 100% !important;
          max-width: 100% !important;
          min-width: 0 !important;
          box-sizing: border-box !important;
          overflow: hidden;
          transition: transform 0.15s ease, box-shadow 0.15s ease;
        }
        .con-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 6px 16px rgba(0,0,0,0.08);
          border-color: #CBD5E1;
        }

        /* Card Top Badges */
        .con-card-top {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 8px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
          flex-wrap: wrap;
        }
        .con-div-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.74rem;
          font-weight: 700;
          flex-shrink: 0;
          white-space: nowrap;
        }
        .con-div-fc {
          background: #EFF6FF;
          color: #2563EB;
          border: 1px solid #DBEAFE;
        }
        .con-div-sec {
          background: #F1F5F9;
          color: #334155;
          border: 1px solid #E2E8F0;
        }
        .con-div-rel {
          background: #F5F3FF;
          color: #7C3AED;
          border: 1px solid #EDE9FE;
        }

        .con-pkwt-badge {
          display: inline-flex;
          align-items: center;
          gap: 4px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.74rem;
          font-weight: 700;
          background: #F5F3FF;
          color: #6D28D9;
          border: 1px solid #DDD6FE;
          flex-shrink: 0;
          white-space: nowrap;
        }

        .con-status-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.74rem;
          font-weight: 700;
          flex-shrink: 0;
          white-space: nowrap;
        }
        .con-status-aktif {
          background: #DCFCE7;
          color: #16A34A;
          border: 1px solid #BBF7D0;
        }
        .con-status-warning {
          background: #FEF3C7;
          color: #B45309;
          border: 1px solid #FDE68A;
        }
        .con-status-danger {
          background: #FEE2E2;
          color: #DC2626;
          border: 1px solid #FCA5A5;
        }
        .con-status-tetap {
          background: #E0E7FF;
          color: #4338CA;
          border: 1px solid #C7D2FE;
        }

        /* Profile Box */
        .con-profile-box {
          display: flex;
          align-items: center;
          gap: 12px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .con-avatar {
          width: 48px;
          height: 48px;
          border-radius: 12px;
          background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
          color: white;
          display: flex;
          align-items: center;
          justify-content: center;
          font-weight: 800;
          font-size: 1.05rem;
          flex-shrink: 0;
          box-shadow: 0 4px 10px rgba(37, 99, 235, 0.2);
        }
        .con-avatar.sec {
          background: linear-gradient(135deg, #334155 0%, #0F172A 100%);
          box-shadow: 0 4px 10px rgba(15, 23, 42, 0.2);
        }
        .con-avatar.rel {
          background: linear-gradient(135deg, #7C3AED 0%, #5B21B6 100%);
          box-shadow: 0 4px 10px rgba(124, 58, 237, 0.2);
        }
        .con-name {
          font-size: 1.02rem;
          font-weight: 800;
          color: #0F172A;
          line-height: 1.25;
          word-break: break-word;
          overflow-wrap: anywhere;
        }
        .con-subtitle {
          font-size: 0.76rem;
          color: #64748B;
          font-weight: 600;
          margin-top: 3px;
        }

        /* Info Container */
        .con-info-card {
          background: #F8FAFC;
          border: 1px solid #F1F5F9;
          border-radius: 10px;
          padding: 11px 12px;
          display: flex;
          flex-direction: column;
          gap: 8px;
          font-size: 0.83rem;
          width: 100%;
          max-width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .con-info-row {
          display: flex;
          align-items: flex-start;
          justify-content: space-between;
          gap: 8px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .con-info-label {
          color: #64748B;
          display: inline-flex;
          align-items: center;
          gap: 5px;
          flex-shrink: 0;
          white-space: nowrap;
          font-size: 0.8rem;
        }
        .con-info-val {
          color: #0F172A;
          font-weight: 600;
          text-align: right;
          min-width: 0;
          flex: 1;
          word-break: break-word;
          overflow-wrap: anywhere;
          font-size: 0.82rem;
        }

        /* Notes box */
        .con-notes-box {
          font-size: 0.78rem;
          color: #475569;
          background: #FFFBEB;
          border: 1px solid #FEF3C7;
          border-radius: 8px;
          padding: 8px 10px;
          line-height: 1.4;
          word-break: break-word;
          overflow-wrap: anywhere;
          width: 100%;
          box-sizing: border-box;
        }

        /* Card footer */
        .con-card-footer {
          display: flex;
          justify-content: space-between;
          align-items: center;
          border-top: 1px solid #F1F5F9;
          padding-top: 11px;
          margin-top: auto;
          gap: 8px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
          flex-wrap: wrap;
        }
        .con-action-btns {
          display: flex;
          align-items: center;
          gap: 6px;
          margin-left: auto;
          flex-shrink: 0;
        }
        .con-icon-btn {
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
        .con-icon-btn:hover {
          background: #F1F5F9;
          color: #0F172A;
          border-color: #CBD5E1;
        }
        .con-icon-btn.delete:hover {
          background: #FEE2E2;
          color: #DC2626;
          border-color: #FCA5A5;
        }

        /* Empty state */
        .con-empty-state {
          grid-column: 1 / -1;
          background: white;
          border: 2px dashed #E2E8F0;
          border-radius: 16px;
          padding: 48px 24px;
          text-align: center;
        }

        /* Mobile Responsive */
        @media (max-width: 768px) {
          .con-page {
            gap: 14px;
            width: 100% !important;
            max-width: 100% !important;
            overflow-x: hidden !important;
          }
          .con-hero-card {
            padding: 16px 14px !important;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
          .con-stats-row {
            grid-template-columns: 1fr 1fr;
            gap: 10px;
          }
          .con-stat-box {
            padding: 10px 12px;
          }
          .con-stat-icon {
            width: 38px;
            height: 38px;
            font-size: 1.1rem;
          }
          .con-stat-val {
            font-size: 1.3rem;
          }
          .con-hero-actions {
            width: 100% !important;
            display: flex !important;
            align-items: center !important;
            gap: 6px !important;
            flex-wrap: nowrap !important;
            overflow-x: auto !important;
            scrollbar-width: none !important;
            padding-bottom: 2px !important;
          }
          .con-hero-actions::-webkit-scrollbar {
            display: none;
          }
          .con-btn-add {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.8rem !important;
            white-space: nowrap !important;
          }
          .con-btn-excel {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.8rem !important;
            white-space: nowrap !important;
          }
          .con-btn-missing {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.8rem !important;
            white-space: nowrap !important;
          }
          .con-btn-refresh {
            width: 36px !important;
            height: 36px !important;
            padding: 0 !important;
            flex-shrink: 0 !important;
            justify-content: center !important;
          }
          .con-controls-card {
            padding: 14px !important;
            gap: 12px;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
          .con-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .con-select {
            width: 100%;
          }
          .con-view-toggle {
            width: 100%;
            display: flex;
          }
          .con-view-btn {
            flex: 1;
            justify-content: center;
          }
          .con-table-card {
            border-radius: 12px;
          }
          .con-cards-grid {
            grid-template-columns: minmax(0, 1fr) !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .con-card {
            padding: 14px !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .con-info-card {
            padding: 10px !important;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
        }
      </style>

      <div class="con-page">

        <!-- Hero Header Card -->
        <div class="con-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">\u{1F4CB}</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Kontrak Kerja Karyawan
                </h1>
              </div>
            </div>

            <!-- Action Buttons 1 Baris Sejajar -->
            <div class="con-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${D?`
                <button id="con-btn-add" class="btn btn-primary con-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Kontrak</span>
                </button>
              `:""}

              ${X(j,"export")?`
              <button id="con-btn-excel" class="btn btn-ghost con-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>
              `:""}

              ${T?`
              <button id="con-btn-missing" class="btn btn-ghost con-btn-missing" style="background:rgba(239,68,68,0.15); border:1px solid rgba(239,68,68,0.4); color:#FCA5A5; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;" title="Cek Karyawan Aktif yang belum memiliki Kontrak Aktif">
                <span>\u{1F50D}</span>
                <span>Cek Selisih</span>
              </button>
              `:""}

              <button id="con-btn-refresh" class="btn btn-ghost con-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="con-stats-row">
            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(59, 130, 246, 0.2); color:#60A5FA;">
                \u{1F4CB}
              </div>
              <div>
                <div class="con-stat-val">${d}</div>
                <div class="con-stat-lbl">Total Kontrak</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#34D399;">
                \u{1F7E2}
              </div>
              <div>
                <div class="con-stat-val" style="color:#34D399;">${y}</div>
                <div class="con-stat-lbl">Kontrak Aktif</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FBBF24;">
                \u26A0\uFE0F
              </div>
              <div>
                <div class="con-stat-val" style="color:#FBBF24;">${k}</div>
                <div class="con-stat-lbl">Habis &lt; 30 Hari</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(99, 102, 241, 0.2); color:#A5B4FC;">
                \u{1F3DB}\uFE0F
              </div>
              <div>
                <div class="con-stat-val" style="color:#A5B4FC;">${M}</div>
                <div class="con-stat-lbl">PKWTT (Tetap)</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Controls / Filter Card -->
        <div class="con-controls-card">
          <div class="con-search-row">
            <div class="con-search-wrap">
              <svg class="con-search-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              <input type="text" id="con-search-input" class="con-search-input" 
                placeholder="Cari nama karyawan, cabang, no PKWT, catatan..." value="${f?String(f).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="con-branch-filter" class="con-select">
              <option value="all">Semua Cabang (${r.length})</option>
              ${r.map(L=>`
                <option value="${L.value}" ${String(u)===String(L.value)?"selected":""}>
                  ${L.label}
                </option>
              `).join("")}
            </select>

            <select id="con-division-filter" class="con-select">
              <option value="all">Semua Divisi</option>
              <option value="FACILITY CARE" ${o==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${o==="SECURITY"?"selected":""}>SECURITY</option>
              <option value="FC - RELIEFER" ${o==="FC - RELIEFER"?"selected":""}>FC - RELIEFER</option>
            </select>

            <select id="con-expiring-filter" class="con-select">
              <option value="all" ${$==="all"?"selected":""}>Semua Durasi Masa Habis</option>
              <option value="7" ${$==="7"?"selected":""}>Akan Habis \u2264 7 Hari</option>
              <option value="14" ${$==="14"?"selected":""}>Akan Habis \u2264 14 Hari</option>
              <option value="30" ${$==="30"?"selected":""}>Akan Habis \u2264 30 Hari</option>
              <option value="60" ${$==="60"?"selected":""}>Akan Habis \u2264 60 Hari</option>
            </select>

            <div class="con-view-toggle">
              <button type="button" class="con-view-btn ${E==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="con-view-btn ${E==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="con-filter-pills">
            <div class="con-pill ${p==="all"&&$==="all"?"active":""}" data-status="all">
              <span>Semua Kontrak</span>
              <span class="con-pill-count">${d}</span>
            </div>
            <div class="con-pill ${p==="Aktif"&&$==="all"?"active":""}" data-status="Aktif">
              <span class="con-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="con-pill-count">${y}</span>
            </div>
            <div class="con-pill ${p==="Expiring30"?"active":""}" data-status="Expiring30">
              <span class="con-pill-dot" style="background:#F59E0B;"></span>
              <span>Akan Habis (&lt;30 Hari)</span>
              <span class="con-pill-count">${k}</span>
            </div>
            <div class="con-pill ${p==="Tetap"?"active":""}" data-status="Tetap">
              <span class="con-pill-dot" style="background:#6366F1;"></span>
              <span>PKWTT (Tetap)</span>
              <span class="con-pill-count">${M}</span>
            </div>
            <div class="con-pill ${p==="Nonaktif"?"active":""}" data-status="Nonaktif">
              <span class="con-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="con-pill-count">${I}</span>
            </div>
          </div>
        </div>

        <!-- Content Area: Pilihan Kartu atau Tabel -->
        <div id="con-list-container">
          ${a.length===0?`
            <div class="con-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada kontrak yang sesuai kriteria</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba atur ulang kata kunci pencarian, filter cabang, atau filter masa berlaku kontrak.</p>
              <button id="con-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
          `:E==="table"?`
            <div class="con-table-card">
              <div class="con-table-responsive">
                <table class="con-table">
                  <thead>
                    <tr>
                      <th style="width:45px; text-align:center;">No</th>
                      <th>Nama Karyawan</th>
                      <th>Cabang & Divisi</th>
                      <th style="text-align:center;">Tingkat PKWT</th>
                      <th>Mulai Kontrak</th>
                      <th>Selesai Kontrak</th>
                      <th>Masa Berlaku</th>
                      <th style="text-align:center;">Status</th>
                      <th style="width:110px; text-align:center;">Aksi</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${a.map((L,R)=>{let V=L.division||"FACILITY CARE",J="con-div-fc",Q="";V==="SECURITY"?(J="con-div-sec",Q="sec"):V==="FC - RELIEFER"&&(J="con-div-rel",Q="rel");let U=!!(L.end_date&&String(L.end_date).trim()&&L.end_date!=="-"),Y=Fe(L.end_date),ee=L.status==="Aktif",W=L.days_remaining!==null&&L.days_remaining!==void 0&&L.days_remaining!==""?Number(L.days_remaining):null,te=U&&!Y&&W!==null&&!isNaN(W)?W:null,de="con-status-aktif",ie="\u{1F7E2} Aktif";Y?(de="con-status-tetap",ie="\u{1F3DB}\uFE0F PKWTT (Tetap)"):ee?te!==null?te<=0?(de="con-status-danger",ie="\u26A0\uFE0F Habis Hari Ini"):te<=30?(de="con-status-warning",ie=`\u26A0\uFE0F Sisa ${te} Hari`):(de="con-status-aktif",ie=`\u{1F7E2} Aktif (${te} Hari)`):(de="con-status-aktif",ie="\u{1F7E2} Aktif"):(de="con-status-danger",ie=`\u26AA ${L.status||"Nonaktif"}`);let ce=m[L.employee_id]||m[String(L.employee_name||"").toLowerCase()]||"";return`
                        <tr data-id="${L.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${R+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="con-avatar ${Q}" style="width:36px; height:36px; font-size:0.8rem; border-radius:9px;">
                                ${ta(L.employee_name)}
                              </div>
                              <div>
                                <div style="font-weight:700; color:#0F172A; white-space:nowrap;">${L.employee_name||"Karyawan Tanpa Nama"}</div>
                                <div style="font-size:0.74rem; color:#64748B;">${L.contract_type||"Kontrak Kerja"} \u2022 ID #${L.id}</div>
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#0F172A; font-size:0.84rem; margin-bottom:4px;">\u{1F4CD} ${L.branch_name||"Cabang Tidak Ditentukan"}</div>
                            <span class="con-div-badge ${J}" style="padding:2px 8px; font-size:0.7rem;">
                              ${V==="SECURITY"?"\u{1F6E1}\uFE0F":V==="FC - RELIEFER"?"\u{1F504}":"\u{1F9F9}"} ${V}
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            ${L.pkwt_number?`
                              <span class="con-pkwt-badge">
                                \u{1F4DC} ${L.pkwt_number}
                              </span>
                            `:'<span style="color:#94A3B8;">-</span>'}
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            ${Ye(L.start_date)}
                          </td>
                          <td style="white-space:nowrap; font-weight:600; ${Y?"color:#4338CA; font-weight:700;":"color:#1E293B;"}">
                            ${Y?"Tetap / PKWTT":U?Ye(L.end_date):"-"}
                          </td>
                          <td style="white-space:nowrap; font-size:0.82rem; font-weight:600; ${Y?"color:#15803D;":te!==null&&te<=30&&ee?"color:#D97706; font-weight:700;":"color:#475569;"}">
                            ${Y?"Karyawan Tetap (Permanen)":te!==null?te<=0?"Masa Kontrak Telah Habis":`${te} Hari Tersisa`:"-"}
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="con-status-badge ${de}">
                              ${ie}
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <div style="display:inline-flex; align-items:center; gap:6px;">
                              ${ce?`
                                <a href="https://wa.me/${ce.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="con-icon-btn" title="Chat WhatsApp Karyawan" style="color:#16A34A;">
                                  <svg width="15" height="15" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                                </a>
                              `:""}
                              ${T?`
                                <button class="con-icon-btn edit" data-action="edit" data-id="${L.id}" title="Edit Data Kontrak">
                                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              `:""}
                              ${_?`
                                <button class="con-icon-btn delete" data-action="delete" data-id="${L.id}" title="Hapus Kontrak">
                                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                                </button>
                              `:""}
                            </div>
                          </td>
                        </tr>
                      `}).join("")}
                  </tbody>
                </table>
              </div>
            </div>
          `:`
            <div class="con-cards-grid">
              ${a.map(L=>{let R=L.division||"FACILITY CARE",V="con-div-fc",J="";R==="SECURITY"?(V="con-div-sec",J="sec"):R==="FC - RELIEFER"&&(V="con-div-rel",J="rel");let Q=!!(L.end_date&&String(L.end_date).trim()&&L.end_date!=="-"),U=Fe(L.end_date),Y=L.status==="Aktif",ee=L.days_remaining!==null&&L.days_remaining!==void 0&&L.days_remaining!==""?Number(L.days_remaining):null,W=Q&&!U&&ee!==null&&!isNaN(ee)?ee:null,te="con-status-aktif",de="\u{1F7E2} Aktif";U?(te="con-status-tetap",de="\u{1F3DB}\uFE0F PKWTT (Tetap)"):Y?W!==null?W<=0?(te="con-status-danger",de="\u26A0\uFE0F Habis Hari Ini"):W<=30?(te="con-status-warning",de=`\u26A0\uFE0F Sisa ${W} Hari`):(te="con-status-aktif",de=`\u{1F7E2} Aktif (${W} Hari)`):(te="con-status-aktif",de="\u{1F7E2} Aktif"):(te="con-status-danger",de=`\u26AA ${L.status||"Nonaktif"}`);let ie=m[L.employee_id]||m[String(L.employee_name||"").toLowerCase()]||"";return`
              <div class="con-card" data-id="${L.id}">
                <!-- Top Badges Row -->
                <div class="con-card-top">
                  <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
                    <span class="con-div-badge ${V}">
                      ${R==="SECURITY"?"\u{1F6E1}\uFE0F":R==="FC - RELIEFER"?"\u{1F504}":"\u{1F9F9}"} ${R}
                    </span>
                    ${L.pkwt_number?`
                      <span class="con-pkwt-badge">
                        \u{1F4DC} ${L.pkwt_number}
                      </span>
                    `:""}
                  </div>
                  <span class="con-status-badge ${te}">
                    ${de}
                  </span>
                </div>

                <!-- Profile Box -->
                <div class="con-profile-box">
                  <div class="con-avatar ${J}">
                    ${ta(L.employee_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="con-name">
                      ${L.employee_name||"Karyawan Tanpa Nama"}
                    </div>
                    <div class="con-subtitle">
                      Personil ${R} \u2022 ${L.contract_type||"Kontrak Kerja"}
                    </div>
                  </div>
                </div>

                <!-- Information Grid -->
                <div class="con-info-card">
                  <div class="con-info-row">
                    <span class="con-info-label">\u{1F4CD} Cabang:</span>
                    <span class="con-info-val">${L.branch_name||"Cabang Tidak Ditentukan"}</span>
                  </div>
                  <div class="con-info-row">
                    <span class="con-info-label">\u{1F5D3}\uFE0F Mulai Kontrak:</span>
                    <span class="con-info-val">${Ye(L.start_date)}</span>
                  </div>
                  <div class="con-info-row">
                    <span class="con-info-label">\u{1F3C1} Selesai Kontrak:</span>
                    <span class="con-info-val" style="${U?"color:#4338CA; font-weight:700;":""}">
                      ${U?"Tetap / PKWTT":Q?Ye(L.end_date):"-"}
                    </span>
                  </div>
                  <div class="con-info-row">
                    <span class="con-info-label">\u23F3 Masa Berlaku:</span>
                    <span class="con-info-val" style="${U?"color:#15803D;":W!==null&&W<=30&&Y?"color:#D97706; font-weight:700;":""}">
                      ${U?"Karyawan Tetap (Permanen)":W!==null?W<=0?"Masa Kontrak Telah Habis":`${W} Hari Tersisa`:"-"}
                    </span>
                  </div>
                </div>

                <!-- Notes box (if any) -->
                ${L.notes?`
                  <div class="con-notes-box">
                    \u{1F4DD} <strong>Catatan:</strong> ${L.notes}
                  </div>
                `:""}

                <!-- Footer Actions -->
                <div class="con-card-footer">
                  <div style="font-size:0.75rem; color:#94A3B8;">
                    ID #${L.id} \u2022 ${L.pkwt_number||"PKWT"}
                  </div>

                  <div class="con-action-btns">
                    ${ie?`
                      <a href="https://wa.me/${ie.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="con-icon-btn" title="Chat WhatsApp Karyawan" style="color:#16A34A;">
                        <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                      </a>
                    `:""}

                    ${T?`
                      <button class="con-icon-btn edit" data-action="edit" data-id="${L.id}" title="Edit Data Kontrak">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                    `:""}
                    ${_?`
                      <button class="con-icon-btn delete" data-action="delete" data-id="${L.id}" title="Hapus Kontrak">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                      </button>
                    `:""}
                  </div>
                </div>
              </div>
            `}).join("")}
            </div>
          `}
        </div>

      </div>
    `,e.querySelectorAll(".con-view-btn").forEach(L=>{L.addEventListener("click",()=>{let R=L.getAttribute("data-view");R&&R!==E&&(E=R,localStorage.setItem("fcms_contracts_view_mode",R),w())})});let K=e.querySelector("#con-search-input");K&&K.addEventListener("input",L=>{f=L.target.value,w()});let B=e.querySelector("#con-branch-filter");B&&B.addEventListener("change",L=>{u=L.target.value,w()});let P=e.querySelector("#con-division-filter");P&&P.addEventListener("change",L=>{o=L.target.value,w()});let b=e.querySelector("#con-expiring-filter");b&&b.addEventListener("change",L=>{$=L.target.value,w()}),e.querySelectorAll(".con-pill").forEach(L=>{L.addEventListener("click",R=>{let V=R.currentTarget.getAttribute("data-status");p=V,V==="Expiring30"?$="30":$="all",w()})});let h=e.querySelector("#con-reset-filter");h&&h.addEventListener("click",()=>{f="",u="all",o="all",p="all",$="all",w()});let S=e.querySelector("#con-btn-add");S&&S.addEventListener("click",()=>c());let O=e.querySelector("#con-btn-excel");O&&O.addEventListener("click",x);let z=e.querySelector("#con-btn-missing");z&&z.addEventListener("click",s);let G=e.querySelector("#con-btn-refresh");G&&G.addEventListener("click",()=>g()),e.querySelectorAll(".con-icon-btn.edit").forEach(L=>{L.addEventListener("click",R=>{let V=parseInt(R.currentTarget.getAttribute("data-id"),10),J=n.find(Q=>Q.id===V);J&&c(J)})}),e.querySelectorAll(".con-icon-btn.delete").forEach(L=>{L.addEventListener("click",R=>{let V=parseInt(R.currentTarget.getAttribute("data-id"),10),J=n.find(Q=>Q.id===V);J&&v(J)})})}await g()}ge();var wt=[{id:1,report_date:"2026-08-27",branch_id:19,branch_name:"019. Cinere",category:"SDM",source:"AM",complaint:"Keluhan AM fc atas nama Malik bau badan, dan sudah dibilangin berulang kali terkait kebersihan seragam dan deodoran.",employee_name:"MALIK",fc_specialist:"Mizwar",solution:"Diberikan teguran lisan, edukasi kebersihan diri, dan pendampingan seragam cadangan di klinik.",status:"Done",completion_date:"2026-08-28",day_count:1},{id:2,report_date:"2026-08-26",branch_id:0,branch_name:"000. Management",category:"SDM",source:"Berlin",complaint:"Security atas nama Reza suhelmi meninggalkan area tugas jaga tanpa personil pengganti saat jam operasional klinik.",employee_name:"REZA SUHELMI",fc_specialist:"Ade Surahman",solution:"Pemanggilan investigasi dan penerbitan Surat Peringatan (SP) sesuai SOP security.",status:"Done",completion_date:"2026-08-27",day_count:1},{id:3,report_date:"2026-07-27",branch_id:30,branch_name:"030. Ciputat",category:"SDM",source:"AM",complaint:"Terinfo kembali permasalahan FC Ciputat, terkait standar pembersihan area dental unit yang belum memenuhi checklist hygiene.",employee_name:"HENDRA",fc_specialist:"Berlin Ariansyah",solution:"Re-training modul pembersihan dental unit dan monitoring harian oleh supervisor.",status:"In Progress",completion_date:null,day_count:null},{id:4,report_date:"2026-06-24",branch_id:30,branch_name:"030. Ciputat",category:"SDM",source:"AM",complaint:"FC atas nama Hendra mengintimidasi rcp pada saat briefing pagi mengenai pembagian area kerja.",employee_name:"HENDRA",fc_specialist:"Berlin Ariansyah",solution:"Mediasi bersama AM dan HR, diberikan pembinaan konseling perilaku kerja.",status:"Done",completion_date:"2026-06-25",day_count:1},{id:5,report_date:"2026-06-05",branch_id:0,branch_name:"000. Management",category:"SDM",source:"Berlin",complaint:"Kinerja kurang dan ada komplain dari cabang mengenai ketidakhadiran berkala personil pengganti.",employee_name:"Indra",fc_specialist:"Ade",solution:"Evaluasi penempatan dan pembinaan kedisiplinan kerja.",status:"Done",completion_date:"2026-06-08",day_count:3},{id:6,report_date:"2026-06-05",branch_id:0,branch_name:"000. Management",category:"SDM",source:"Ade",complaint:"Sering tidak masuk kerja dan memalsukan surat dokter saat izin sakit.",employee_name:"Kristian",fc_specialist:"Ade Surahman",solution:"Investigasi ke klinik dokter terkait dan eskalasi ke HRD untuk tindakan tegas.",status:"Done",completion_date:"2026-06-07",day_count:2},{id:7,report_date:"2026-05-25",branch_id:19,branch_name:"019. Cinere",category:"SDM",source:"SPV",complaint:"Tidak masuk kerja tanpa kabar (alpa), dan sudah di-one on one oleh SPV terkait absensi.",employee_name:"Malik",fc_specialist:"Mizwar",solution:"Pemberian SP 1 dan penandatanganan komitmen kehadiran kerja.",status:"Done",completion_date:"2026-05-26",day_count:1}],na=["SDM","Cleaning","Aset","K3","Lainnya"],wn=["Open","In Progress","Done"],En=["SPV","AM","RCP","Perawat","FC","Security","Berlin","Ade","Pattrel","Dentrel"];function ia(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Sn(e){if(!e||e==="-")return"\u26A0\uFE0F";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function ra(e,t){let n=[],r=[],i=[],m="all",f="all",u="all",o="",p=!1;(t?t.get("dash_filter"):null)==="open"&&(m="Open");let F=ae(),E=X(F,"issues","write");try{[r,i]=await Promise.all([se(),be()])}catch(s){console.error("Error fetching cache:",s)}async function C(){p=!0,A();try{let s=await q("/api/issues?limit=10000");if(s.ok){let l=s.data?.data||s.data||[];n=Array.isArray(l)&&l.length>0?l:wt}else n=wt}catch(s){console.warn("Gagal load issues, menggunakan data lokal:",s),n=wt}finally{p=!1,A()}}async function j(){if(!X(F,"export")){N("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:s}=await Promise.resolve().then(()=>(pe(),fe)),l=n.map((c,v)=>({No:v+1,"Tanggal Info":c.report_date||"",Cabang:c.branch_name||"",Kategori:c.category||"","Sumber Laporan":c.source||"",Keluhan:c.complaint||"","Nama FC / Security":c.employee_name||"","FC Spesialis":c.fc_specialist||"","Solusi / Tindak Lanjut":c.solution||"",Status:c.status||"","Tanggal Selesai":c.completion_date||"","Durasi Hari":c.day_count!=null?c.day_count:""}));s(l,`Data_Permasalahan_FCMS_${new Date().toISOString().split("T")[0]}`),H("Berhasil mengekspor data permasalahan ke Excel")}catch(s){N("Gagal ekspor Excel: "+s.message)}}async function T(s,l){try{let c=l==="Done"?s.completion_date||new Date().toISOString().split("T")[0]:l==="Open"?null:s.completion_date,v=s.day_count;if(s.report_date&&c){let a=new Date(s.report_date),d=new Date(c);v=Math.max(0,Math.floor((d-a)/(1e3*60*60*24)))}else l==="Open"&&(v=null);let x={...s,status:l,completion_date:c,day_count:v},w=await q(`/api/issues/${s.id}`,{method:"PUT",body:JSON.stringify(x)});w.ok?(H(`Status masalah diubah ke "${l}"`),await C()):N(w.data?.error||"Gagal memperbarui status")}catch(c){N("Terjadi kesalahan: "+c.message)}}function D(s=null){let l=!!s,c=new Date().toISOString().split("T")[0],v=s?.report_date||c,x=s?.category||"SDM",w=s?.source||"AM",a=s?.status||"Open",d=document.createElement("div");d.className="issue-modal-body",d.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Info <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="iss-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${v}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${r.map(B=>`
                <option value="${B.value}" ${String(s?.branch_id)===String(B.value)?"selected":""}>
                  ${B.label}
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
              ${na.map(B=>`
                <option value="${B}" ${x===B?"selected":""}>${B}</option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Sumber Laporan
            </label>
            <input list="iss-source-options" id="iss-modal-source" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: AM, SPV, Berlin..." value="${w}" />
            <datalist id="iss-source-options">
              ${En.map(B=>`<option value="${B}">`).join("")}
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
              ${i.map(B=>`<option value="${B.label}">`).join("")}
            </datalist>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              FC Spesialis / Penanggung Jawab
            </label>
            <input list="iss-spec-options" id="iss-modal-spec" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: Mizwar, Ade, Berlin..." value="${s?.fc_specialist?String(s.fc_specialist).replace(/"/g,"&quot;"):""}" />
            <datalist id="iss-spec-options">
              ${i.map(B=>`<option value="${B.label}">`).join("")}
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
              ${wn.map(B=>`
                <option value="${B}" ${a===B?"selected":""}>${B}</option>
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
    `;let y=d.querySelector("#iss-modal-date"),k=d.querySelector("#iss-modal-completion"),M=d.querySelector("#iss-modal-daycount"),I=d.querySelector("#iss-modal-status");function K(){if(y.value&&k.value){let B=new Date(y.value),P=new Date(k.value),b=Math.floor((P-B)/(1e3*60*60*24));M.value=isNaN(b)?"":Math.max(0,b)}else M.value=""}y.addEventListener("change",K),k.addEventListener("change",K),I.addEventListener("change",()=>{I.value==="Done"&&!k.value?(k.value=new Date().toISOString().split("T")[0],K()):I.value==="Open"&&(k.value="",M.value="")}),Z({title:l?"\u270F\uFE0F Edit Data Permasalahan":"\u2795 Tambah Data Permasalahan",content:d,size:"md",confirmText:l?"Simpan Perubahan":"Simpan Permasalahan",confirmClass:"btn-primary",onConfirm:async(B,P)=>{let b=d.querySelector("#iss-modal-date").value,h=d.querySelector("#iss-modal-branch").value,S=d.querySelector("#iss-modal-category").value,O=(d.querySelector("#iss-modal-source").value||"").trim(),z=(d.querySelector("#iss-modal-complaint").value||"").trim(),G=(d.querySelector("#iss-modal-employee").value||"").trim(),L=(d.querySelector("#iss-modal-spec").value||"").trim(),R=(d.querySelector("#iss-modal-solution").value||"").trim(),V=d.querySelector("#iss-modal-status").value,J=d.querySelector("#iss-modal-completion").value||null;if(!b||!h||!z||!S){N("Harap lengkapi Tanggal, Cabang, Kategori, dan Keluhan!");return}let Q=null;if(b&&J){let Y=new Date(b),ee=new Date(J);Q=Math.max(0,Math.floor((ee-Y)/(1e3*60*60*24)))}let U={report_date:b,branch_id:Number(h),category:S,source:O,complaint:z,employee_name:G,fc_specialist:L,solution:R,status:V,completion_date:J,day_count:Q};try{let Y;l?Y=await q(`/api/issues/${s.id}`,{method:"PUT",body:JSON.stringify(U)}):Y=await q("/api/issues",{method:"POST",body:JSON.stringify(U)}),Y.ok?(H(l?"Permasalahan berhasil diperbarui!":"Permasalahan baru berhasil ditambahkan!"),P(),await C()):N(Y.data?.error||"Gagal menyimpan permasalahan")}catch(Y){N("Terjadi kesalahan: "+Y.message)}}})}async function _(s){re({title:"Hapus Permasalahan?",message:`Apakah Anda yakin ingin menghapus tiket permasalahan dari <strong>${s.branch_name||"cabang"}</strong>:<br><br><em>"${s.complaint}"</em>?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let l=await q(`/api/issues/${s.id}`,{method:"DELETE"});l.ok?(H("Permasalahan berhasil dihapus"),await C()):N(l.data?.error||"Gagal menghapus data")}catch(l){N("Terjadi kesalahan: "+l.message)}}})}function g(){return n.filter(s=>{if(o){let l=o.toLowerCase(),c=(s.complaint||"").toLowerCase().includes(l),v=(s.employee_name||"").toLowerCase().includes(l),x=(s.branch_name||"").toLowerCase().includes(l),w=(s.source||"").toLowerCase().includes(l),a=(s.category||"").toLowerCase().includes(l),d=(s.solution||"").toLowerCase().includes(l);if(!c&&!v&&!x&&!w&&!a&&!d)return!1}return!(m!=="all"&&s.status!==m||f!=="all"&&s.category!==f||u!=="all"&&String(s.branch_id)!==String(u))})}function A(){let s=g(),l=n.length,c=n.filter(z=>z.status==="Open"||z.status==="In Progress").length,v=n.filter(z=>z.status==="Done").length,x=new Date,w=x.getFullYear(),a=x.getMonth(),d=n.filter(z=>{if(!z.report_date)return!1;let G=new Date(z.report_date);return!isNaN(G.getTime())&&G.getFullYear()===w&&G.getMonth()===a}).length,y=n.length,k=n.filter(z=>z.status==="Open").length,M=n.filter(z=>z.status==="In Progress").length,I=n.filter(z=>z.status==="Done").length;e.innerHTML=`
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
        .issues-hero-actions {
          display: flex;
          align-items: center;
          gap: 8px;
          flex-wrap: wrap;
        }
        .issues-hero-actions .btn {
          height: 38px;
          box-sizing: border-box;
          transition: all 0.2s ease;
        }
        .issues-btn-add:hover {
          background: #1D4ED8 !important;
          transform: translateY(-1px);
        }
        .issues-btn-excel:hover, .issues-btn-refresh:hover {
          background: rgba(255, 255, 255, 0.16) !important;
          border-color: rgba(255, 255, 255, 0.3) !important;
          transform: translateY(-1px);
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
            gap: 6px !important;
            flex-wrap: nowrap !important;
            overflow-x: auto !important;
            scrollbar-width: none !important;
            padding-bottom: 2px !important;
          }
          .issues-hero-actions::-webkit-scrollbar {
            display: none;
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
            grid-template-columns: minmax(0, 1fr) !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .issues-card {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
            overflow: hidden;
            padding: 14px !important;
          }
        }
      </style>

      <div class="issues-page">

        <!-- Hero Header Card -->
        <div class="issues-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">\u26A0\uFE0F</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Log Permasalahan Operasional
                </h1>
              </div>
            </div>

            <div class="issues-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${E?`
                <button id="issues-btn-add" class="btn btn-primary issues-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Permasalahan</span>
                </button>
              `:""}
              
              ${X(F,"export")?`<button id="issues-btn-excel" class="btn btn-ghost issues-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="issues-btn-refresh" class="btn btn-ghost issues-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${p?"spin-icon":""}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="issues-stats-row">
            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(239, 68, 68, 0.2); color:#FCA5A5;">\u26A0\uFE0F</div>
              <div>
                <div class="issues-stat-val">${l}</div>
                <div class="issues-stat-lbl">Total Permasalahan</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FCD34D;">\u23F3</div>
              <div>
                <div class="issues-stat-val">${c}</div>
                <div class="issues-stat-lbl">Open & In Progress</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">\u2705</div>
              <div>
                <div class="issues-stat-val">${v}</div>
                <div class="issues-stat-lbl">Tuntas (Done)</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(59, 130, 246, 0.2); color:#93C5FD;">\u{1F4C5}</div>
              <div>
                <div class="issues-stat-val">${d}</div>
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
                placeholder="Cari keluhan, nama personil, sumber, cabang..." value="${o?String(o).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="issues-branch-filter" class="issues-select">
              <option value="all">Semua Cabang (${r.length})</option>
              ${r.map(z=>`
                <option value="${z.value}" ${String(u)===String(z.value)?"selected":""}>
                  ${z.label}
                </option>
              `).join("")}
            </select>

            <select id="issues-category-filter" class="issues-select">
              <option value="all">Semua Kategori</option>
              ${na.map(z=>`
                <option value="${z}" ${f===z?"selected":""}>
                  Kategori: ${z}
                </option>
              `).join("")}
            </select>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="issues-filter-pills">
            <div class="issues-pill ${m==="all"?"active":""}" data-status="all">
              <span>Semua Status</span>
              <span class="issues-pill-count">${y}</span>
            </div>
            <div class="issues-pill ${m==="Open"?"active":""}" data-status="Open">
              <span class="issues-pill-dot" style="background:#F59E0B;"></span>
              <span>Open</span>
              <span class="issues-pill-count">${k}</span>
            </div>
            <div class="issues-pill ${m==="In Progress"?"active":""}" data-status="In Progress">
              <span class="issues-pill-dot" style="background:#3B82F6;"></span>
              <span>In Progress</span>
              <span class="issues-pill-count">${M}</span>
            </div>
            <div class="issues-pill ${m==="Done"?"active":""}" data-status="Done">
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
          `:s.map(z=>{let G=z.category||"Lainnya",L=`issues-category-${G.toLowerCase().replace(/[^a-z0-9]/g,"")}`||"issues-category-lainnya",R=z.status||"Open",V="issues-status-open",J="\u26A0\uFE0F Open";R==="In Progress"?(V="issues-status-inprogress",J="\u{1F504} In Progress"):R==="Done"&&(V="issues-status-done",J="\u2705 Done");let Q=R==="Done"?"status-done":R==="In Progress"?"status-inprogress":"";return`
              <div class="issues-card" data-id="${z.id}">
                <!-- Top Badges Row -->
                <div class="issues-card-top">
                  <span class="issues-category-badge ${L}">
                    \u{1F3F7}\uFE0F ${G}
                  </span>
                  <div style="display:flex; align-items:center; gap:8px;">
                    <span class="issues-status-badge ${V}">
                      ${J}
                    </span>
                    <span style="font-size:0.8rem; color:#64748B; display:flex; align-items:center; gap:4px;">
                      \u{1F5D3}\uFE0F ${ia(z.report_date)}
                    </span>
                  </div>
                </div>

                <!-- Complaint Box -->
                <div class="issues-complaint-box ${Q}">
                  "${z.complaint||"-"}"
                </div>

                <!-- Person & Location Box -->
                <div class="issues-person-box">
                  <div class="issues-avatar-box">
                    ${Sn(z.employee_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="issues-person-sub">PERSONIL FC / SECURITY</div>
                    <div class="issues-person-name" style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${z.employee_name||"Tidak ada personil spesifik"}
                    </div>
                  </div>
                </div>

                <!-- Details Grid -->
                <div class="issues-details-grid">
                  <div class="issues-detail-item" style="grid-column: 1 / -1;">
                    <span style="color:#64748B;">\u{1F4CD} Cabang:</span>
                    <strong style="color:#0F172A;">${z.branch_name||"Cabang Tidak Diketahui"}</strong>
                  </div>
                  <div class="issues-detail-item">
                    <span style="color:#64748B;">\u{1F5E3}\uFE0F Sumber:</span>
                    <strong>${z.source||"-"}</strong>
                  </div>
                  <div class="issues-detail-item">
                    <span style="color:#64748B;">\u{1F6E1}\uFE0F Spesialis:</span>
                    <strong>${z.fc_specialist||"-"}</strong>
                  </div>
                </div>

                <!-- Solution & Completion Box (if available) -->
                ${z.solution||z.completion_date?`
                  <div class="issues-solution-box">
                    ${z.solution?`
                      <div class="issues-solution-title">
                        <span>\u{1F4A1} Tindak Lanjut / Solusi:</span>
                      </div>
                      <div style="margin-bottom:${z.completion_date?"6px":"0"};">
                        ${z.solution}
                      </div>
                    `:""}
                    ${z.completion_date?`
                      <div style="font-size:0.76rem; color:#15803D; font-weight:600; display:flex; align-items:center; gap:4px;">
                        <span>\u23F1\uFE0F Selesai: ${ia(z.completion_date)}</span>
                        ${z.day_count!=null?`<span>(Durasi: ${z.day_count} Hari)</span>`:""}
                      </div>
                    `:""}
                  </div>
                `:""}

                <!-- Footer Actions -->
                <div class="issues-card-footer">
                  <div>
                    ${E?R==="Open"?`
                        <button class="issues-quick-btn start" data-action="quick-progress" data-id="${z.id}" title="Ubah status ke In Progress">
                          <span>\u{1F504} Mulai Tangani</span>
                        </button>
                      `:R==="In Progress"?`
                        <button class="issues-quick-btn complete" data-action="quick-done" data-id="${z.id}" title="Tandai selesai">
                          <span>\u2705 Selesaikan</span>
                        </button>
                      `:`
                        <button class="issues-quick-btn reopen" data-action="quick-reopen" data-id="${z.id}" title="Buka kembali tiket masalah">
                          <span>\u21A9\uFE0F Buka Kembali</span>
                        </button>
                      `:""}
                  </div>

                  <div class="issues-action-btns">
                    ${E?`
                      <button class="issues-icon-btn edit" data-action="edit" data-id="${z.id}" title="Edit Masalah">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                      <button class="issues-icon-btn delete" data-action="delete" data-id="${z.id}" title="Hapus Masalah">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                      </button>
                    `:""}
                  </div>
                </div>
              </div>
            `}).join("")}
        </div>

      </div>
    `;let K=e.querySelector("#issues-search-input");K&&K.addEventListener("input",z=>{o=z.target.value,A();let G=e.querySelector("#issues-search-input");G&&(G.focus(),G.setSelectionRange(G.value.length,G.value.length))});let B=e.querySelector("#issues-branch-filter");B&&B.addEventListener("change",z=>{u=z.target.value,A()});let P=e.querySelector("#issues-category-filter");P&&P.addEventListener("change",z=>{f=z.target.value,A()}),e.querySelectorAll(".issues-pill").forEach(z=>{z.addEventListener("click",()=>{m=z.dataset.status,A()})});let b=e.querySelector("#issues-reset-filter");b&&b.addEventListener("click",()=>{o="",u="all",f="all",m="all",A()});let h=e.querySelector("#issues-btn-add");h&&h.addEventListener("click",()=>D());let S=e.querySelector("#issues-btn-excel");S&&S.addEventListener("click",()=>j());let O=e.querySelector("#issues-btn-refresh");O&&O.addEventListener("click",()=>C()),e.querySelectorAll(".issues-card").forEach(z=>{let G=Number(z.dataset.id),L=n.find(Y=>Y.id===G);if(!L)return;let R=z.querySelector('[data-action="edit"]');R&&R.addEventListener("click",()=>D(L));let V=z.querySelector('[data-action="delete"]');V&&V.addEventListener("click",()=>_(L));let J=z.querySelector('[data-action="quick-progress"]');J&&J.addEventListener("click",()=>T(L,"In Progress"));let Q=z.querySelector('[data-action="quick-done"]');Q&&Q.addEventListener("click",()=>T(L,"Done"));let U=z.querySelector('[data-action="quick-reopen"]');U&&U.addEventListener("click",()=>T(L,"Open"))})}await C()}var Me=[];function Fn(e,t){let n=String(e.status||"").toLowerCase();return t==="pending"?n==="pending":!1}async function oa(e,t){let n=t?t.get("dash_filter"):null;Me=await se();let r=await be(),i=["Ade","Berlin"],m=u=>u&&!r.find(o=>o.value===u)?[...r,{value:u,label:u}]:r,f=u=>u&&!i.find(o=>(typeof o=="object"?o.value:o)===u)?[...i,u]:i;me({container:e,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!1,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:u=>n?u.filter(o=>Fn(o,n)):u,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:u=>window.formatDate(u)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:u=>`<span title="${u||""}">${u?.length>50?u.slice(0,50)+"\u2026":u||"-"}</span>`},{key:"solution",label:"Solusi",render:u=>`<span title="${u||""}">${u?.length>40?u.slice(0,40)+"\u2026":u||"-"}</span>`},{key:"status",label:"Status",render:u=>he(u)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:u=>window.formatDate(u)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:u=>u?`<a href="${u}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:Me},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async u=>{let o=new URLSearchParams(u||{}).toString(),p=await q(`/api/one-on-one?limit=10000&${o}`);if(p.ok){let $=p.data.data.map(E=>({Tanggal:E.meeting_date||"",Cabang:E.branch_name||"","Nama Karyawan":E.employee_name||"",PIC:E.pic||"",Masalah:E.problem||"",Solusi:E.solution||"",Status:E.status||"","Tgl Selesai":E.completion_date||"",Dokumen:E.document_link||""})),{downloadExcel:F}=await Promise.resolve().then(()=>(pe(),fe));F($,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let u=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:o}=await Promise.resolve().then(()=>(pe(),fe));o(u,"Template_Import_OneOnOne")},onImport:async u=>{let o=E=>{if(!E)return null;let C=String(E||"").toLowerCase(),j=Me.find(T=>String(T.label||"").toLowerCase()===C);return j?j.value:null},p=E=>{if(!E)return"";if(E instanceof Date&&!isNaN(E.getTime()))return E.toISOString().slice(0,10);let C=String(E).trim();if(/^\d{4,5}$/.test(C)){let T=Number(C);if(T>2e4&&T<99999){let D=new Date(Date.UTC(1899,11,30)+T*864e5);return isNaN(D.getTime())?"":D.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(C))return C.slice(0,10);let j=C.split(/[\/\-\.]/);if(j.length===3){let[T,D,_]=j.map(g=>g.trim());if(T.length===4&&D.length<=2&&_.length<=2)return`${T}-${D.padStart(2,"0")}-${_.padStart(2,"0")}`;if(_.length===4&&D.length<=2&&T.length<=2)return`${_}-${D.padStart(2,"0")}-${T.padStart(2,"0")}`}return C},$=u.map(E=>({meeting_date:p(E.Tanggal),employee_name:String(E["Nama Karyawan"]||"").trim(),branch_id:o(String(E.Cabang||"").trim()),pic:String(E.PIC||"").trim(),problem:String(E.Masalah||"").trim(),solution:String(E.Solusi||"").trim(),status:String(E.Status||"").trim(),completion_date:p(E["Tgl Selesai"]),document_link:String(E.Dokumen||"").trim()})).filter(E=>E.meeting_date&&E.employee_name&&E.branch_id),F=await q("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:$,onDuplicate:"update"})});if(!F.ok)throw new Error(F.data?.error||"Import gagal");return F.data}},formFields:u=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:u?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:u?.branch_id&&!Me.find(o=>o.value==u.branch_id)?[...Me,{value:u.branch_id,label:u.branch_name||u.branch_id}]:Me,createApi:{path:"/api/branches",field:"full_name"},value:u?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:m(u?.employee_name),value:u?.employee_name},{name:"pic",label:"PIC",type:"select",options:f(u?.pic),createApi:{path:"/api/pic",field:"name"},value:u?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:u?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:u?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:u?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:u?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:u?.document_link}]})}var Ae={};function Ne(e){if(Ae[e]){try{Ae[e].destroy()}catch{}delete Ae[e]}}function Cn(){Object.keys(Ae).forEach(Ne)}var ye=(e,t=0)=>{let n=Number(e);return isNaN(n)||e===null||e===void 0?t:n},je=(e,t="\u2014")=>{if(e==null||e==="")return t;let n=String(e).trim();return n===""||n==="[object Object]"?t:n};function dt(e,t,n=900){if(!e)return;e._animFrame&&cancelAnimationFrame(e._animFrame);let r=Math.max(0,Math.round(ye(t))),i=e.textContent?e.textContent.replace(/[^\d]/g,""):"",m=i!==""&&parseInt(i)||0;if(r===m){e.textContent=r.toLocaleString("id-ID"),e._animFrame=null;return}let f=Date.now(),u=()=>{let o=Math.min((Date.now()-f)/n,1),p=1-Math.pow(1-o,3);e.textContent=Math.round(m+p*(r-m)).toLocaleString("id-ID"),o<1?e._animFrame=requestAnimationFrame(u):(e.textContent=r.toLocaleString("id-ID"),e._animFrame=null)};e._animFrame=requestAnimationFrame(u)}var $n={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},_n=e=>{let t=je(e,"\u2014");return`<span class="status-pill ${$n[t]||"pill-neutral"}">${t}</span>`};var ke={family:"Inter",size:11},Ce="#94A3B8",Ke="#F1F5F9",Et=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],An=()=>window.innerWidth<768;function pt(e={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:An()?"bottom":"top",labels:{font:ke,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ke,titleFont:{...ke,weight:"700"}}},scales:{x:{grid:{color:Ke},ticks:{font:ke,color:Ce,maxRotation:0}},y:{grid:{color:Ke},ticks:{font:ke,color:Ce},beginAtZero:!0}},...e}}var Bn=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),Tn=()=>Array(5).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function sa(e=3){return Array(e).fill(0).map((t,n)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${n<e-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function ve(e,t,n=8e3){try{let r=new AbortController,i=setTimeout(()=>r.abort(),n),m=await q(e,{signal:r.signal}).catch(()=>null);if(clearTimeout(i),!m||!m.ok)return t;let f=m.data;return f?f.data!==void 0?f.data??t:f:t}catch{return t}}function Dn(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(r=>{let i=document.getElementById(r);i&&(i.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(r=>{let i=document.getElementById(r);if(i&&i.style.display==="none"){i.style.display="block";let m=i.parentElement;if(m&&!m.querySelector(".chart-empty")){let f=document.createElement("div");f.className="chart-empty",f.textContent="Belum ada data",i.style.display="none",m.appendChild(f)}}});let e=document.getElementById("kpi-row");e&&e.querySelector(".skeleton")&&pa({});let t=document.getElementById("mini-stats-row");t&&t.querySelector(".skeleton")&&ca({}),["table-contracts","table-issues"].forEach(r=>{let i=document.getElementById(r);i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada data</div>')});let n=document.getElementById("activity-log");n&&n.querySelector(".skeleton")&&(n.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function da(e){Cn(),e._dashRefresh&&clearInterval(e._dashRefresh),e._skelTimeout&&clearTimeout(e._skelTimeout);let t=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});e.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">
      
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h2 style="margin:0; font-size:1.2rem; color:var(--text-1);">Dashboard Overview</h2>
        <button id="btn-dash-refresh" class="btn-primary" style="padding:6px 12px; font-size:0.8rem; display:flex; align-items:center; gap:6px; cursor:pointer; background:var(--primary); color:#fff; border:none; border-radius:6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
          Refresh Data
        </button>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${Bn()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${Tn()}</div>

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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${sa(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${sa(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>la(e)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async n=>{let r=n.target.value,i=document.getElementById("jadwal-year-label");i&&(i.textContent=r);let m=document.getElementById("skel-jadwal"),f=document.getElementById("chart-jadwal");m&&(m.style.display="block",m.style.position="absolute"),f&&(f.style.display="none");let u=await ve(`/api/dashboard/schedule-chart?year=${r}`,{},8e3);try{ua(u)}catch(o){console.warn("ScheduleChart render:",o),we("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async n=>{let r=n.target.value,i=r?`/api/dashboard/inspection-bar?month=${r}`:"/api/dashboard/inspection-bar",m=document.getElementById("skel-insp"),f=document.getElementById("chart-insp");m&&(m.style.display="block",m.style.position="absolute"),f&&(f.style.display="none");let u=await ve(i,{},8e3);try{ga(u)}catch(o){console.warn("InspBar render:",o),we("skel-insp","chart-insp")}}),e._skelTimeout=setTimeout(()=>Dn(),5e3),await la(e)}async function la(e){e._skelTimeout&&(clearTimeout(e._skelTimeout),e._skelTimeout=null);let[t,n,r,i,m,f,u,o]=await Promise.all([ve("/api/dashboard/kpi",{},8e3),ve("/api/dashboard/issues-trend",{},8e3),ve("/api/dashboard/issues-summary",{},8e3),ve("/api/dashboard/stats",{},8e3),ve("/api/dashboard/calendar",[],8e3),ve("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),ve(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),ve("/api/dashboard/dropdown-aggregates",{},8e3)]),p=document.getElementById("filter-insp-month"),$=p?p.value:"",F=$?`/api/dashboard/inspection-bar?month=${$}`:"/api/dashboard/inspection-bar",E=await ve(F,{},8e3);window.dashboardAggregates=o||{};try{pa(t)}catch(C){console.warn("KPI render:",C)}try{ca(t)}catch(C){console.warn("MiniStats render:",C)}try{ua(u)}catch(C){console.warn("ScheduleChart render:",C),we("skel-jadwal","chart-jadwal")}try{Pn(Array.isArray(r?.by_category)?r.by_category:[])}catch(C){console.warn("Donut render:",C),we("skel-donut","chart-donut")}try{Ln(n)}catch(C){console.warn("Trend render:",C),we("skel-trend","chart-trend")}try{ga(E)}catch(C){console.warn("InspBar render:",C),we("skel-insp","chart-insp")}try{let C=Array.isArray(i)?i:Array.isArray(i?.recent_issues)?i.recent_issues:[];zn(C)}catch(C){console.warn("IssuesTable render:",C)}try{let C=Array.isArray(i?.expiring_contracts)?i.expiring_contracts:[];In(f)}catch(C){console.warn("ContractsTable render:",C)}try{Mn(Array.isArray(m)?m:[])}catch(C){console.warn("Agenda render:",C)}try{jn()}catch(C){console.warn("Quick Actions render:",C)}}function pa(e){let t=document.getElementById("kpi-row");if(!t)return;e=e||{};let n=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,r=window.dashboardAggregates||{},i=o=>{let p=window.dashboardAggregates?.schedule_quarter_details?.[o];return p?`Insp: ${p.insp} \u2022 GC: ${p.gc} \u2022 DC: ${p.dc} \u2022 Fog: ${p.fog}`:""},m=o=>{let p=window.dashboardAggregates?.schedule_quarter_details?.[o];return p?`Rincian ${o}: Inspeksi: ${p.insp} | General Cleaning: ${p.gc} | Deep Cleaning: ${p.dc} | Fogging: ${p.fog}`:`Total Kegiatan ${o}`},f=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
          <option value="Q1" ${n==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${n==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${n==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${n==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:r.schedule_by_quarter?.[n]??e.schedule?.current??0,href:`#/timeline?dash_filter=period_${n.toLowerCase()}`,color:"kpi-blue",sub:i(n),subId:"mini-jadwal-detail",tooltip:m(n)}];t.innerHTML=f.map(o=>{let p=o.val!==void 0?ye(o.val,0):ye(e[o.key]?.current,0);return o.dropdown?`
        <div class="kpi-card ${o.color}" ${o.id?`id="${o.id}"`:""} ${o.tooltip?`title="${o.tooltip}"`:""} data-href="${o.href}" tabindex="0" role="link" style="cursor:pointer;padding:10px 12px">
          <div style="display:flex; gap:10px; align-items:center;">
            <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${o.icon}</span></div>
            <div style="flex:1;min-width:0;">
              <div style="display:flex; align-items:baseline; gap:3px;">
                <div class="kpi-value" data-target="${p}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${p}</div>
                ${o.dropdown}
              </div>
              <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${o.label}</div>
              <div class="kpi-subtitle" ${o.subId?`id="${o.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${o.tooltip||""}">${o.sub}</div>
            </div>
          </div>
        </div>`:`
      <a href="${o.href}" class="kpi-card ${o.color}" ${o.id?`id="${o.id}"`:""} ${o.tooltip?`title="${o.tooltip}"`:""} style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${o.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="kpi-value" data-target="${p}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${p}</div>
              ${o.dropdown?o.dropdown:""}
            </div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${o.label}</div>
            <div class="kpi-subtitle" ${o.subId?`id="${o.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${o.tooltip||""}">${o.sub}</div>
          </div>
        </div>
      </a>`}).join(""),t.querySelectorAll(".kpi-card[data-href]").forEach(o=>{o.addEventListener("click",p=>{p.target.closest("select")||o.dataset.href&&(window.location.hash=o.dataset.href)}),o.addEventListener("keydown",p=>{p.key==="Enter"&&!p.target.closest("select")&&o.dataset.href&&(window.location.hash=o.dataset.href)})}),t.querySelectorAll(".kpi-value").forEach(o=>{dt(o,parseInt(o.dataset.target)||0)});let u=document.getElementById("dash-jadwal-period");if(u){let o=p=>{let F=(window.dashboardAggregates?.schedule_by_quarter||{})[p]||0,E=document.querySelector("#mini-jadwal .kpi-value");E&&(E.dataset.target=F,dt(E,F,400));let C=document.getElementById("mini-jadwal-detail");C&&(C.textContent=i(p),C.title=m(p));let j=document.getElementById("mini-jadwal");if(j){let T=`#/timeline?dash_filter=period_${p.toLowerCase()}`;j.tagName==="A"?j.href=T:j.dataset.href=T,j.title=m(p)}};u.addEventListener("change",p=>o(p.target.value))}}function ca(e){let t=document.getElementById("mini-stats-row");if(!t)return;e=e||{};let n=new Date().getFullYear(),r=String(new Date().getMonth()+1).padStart(2,"0"),i=`${n}-${r}`,m=p=>`
    <select id="${p}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
      <option value="${n}-01" ${i===`${n}-01`?"selected":""}>Jan</option>
      <option value="${n}-02" ${i===`${n}-02`?"selected":""}>Feb</option>
      <option value="${n}-03" ${i===`${n}-03`?"selected":""}>Mar</option>
      <option value="${n}-04" ${i===`${n}-04`?"selected":""}>Apr</option>
      <option value="${n}-05" ${i===`${n}-05`?"selected":""}>Mei</option>
      <option value="${n}-06" ${i===`${n}-06`?"selected":""}>Jun</option>
      <option value="${n}-07" ${i===`${n}-07`?"selected":""}>Jul</option>
      <option value="${n}-08" ${i===`${n}-08`?"selected":""}>Agu</option>
      <option value="${n}-09" ${i===`${n}-09`?"selected":""}>Sep</option>
      <option value="${n}-10" ${i===`${n}-10`?"selected":""}>Okt</option>
      <option value="${n}-11" ${i===`${n}-11`?"selected":""}>Nov</option>
      <option value="${n}-12" ${i===`${n}-12`?"selected":""}>Des</option>
    </select>
  `,f=window.dashboardAggregates||{},u=[{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:m("dash-inspeksi-month"),val:f.inspeksi_by_month?.[i]??e.inspection_month?.current??0,href:`#/timeline?dash_filter=inspeksi&month=${i}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:m("dash-gcdc-month"),val:f.gcdc_by_month?.[i]??e.cleaning_month?.current??0,href:`#/timeline?dash_filter=gcdc&month=${i}`,color:"mini-green"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:m("dash-fogging-month"),val:f.fogging_by_month?.[i]??e.fogging_month?.current??0,href:`#/reports/fogging?dash_filter=fogging&month=${i}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:e.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:e.branches?.current,href:"#/branches",color:"mini-teal"}];t.innerHTML=u.map(p=>p.dropdown?`
        <div class="mini-stat ${p.color}" style="cursor:pointer;" id="${p.id||""}" data-href="${p.href}" tabindex="0" role="link">
          <div class="mini-stat-icon">${p.icon}</div>
          <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="mini-stat-value" data-target="${ye(p.val)}">0</div>
              ${p.dropdown}
            </div>
            <div class="mini-stat-text">${p.label}</div>
            ${p.extra||""}
          </div>
        </div>`:`
      <a href="${p.href}" class="mini-stat ${p.color}" style="text-decoration:none" id="${p.id||""}">
        <div class="mini-stat-icon">${p.icon}</div>
        <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
          <div style="display:flex; align-items:baseline; gap:3px;">
            <div class="mini-stat-value" data-target="${ye(p.val)}">0</div>
            ${p.dropdown?p.dropdown:""}
          </div>
          <div class="mini-stat-text">${p.label}</div>
          ${p.extra||""}
        </div>
      </a>`).join(""),t.querySelectorAll(".mini-stat[data-href]").forEach(p=>{p.addEventListener("click",$=>{$.target.closest("select")||p.dataset.href&&(window.location.hash=p.dataset.href)}),p.addEventListener("keydown",$=>{$.key==="Enter"&&!$.target.closest("select")&&p.dataset.href&&(window.location.hash=p.dataset.href)})}),t.querySelectorAll(".mini-stat-value").forEach(p=>dt(p,parseInt(p.dataset.target)||0,700));let o=(p,$,F,E)=>{let C=document.getElementById(p);if(C){let j=T=>{let _=(window.dashboardAggregates?.[F]||{})[T]||0,g=document.querySelector(`#${$} .mini-stat-value`);g&&(g.dataset.target=_,dt(g,_,400));let A=document.getElementById($);if(A){let s=`${E}&month=${T}`;A.tagName==="A"?A.href=s:A.dataset.href=s}};C.addEventListener("change",T=>j(T.target.value))}};o("dash-inspeksi-month","mini-inspeksi","inspeksi_by_month","#/timeline?dash_filter=inspeksi"),o("dash-gcdc-month","mini-gcdc","gcdc_by_month","#/timeline?dash_filter=gcdc"),o("dash-fogging-month","mini-fogging","fogging_by_month","#/reports/fogging?dash_filter=fogging")}function Pn(e){we("skel-donut","chart-donut");let t=document.getElementById("chart-donut"),n=document.getElementById("donut-legend");if(!t||!n)return;Ne("donut");let r=(e||[]).filter(o=>ye(o.count)>0);if(!r.length){Ve(t,"Belum ada data permasalahan");return}let i=r.map(o=>`${je(o.category,"Lainnya")}`),m=r.map(o=>ye(o.count)),f=m.reduce((o,p)=>o+p,0);n.innerHTML=r.map((o,p)=>{let $=Et[p%Et.length],F=f>0?Math.round(o.count/f*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${$}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${o.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${F}%)</span></div>
          <div class="donut-legend-label">${i[p]}</div>
        </div>
      </div>
    `}).join("");let u={id:"centerText",beforeDraw:function(o){let p=o.width,$=o.height,F=o.ctx;F.restore();let E=($/80).toFixed(2);F.font="bold "+E+"em Inter",F.textBaseline="middle",F.fillStyle="#1E293B";let C=f.toString(),j=Math.round((p-F.measureText(C).width)/2),T=$/2;F.fillText(C,j,T-4),F.font="600 "+(E*.35).toFixed(2)+"em Inter",F.fillStyle="#64748B";let D="Total",_=Math.round((p-F.measureText(D).width)/2);F.fillText(D,_,T+10),F.save()}};Ae.donut=new Chart(t,{type:"doughnut",data:{labels:i,datasets:[{data:m,backgroundColor:Et,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ke,titleFont:{...ke,weight:"700"},callbacks:{label:o=>` ${o.label}: ${o.parsed} kasus`}}},cutout:"75%"},plugins:[u]})}function Ln(e){we("skel-trend","chart-trend");let t=document.getElementById("chart-trend");if(!t)return;Ne("trend"),e=e||{};let n=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],r=(e.labels||[]).map(f=>{if(!f||typeof f!="string")return"";try{let[u,o]=f.split("-");return(n[Number(o)-1]||o)+" "+String(u).slice(-2)}catch{return f}}),i=(e.open||[]).map(f=>ye(f)),m=(e.closed||[]).map(f=>ye(f));if(!r.length){Ve(t,"Belum ada data trend");return}Ae.trend=new Chart(t,{type:"line",data:{labels:r,datasets:[{label:"Open",data:i,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:m,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:pt({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Ce,maxRotation:45,autoSkip:!0}},y:{grid:{color:Ke},ticks:{font:{family:"Inter",size:9},color:Ce},beginAtZero:!0}}})})}function ua(e){we("skel-jadwal","chart-jadwal");let t=document.getElementById("chart-jadwal");if(!t)return;Ne("jadwal"),e=e||{};let n=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(e).some(o=>Array.isArray(o)&&o.some(p=>p>0))){Ve(t,"Belum ada data jadwal");return}let i=e["Inspeksi Hygiene"]||Array(12).fill(0),m=e["General Cleaning"]||Array(12).fill(0),f=e["Deep Cleaning"]||Array(12).fill(0),u=e.Fogging||Array(12).fill(0);Ae.jadwal=new Chart(t,{type:"bar",data:{labels:n,datasets:[{label:"Inspeksi",data:i,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:m,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:f,backgroundColor:"#F59E0B"},{label:"Fogging",data:u,backgroundColor:"#EF4444"}]},options:pt({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Ce,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Ke},ticks:{font:{family:"Inter",size:9},color:Ce},min:0}}})})}function ga(e){we("skel-insp","chart-insp");let t=document.getElementById("chart-insp");if(!t)return;Ne("inspBar"),e=e||{};let n=e.labels||[],r=(e.fc||[]).map(m=>ye(m)),i=(e.spv||[]).map(m=>ye(m));if(!n.length){Ve(t,"Belum ada data inspeksi");return}Ae.inspBar=new Chart(t,{type:"bar",data:{labels:n,datasets:[{label:"Skor FC",data:r,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:i,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:pt({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ke,color:Ce,maxRotation:45,minRotation:45}},y:{grid:{color:Ke},ticks:{font:ke,color:Ce},min:0,max:100}}})})}function In(e){we("skel-contract-mini","chart-contract-mini");let t=document.getElementById("chart-contract-mini");if(!t)return;Ne("contractMiniBar"),e=e||{};let n={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},r=(e.labels||[]).map(f=>{let u=f.split("-")[1];return n[u]||f}),i=(e.data||[]).map(f=>ye(f));if(!r.length){Ve(t,"Belum ada data");return}let m=t.getContext("2d");Ae.contractMiniBar=new Chart(t,{type:"bar",data:{labels:r,datasets:[{label:"Kontrak Habis",data:i,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:pt({onClick:(f,u)=>{if(u&&u.length>0){let o=u[0].index,p=(e.labels||[])[o];p&&(window.location.hash="#/contracts?month_expiry="+p)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ke,color:Ce,maxRotation:0,autoSkip:!1}},y:{grid:{color:Ke,borderDash:[4,4],drawBorder:!1},ticks:{font:ke,color:Ce,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function zn(e){let t=document.getElementById("table-issues");if(!t)return;let n=(e||[]).slice(0,8);if(!n.length){t.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}t.innerHTML=`
    <div class="dash-list">
      ${n.map(r=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${_n(r.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${je(r.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${je(r.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function Mn(e){let t=document.getElementById("widget-agenda");if(!t)return;let n=new Date,r=`${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,"0")}-${String(n.getDate()).padStart(2,"0")}`,m=(e||[]).filter(f=>(f.event_date||"").startsWith(r)).slice(0,10);if(!m.length){t.innerHTML="";return}t.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${m.map(f=>{let u="#3B82F6",o="#EFF6FF",p="Agenda",$=(f.title||"").toLowerCase();return $.includes("inspeksi")?(u="#10B981",o="#ECFDF5",p="Inspeksi"):$.includes("cleaning")||$.includes("gcdc")?(u="#3B82F6",o="#EFF6FF",p="Cleaning"):$.includes("reliefer")?(u="#F59E0B",o="#FFFBEB",p="Reliefer"):$.includes("fogging")&&(u="#8B5CF6",o="#F5F3FF",p="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(f.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${u};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${je(f.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${je(f.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${o};color:${u}">${p}</div>
        </div>
      `}).join("")}
    </div>
  `}function jn(){let e=document.getElementById("quick-actions");if(!e)return;let t=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];e.innerHTML=t.map(n=>`
    <a href="${n.href}" class="action-btn">
      <div class="action-icon" style="background:${n.bg}">${n.icon}</div>
      ${n.label}
    </a>
  `).join("")}function we(e,t){let n=document.getElementById(e),r=document.getElementById(t);if(n&&(n.style.display="none",n.style.position=""),r){r.style.display="block";let i=r.parentElement;if(i){let m=i.querySelector(".chart-empty");m&&m.remove()}}}function Ve(e,t="Belum ada data"){if(!e)return;e.style.display="none";let n=e.parentElement;if(!n)return;if(!n.querySelector(".chart-empty")){let i=document.createElement("div");i.className="chart-empty",i.textContent=t,n.appendChild(i)}}async function ma(e){document.getElementById("app").innerHTML=`
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
  `;let t=document.getElementById("login-form"),n=document.getElementById("login-error"),r=document.getElementById("login-btn"),i=document.getElementById("toggle-password"),m=document.getElementById("login-password");i?.addEventListener("click",()=>{let f=m.type==="text";m.type=f?"password":"text",i.style.color=f?"":"var(--primary)"}),t?.addEventListener("submit",async f=>{f.preventDefault(),n.style.display="none";let u=t.username.value.trim(),o=t.password.value;if(!u||!o){n.textContent="Username dan password wajib diisi.",n.style.display="block";return}r.querySelector(".btn-text").style.display="none",r.querySelector(".btn-spinner").style.display="",r.disabled=!0;try{let p=await q("/api/auth/login",{method:"POST",body:JSON.stringify({username:u,password:o})});p.ok&&p.data.success?(Jt(p.data.data.token),tt(p.data.data.user),H("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(n.textContent=p.data.error||"Username atau password salah.",n.style.display="block",r.classList.add("shake"),setTimeout(()=>r.classList.remove("shake"),600))}catch{n.textContent="Gagal terhubung ke server. Periksa koneksi internet.",n.style.display="block"}finally{r.querySelector(".btn-text").style.display="",r.querySelector(".btn-spinner").style.display="none",r.disabled=!1}})}ge();var St=[{id:1,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:6,branch_name:"006. Kelapa Gading MOI",trainer:"Berlin Ariansyah",participants:"MUHAMMAD PRAKAS AGAZI",score:90,document_link:""},{id:2,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:11,branch_name:"011. Kuningan",trainer:"Berlin Ariansyah",participants:"KAMALLUDIN",score:88,document_link:""},{id:3,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:42,branch_name:"042. Muara Karang",trainer:"Berlin Ariansyah",participants:"SAHRUL ICHSANI",score:85,document_link:""},{id:4,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:51,branch_name:"051. Bintaro Sektor 9",trainer:"Berlin Ariansyah",participants:"MUHAMAD VICKRI FATHURROHMAN",score:92,document_link:""},{id:5,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:45,branch_name:"045. FX Mall",trainer:"Berlin Ariansyah",participants:"DEDE AHDORI",score:86,document_link:""},{id:6,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:54,branch_name:"054. Bogor Gunung Batu",trainer:"Berlin Ariansyah",participants:"ADITYA PRATAMA",score:95,document_link:""},{id:7,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:55,branch_name:"055. Gajah Mada",trainer:"Berlin Ariansyah",participants:"ELVAN JANGGA PRAMANA",score:89,document_link:""}],ba=["Batch 5","Batch 4","Batch 3","Batch 2","Batch 1"],ha=["Berlin Ariansyah","Ade Surahman","Mizwar","Fajar","Ade","Berlin"];function Kn(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Te(e){if(!e)return"-";try{let t=JSON.parse(e);if(Array.isArray(t))return t.join(", ")}catch{}return String(e)}async function fa(e){let t=[],n=[],r=[],i="all",m="all",f="all",u="",o=!1,p=ae(),$=X(p,"training","write");try{[n,r]=await Promise.all([se(),be()])}catch(g){console.error("Error fetching cache:",g)}async function F(){o=!0,D();try{let g=await q("/api/training?limit=10000");if(g.ok){let A=g.data?.data||g.data||[];t=Array.isArray(A)&&A.length>0?A:St}else t=St}catch(g){console.warn("Gagal load training, menggunakan data lokal:",g),t=St}finally{o=!1,D()}}async function E(){if(!X(p,"export")){N("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:g}=await Promise.resolve().then(()=>(pe(),fe)),A=t.map((s,l)=>({No:l+1,Tanggal:s.training_date||"",Batch:s.batch||"",Materi:s.subject||"",Cabang:s.branch_name||"",Trainer:s.trainer||"",Peserta:Te(s.participants),Nilai:s.score!=null?s.score:"",Dokumen:s.document_link||"",Catatan:s.notes||""}));g(A,`Data_Training_${new Date().toISOString().slice(0,10)}`),H("Data training berhasil diekspor ke Excel!")}catch(g){N("Gagal ekspor Excel: "+g.message)}}async function C(g){if(!$)return;let A=t.find(s=>s.id===g);re({title:"Hapus Data Training",message:`Apakah Anda yakin ingin menghapus data sesi training "${A?.subject||"ini"}" untuk ${Te(A?.participants)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let s=await q(`/api/training/${g}`,{method:"DELETE"});s.ok?(H("Data training berhasil dihapus"),t=t.filter(l=>l.id!==g),D()):N(s.data?.error||"Gagal menghapus data training")}catch(s){N("Gagal menghapus training: "+s.message)}}})}function j(g=null){if(!$)return;let A=!!g,s=g?.training_date?g.training_date.slice(0,10):new Date().toISOString().slice(0,10),l=g?.batch||"Batch 5",c=g?.trainer||"Berlin Ariansyah",v=Te(g?.participants)!=="-"?Te(g?.participants):"",x=document.createElement("div");x.innerHTML=`
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
              placeholder="Pilih atau ketik batch..." value="${l}" required />
            <datalist id="tr-batch-options">
              ${ba.map(w=>`<option value="${w}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Materi / Topik Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="tr-modal-subject" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Basic Knowledge & Practical Skills" value="${g?.subject?String(g.subject).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="tr-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${n.map(w=>`
                <option value="${w.value}" ${String(g?.branch_id)===String(w.value)?"selected":""}>
                  ${w.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Trainer / Instruktur <span style="color:#DC2626;">*</span>
            </label>
            <input list="tr-trainer-options" id="tr-modal-trainer" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih trainer..." value="${c}" required />
            <datalist id="tr-trainer-options">
              ${ha.map(w=>`<option value="${w}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Peserta Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="tr-modal-participants" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Ketik nama karyawan peserta training..." required>${v}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nilai / Score Evaluasi (0 - 100)
            </label>
            <input type="number" min="0" max="100" step="1" id="tr-modal-score" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: 90" value="${g?.score!=null?g.score:""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Link Dokumen / Sertifikat (Opsional)
            </label>
            <input type="url" id="tr-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="https://drive.google.com/..." value="${g?.document_link?String(g.document_link).replace(/"/g,"&quot;"):""}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan Pelatihan
          </label>
          <textarea id="tr-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Catatan hasil pelatihan, rekomendasi materi lanjutan...">${g?.notes||""}</textarea>
        </div>
      </div>
    `,Z({title:A?"\u270F\uFE0F Edit Data Sesi Training":"\u2795 Tambah Data Sesi Training",content:x,size:"md",confirmText:A?"Simpan Perubahan":"Simpan Training",confirmClass:"btn-primary",onConfirm:async(w,a)=>{let d=x.querySelector("#tr-modal-date").value,y=(x.querySelector("#tr-modal-batch").value||"").trim(),k=(x.querySelector("#tr-modal-subject").value||"").trim(),M=x.querySelector("#tr-modal-branch").value,I=(x.querySelector("#tr-modal-trainer").value||"").trim(),K=(x.querySelector("#tr-modal-participants").value||"").trim(),B=x.querySelector("#tr-modal-score").value,P=B!==""?Number(B):null,b=(x.querySelector("#tr-modal-link").value||"").trim(),h=(x.querySelector("#tr-modal-notes").value||"").trim();if(!d){N("Tanggal training wajib diisi");return}if(!k){N("Materi / topik pelatihan wajib diisi");return}if(!M){N("Cabang klinik wajib dipilih");return}if(!I){N("Trainer wajib diisi");return}if(!K){N("Nama peserta training wajib diisi");return}let S=n.find(z=>String(z.value)===String(M)),O={training_date:d,batch:y,subject:k,branch_id:Number(M),branch_name:S?S.label:"Cabang",trainer:I,participants:K,score:P,document_link:b,notes:h};try{if(A){let z=await q(`/api/training/${g.id}`,{method:"PUT",body:JSON.stringify(O)});if(z.ok){H("Data training berhasil diperbarui");let G=t.findIndex(L=>L.id===g.id);G!==-1&&(t[G]={...t[G],...O}),a(),D()}else N(z.data?.error||"Gagal memperbarui data training")}else{let z=await q("/api/training",{method:"POST",body:JSON.stringify(O)});if(z.ok){H("Sesi training baru berhasil ditambahkan");let G=z.data?.id||Date.now();t.unshift({id:G,...O}),a(),D()}else N(z.data?.error||"Gagal menambahkan training")}}catch(z){N("Terjadi kesalahan: "+z.message)}}})}function T(){return t.filter(g=>{if(u){let A=u.toLowerCase(),s=(g.subject||"").toLowerCase().includes(A),l=(g.trainer||"").toLowerCase().includes(A),c=Te(g.participants).toLowerCase().includes(A),v=(g.branch_name||"").toLowerCase().includes(A),x=(g.batch||"").toLowerCase().includes(A);if(!s&&!l&&!c&&!v&&!x)return!1}return!(i!=="all"&&g.batch!==i||m!=="all"&&String(g.branch_id)!==String(m)||f!=="all"&&g.trainer!==f)})}function D(){let g=T(),A=t.length,s=new Set;t.forEach(y=>{Te(y.participants).split(",").forEach(M=>{let I=M.trim();I&&I!=="-"&&s.add(I)})});let l=s.size,v=new Set(t.map(y=>y.batch).filter(Boolean)).size,x=new Date,w=x.getFullYear(),a=x.getMonth(),d=t.filter(y=>{if(!y.training_date)return!1;let k=new Date(y.training_date);return!isNaN(k.getTime())&&k.getFullYear()===w&&k.getMonth()===a}).length;e.innerHTML=`
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
        .training-hero-actions {
          display: flex;
          align-items: center;
          gap: 8px;
          flex-wrap: wrap;
        }
        .training-hero-actions .btn {
          height: 38px;
          box-sizing: border-box;
          transition: all 0.2s ease;
        }
        .training-btn-add:hover {
          background: #1D4ED8 !important;
          transform: translateY(-1px);
        }
        .training-btn-excel:hover, .training-btn-refresh:hover {
          background: rgba(255, 255, 255, 0.16) !important;
          border-color: rgba(255, 255, 255, 0.3) !important;
          transform: translateY(-1px);
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
            gap: 6px !important;
            flex-wrap: nowrap !important;
            overflow-x: auto !important;
            scrollbar-width: none !important;
            padding-bottom: 2px !important;
          }
          .training-hero-actions::-webkit-scrollbar {
            display: none;
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
            grid-template-columns: minmax(0, 1fr) !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .training-card {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
            overflow: hidden;
            padding: 14px !important;
          }
        }
      </style>

      <div class="training-page">

        <!-- Hero Header Card -->
        <div class="training-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">\u{1F393}</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Training Karyawan
                </h1>
              </div>
            </div>

            <div class="training-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${$?`
                <button id="tr-btn-add" class="btn btn-primary training-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Training</span>
                </button>
              `:""}
              
              ${X(p,"export")?`<button id="tr-btn-excel" class="btn btn-ghost training-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="tr-btn-refresh" class="btn btn-ghost training-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
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
                <div class="training-stat-val">${A}</div>
                <div class="training-stat-lbl">Sesi Training</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F465}
              </div>
              <div>
                <div class="training-stat-val">${l}</div>
                <div class="training-stat-lbl">Peserta Terlatih</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4DA}
              </div>
              <div>
                <div class="training-stat-val">${v} Batch</div>
                <div class="training-stat-lbl">Total Angkatan</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="training-stat-val">${d}</div>
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
              <option value="all">Semua Cabang (${n.length})</option>
              ${n.map(y=>`<option value="${y.value}" ${String(m)===String(y.value)?"selected":""}>${y.label}</option>`).join("")}
            </select>

            <select id="tr-filter-trainer" class="training-select">
              <option value="all">Semua Trainer</option>
              ${ha.map(y=>`<option value="${y}" ${f===y?"selected":""}>${y}</option>`).join("")}
            </select>
          </div>

          <!-- Batch Filter Pills (Swipeable on Mobile) -->
          <div class="training-filter-pills">
            <div class="training-pill ${i==="all"?"active":""}" data-batch="all">
              <span>Semua Batch</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${A}</span>
            </div>
            ${ba.map(y=>{let k=t.filter(M=>M.batch===y).length;return`
                <div class="training-pill ${i===y?"active":""}" data-batch="${y}">
                  <span>\u{1F393}</span>
                  <span>${y}</span>
                  ${k>0?`<span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${k}</span>`:""}
                </div>
              `}).join("")}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="training-list-container">
          ${g.length===0?`
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
              ${g.map(y=>{let k=Te(y.participants),I=k.split(",")[0].trim().split(" ").map(K=>K[0]).filter(Boolean).slice(0,2).join("").toUpperCase()||"TR";return`
                  <div class="training-card" data-id="${y.id}">
                    <div>
                      <!-- Card Header: Batch & Nilai/Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="training-batch-badge">
                          <span>\u{1F393}</span>
                          <span>${y.batch||"Reguler"}</span>
                        </span>
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${y.score!=null?`
                            <span class="training-score-badge">
                              <span>\u2B50</span>
                              <span>${y.score}</span>
                            </span>
                          `:""}
                          <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                            \u{1F5D3}\uFE0F ${Kn(y.training_date)}
                          </span>
                        </div>
                      </div>

                      <!-- Subject / Materi Title -->
                      <h4 style="margin:0 0 10px 0; font-size:1.05rem; font-weight:700; color:#0F172A; line-height:1.4;">
                        ${y.subject||"Materi Pelatihan"}
                      </h4>

                      <!-- Participant & Branch Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:10px;">
                        <div class="training-avatar-box">
                          ${I}
                        </div>
                        <div style="min-width:0;">
                          <div style="font-size:0.72rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:1px;">Peserta Training</div>
                          <div style="font-size:0.95rem; font-weight:700; color:#1E293B; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${k}">
                            ${k}
                          </div>
                        </div>
                      </div>

                      <!-- Detail Box: Cabang & Trainer -->
                      <div class="training-info-box">
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#334155;">
                          <span>\u{1F4CD}</span>
                          <span style="font-weight:600;">${y.branch_name||"Cabang Belum Ditentukan"}</span>
                        </div>
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#64748B;">
                          <span>\u{1F468}\u200D\u{1F3EB}</span>
                          <span>Trainer: <strong style="color:#0F172A;">${y.trainer||"-"}</strong></span>
                        </div>
                      </div>

                      ${y.notes?`
                        <div style="font-size:0.78rem; color:#64748B; font-style:italic; line-height:1.4; margin-bottom:8px;">
                          \u{1F4DD} ${y.notes}
                        </div>
                      `:""}
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:8px;">
                      <div>
                        ${y.document_link?`
                          <a href="${y.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>\u{1F4C4} Materi / Sertifikat</span>
                          </a>
                        `:`
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${$?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          <button class="training-icon-btn training-edit-btn" title="Edit Data Training" data-id="${y.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                          <button class="training-icon-btn delete training-delete-btn" title="Hapus Training" data-id="${y.id}">
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
    `,_()}function _(){let g=e.querySelector("#tr-btn-add");g&&(g.onclick=()=>j());let A=e.querySelector("#tr-btn-excel");A&&(A.onclick=()=>E());let s=e.querySelector("#tr-btn-refresh");s&&(s.onclick=()=>F());let l=e.querySelector("#tr-btn-reset-filter");l&&(l.onclick=()=>{u="",m="all",i="all",f="all",D()});let c=e.querySelector("#tr-search-input");c&&(c.oninput=w=>{u=w.target.value,D();let a=e.querySelector("#tr-search-input");a&&(a.focus(),a.setSelectionRange(a.value.length,a.value.length))});let v=e.querySelector("#tr-filter-branch");v&&(v.onchange=w=>{m=w.target.value,D()});let x=e.querySelector("#tr-filter-trainer");x&&(x.onchange=w=>{f=w.target.value,D()}),e.querySelectorAll(".training-pill").forEach(w=>{w.onclick=()=>{i=w.getAttribute("data-batch")||"all",D()}}),e.querySelectorAll(".training-edit-btn").forEach(w=>{w.onclick=a=>{a.stopPropagation();let d=Number(w.getAttribute("data-id")),y=t.find(k=>k.id===d);y&&j(y)}}),e.querySelectorAll(".training-delete-btn").forEach(w=>{w.onclick=a=>{a.stopPropagation();let d=Number(w.getAttribute("data-id"));C(d)}})}await F()}ge();pe();function xa({container:e,title:t,icon:n,apiPath:r,columns:i,formFields:m,filterFields:f,defaultFilters:u={},itemLabel:o="Data",canCreate:p=!0,canEdit:$=!0,canDelete:F=!0,onBeforeSubmit:E,onAfterLoad:C,onDataLoaded:j,extraActions:T=[],initialSearch:D="",exportOptions:_=null,bulkDelete:g=!1,paginationMode:A="server"}){let s=ae();s&&(s.role==="superadmin"||s.role==="admin")?X(s,"delete")||(F=!1,g=!1):(_=null,g=!1,s&&s.role==="editor_khusus"?r!=="/api/issues"&&r!=="/api/overtime"&&(p=!1,$=!1,F=!1):s&&s.role==="input_lembur"?(r==="/api/overtime"||(p=!1,$=!1),F=!1):(p=!1,$=!1,F=!1));let c=1,v={...u};D&&(v.search=D);let x=new Set;e.innerHTML=`
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
      <h1 class="page-title">${n} ${t}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${p?`<button class="btn btn-primary" id="btn-create">+ Tambah ${o}</button>`:""}
        ${_?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${_.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${_.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${_.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${_.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
    

    ${f&&f.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${f.filter(b=>b.type==="search").map(b=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${b.placeholder||"Cari..."}" id="filter-search" value="${v.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${f.filter(b=>b.type!=="search").map(b=>{if(b.type==="select"||b.type==="combobox"){let h=(b.label||"").startsWith("Pilih")?b.label:`Pilih ${b.label||""}`;return`<select class="filter-select" name="${b.name}" id="filter-${b.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${h}</option>${(b.options||[]).map(S=>`<option value="${typeof S=="object"?S.value:S}" ${v[b.name]===(typeof S=="object"?S.value:S)?"selected":""}>${typeof S=="object"?S.label:S}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${f.filter(b=>b.type!=="search").map(b=>{if(b.type==="select"||b.type==="combobox"){let h=(b.label||"").startsWith("Pilih")?b.label:`Pilih ${b.label||""}`;return`<select class="filter-select filter-select-sheet" name="${b.name}-sheet" id="filter-sheet-${b.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${h}</option>${(b.options||[]).map(S=>`<option value="${typeof S=="object"?S.value:S}" ${v[b.name]===(typeof S=="object"?S.value:S)?"selected":""}>${typeof S=="object"?S.label:S}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function w(){let b=document.getElementById("bulk-toolbar");if(!b)return;let h=document.getElementById("bulk-count"),S=document.getElementById("btn-bulk-delete"),O=document.getElementById("btn-bulk-cancel");h.textContent=`${x.size} item dipilih`,x.size>0?(b.style.display="flex",S.disabled=!1,O.disabled=!1):(b.style.display="none",S.disabled=!0,O.disabled=!0);let z=document.getElementById("select-all-checkbox");if(z){let G=document.querySelectorAll(".row-checkbox");if(G.length>0){let L=[...G].every(V=>V.checked),R=[...G].some(V=>V.checked);z.checked=L,z.indeterminate=R&&!L}else z.checked=!1,z.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{x.clear(),document.querySelectorAll(".row-checkbox").forEach(h=>h.checked=!1);let b=document.getElementById("select-all-checkbox");b&&(b.checked=!1),w()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(x.size===0)return;let b=[...x],h=document.createElement("div");h.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",h.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${b.length} ${o}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${b.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(h),h.querySelector("#bulk-cancel-btn").addEventListener("click",()=>h.remove()),h.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let S=h.querySelector("#bulk-confirm-btn");S.disabled=!0,S.textContent="Menghapus...";let O=await q(`${r}/bulk`,{method:"DELETE",body:JSON.stringify({ids:b})});h.remove(),O.ok?(H(`${b.length} ${o} berhasil dihapus.`),x.clear(),w(),xe(r),I()):N(O.data?.error||"Gagal menghapus data.")})});let a=document.getElementById("filter-search"),d;a?.addEventListener("input",b=>{clearTimeout(d),d=setTimeout(()=>{v.search=b.target.value,c=1,x.clear(),w(),I()},400)}),f?.forEach(b=>{(b.type==="select"||b.type==="combobox")&&(document.getElementById(`filter-${b.name}`)?.addEventListener("change",h=>{v[b.name]=h.target.value;let S=document.getElementById(`filter-sheet-${b.name}`);S&&(S.value=h.target.value),c=1,x.clear(),w(),I()}),document.getElementById(`filter-sheet-${b.name}`)?.addEventListener("change",h=>{v[b.name]=h.target.value;let S=document.getElementById(`filter-${b.name}`);S&&(S.value=h.target.value),c=1,x.clear(),w(),I(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{v={},a&&(a.value=""),f?.forEach(b=>{let h=document.getElementById(`filter-${b.name}`);h&&(h.value="");let S=document.getElementById(`filter-sheet-${b.name}`);S&&(S.value="")}),c=1,x.clear(),w(),I()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{v={},a&&(a.value=""),f?.forEach(b=>{let h=document.getElementById(`filter-${b.name}`);h&&(h.value="");let S=document.getElementById(`filter-sheet-${b.name}`);S&&(S.value="")}),c=1,x.clear(),w(),I(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>B(null)),_&&document.addEventListener("click",function(b){let h=document.getElementById("aksi-menu-main"),S=document.getElementById("btn-aksi-main");h&&S&&!S.contains(b.target)&&!h.contains(b.target)&&h.classList.remove("show-aksi-menu")});let y=document.getElementById("btn-mobile-filter"),k=document.getElementById("filter-options-wrapper"),M=document.getElementById("btn-close-filter-sheet");if(y&&k&&(y.addEventListener("click",b=>{b.preventDefault(),k.classList.add("sheet-open")}),M&&M.addEventListener("click",b=>{b.preventDefault(),k.classList.remove("sheet-open")})),_){document.getElementById(`btn-export-${_.moduleName}`)?.addEventListener("click",async h=>{let S=h.target,O=S.innerHTML;S.innerHTML="\u23F3 Loading...",S.disabled=!0;try{await _.onExport()}catch{N("Gagal export data")}finally{S.innerHTML=O,S.disabled=!1}}),document.getElementById(`btn-template-${_.moduleName}`)?.addEventListener("click",()=>{_.onTemplate()});let b=document.getElementById(`input-import-${_.moduleName}`);b?.addEventListener("change",async h=>{let S=h.target.files[0];if(!S)return;let O=document.getElementById(`label-import-${_.moduleName}`),z=O?O.querySelector(".import-text"):null,G=z?z.innerText:"";z&&(z.innerText="\u231B Memproses..."),O&&(O.style.pointerEvents="none"),b.disabled=!0;try{let L=await Ge(S);if(L.length===0)throw new Error("File kosong atau format salah");await _.onImport(L),H("Import berhasil!"),xe(r),I()}catch(L){N(L.message||"Gagal import data")}finally{z&&(z.innerText=G),O&&(O.style.pointerEvents="auto"),b.disabled=!1,b.value=""}})}async function I(){w();let b=document.getElementById("table-container");if(!b)return;b.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let h=A==="client",S=h?1:c,O=h?Le:20,z=new URLSearchParams({page:S,limit:O,...Object.fromEntries(Object.entries(v).filter(([,U])=>U))}),G=await q(`${r}?${z}`);if(!G.ok){b.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${G.data?.error||"Error"}</p></div>`;return}let L=G.data?.data||G.data||[],R=G.data?.pagination,V=L.length;if(h){L=j(L);let U=L.length,Y=20,ee=Math.ceil(U/Y);c>ee&&ee>0&&(c=ee);let W=(c-1)*Y,te=c*Y;L=L.slice(W,te),R={page:c,limit:Y,total:U,pages:ee}}!1,C&&C(L);let J=nt({columns:i,data:L,onEdit:$?U=>B(U):null,actions:T.map(U=>({...U,handler:Y=>U.handler(Y,I)})),emptyText:`Tidak ada ${String(o||"").toLowerCase()}`,bulkSelect:g?{selectedIds:x,onToggle:w}:null});b.innerHTML="",b.appendChild(J);let Q=document.getElementById("pagination-container");if(Q&&(Q.innerHTML="",R&&R.pages>1)){let U=it({page:R.page,pages:R.pages,total:R.total,limit:R.limit,onPage:Y=>{c=Y,I()}});U&&Q.appendChild(U)}}function K(b){let h=typeof m=="function"?m(b):m;return Ue(h)}function B(b){let h=!!b,S=document.createElement("form");if(S.noValidate=!0,S.innerHTML=K(b),h){let z=typeof m=="function"?m(b):m;ot(S,b)}let{close:O}=Z({title:h?`Edit ${o}`:`Tambah ${o}`,content:S,size:"lg",confirmText:h?"Simpan Perubahan":`Tambah ${o}`,onConfirm:async(z,G)=>{if(!S.reportValidity())return;let L=z.querySelector(".modal-confirm");L.disabled=!0,L.textContent="Menyimpan...";let R=rt(S),V=typeof m=="function"?m(b):m,J=async ee=>{for(let W of ee)if(W.type==="row")await J(W.fields);else if(W.type==="combobox"&&R[W.name]){let te=R[W.name],de=(W.options||[]).find(ie=>{let ce=String(typeof ie=="object"?ie.value:ie),Pe=String(typeof ie=="object"?ie.label:ie);return ce===te||Pe===te});if(de)R[W.name]=typeof de=="object"?de.value:de;else if(W.createApi){let ie={};ie[W.createApi.field]=te,W.createApi.extra&&Object.assign(ie,W.createApi.extra);let ce=await q(W.createApi.path,{method:"POST",body:JSON.stringify(ie)});if(ce.ok&&ce.data?.id)R[W.name]=ce.data.id;else if(ce.ok&&!ce.data?.id)R[W.name]=te;else throw new Error(`Gagal membuat master data: ${ce.data?.error||"Unknown error"}`)}}};try{await J(V)}catch(ee){N(ee.message),L.disabled=!1,L.textContent=h?"Simpan Perubahan":`Tambah ${o}`;return}E&&(R=await E(R,b));let Q=h?"PUT":"POST",U=h?`${r}/${b.id}`:r,Y=await q(U,{method:Q,body:JSON.stringify(R)});Y.ok?(H(h?`${o} berhasil diperbarui.`:`${o} berhasil ditambahkan.`),G(),xe(r),I()):(N(Y.data?.error||"Gagal menyimpan data."),L.disabled=!1,L.textContent=h?"Simpan Perubahan":`Tambah ${o}`)}})}function P(b){re(`Hapus ${o} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let h=await q(`${r}/${b.id}`,{method:"DELETE"});h.ok?(H(`${o} berhasil dihapus.`),xe(r),I()):N(h.data?.error||"Gagal menghapus.")},`Hapus ${o}`)}return I(),I}pe();async function ya(e,t){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let n=await se(),r=await be(),i=t?t.get("dash_filter"):null,m={};if(i==="reliever"){let E=new Date,C=`${E.getFullYear()}-${String(E.getMonth()+1).padStart(2,"0")}`;m={status:"Done",month:t&&t.get("month")?t.get("month"):C}}console.log("RAW",await Be()),console.log("OPTIONS",r);let f=E=>E&&!r.find(C=>C.value===E)?[...r,{value:E,label:E}]:r,u=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],o=E=>E&&!u.includes(E)?[...u,E]:u,p=new Date().getFullYear(),F=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((E,C)=>{let j=String(C+1).padStart(2,"0");return{value:`${p}-${j}`,label:`${E} ${p}`}});xa({container:e,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!1,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:m,onDataLoaded:E=>E.sort((C,j)=>{let T=C.backup_date?new Date(C.backup_date).getTime():0;return(j.backup_date?new Date(j.backup_date).getTime():0)-T}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:E=>Se(E)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:E=>window.formatDate(E)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:E=>window.formatDate(E)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:E=>E?`<span class="badge badge-info">${E}</span>`:"-"},{key:"status",label:"Status",render:E=>he(E)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:u},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"month",label:"Bulan",options:F},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:E=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:n,value:E?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:E?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:f(E?.original_fc_name),value:E?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:o(E?.reliever_name),value:E?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:E?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:E?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:E?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:E?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:E?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let E=await q(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(E.ok){let C=E.data.data.map(j=>({Cabang:j.branch_name||"","Nama Facility care":j.original_fc_name||"",Periode:j.period||"",Relifer:j.reliever_name||"","Tanggal Back Up":j.backup_date||"","Tanggal Selesai":j.completion_date||"",Keterangan:j.reason||"",Shift:j.shift||"",Status:j.status||""}));C.length===0&&C.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),ne(C,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async E=>{let j=(await q("/api/branches?all=1")).data?.data||[],T=g=>{if(!g)return null;let A=String(g||"").toLowerCase(),s=j.find(l=>String(l.full_name||"").toLowerCase()===A||String(l.code||"").toLowerCase()===A||String(l.name||"").toLowerCase()===A);return s?s.id:null},D=E.map(g=>({branch_name:String(g.Cabang||"").trim(),backup_date:String(g["Tanggal Back Up"]||g["Tanggal Backup"]||"").trim(),original_fc_name:String(g["Nama Facility care"]||g["FC Digantikan"]||"").trim(),reliever_name:String(g.Relifer||g.Reliefer||"").trim(),period:String(g.Periode||"").trim(),reason:String(g.Keterangan||"").trim(),shift:String(g.Shift||"").trim(),completion_date:String(g["Tanggal Selesai"]||"").trim(),status:String(g.Status||"").trim()})).filter(g=>g.reliever_name&&g.backup_date),_=await q("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:D,onDuplicate:"update"})});if(!_.ok)throw new Error(_.data?.error||"Import gagal");return _.data}}})}pe();async function va(e){let t=await se(),n=Array.from({length:4},(r,i)=>String(new Date().getFullYear()-i));me({container:e,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:r=>Se(r)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"fc_score",label:"Point FC",render:r=>r!=null?`<strong class="${r>=80?"text-success":r>=60?"text-warning":"text-danger"}">${r}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:r=>r!=null?`<strong>${r}</strong>`:"-"},{key:"status",label:"Status",render:r=>he(r)},{key:"document_link",label:"Dokumen",render:r=>r?`<a href="${r}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:r=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:r?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:r?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:r?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:r?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:r?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:r?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:r?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async r=>{let i=new URLSearchParams(r||{}).toString(),m=await q(`/api/reports/inspection?limit=10000&${i}`);if(m.ok){let f=m.data.data.map(u=>({Cabang:u.branch_name||"",Periode:u.period||"",Tanggal:u.inspection_date||"","Point FC":u.fc_score!==null&&u.fc_score!==void 0?u.fc_score:"","Point SPV":u.spv_score!==null&&u.spv_score!==void 0?u.spv_score:"",Status:u.status||"","Link Dokumen":u.document_link||""}));ne(f,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async r=>{let i=o=>{if(!o)return null;let p=String(o||"").toLowerCase(),$=t.find(F=>String(F.label||"").toLowerCase()===p);return $?$.value:null},m=o=>{if(o==null||o==="")return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let p=String(o).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let F=Number(p);if(F>2e4&&F<99999){let E=new Date(Date.UTC(1899,11,30)+F*864e5);return isNaN(E.getTime())?"":E.toISOString().slice(0,10)}}let $=p.split(/[\/\-\.]/);if($.length===3){let[F,E,C]=$.map(j=>j.trim());if(F.length===4&&E.length<=2&&C.length<=2)return`${F}-${E.padStart(2,"0")}-${C.padStart(2,"0")}`;if(C.length===4&&E.length<=2&&F.length<=2)return`${C}-${E.padStart(2,"0")}-${F.padStart(2,"0")}`}return p},f=r.map(o=>({branch_id:i(String(o.Cabang||"").trim()),period:String(o.Periode||"").trim(),inspection_date:m(o.Tanggal),fc_score:o["Point FC"]!==void 0&&o["Point FC"]!==""?Number(o["Point FC"]):null,spv_score:o["Point SPV"]!==void 0&&o["Point SPV"]!==""?Number(o["Point SPV"]):null,status:String(o.Status||"").trim(),document_link:String(o["Link Dokumen"]||"").trim(),notes:String(o.Catatan||o.Keterangan||"").trim()})).filter(o=>o.branch_id&&o.period&&o.inspection_date),u=await q("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:f,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}pe();async function ka(e){let t=await se(),n=Array.from({length:4},(r,i)=>String(new Date().getFullYear()-i));me({container:e,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:r=>`<span class="badge ${r==="Deep Cleaning"?"badge-purple":"badge-success"}">${r}</span>`},{key:"period",label:"Periode",render:r=>Se(r)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"status",label:"Status",render:r=>he(r)},{key:"document_link",label:"Dokumen",render:r=>r?`<a href="${r}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:r=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t,value:r?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:r?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:r?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:r?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:r?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:r?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async r=>{let i=new URLSearchParams(r||{}).toString(),m=await q(`/api/reports/cleaning?limit=10000&${i}`);if(m.ok){let f=m.data.data.map(u=>({Cabang:u.branch_name||"",Jenis:u.activity_type||"",Periode:u.period||"",Tanggal:u.activity_date||"",Status:u.status||"","Link Dokumen":u.document_link||""}));ne(f,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async r=>{let i=o=>{if(!o)return null;let p=String(o||"").toLowerCase(),$=t.find(F=>String(F.label||"").toLowerCase()===p);return $?$.value:null},m=o=>{if(o==null||o==="")return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let p=String(o).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let F=Number(p);if(F>2e4&&F<99999){let E=new Date(Date.UTC(1899,11,30)+F*864e5);return isNaN(E.getTime())?"":E.toISOString().slice(0,10)}}let $=p.split(/[\/\-\.]/);if($.length===3){let[F,E,C]=$.map(j=>j.trim());if(F.length===4&&E.length<=2&&C.length<=2)return`${F}-${E.padStart(2,"0")}-${C.padStart(2,"0")}`;if(C.length===4&&E.length<=2&&F.length<=2)return`${C}-${E.padStart(2,"0")}-${F.padStart(2,"0")}`}return p},f=r.map(o=>({branch_id:i(String(o.Cabang||"").trim()),activity_type:String(o.Jenis||o.Kegiatan||"").trim(),period:String(o.Periode||"").trim(),activity_date:m(o.Tanggal),status:String(o.Status||"").trim(),document_link:String(o["Link Dokumen"]||"").trim(),notes:String(o.Catatan||o.Keterangan||"").trim()})).filter(o=>o.branch_id&&o.activity_type&&o.period&&o.activity_date),u=await q("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:f,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}pe();async function wa(e,t){let n=await se(),r=Array.from({length:4},(f,u)=>String(new Date().getFullYear()-u)),i=t?t.get("dash_filter"):null,m={};if(i==="fogging"){let f=new Date,u=String(f.getMonth()+1).padStart(2,"0"),o=String(f.getFullYear()),p=t?t.get("month"):null;p&&p.length===7&&(o=p.split("-")[0],u=p.split("-")[1]),m={status:"Done",month:u,year:o}}me({container:e,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:m,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:f=>`<span class="badge badge-warning">${f}</span>`},{key:"period",label:"Periode",render:f=>Se(f)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:f=>window.formatDate(f)},{key:"status",label:"Status",render:f=>he(f)},{key:"document_link",label:"Dokumen",render:f=>f?`<a href="${f}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:f=>f||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:f=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:n,value:f?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:f?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:f?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:f?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:f?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:f?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async f=>{let u=new URLSearchParams(f||{}).toString(),o=await q(`/api/reports/fogging?limit=10000&${u}`);if(o.ok){let p=o.data.data.map($=>({Cabang:$.branch_name||"",Jenis:$.activity_type||"Fogging",Periode:$.period||"",Tanggal:$.activity_date||"",Status:$.status||"","Link Dokumen":$.document_link||""}));ne(p,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async f=>{let u=F=>{if(!F)return null;let E=String(F||"").toLowerCase(),C=n.find(j=>String(j.label||"").toLowerCase()===E);return C?C.value:null},o=F=>{if(F==null||F==="")return"";if(F instanceof Date&&!isNaN(F.getTime()))return F.toISOString().slice(0,10);let E=String(F).trim();if(E===""||E==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(E))return E.slice(0,10);if(/^\d{4,5}$/.test(E)){let j=Number(E);if(j>2e4&&j<99999){let T=new Date(Date.UTC(1899,11,30)+j*864e5);return isNaN(T.getTime())?"":T.toISOString().slice(0,10)}}let C=E.split(/[\/\-\.]/);if(C.length===3){let[j,T,D]=C.map(_=>_.trim());if(j.length===4&&T.length<=2&&D.length<=2)return`${j}-${T.padStart(2,"0")}-${D.padStart(2,"0")}`;if(D.length===4&&T.length<=2&&j.length<=2)return`${D}-${T.padStart(2,"0")}-${j.padStart(2,"0")}`}return E},p=f.map(F=>({branch_id:u(String(F.Cabang||"").trim()),activity_type:String(F.Jenis||F.Kegiatan||"Fogging").trim(),period:String(F.Periode||"").trim(),activity_date:o(F.Tanggal),status:String(F.Status||"").trim(),document_link:String(F["Link Dokumen"]||"").trim(),notes:String(F.Catatan||F.Keterangan||"").trim()})).filter(F=>F.branch_id&&F.period&&F.activity_date),$=await q("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(p)});if(!$.ok)throw new Error($.data?.error||"Import gagal");return $.data}}})}pe();async function Ea(e){let t=await se(),n=await be(),r=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],i=Array.from({length:4},(u,o)=>String(new Date().getFullYear()-o)),m=u=>u&&!n.find(o=>o.value===u)?[...n,{value:u,label:u}]:n,f=u=>u&&!r.find(o=>o.value===u)?[...r,{value:u,label:u}]:r;me({container:e,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:u=>window.formatDate(u)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:u=>`<span title="${u||""}">${u?.length>60?u.slice(0,60)+"\u2026":u||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:u=>window.formatDate(u)},{key:"status",label:"Status",render:u=>he(u)},{key:"notes",label:"Keterangan",render:u=>u?.length>40?u.slice(0,40)+"\u2026":u||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:u=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:u?.branch_id},{name:"pic",label:"PIC",type:"select",options:f(u?.pic),value:u?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:u?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:u?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:u?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:u?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:u?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async u=>{let o=new URLSearchParams(u||{}).toString(),p=await q(`/api/reports/basecamp?limit=10000&${o}`);if(p.ok){let $=p.data.data.map(F=>({"Tgl Info":F.info_date||"",Cabang:F.branch_name||"",Permasalahan:F.problem||"",PIC:F.pic||"","Tgl Done":F.done_date||"",Status:F.status||"",Keterangan:F.notes||""}));ne($,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async u=>{let o=E=>{if(!E)return null;let C=String(E||"").toLowerCase(),j=t.find(T=>String(T.label||"").toLowerCase()===C);return j?j.value:null},p=E=>{if(E==null||E==="")return"";if(E instanceof Date&&!isNaN(E.getTime()))return E.toISOString().slice(0,10);let C=String(E).trim();if(C===""||C==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(C))return C.slice(0,10);if(/^\d{4,5}$/.test(C)){let T=Number(C);if(T>2e4&&T<99999){let D=new Date(Date.UTC(1899,11,30)+T*864e5);return isNaN(D.getTime())?"":D.toISOString().slice(0,10)}}let j=C.split(/[\/\-\.]/);if(j.length===3){let[T,D,_]=j.map(g=>g.trim());if(T.length===4&&D.length<=2&&_.length<=2)return`${T}-${D.padStart(2,"0")}-${_.padStart(2,"0")}`;if(_.length===4&&D.length<=2&&T.length<=2)return`${_}-${D.padStart(2,"0")}-${T.padStart(2,"0")}`}return C},$=u.map(E=>({info_date:p(E["Tgl Info"]||E["Tanggal Info"]),branch_id:o(String(E.Cabang||"").trim()),problem:String(E.Permasalahan||"").trim(),pic:String(E.PIC||"").trim(),done_date:p(E["Tgl Done"]||E["Tanggal Done"]),status:String(E.Status||"").trim(),notes:String(E.Keterangan||E.Catatan||"").trim()})).filter(E=>E.info_date&&E.branch_id&&E.problem),F=await q("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify($)});if(!F.ok)throw new Error(F.data?.error||"Import gagal");return F.data}}})}ge();var Nn=[{id:1,title:"Dusting (Pembersihan Debu)",category:"Pembersihan Rutin",icon:"\u{1F9F9}",color:"#2563EB",bg:"#EFF6FF",target:"Meja resepsionis, meja pasien, meja dokter, furniture, kaca, monitor & keyboard",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Kain Lap Microfiber Biru (Area Umum)","Kanebo (Kaca & Layar Monitor)","Kemoceng","Botol Spray Glass Cleaner","Kuas 4 Inch (Keyboard & Sela Sempit)"],steps:["Siapkan larutan glass cleaner di botol spray khusus kaca dan perabot.","Lipat kain microfiber menjadi 8 bidang kerja agar efisien dan bersih merata.","Semprotkan cairan pembersih ke permukaan kain lap, bukan langsung ke monitor atau bodi elektronik.","Usap permukaan perabot dari bagian paling atas ke bagian bawah secara searah atau pola S.","Gunakan kuas cat 4 inch untuk membersihkan sela-sela keyboard dan lubang ventilasi.","Keringkan kaca atau monitor menggunakan lap microfiber kering atau kanebo bersih."],notes:"Khusus layar monitor medis & dental unit, dilarang menyemprotkan cairan langsung ke layar."},{id:2,title:"Sweeping (Lobby Duster & Sapu Nylon)",category:"Pembersihan Rutin",icon:"\u{1F33E}",color:"#D97706",bg:"#FFFBEB",target:"Lobby utama, lorong, ruang tindakan, balkon, pantry, teras",chemical_ratio:"-",tools:["Lobby Duster Set 60cm (Area Lobby & Lorong)","Lobby Duster Set 40cm (Area Ruang Tindakan)","Sapu & Pengki Nagata 742s","Kapi Scrapper (Kerak/Permen Karet)"],steps:["Pastikan kain lobby duster dalam keadaan kering, bersih, dan bebas pasir kasar.","Dorong lobby duster menyusuri tepi dinding ruangan terlebih dahulu membentuk perimeter.","Lanjutkan mendorong dari sudut terjauh menuju arah pintu keluar tanpa mengangkat kain pel lobi dari lantai.","Kumpulkan debu dan kotoran di dekat pintu keluar, lalu angkat menggunakan sapu dan pengki Nagata.","Jika ada noda kerak atau bekas stiker/permen karet menempel di lantai, kikis perlahan menggunakan kapi scrapper.","Kibaskan kain duster di luar ruangan atau sedot sisa debu pada kain duster."],notes:"Jangan pernah menggunakan lobby duster di area lantai yang masih basah atau tergenang air."},{id:3,title:"Mopping (Pengepelan Lantai)",category:"Pembersihan Rutin",icon:"\u{1FAA3}",color:"#059669",bg:"#ECFDF5",target:"Seluruh lantai ruangan, koridor, dan area tunggu klinik",chemical_ratio:"25 ml Floor Cleaner per 1 Liter air bersih",tools:["Single Bucket Mop Wringer 20L / Double Bucket 46L","Mopping Stick / Tongkat & Kain Pel","Chemical Floor Cleaner Prokleen Lavender","Warning Sign Lantai Licin"],steps:['Pasang papan peringatan "Caution Wet Floor" di titik strategis yang mudah terlihat.',"Campurkan cairan Floor Cleaner ke dalam ember air bersih dengan takaran 25 ml per 1 Liter air.","Celupkan kain pel, masukkan ke dalam alat pemeras (wringer), dan peras hingga kondisi kain pel lembap (tidak menetes).","Mulai mengepel dari sudut terjauh ruangan mundur ke arah pintu keluar dengan gerakan membentuk pola angka delapan (8).","Bilas kain pel secara berkala apabila air pada ember bilas sudah terlihat keruh.","Biarkan lantai mengering sempurna secara alami sebelum papan peringatan wet floor dilepas."],notes:'Wajib pasang papan peringatan "Caution Wet Floor" sebelum mulai mengepel untuk mencegah risiko terpeleset.'},{id:4,title:"Membersihkan Toilet & Wastafel",category:"Sanitasi Toilet",icon:"\u{1F6BD}",color:"#DC2626",bg:"#FEF2F2",target:"Toilet pasien, toilet tim, closet, wastafel, dinding keramik, floor drain",chemical_ratio:"Karbol 25 ml:1 L Air / Bowl Cleaner 100 ml spray (kerak) / Bubuk KIFA",tools:["Sarung Tangan Karet (Wajib)","Sikat Kloset Scotch-Brite 3M","Kain Lap Microfiber Merah (Khusus Toilet)","Sponge Scotch Brite & Sikat Lantai ID-51","Floor Squeegee 75cm","Chemical Karbol Pine, Bowl Cleaner, Bubuk KIFA"],steps:["Kenakan sarung tangan karet dan siapkan perlengkapan sanitasi toilet.","Flush kloset, tuangkan Toilet Bowl Cleaner ke bagian dalam dinding kloset dan leher angsa, diamkan 5\u201310 menit.","Sikat bagian dalam kloset secara menyeluruh dengan sikat kloset, lalu bilas (flush) kembali hingga bersih.","Seka dudukan, penutup, dan tabung luar kloset menggunakan kain lap microfiber merah yang dibasahi larutan disinfektan karbol.","Bersihkan wastafel dan kran stainless menggunakan sponge tapas dengan bubuk pembersih KIFA (jangan gunakan Bowl Cleaner pada kran stainless).","Sikat lantai kamar mandi dengan sikat lantai ID-51 dan cairan karbol disinfektan.","Tarik genangan air ke arah floor drain menggunakan floor squeegee hingga lantai tidak licin dan cepat kering.","Pasang kamper/urinal block, isi ulang sabun cuci tangan, dan pasang plastik sampah baru."],notes:"Kain lap warna MERAH mutlak hanya untuk toilet dan mangkuk spittoon. Dilarang mencampur cairan kimia sembarangan."},{id:5,title:"Pembersihan Area Parkir & Luar",category:"Pembersihan Rutin",icon:"\u{1F17F}\uFE0F",color:"#4F46E5",bg:"#EEF2FF",target:"Halaman parkir mobil & motor, saluran air luar, teras depan",chemical_ratio:"-",tools:["Sapu Lidi Gagang Panjang","Pengki Nagata","Plastik Sampah Hitam Besar (90x120)","Selang Air"],steps:["Sapu daun kering, puntung rokok, plastik, dan pasir dari batas trotoar/pagar menuju satu titik.","Kumpulkan dan angkat sampah ke dalam pengki, masukkan langsung ke plastik sampah hitam besar.","Periksa dan bersihkan saluran drainase air di area parkir dari sumbatan sampah padat.","Bila ada tumpahan oli atau kotoran tanah membandel, siram dengan air dan sikat.","Rapikan pembatas parkir dan pastikan area depan klinik terlihat bersih dan rapi menyambut pasien."],notes:"Lakukan sebelum jam operasional klinik dibuka (pukul 07.30 - 08.00)."},{id:6,title:"Pembersihan Kaca & Jendela",category:"Pembersihan Rutin",icon:"\u{1FA9F}",color:"#0891B2",bg:"#ECFEFF",target:"Pintu kaca utama, jendela kaca klinik, partisi kaca, cermin toilet",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Botol Spray Glass Cleaner Prokleen","Window Wiper / Window Squeegee","Kain Microfiber Kering / Kanebo","Tongkat Teleskopik (Kaca Tinggi)"],steps:["Semprotkan larutan glass cleaner merata pada permukaan kaca yang hendak dibersihkan.","Tarik wiper karet dari arah atas ke bawah secara lurus vertikal atau zigzag horizontal tanpa terputus.","Lap bilah karet wiper setiap selesai satu tarikan menggunakan kain lap bersih agar tidak meninggalkan garis bekas air.","Gunakan kain microfiber kering untuk membersihkan sisa tetesan air di sudut-sudut lis bingkai kaca.","Untuk kaca bagian luar yang tinggi, pasang wiper pada tongkat sambungan teleskopik."],notes:"Hindari membersihkan kaca saat terpapar terik sinar matahari langsung karena cairan akan cepat mengering dan membekas noda."},{id:7,title:"Mencuci Bola Mandi Bola (Playground)",category:"Pembersihan Khusus",icon:"\u{1F939}",color:"#7C3AED",bg:"#F5F3FF",target:"Area bermain anak (Kids Playground) dan bola-bola mandi bola",chemical_ratio:"30 ml Dishwashing Liquid per 1 Liter Air + Karbol Disinfektan",tools:["Baskom Plastik Besar 55cm","Dishwashing Liquid Prokleen","Karbol Disinfektan Pine","Lap Microfiber Bersih","Karung / Wadah Pengering Jaring"],steps:["Kumpulkan seluruh bola mandi bola ke dalam wadah atau karung bersih.","Siapkan baskom plastik besar berisi air hangat kuku dicampur Dishwashing Liquid dan sedikit karbol disinfektan.","Rendam bola dan cuci secara bertahap sambil digosok perlahan menggunakan sponge/lap bersih.","Bilas bola dengan air bersih mengalir hingga tidak ada sisa busa sabun yang menempel.","Tiriskan dan angin-anginkan bola di tempat bersih beralaskan terpal/kain kering hingga benar-benar kering.","Seka dan sterilkan area kolam mandi bola sebelum bola dimasukkan kembali."],notes:"Jadwal pencucian berkala dilakukan minimal 1 kali seminggu atau sesuai kebutuhan kebersihan arena anak."},{id:8,title:"Membersihkan Filter Air Purifier",category:"Elektronik & Utility",icon:"\u{1F4A8}",color:"#0D9488",bg:"#F0FDFA",target:"Unit Air Purifier di ruang tunggu, ruang dokter, dan koridor",chemical_ratio:"-",tools:["Vacuum Cleaner (Nozzle Sikat Halus)","Kain Microfiber Kering","Kuas Cat Halus"],steps:["Matikan saklar power unit air purifier dan cabut kabel daya dari stop kontak dinding.","Buka penutup bodi (casing) depan atau belakang air purifier secara hati-hati.","Keluarkan pre-filter jaring nilon dari unit.","Bersihkan debu tebal pada pre-filter menggunakan vacuum cleaner nozzle sikat halus atau cuci dengan air mengalir (jika tipe wash-able).","Untuk filter HEPA dan Karbon Aktif: HANYA sedot debu permukaannya secara perlahan dengan vacuum, JANGAN DICUCI DENGAN AIR.","Lap bodi luar dan kisi-kisi udara masuk menggunakan lap microfiber kering.","Pastikan pre-filter benar-benar kering sebelum dipasang kembali ke posisinya semula."],notes:"Jangan pernah mencuci filter HEPA dengan air karena akan merusak serat mikro penyaring bakteri/virus."},{id:9,title:"Membersihkan Filter & Bodi AC",category:"Elektronik & Utility",icon:"\u2744\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Filter udara dan penutup indoor unit AC split di seluruh ruangan",chemical_ratio:"-",tools:["Tangga Besi Lipat","Kain Lap Microfiber Lembap","Kuas Cat Halus","Kran Air Mengalir"],steps:["Matikan AC melalui remote dan matikan MCB/saklar AC bila diperlukan.","Gunakan tangga besi lipat yang stabil di lantai datar.","Buka tutup cover depan unit indoor AC ke atas hingga terkunci.","Lepaskan kedua kisi filter saringan udara plastik secara perlahan agar debu tidak berjatuhan.","Cuci saringan filter di bawah kran air mengalir dari arah bagian dalam ke bagian luar.","Keringkan filter di tempat teduh (jangan dijemur langsung di bawah terik matahari).","Lap bodi casing dan louver/swing AC dengan kain microfiber lembap.","Pasang kembali filter yang sudah kering, tutup cover AC, dan nyalakan AC untuk pengetesan."],notes:"Pembersihan rutin mandiri dilakukan 2 minggu sekali. Deep cleaning / cuci evaporator dilakukan oleh teknisi AC vendor."},{id:10,title:"Membersihkan Dental Unit",category:"Klinis & Medis",icon:"\u{1F9B7}",color:"#E11D48",bg:"#FFF1F2",target:"Kursi dental pasien, spittoon bowl, meja instrumen, lampu dental, cover bawah",chemical_ratio:"Glass Cleaner (netral) / Karbol Disinfektan (mangkuk) / Bowl Cleaner (kerak)",tools:["Sarung Tangan Karet (Wajib)","Kain Lap Microfiber Biru (Khusus Bodi Luar)","Kain Lap Microfiber Merah (Khusus Mangkuk Spittoon)","Botol Spray Glass Cleaner (25ml + Air)","Botol Spray Karbol (20ml + Air)","Botol Spray Bowl Cleaner (100ml + Air, bila berkerak)"],steps:["Pastikan saklar utama dental unit dalam posisi OFF dan kursi dinaikkan ke posisi tertinggi untuk memudahkan akses.","Lap seluruh jok kulit, sandaran kepala, sandaran tangan, dan tiang lampu menggunakan lap microfiber biru dan cairan pembersih netral/glass cleaner.","Bersihkan mangkuk kumur (spittoon bowl) menggunakan kain lap merah dan cairan karbol disinfektan.","Jika ditemukan kerak sisa darah/plak membandel di mangkuk spittoon, gunakan sedikit bowl cleaner, sikat perlahan, lalu bilas hingga bersih.","Lap permukaan meja instrumen meja dokter. Bila ada instrumen steril, geser dengan hati-hati dan jangan mengotori instrumen.","Buka penutup cover bawah unit dental, lap debu yang terjangkau dan pastikan kondisi area selang dalam keadaan kering."],notes:"Kain lap warna MERAH hanya untuk mangkuk spittoon. Dilarang keras menggosok jok kulit dental unit dengan bahan kimia asam keras."},{id:11,title:"Membersihkan Wastafel & Keran Air",category:"Sanitasi Toilet",icon:"\u{1F6B0}",color:"#0284C7",bg:"#F0F9FF",target:"Wastafel keramik ruang tindakan, wastafel lorong, kran stainless steel, afur sink",chemical_ratio:"Bubuk Pembersih KIFA / Dishwashing Liquid Prokleen",tools:["Sponge Scotch Brite Hijau/Kuning","Kain Microfiber Hijau (Area Basah)","Bubuk Pembersih KIFA","Sikat Gigi Bekas / Sikat Kecil"],steps:["Basahi seluruh mangkuk wastafel keramik dan kran air stainless steel.","Taburkan sedikit Bubuk KIFA pada tapas sponge yang basah.","Gosok perlahan noda sabun dan kerak air putih pada kran stainless steel searah guratan logam hingga mengkilap kembali.","Gunakan sikat kecil untuk membersihkan sela leher kran dan sudut sambungan afur pembuangan wastafel.","Bilas mangkuk wastafel dan kran dengan air bersih mengalir hingga seluruh sisa residu bubuk KIFA hilang.","Keringkan kran air stainless steel dan cermin wastafel menggunakan lap microfiber hijau kering."],notes:"DILARANG KERAS menggunakan Toilet Bowl Cleaner pada kran air stainless steel karena menyebabkan karat dan warna menghitam permanen."},{id:12,title:"Membersihkan Kulkas Klinik & Pantry",category:"Elektronik & Utility",icon:"\u{1F9CA}",color:"#475569",bg:"#F8FAFC",target:"Kulkas pantry staf dan kulkas penyimpan obat/material medis (eksterior & rak)",chemical_ratio:"30 ml Dishwashing Liquid : 1 Liter Air Hangat + Lap Kering",tools:["Kain Microfiber Biru & Hijau","Dishwashing Liquid","Ember Air Hangat","Sponge Scotch Brite Halus"],steps:["Keluarkan makanan/minuman yang telah kedaluwarsa atau berbau tidak sedap.","Lepas rak-rak kaca/plastik interior kulkas yang dapat dilepas, cuci dengan sabun cuci piring dan bilas air bersih.","Seka dinding interior kulkas menggunakan kain microfiber yang dicelupkan ke dalam air hangat bersabun lembut.","Lap kering seluruh permukaan dalam kulkas dengan microfiber kering bersih.","Pasang kembali rak interior dan susun kembali bahan/makanan secara tertib dan rapi.","Bersihkan karet pintu kulkas dari noda hitam/jamur menggunakan sikat halus dan lap kering.","Lap bodi luar kulkas dan pegangan pintu menggunakan cairan pembersih netral."],notes:"Untuk kulkas penyimpanan obat/material dental, koordinasikan terlebih dahulu dengan Perawat/PIC Obat sebelum membersihkan rak interior."},{id:13,title:"Membersihkan Sofa & Kursi Ruang Tunggu",category:"Pembersihan Khusus",icon:"\u{1F6CB}\uFE0F",color:"#9333EA",bg:"#FAF5FF",target:"Sofa ruang tunggu pasien, kursi reception, kursi kerja tim",chemical_ratio:"Baking Soda (penghilang bau) / Pembersih Upholstery secukupnya",tools:["Vacuum Cleaner (Nozzle Kecil & Sikat Lembut)","Kain Microfiber Bersih","Sikat Bulu Halus","Ember Air Hangat","Cairan Pembersih Khusus Sofa (Upholstery Cleaner)","Baking Soda (Penghilang Bau Tidak Sedap)","Sarung Tangan Karet"],steps:["Kenali bahan sofa (kain fabric, beludru, kulit asli, atau kulit sintetis/oscar).","Sedot debu di seluruh permukaan, lipatan jahitan, dan celah bantal sofa menggunakan vacuum cleaner nozzle kecil.","Gunakan sikat bulu halus untuk mengangkat rambut atau kotoran yang menempel kuat di serat kain.","Jika ada noda bercak (kopi, makanan, tinta), semprotkan pembersih upholstery secukupnya, gosok memutar perlahan dengan microfiber.","Celupkan kain microfiber ke air hangat suam-suam kuku, peras hingga lembap, lalu seka seluruh permukaan sofa.","Lap kembali dengan kain kering bersih untuk menyerap sisa kelembapan.","Jika sofa berbau apek: taburkan baking soda tipis di permukaan, diamkan 20 menit, lalu sedot kembali dengan vacuum cleaner.","Biarkan sofa mengering secara alami di dalam ruangan berventilasi atau ber-AC."],notes:"Jangan menjemur sofa langsung di bawah terik matahari karena dapat menyebabkan bahan kulit pecah-pecah dan warna kain memudar."}],qn=[{no:1,name:"Sapu Lidi Gagang",qty:"1 unit",area:"Area parkir mobil & motor klinik"},{no:2,name:"Sapu & Pengki Nagata 742s",qty:"1 set",area:"Area pantry, balkon, dak lantai atas, atau teras"},{no:3,name:"Lobby Duster Set 60 cm",qty:"1 set",area:"Lobby utama dan koridor/lorong klinik"},{no:4,name:"Lobby Duster Set 40 cm / Refill",qty:"1 unit",area:"Area ruang tindakan dokter"},{no:5,name:"Single Bucket Mop Wringer 20L",qty:"1 paket",area:"Ember pel dengan pemeras + tongkat pel"},{no:6,name:"Double Bucket Wringer Trolley 46L",qty:"1 paket",area:"Khusus klinik besar / klinik dengan fasilitas lift"},{no:7,name:"Kemoceng Bulu Lembut",qty:"1 unit",area:"Membersihkan debu kasar meja dan perabot"},{no:8,name:"Lap Microfiber Warna Merah",qty:"2 pcs",area:"Khusus area toilet, sanitasi, dan mangkuk spittoon"},{no:9,name:"Lap Microfiber Warna Biru",qty:"2 pcs",area:"Area umum kering (meja, kursi, bodi peralatan)"},{no:10,name:"Lap Microfiber Warna Hijau",qty:"2 pcs",area:"Area umum basah (pantry, wastafel luar)"},{no:11,name:"Kanebo / Synthetic Chamois",qty:"1 pcs",area:"Area umum, cermin, dan pengering kaca"},{no:12,name:"Caddy Bag Cleaning",qty:"1 unit",area:"Tas perlengkapan pinggang petugas cleaning"},{no:13,name:"Sponge Scotch Brite 3M",qty:"2 pcs",area:"Pembersihan noda lantai dan kerak ringan (spotting)"},{no:14,name:"Scotch-Brite 3M Sikat Toilet",qty:"1 unit",area:"Sikat khusus kloset dan leher angsa"},{no:15,name:"Kapi Gagang Karet / Scrapper",qty:"1 unit",area:"Pembersihan kerak lantai, noda permen karet, stiker"},{no:16,name:"Kuas Cat 4 Inch",qty:"2 pcs",area:"Pembersih sela keyboard, ventilasi, dan sudut sempit"},{no:17,name:"Warning Sign (Lantai Licin)",qty:"1 unit",area:"Digunakan saat mopping dan general cleaning"},{no:18,name:"Teleskopik + Pembersih Kaca/Langit",qty:"1 set",area:"Membersihkan jendela tinggi, plafon, dan sarang laba-laba"},{no:19,name:"Sarung Tangan Karet 3M",qty:"1 pasang",area:"Perlindungan tangan area toilet dan bahan kimia"},{no:20,name:"Botol Spray 500ml",qty:"5 unit",area:"Wadah larutan chemical (diberi label masing-masing)"},{no:21,name:"Pengharum Ruangan Stella Aerosol",qty:"1 kaleng",area:"Digunakan darurat apabila timbul bau menyengat"},{no:22,name:"Kleaner Gantungan Sapu & Pel",qty:"1 set",area:"Gantungan penyimpanan alat cleaning di janitor"},{no:23,name:"Floor Squeegee 75 cm",qty:"1 unit",area:"Pendorong genangan air lantai toilet/teras"},{no:24,name:"Sikat Botol Ikea",qty:"1 unit",area:"Di setiap ruang tindakan dokter gigi"},{no:25,name:"Scotch-Brite 3M Sikat Lantai ID-51",qty:"1 unit",area:"Sikat lantai kamar mandi dan teras"},{no:26,name:"Kamper Toilet / Urinal Block",qty:"1 pack",area:"Pengharum di toilet dan urinoir"},{no:27,name:"Plastik Sampah Hitam Sedang (60x100)",qty:"1 pack",area:"Tempat sampah indoor ruangan"},{no:28,name:"Plastik Sampah Hitam Besar (90x120)",qty:"1 pack",area:"Tempat sampah luar & penampung sampah akhir"},{no:29,name:"Plastik Sampah Medis Kuning (90x120)",qty:"1 pack",area:"Limbah medis padat B3 klinik"},{no:30,name:"Plastik Sampah Medis Kuning (40x60)",qty:"1 pack",area:"Tempat sampah medis kecil di ruang tindakan"},{no:31,name:"Tangga Besi Lipat",qty:"1 unit",area:"Alat bantu membersihkan plafon, lampu, dan AC"},{no:32,name:"Baskom Plastik Besar 55cm",qty:"1 unit",area:"Mencuci bola kids playground"}],Rn=[{no:1,name:"Floor Cleaner Prokleen Lavender",size:"5 Liter",ratio:"25 ml : 1 Liter Air",function:"Pembersihan dan pengepelan rutin seluruh lantai ruangan"},{no:2,name:"Karbol Prokleen Pine Disinfektan",size:"5 Liter",ratio:"25 ml : 1 Liter Air (Mop) / 20 ml per botol spray",function:"Mengepel lantai toilet, disinfeksi area basah dan berbau tak sedap"},{no:3,name:"Glass Cleaner Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air / Siap pakai di botol spray",function:"Membersihkan pintu kaca, jendela, cermin, dan layar monitor"},{no:4,name:"Manual Dishwashing Liquid Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air",function:"Mencuci piring, gelas, sendok pantry, dan mencuci bola playground"},{no:5,name:"Hand Soap Prokleen Green Tea",size:"5 Liter",ratio:"Siap pakai (tanpa campur)",function:"Refill sabun pencuci tangan di wastafel pasien dan tim"},{no:6,name:"Liquid Detergent Prokleen",size:"5 Liter",ratio:"Sesuai kapasitas mesin",function:"Mencuci handuk, kain microfiber, dan seragam"},{no:7,name:"Toilet Bowl Cleaner Prokleen",size:"5 Liter",ratio:"100 ml per botol spray / Khusus kerak",function:"Pembersih kerak membandel pada kloset, wastafel keramik, dan nat keramik"},{no:8,name:"Bubuk Pembersih KIFA",size:"500 gram",ratio:"Taburkan tipis pada tapas basah",function:"Membersihkan kerak air pada kran air stainless steel, jet washer, dan sink"}],_t="fcms_sop_fc_ik_v1",At="fcms_sop_fc_tools_v1",Bt="fcms_sop_fc_chem_v1";function Sa(){try{let e=localStorage.getItem(_t);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK",e)}return JSON.parse(JSON.stringify(Nn))}function Ft(e){localStorage.setItem(_t,JSON.stringify(e))}function Fa(){try{let e=localStorage.getItem(At);if(e)return JSON.parse(e)}catch(e){console.error("Error load tools",e)}return JSON.parse(JSON.stringify(qn))}function Ct(e){localStorage.setItem(At,JSON.stringify(e))}function Ca(){try{let e=localStorage.getItem(Bt);if(e)return JSON.parse(e)}catch(e){console.error("Error load chem",e)}return JSON.parse(JSON.stringify(Rn))}function $t(e){localStorage.setItem(Bt,JSON.stringify(e))}async function $a(e){let t=ae(),n=X(t,"sop","write"),r=X(t,"sop","delete"),i="ik",m="",f="all",u="tools",o="",p=Sa(),$=Fa(),F=Ca();function E(){e.innerHTML=`
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
          ${r?`
          <button id="btn-reset-default" class="btn btn-outline" title="Kembalikan data ke dokumen standar asli" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F504} Reset ke Standar
          </button>
          `:""}
          <button id="btn-print-sop" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F5A8}\uFE0F Cetak Ringkasan SOP
          </button>
          ${X(t,"export")?`<button id="btn-export-tools" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F4E5} Download Data Excel
          </button>`:""}
        </div>
      </div>

      <!-- Navigation Tabs -->
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto; flex-wrap:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none;">
        <button class="sop-tab-btn ${i==="ik"?"active":""}" data-tab="ik" style="${C(i==="ik")}">
          \u{1F4CB} Instruksi Kerja (IK) <span class="badge" style="${j(i==="ik")}">${p.length}</span>
        </button>
        <button class="sop-tab-btn ${i==="tools"?"active":""}" data-tab="tools" style="${C(i==="tools")}">
          \u{1F6E0}\uFE0F Alat & Chemical <span class="badge" style="${j(i==="tools")}">${$.length+F.length}</span>
        </button>
        <button class="sop-tab-btn ${i==="policy"?"active":""}" data-tab="policy" style="${C(i==="policy")}">
          \u{1F4D6} Pedoman & Ketentuan
        </button>
        <button class="sop-tab-btn ${i==="links"?"active":""}" data-tab="links" style="${C(i==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sop-content-area"></div>
    `,T(),D()}function C(d){return d?"border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function j(d){return d?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function T(){e.querySelectorAll(".sop-tab-btn").forEach(M=>{M.addEventListener("click",()=>{i=M.dataset.tab,E()})});let d=e.querySelector("#btn-reset-default");d&&d.addEventListener("click",()=>{re("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja, Alat, dan Chemical ke data standar dokumen asli?",()=>{localStorage.removeItem(_t),localStorage.removeItem(At),localStorage.removeItem(Bt),p=Sa(),$=Fa(),F=Ca(),H("Data SOP Facility Care berhasil di-reset ke standar asli"),E()},"Konfirmasi Reset Standar")});let y=e.querySelector("#btn-print-sop");y&&y.addEventListener("click",()=>{window.print()});let k=e.querySelector("#btn-export-tools");k&&k.addEventListener("click",()=>{w()})}function D(){let d=e.querySelector("#sop-content-area");d&&(i==="ik"?_(d):i==="tools"?s(d):i==="policy"?v(d):i==="links"&&x(d))}function _(d){d.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="ik-search" class="form-control" placeholder="Cari pekerjaan, alat, atau chemical..." value="${m}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="cat-pills-wrap">
              <button class="cat-pill ${f==="all"?"active":""}" data-cat="all">Semua (${p.length})</button>
              <button class="cat-pill ${f==="Pembersihan Rutin"?"active":""}" data-cat="Pembersihan Rutin">Rutin (${p.filter(P=>P.category==="Pembersihan Rutin").length})</button>
              <button class="cat-pill ${f==="Sanitasi Toilet"?"active":""}" data-cat="Sanitasi Toilet">Toilet & Wastafel (${p.filter(P=>P.category==="Sanitasi Toilet").length})</button>
              <button class="cat-pill ${f==="Klinis & Medis"?"active":""}" data-cat="Klinis & Medis">Klinis Dental (${p.filter(P=>P.category==="Klinis & Medis").length})</button>
              <button class="cat-pill ${f==="Elektronik & Utility"?"active":""}" data-cat="Elektronik & Utility">Elektronik & AC (${p.filter(P=>P.category==="Elektronik & Utility").length})</button>
              <button class="cat-pill ${f==="Pembersihan Khusus"?"active":""}" data-cat="Pembersihan Khusus">Khusus (${p.filter(P=>P.category==="Pembersihan Khusus").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${p.length} dari ${p.length} Instruksi Kerja
            </div>
            ${n?`
            <button id="btn-add-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
              \u2795 Tambah IK Baru
            </button>
            `:""}
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="ik-cards-grid-container"></div>
    `;let y=document.createElement("style");y.innerHTML=`
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
    `,d.appendChild(y);let k=d.querySelector("#ik-cards-grid-container"),M=d.querySelector("#ik-count-label");function I(){let P=p.filter(b=>{let h=f==="all"||b.category===f,S=m.toLowerCase().trim(),O=!S||b.title.toLowerCase().includes(S)||b.target.toLowerCase().includes(S)||b.tools.some(z=>z.toLowerCase().includes(S))||b.steps.some(z=>z.toLowerCase().includes(S));return h&&O});if(M&&(M.textContent=`Menampilkan ${P.length} dari ${p.length} Instruksi Kerja`),P.length===0){k.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}k.innerHTML=`
        <div class="sop-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${P.map(b=>`
            <div class="card ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${b.bg||"#EFF6FF"}; border-radius:10px;">${b.icon||"\u{1F4CB}"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        ${b.id}. ${b.title}
                      </h3>
                      <span class="badge" style="background:${b.bg||"#EFF6FF"}; color:${b.color||"#2563EB"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${b.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    ${n?`
                    <button class="btn btn-sm btn-ghost btn-edit-ik" data-id="${b.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    `:""}
                    ${r?`
                    <button class="btn btn-sm btn-ghost btn-delete-ik" data-id="${b.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                    `:""}
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${b.color||"#2563EB"};">
                  <strong>Area:</strong> ${b.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Bahan Utama:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${b.tools.slice(0,3).map(h=>`<li style="margin-bottom:2px;">${h}</li>`).join("")}
                    ${b.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${b.tools.length-3} alat lainnya...</li>`:""}
                  </ul>
                </div>

                ${b.chemical_ratio&&b.chemical_ratio!=="-"?`
                  <div style="font-size:0.8rem; background:#FEF3C7; color:#92400E; padding:6px 10px; border-radius:6px; margin-bottom:14px; font-weight:600;">
                    \u{1F9EA} Takaran: ${b.chemical_ratio}
                  </div>
                `:""}
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${b.steps.length} Langkah Kerja
                </span>
                <div style="display:flex; gap:6px;">
                  <button class="btn btn-sm btn-primary btn-view-ik" data-id="${b.id}" style="padding:6px 14px; font-weight:600; border-radius:6px;">
                    Lihat Panduan \u2794
                  </button>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,k.querySelectorAll(".btn-view-ik").forEach(b=>{b.addEventListener("click",()=>{let h=parseInt(b.dataset.id,10),S=p.find(O=>O.id===h);S&&g(S)})}),k.querySelectorAll(".btn-edit-ik").forEach(b=>{b.addEventListener("click",()=>{let h=parseInt(b.dataset.id,10),S=p.find(O=>O.id===h);S&&A(S)})}),k.querySelectorAll(".btn-delete-ik").forEach(b=>{b.addEventListener("click",()=>{let h=parseInt(b.dataset.id,10),S=p.find(O=>O.id===h);S&&re(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${S.title}"?`,()=>{p=p.filter(O=>O.id!==h),Ft(p),H("Instruksi Kerja berhasil dihapus"),E()},"Hapus Instruksi Kerja")})})}I();let K=d.querySelector("#ik-search");K&&K.addEventListener("input",P=>{m=P.target.value,I()}),d.querySelectorAll(".cat-pill").forEach(P=>{P.addEventListener("click",()=>{d.querySelectorAll(".cat-pill").forEach(b=>b.classList.remove("active")),P.classList.add("active"),f=P.dataset.cat,I()})});let B=d.querySelector("#btn-add-ik");B&&B.addEventListener("click",()=>{A()})}function g(d){Z({title:`${d.icon||"\u{1F4CB}"} IK-${d.id}: ${d.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${d.bg||"#EFF6FF"}; color:${d.color||"#2563EB"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${d.category}</span>
            <span>ID: IK-${String(d.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${d.color||"#2563EB"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${d.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Bahan yang Dibutuhkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${d.tools.map(y=>`<li>${y}</li>`).join("")}
              </ul>
            </div>
          </div>

          ${d.chemical_ratio&&d.chemical_ratio!=="-"?`
            <div style="background:#FEF3C7; border:1px solid #FCD34D; border-radius:8px; padding:10px 14px;">
              <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:2px;">\u{1F9EA} Rasio & Takaran Chemical:</strong>
              <span style="color:#78350F; font-size:0.9rem; font-weight:600;">${d.chemical_ratio}</span>
            </div>
          `:""}

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah-Langkah Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${d.steps.map(y=>`<li style="margin-bottom:6px;">${y}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${d.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Penting & Safety:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${d.notes}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(y,k)=>{k()}})}function A(d=null){if(!n)return;let y=!!d,k=d?d.tools.join(`
`):"",M=d?d.steps.join(`
`):"",I=`
      <form id="ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-ik-title" class="form-control" required value="${d?a(d.title):""}" placeholder="Contoh: Membersihkan Dental Unit" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-ik-icon" class="form-control" value="${d?d.icon:"\u{1F4CB}"}" placeholder="\u{1F9F9}" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
            <select id="form-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
              <option value="Pembersihan Rutin" ${d?.category==="Pembersihan Rutin"?"selected":""}>Pembersihan Rutin</option>
              <option value="Sanitasi Toilet" ${d?.category==="Sanitasi Toilet"?"selected":""}>Sanitasi Toilet</option>
              <option value="Klinis & Medis" ${d?.category==="Klinis & Medis"?"selected":""}>Klinis & Medis</option>
              <option value="Elektronik & Utility" ${d?.category==="Elektronik & Utility"?"selected":""}>Elektronik & Utility</option>
              <option value="Pembersihan Khusus" ${d?.category==="Pembersihan Khusus"?"selected":""}>Pembersihan Khusus</option>
            </select>
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Takaran / Rasio Chemical</label>
            <input type="text" id="form-ik-ratio" class="form-control" value="${d?a(d.chemical_ratio||""):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-ik-target" class="form-control" required value="${d?a(d.target):""}" placeholder="Contoh: Kursi pasien, meja instrumen, mangkuk spittoon" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Bahan yang Digunakan (1 item per baris) *</label>
          <textarea id="form-ik-tools" class="form-control" rows="4" placeholder="Kain Microfiber Biru&#10;Sarung Tangan Karet&#10;Botol Spray Glass Cleaner" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${a(k)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah-Langkah Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-ik-steps" class="form-control" rows="6" placeholder="1. Pastikan posisi saklar OFF.&#10;2. Lap jok dental unit dengan kain microfiber biru.&#10;3. Sikat mangkuk spittoon..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${a(M)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Penting / Hal yang Perlu Diperhatikan (Safety)</label>
          <textarea id="form-ik-notes" class="form-control" rows="2" placeholder="Catatan larangan atau peringatan keselamatan..." style="width:100%; border-radius:6px; font-size:0.85rem;">${d?a(d.notes||""):""}</textarea>
        </div>
      </form>
    `;Z({title:y?`\u270F\uFE0F Edit IK-${d.id}: ${d.title}`:"\u2795 Tambah Instruksi Kerja Baru",content:I,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(K,B)=>{let P=K.querySelector("#form-ik-title").value.trim(),b=K.querySelector("#form-ik-icon").value.trim()||"\u{1F4CB}",h=K.querySelector("#form-ik-cat").value,S=K.querySelector("#form-ik-ratio").value.trim(),O=K.querySelector("#form-ik-target").value.trim(),z=K.querySelector("#form-ik-tools").value.trim(),G=K.querySelector("#form-ik-steps").value.trim(),L=K.querySelector("#form-ik-notes").value.trim();if(!P||!O||!z||!G){N("Harap lengkapi seluruh field wajib (*)");return}let R=z.split(`
`).map(U=>U.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),V=G.split(`
`).map(U=>U.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),Q={"Pembersihan Rutin":{color:"#2563EB",bg:"#EFF6FF"},"Sanitasi Toilet":{color:"#DC2626",bg:"#FEF2F2"},"Klinis & Medis":{color:"#E11D48",bg:"#FFF1F2"},"Elektronik & Utility":{color:"#0D9488",bg:"#F0FDFA"},"Pembersihan Khusus":{color:"#7C3AED",bg:"#F5F3FF"}}[h]||{color:"#2563EB",bg:"#EFF6FF"};if(y)d.title=P,d.icon=b,d.category=h,d.chemical_ratio=S,d.target=O,d.tools=R,d.steps=V,d.notes=L,d.color=Q.color,d.bg=Q.bg,Ft(p),H(`Instruksi Kerja "${P}" berhasil diperbarui`);else{let U=p.length>0?Math.max(...p.map(ee=>ee.id))+1:1,Y={id:U,title:P,category:h,icon:b,color:Q.color,bg:Q.bg,target:O,chemical_ratio:S||"-",tools:R,steps:V,notes:L};p.push(Y),Ft(p),H(`Instruksi Kerja baru IK-${U} berhasil ditambahkan`)}B(),E()}})}function s(d){d.innerHTML=`
      <!-- Sub-tab Switcher & Actions -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-tools-btn" class="btn btn-sm ${u==="tools"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9F9} Alat Kebersihan (${$.length} Item)
          </button>
          <button id="view-chem-btn" class="btn btn-sm ${u==="chemical"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9EA} Bahan Chemical (${F.length} Item)
          </button>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="tool-search" class="form-control" placeholder="Cari nama / peruntukan / rasio..." value="${o}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
            ${n?`
          <button id="btn-add-tool-chem" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u2795 ${u==="tools"?"Tambah Alat Baru":"Tambah Chemical Baru"}
          </button>
            `:""}
        </div>
      </div>

      <!-- Table Container -->
      <div id="tool-table-container"></div>
    `;let y=d.querySelector("#tool-table-container"),k=d.querySelector("#tool-search");function M(){let I=o.toLowerCase().trim();if(u==="tools"){let K=$.filter(B=>!I||B.name.toLowerCase().includes(I)||B.area.toLowerCase().includes(I));y.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Standar Peralatan Kebersihan (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-primary">${K.length} Item</span>
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
                  ${K.map((B,P)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${P%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${B.no||P+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${a(B.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#E0F2FE; color:#0369A1; font-weight:600;">${a(B.qty)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${a(B.area)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                    ${n?`
                        <button class="btn btn-sm btn-ghost btn-edit-tool" data-no="${B.no}" title="Edit Alat" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                    `:""}
                    ${r?`
                        <button class="btn btn-sm btn-ghost btn-delete-tool" data-no="${B.no}" title="Hapus Alat" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    `:""}
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,y.querySelectorAll(".btn-edit-tool").forEach(B=>{B.addEventListener("click",()=>{let P=parseInt(B.dataset.no,10),b=$.find(h=>h.no===P);b&&l(b)})}),y.querySelectorAll(".btn-delete-tool").forEach(B=>{B.addEventListener("click",()=>{let P=parseInt(B.dataset.no,10),b=$.find(h=>h.no===P);b&&re(`Apakah Anda yakin ingin menghapus alat: "${b.name}"?`,()=>{$=$.filter(h=>h.no!==P),$.forEach((h,S)=>{h.no=S+1}),Ct($),H("Alat kebersihan berhasil dihapus"),E()},"Hapus Alat Kebersihan")})})}else{let K=F.filter(B=>!I||B.name.toLowerCase().includes(I)||B.function.toLowerCase().includes(I));y.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Bahan Kimia & Standar Takaran (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-success">${K.length} Chemical</span>
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
                  ${K.map((B,P)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${P%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${B.no||P+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${a(B.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#FEF3C7; color:#92400E; font-weight:600;">${a(B.size)}</span></td>
                      <td style="padding:10px 14px;"><span style="background:#EFF6FF; color:#1E40AF; padding:3px 8px; border-radius:6px; font-weight:600; font-size:0.82rem;">${a(B.ratio)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${a(B.function)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                    ${n?`
                        <button class="btn btn-sm btn-ghost btn-edit-chem" data-no="${B.no}" title="Edit Chemical" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                    `:""}
                    ${r?`
                        <button class="btn btn-sm btn-ghost btn-delete-chem" data-no="${B.no}" title="Hapus Chemical" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    `:""}
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,y.querySelectorAll(".btn-edit-chem").forEach(B=>{B.addEventListener("click",()=>{let P=parseInt(B.dataset.no,10),b=F.find(h=>h.no===P);b&&c(b)})}),y.querySelectorAll(".btn-delete-chem").forEach(B=>{B.addEventListener("click",()=>{let P=parseInt(B.dataset.no,10),b=F.find(h=>h.no===P);b&&re(`Apakah Anda yakin ingin menghapus chemical: "${b.name}"?`,()=>{F=F.filter(h=>h.no!==P),F.forEach((h,S)=>{h.no=S+1}),$t(F),H("Bahan chemical berhasil dihapus"),E()},"Hapus Bahan Chemical")})})}}M(),d.querySelector("#view-tools-btn").addEventListener("click",()=>{u="tools",s(d)}),d.querySelector("#view-chem-btn").addEventListener("click",()=>{u="chemical",s(d)}),k&&k.addEventListener("input",I=>{o=I.target.value,M()}),d.querySelector("#btn-add-tool-chem").addEventListener("click",()=>{u==="tools"?l():c()})}function l(d=null){if(!n)return;let y=!!d,k=`
      <form id="tool-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Alat / Item *</label>
          <input type="text" id="form-tool-name" class="form-control" required value="${d?a(d.name):""}" placeholder="Contoh: Lobby Duster Set 60 cm" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kuantitas Standar *</label>
          <input type="text" id="form-tool-qty" class="form-control" required value="${d?a(d.qty):"1 unit"}" placeholder="Contoh: 1 set, 2 pcs, 1 paket" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Peruntukan Area & Kegunaan *</label>
          <textarea id="form-tool-area" class="form-control" rows="3" required placeholder="Contoh: Lobby utama dan koridor/lorong klinik" style="width:100%; border-radius:6px; font-size:0.85rem;">${d?a(d.area):""}</textarea>
        </div>
      </form>
    `;Z({title:y?`\u270F\uFE0F Edit Alat: ${d.name}`:"\u2795 Tambah Alat Kebersihan Baru",content:k,confirmText:"Simpan",onConfirm:(M,I)=>{let K=M.querySelector("#form-tool-name").value.trim(),B=M.querySelector("#form-tool-qty").value.trim(),P=M.querySelector("#form-tool-area").value.trim();if(!K||!B||!P){N("Harap lengkapi semua kolom");return}if(y)d.name=K,d.qty=B,d.area=P,Ct($),H(`Alat "${K}" berhasil diperbarui`);else{let b=$.length>0?Math.max(...$.map(h=>h.no))+1:1;$.push({no:b,name:K,qty:B,area:P}),Ct($),H(`Alat baru "${K}" berhasil ditambahkan`)}I(),E()}})}function c(d=null){if(!n)return;let y=!!d,k=`
      <form id="chem-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Produk Chemical *</label>
          <input type="text" id="form-chem-name" class="form-control" required value="${d?a(d.name):""}" placeholder="Contoh: Floor Cleaner Prokleen Lavender" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kemasan Standar *</label>
            <input type="text" id="form-chem-size" class="form-control" required value="${d?a(d.size):"5 Liter"}" placeholder="Contoh: 5 Liter, 500 gram" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Rasio & Takaran Standar *</label>
            <input type="text" id="form-chem-ratio" class="form-control" required value="${d?a(d.ratio):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Fungsi & Cara Penggunaan *</label>
          <textarea id="form-chem-function" class="form-control" rows="3" required placeholder="Contoh: Pembersihan dan pengepelan rutin seluruh lantai ruangan" style="width:100%; border-radius:6px; font-size:0.85rem;">${d?a(d.function):""}</textarea>
        </div>
      </form>
    `;Z({title:y?`\u270F\uFE0F Edit Chemical: ${d.name}`:"\u2795 Tambah Chemical Baru",content:k,confirmText:"Simpan",onConfirm:(M,I)=>{let K=M.querySelector("#form-chem-name").value.trim(),B=M.querySelector("#form-chem-size").value.trim(),P=M.querySelector("#form-chem-ratio").value.trim(),b=M.querySelector("#form-chem-function").value.trim();if(!K||!B||!P||!b){N("Harap lengkapi semua kolom");return}if(y)d.name=K,d.size=B,d.ratio=P,d.function=b,$t(F),H(`Chemical "${K}" berhasil diperbarui`);else{let h=F.length>0?Math.max(...F.map(S=>S.no))+1:1;F.push({no:h,name:K,size:B,ratio:P,function:b}),$t(F),H(`Chemical baru "${K}" berhasil ditambahkan`)}I(),E()}})}function v(d){d.innerHTML=`
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
    `}function x(d){d.innerHTML=`
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
    `}async function w(){try{let{downloadExcel:d}=await Promise.resolve().then(()=>(pe(),fe)),y=p.map(M=>({"No IK":`IK-${M.id}`,"Judul Pekerjaan":M.title,Kategori:M.category,"Sasaran Target Area":M.target,"Rasio Chemical":M.chemical_ratio||"-","Alat & Bahan":M.tools.join(", "),"Langkah Kerja":M.steps.map((I,K)=>`${K+1}. ${I}`).join(" | "),"Catatan Safety":M.notes||"-"})),k=new Date().toISOString().slice(0,10);d(y,`SOP_Facility_Care_Instruksi_Kerja_${k}`),H("Data Excel Instruksi Kerja berhasil diunduh")}catch(d){console.error(d),N("Gagal mendownload data Excel: "+d.message)}}function a(d){return d?String(d).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}E()}ge();var On=[{id:1,title:"Pengamanan Area & Aset Perusahaan",category:"Pengamanan Fisik & Patroli",icon:"\u{1F6E1}\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Seluruh area dalam & luar Head Quarter dan Klinik Cabang Audy Dental",tools:["Google Form Checklist Monitoring Klinik: https://forms.gle/1EAvkqDJVdctXEGn7","Senter Patroli","Komputer dengan Akses Monitor CCTV (Live View)","Alat Komunikasi / Handphone / HT"],steps:["Lakukan Safety Patrol secara rutin mengelilingi seluruh area sesuai jadwal yang telah ditentukan (Patrol 1, 2, dan 3).","Catat hasil temuan dan kondisi keamanan secara berkala ke dalam Google Form Safety Patrol resmi.","Awasi kendaraan yang terparkir di area luar klinik; pastikan khusus sepeda motor dalam keadaan terkunci ganda.","Pantau situasi lingkungan dan area sensitif secara berkala melalui layar monitor CCTV saat berada di pos security.","Segera lakukan tindakan persuasif ramah apabila ditemukan potensi gangguan keamanan, dan segera hubungi pihak Kepolisian terdekat bila situasi darurat tidak dapat diatasi.","Bila ditemukan permasalahan fisik, kerusakan fasilitas, atau kendala di area klinik, segera dokumentasikan (foto) dan laporkan langsung ke SPV Klinik dan Leader Facility Care.","Disaat malam hari atau saat klinik telah tutup operasional, pastikan seluruh AC dan lampu ruangan dalam keadaan mati, kecuali lampu keamanan tertentu yang memang harus dibiarkan menyala terang."],notes:"Prioritaskan keselamatan diri dan tindakan persuasif. Dilarang meninggalkan area pengawasan tanpa personil pengganti yang jelas."},{id:2,title:"Penertiban & Pengaturan Parkir Kendaraan",category:"Operasional Parkir & Luar",icon:"\u{1F17F}\uFE0F",color:"#D97706",bg:"#FFFBEB",target:"Area parkir mobil, motor, akses keluar masuk, dan teras depan klinik",tools:["Lampu Lalu Lintas (Lampu Lalin)","Cone / Kerucut Pembatas Parkir","Peluit Parkir"],steps:["Berdiri di posisi strategis dan bantu mengarahkan setiap kendaraan pasien atau tamu yang datang secara tepat dan rapi ke slot parkir yang tersedia.","Berikan informasi dan solusi dengan sikap ramah serta santun apabila seluruh slot parkir klinik dalam kondisi penuh.","Dalam kondisi tertentu (bila ada mobil yang terhalang mobil lain dan harus dipindahkan, sementara pasien sedang dalam tindakan perawatan): bantu memindahkan kendaraan HANYA JIKA petugas security memiliki kemampuan mengemudi dan memiliki SIM aktif.","Apabila petugas security tidak memiliki kemampuan mengemudikan mobil: sampaikan permohonan maaf dengan tulus kepada pasien/tamu, minta untuk menunggu sejenak, dan segera minta bantuan kepada staf Resepsionis untuk menghubungi pemilik mobil yang menghalangi.","Bantu pasien atau tamu yang akan keluar parkir dengan memberikan aba-aba dan arahan lalu lintas jalan raya yang aman dan tepat."],notes:"Dilarang memaksakan memindahkan mobil pasien jika tidak memiliki lisensi mengemudi demi menghindari risiko gesekan atau kerusakan kendaraan."},{id:3,title:"Penerimaan Tamu & Pasien Klinik",category:"Layanan & Hospitalitas",icon:"\u{1F91D}",color:"#059669",bg:"#ECFDF5",target:"Pintu masuk utama klinik, lobby, area resepsionis, dan pos penerimaan tamu",tools:["Alat Tulis Kantor (Pulpen)","Buku Tamu / Log Book Keamanan","Alat Komunikasi / Handphone"],steps:['Berdiri dengan tegap dan sigap, bukakan pintu untuk setiap pasien/tamu yang datang dengan menerapkan budaya 5S (Senyum, Salam, Sapa, Sopan, Santun) dan menyapa: "Selamat pagi / siang / sore, ada yang bisa kami bantu?".',"Setelah pasien atau tamu memberitahukan tujuannya, persilakan dan arahkan pasien langsung menuju meja Resepsionis.","Apabila yang datang adalah tamu instansi, vendor, atau pihak Manajemen/BOD: persilakan mengisi Buku Tamu (Log Book) terlebih dahulu dan segera hubungi staf/karyawan yang ingin ditemui.",'Apabila pasien langsung melangkah ke arah tujuannya, berikan gestur tubuh hormat menyilangkan kedua tangan di depan, dan jika pasien/tamu mengucapkan terima kasih, jawab dengan ramah: "Dengan Senang Hati".',"Dalam kondisi khusus (pasien lansia, ibu hamil, atau pengguna kursi roda): petugas security wajib mendampingi, mengawal, dan membantu mendorong kursi roda sampai ke depan meja resepsionis.",'Buka kembali pintu untuk setiap pasien yang telah selesai melakukan perawatan saat hendak keluar klinik sambil mengucapkan ucapan hangat: "Terima Kasih, semoga lekas sehat kembali".'],notes:"Jaga senyum tulus dan keramahan sebagai garda terdepan representasi citra Audy Dental."}],Hn=[{no:1,type:"Aset Tetap",item:"Rumah & Bangunan",description:"Gedung klinik, pos satpam, pagar, kanopi, teras beton",pic:"Building Maintenance & Security"},{no:2,type:"Aset Tetap",item:"Genset & Lift",description:"Mesin generator listrik cadangan dan elevator penumpang tetap",pic:"Teknisi & Security"},{no:3,type:"Aset Tetap",item:"AC Central & Split Tetap",description:"Outdoor dan indoor AC yang terpasang permanen di dinding/plafon",pic:"Teknisi AC & Security"},{no:4,type:"Aset Tetap",item:"Dental Unit Permanen",description:"Kursi dental unit terpasang instalasi air dan kompresor",pic:"Tim Klinik & Security"},{no:5,type:"Aset Tetap",item:"Instalasi Listrik, Pipa & Jaringan IT Permanen",description:"Panel MCB, kabel fiber optik, pipa drainase, server rack",pic:"IT & Teknisi"},{no:6,type:"Aset Tetap",item:"Furniture Melekat (Built-in)",description:"Meja resepsionis cor/melekat, partisi kaca mati, kitchen set pantry",pic:"General Support"},{no:7,type:"Aset Bergerak",item:"Komputer, Laptop & Printer",description:"PC kasir/resepsionis, laptop staf, printer rontgen/dokumen",pic:"Staf Terkait & IT"},{no:8,type:"Aset Bergerak",item:"Kendaraan Operasional",description:"Mobil operasional dan sepeda motor inventaris perusahaan",pic:"Security & Driver"},{no:9,type:"Aset Bergerak",item:"Furniture Lepasan",description:"Meja kerja, kursi resepsionis, kursi tunggu pasien, lemari file",pic:"Facility Care & Staf"},{no:10,type:"Aset Bergerak",item:"Hand Tools, Alat Kerja & Alat Dental Portabel",description:"Handpiece dental, alat rontgen portable, toolbox teknisi",pic:"Dokter/Perawat & Teknisi"},{no:11,type:"Low Value Asset",item:"Peralatan Kantor (Telepon, ATK)",description:"Pesawat telepon, kalkulator, stapler, perlengkapan meja",pic:"Resepsionis"},{no:12,type:"Low Value Asset",item:"Alat-alat Kebersihan",description:"Lobby duster, ember wringer pel, sapu, tangga, botol spray",pic:"Facility Care"},{no:13,type:"Low Value Asset",item:"Perlengkapan Security",description:"Lampu lalin, cone parkir, peluit, senter patroli, jas hujan",pic:"Security"}],_a={cabang:[{zone:"Zona Hijau",level:"Akses Publik Bebas",color:"#059669",bg:"#ECFDF5",rooms:["Area Parkir Kendaraan","Lobby Lantai 1","Teras Luar Klinik"],desc:"Area dengan pengamanan umum yang bebas diakses pasien dan tamu setelah pengarahan security."},{zone:"Zona Kuning",level:"Akses Terbatas / Pendampingan",color:"#D97706",bg:"#FFFBEB",rooms:["Koridor Ruang Tindakan","Tangga & Lift Klinik","Ruang Tunggu Lantai 2 & 3","Kids Playground","Toilet Lantai 1 & 2","Mushola Staf & Pasien"],desc:"Area terbatas, tamu wajib didampingi petugas atau staf klinik terkait."},{zone:"Zona Merah",level:"Area Terlarang / Restricted",color:"#DC2626",bg:"#FEF2F2",rooms:["Area Meja Receptionist / Kasir","Seluruh Ruang Dental Unit (DU 101, 102, 201, 202, 301, 302)","Ruang Sterilisasi Alat","Ruang Rontgen Radiologi","Ruang Suction & Kompresor","Ruang Tim Staf & Gudang Material Medis"],desc:"Area sangat sensitif, hanya staf medis bertugas dan atasan pemilik area. Tamu wajib izin Manajemen/BOD dan didampingi penuh."}],hq:[{floor:"Lantai 1",hijau:["Area Parkir Mobil/Motor","Lobby Resepsionis Lantai 1"],kuning:["Seluruh Ruang Meeting Lantai 1","Ruang Training Lantai 1","Koridor Lift & Toilet"],merah:["Ruang Customer Services","Gudang Utama Lantai 1","Gudang Facility Care","Gudang Teknisi & Ruang Teknisi","Pos Satpam & Ruang Genset"]},{floor:"Lantai 2",hijau:["-"],kuning:["Ruang Meeting Lantai 2","Koridor Lift & Toilet","Pantry"],merah:["Ruang CEO","Ruang Sekretaris CEO","Ruang Human Capital","Ruang Dentist Relation","Ruang General Support","Ruang Eks. Patrel"]},{floor:"Lantai 3",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry","Mushola & Tempat Wudhu","Ruang Santai Karyawan"],merah:["Ruang Marketing","Studio Konten Marketing","Gudang Marketing","Ruang Patient Relation","Ruang OPS","Ruang Training Lantai 3","Gudang Lantai 3"]},{floor:"Lantai 4",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry"],merah:["Ruang Finance & Accounting","Ruang VP Finance & Accounting","Ruang Audit Internal","Ruang Training Lantai 4","Gudang Lantai 4"]},{floor:"Lantai 5",hijau:["-"],kuning:["Lobby Lantai 5","Pantry Lantai 5","Toilet Lantai 5","Ruang Training Besar Lantai 5"],merah:["Area Dak Toren Air & Rooftop"]}]},Pt="fcms_sop_sec_ik_v1",Lt="fcms_sop_sec_assets_v1";function Aa(){try{let e=localStorage.getItem(Pt);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK Sec",e)}return JSON.parse(JSON.stringify(On))}function Tt(e){localStorage.setItem(Pt,JSON.stringify(e))}function Ba(){try{let e=localStorage.getItem(Lt);if(e)return JSON.parse(e)}catch(e){console.error("Error load assets sec",e)}return JSON.parse(JSON.stringify(Hn))}function Dt(e){localStorage.setItem(Lt,JSON.stringify(e))}async function Ta(e){let t=ae(),n=X(t,"sop","write"),r=X(t,"sop","delete"),i="ik",m="",f="all",u="",o="cabang",p=Aa(),$=Ba();function F(){e.innerHTML=`
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
          ${r?`
          <button id="btn-reset-sec-default" class="btn btn-outline" title="Kembalikan data ke dokumen standar asli" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F504} Reset ke Standar
          </button>
          `:""}
          <button id="btn-print-sec" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F5A8}\uFE0F Cetak Ringkasan SOP
          </button>
          ${X(t,"export")?`<button id="btn-export-sec" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F4E5} Download Data Excel
          </button>`:""}
        </div>
      </div>

      <!-- Navigation Tabs -->
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto; flex-wrap:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none;">
        <button class="sec-tab-btn ${i==="ik"?"active":""}" data-tab="ik" style="${E(i==="ik")}">
          \u{1F6E1}\uFE0F Instruksi Kerja (IK) <span class="badge" style="${C(i==="ik")}">${p.length}</span>
        </button>
        <button class="sec-tab-btn ${i==="zonasi"?"active":""}" data-tab="zonasi" style="${E(i==="zonasi")}">
          \u{1F5FA}\uFE0F Zonasi Akses & Denah
        </button>
        <button class="sec-tab-btn ${i==="assets"?"active":""}" data-tab="assets" style="${E(i==="assets")}">
          \u{1F4E6} Klasifikasi Aset <span class="badge" style="${C(i==="assets")}">${$.length}</span>
        </button>
        <button class="sec-tab-btn ${i==="policy"?"active":""}" data-tab="policy" style="${E(i==="policy")}">
          \u{1F4CB} Shift, Patrol & Insiden
        </button>
        <button class="sec-tab-btn ${i==="links"?"active":""}" data-tab="links" style="${E(i==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sec-content-area"></div>
    `,j(),T()}function E(a){return a?"border:none; background:#1E40AF; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function C(a){return a?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function j(){e.querySelectorAll(".sec-tab-btn").forEach(k=>{k.addEventListener("click",()=>{i=k.dataset.tab,F()})});let a=e.querySelector("#btn-reset-sec-default");a&&a.addEventListener("click",()=>{re("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja dan Klasifikasi Aset Security ke standar dokumen asli?",()=>{localStorage.removeItem(Pt),localStorage.removeItem(Lt),p=Aa(),$=Ba(),H("Data SOP Security berhasil di-reset ke standar dokumen"),F()},"Reset ke Standar Dokumen")});let d=e.querySelector("#btn-print-sec");d&&d.addEventListener("click",()=>{window.print()});let y=e.querySelector("#btn-export-sec");y&&y.addEventListener("click",()=>{x()})}function T(){let a=e.querySelector("#sec-content-area");a&&(i==="ik"?D(a):i==="zonasi"?A(a):i==="assets"?s(a):i==="policy"?c(a):i==="links"&&v(a))}function D(a){a.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="sec-ik-search" class="form-control" placeholder="Cari instruksi kerja security..." value="${m}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="sec-cat-pills-wrap">
              <button class="sec-cat-pill ${f==="all"?"active":""}" data-cat="all">Semua (${p.length})</button>
              <button class="sec-cat-pill ${f==="Pengamanan Fisik & Patroli"?"active":""}" data-cat="Pengamanan Fisik & Patroli">Patroli & Fisik (${p.filter(B=>B.category==="Pengamanan Fisik & Patroli").length})</button>
              <button class="sec-cat-pill ${f==="Operasional Parkir & Luar"?"active":""}" data-cat="Operasional Parkir & Luar">Parkir & Luar (${p.filter(B=>B.category==="Operasional Parkir & Luar").length})</button>
              <button class="sec-cat-pill ${f==="Layanan & Hospitalitas"?"active":""}" data-cat="Layanan & Hospitalitas">Tamu & Pasien (${p.filter(B=>B.category==="Layanan & Hospitalitas").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="sec-ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${p.length} Instruksi Kerja
            </div>
            ${n?`
            <button id="btn-add-sec-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
              \u2795 Tambah IK Baru
            </button>
            `:""}
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="sec-ik-cards-grid-container"></div>
    `;let d=document.createElement("style");d.innerHTML=`
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
    `,a.appendChild(d);let y=a.querySelector("#sec-ik-cards-grid-container"),k=a.querySelector("#sec-ik-count-label");function M(){let B=p.filter(P=>{let b=f==="all"||P.category===f,h=m.toLowerCase().trim(),S=!h||P.title.toLowerCase().includes(h)||P.target.toLowerCase().includes(h)||P.tools.some(O=>O.toLowerCase().includes(h))||P.steps.some(O=>O.toLowerCase().includes(h));return b&&S});if(k&&(k.textContent=`Menampilkan ${B.length} dari ${p.length} Instruksi Kerja`),B.length===0){y.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}y.innerHTML=`
        <div class="sec-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${B.map(P=>`
            <div class="card sec-ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${P.bg||"#EFF6FF"}; border-radius:10px;">${P.icon||"\u{1F6E1}\uFE0F"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        IK-${String(P.id).padStart(2,"0")}. ${P.title}
                      </h3>
                      <span class="badge" style="background:${P.bg||"#EFF6FF"}; color:${P.color||"#1E40AF"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${P.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    ${n?`
                    <button class="btn btn-sm btn-ghost btn-edit-sec-ik" data-id="${P.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    `:""}
                    ${r?`
                    <button class="btn btn-sm btn-ghost btn-delete-sec-ik" data-id="${P.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                    `:""}
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${P.color||"#1E40AF"};">
                  <strong>Area Sasaran:</strong> ${P.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Perlengkapan:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${P.tools.slice(0,3).map(b=>`<li style="margin-bottom:2px;">${w(b)}</li>`).join("")}
                    ${P.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${P.tools.length-3} perlengkapan lainnya...</li>`:""}
                  </ul>
                </div>
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${P.steps.length} Langkah Prosedur
                </span>
                <button class="btn btn-sm btn-primary btn-view-sec-ik" data-id="${P.id}" style="padding:6px 14px; font-weight:600; border-radius:6px; background:#1E40AF; border-color:#1E40AF;">
                  Lihat Panduan Lengkap \u2794
                </button>
              </div>
            </div>
          `).join("")}
        </div>
      `,y.querySelectorAll(".btn-view-sec-ik").forEach(P=>{P.addEventListener("click",()=>{let b=parseInt(P.dataset.id,10),h=p.find(S=>S.id===b);h&&_(h)})}),y.querySelectorAll(".btn-edit-sec-ik").forEach(P=>{P.addEventListener("click",()=>{let b=parseInt(P.dataset.id,10),h=p.find(S=>S.id===b);h&&g(h)})}),y.querySelectorAll(".btn-delete-sec-ik").forEach(P=>{P.addEventListener("click",()=>{let b=parseInt(P.dataset.id,10),h=p.find(S=>S.id===b);h&&re(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${h.title}"?`,()=>{p=p.filter(S=>S.id!==b),Tt(p),H("Instruksi Kerja Security berhasil dihapus"),F()},"Hapus Instruksi Kerja")})})}M();let I=a.querySelector("#sec-ik-search");I&&I.addEventListener("input",B=>{m=B.target.value,M()}),a.querySelectorAll(".sec-cat-pill").forEach(B=>{B.addEventListener("click",()=>{a.querySelectorAll(".sec-cat-pill").forEach(P=>P.classList.remove("active")),B.classList.add("active"),f=B.dataset.cat,M()})});let K=a.querySelector("#btn-add-sec-ik");K&&K.addEventListener("click",()=>{g()})}function _(a){Z({title:`${a.icon||"\u{1F6E1}\uFE0F"} IK-${String(a.id).padStart(2,"0")}: ${a.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${a.bg||"#EFF6FF"}; color:${a.color||"#1E40AF"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${a.category}</span>
            <span>Kode: IK-SEC-${String(a.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${a.color||"#1E40AF"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${a.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Perlengkapan yang Disiapkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${a.tools.map(d=>{if(d.includes("http")){let y=d.split(": ");return`<li>${y[0]}: <a href="${y.slice(1).join(": ")}" target="_blank" rel="noopener" style="color:#2563EB; font-weight:600; text-decoration:underline;">Buka Tautan Form \u2794</a></li>`}return`<li>${w(d)}</li>`}).join("")}
              </ul>
            </div>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah & Prosedur Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${a.steps.map(d=>`<li style="margin-bottom:6px;">${w(d)}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${a.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Khusus & Ketentuan Penting:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${w(a.notes)}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(d,y)=>{y()}})}function g(a=null){if(!n)return;let d=!!a,y=a?a.tools.join(`
`):"",k=a?a.steps.join(`
`):"",M=`
      <form id="sec-ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-sec-ik-title" class="form-control" required value="${a?w(a.title):""}" placeholder="Contoh: Pengamanan Area & Aset Perusahaan" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-sec-ik-icon" class="form-control" value="${a?a.icon:"\u{1F6E1}\uFE0F"}" placeholder="\u{1F6E1}\uFE0F" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
          <select id="form-sec-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Pengamanan Fisik & Patroli" ${a?.category==="Pengamanan Fisik & Patroli"?"selected":""}>Pengamanan Fisik & Patroli</option>
            <option value="Operasional Parkir & Luar" ${a?.category==="Operasional Parkir & Luar"?"selected":""}>Operasional Parkir & Luar</option>
            <option value="Layanan & Hospitalitas" ${a?.category==="Layanan & Hospitalitas"?"selected":""}>Layanan & Hospitalitas</option>
          </select>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-sec-ik-target" class="form-control" required value="${a?w(a.target):""}" placeholder="Contoh: Seluruh area luar klinik dan akses parkir" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Perlengkapan (1 per baris) *</label>
          <textarea id="form-sec-ik-tools" class="form-control" rows="3" placeholder="Senter Patroli&#10;Lampu Lalin&#10;Peluit" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${w(y)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah & Prosedur Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-sec-ik-steps" class="form-control" rows="5" placeholder="1. Lakukan Safety Patrol sesuai jadwal.&#10;2. Arahkan kendaraan pasien dengan ramah..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${w(k)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Khusus / Ketentuan Penting</label>
          <textarea id="form-sec-ik-notes" class="form-control" rows="2" placeholder="Peringatan keselamatan atau pedoman khusus..." style="width:100%; border-radius:6px; font-size:0.85rem;">${a?w(a.notes||""):""}</textarea>
        </div>
      </form>
    `;Z({title:d?`\u270F\uFE0F Edit IK-${a.id}: ${a.title}`:"\u2795 Tambah Instruksi Kerja Security Baru",content:M,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(I,K)=>{let B=I.querySelector("#form-sec-ik-title").value.trim(),P=I.querySelector("#form-sec-ik-icon").value.trim()||"\u{1F6E1}\uFE0F",b=I.querySelector("#form-sec-ik-cat").value,h=I.querySelector("#form-sec-ik-target").value.trim(),S=I.querySelector("#form-sec-ik-tools").value.trim(),O=I.querySelector("#form-sec-ik-steps").value.trim(),z=I.querySelector("#form-sec-ik-notes").value.trim();if(!B||!h||!S||!O){N("Harap lengkapi semua kolom bertanda bintang (*)");return}let G=S.split(`
`).map(J=>J.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),L=O.split(`
`).map(J=>J.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),V={"Pengamanan Fisik & Patroli":{color:"#2563EB",bg:"#EFF6FF"},"Operasional Parkir & Luar":{color:"#D97706",bg:"#FFFBEB"},"Layanan & Hospitalitas":{color:"#059669",bg:"#ECFDF5"}}[b]||{color:"#1E40AF",bg:"#EFF6FF"};if(d)a.title=B,a.icon=P,a.category=b,a.target=h,a.tools=G,a.steps=L,a.notes=z,a.color=V.color,a.bg=V.bg,Tt(p),H(`Instruksi Kerja "${B}" berhasil diperbarui`);else{let J=p.length>0?Math.max(...p.map(U=>U.id))+1:1,Q={id:J,title:B,category:b,icon:P,color:V.color,bg:V.bg,target:h,tools:G,steps:L,notes:z};p.push(Q),Tt(p),H(`Instruksi Kerja baru IK-${J} berhasil ditambahkan`)}K(),F()}})}function A(a){a.innerHTML=`
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:20px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-cabang-btn" class="btn btn-sm ${o==="cabang"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F3E5} Zonasi Klinik / Cabang
          </button>
          <button id="view-hq-btn" class="btn btn-sm ${o==="hq"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
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
    `;let d=a.querySelector("#zonasi-content-container");o==="cabang"?d.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
          ${_a.cabang.map(y=>`
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; border-bottom:1px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.05rem; font-weight:700; color:${y.color}; display:flex; align-items:center; gap:8px;">
                  <span style="width:12px; height:12px; border-radius:50%; background:${y.color}; display:inline-block;"></span>
                  ${y.zone}
                </h3>
                <span class="badge" style="background:${y.bg}; color:${y.color}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:6px;">
                  ${y.level}
                </span>
              </div>
              <p style="font-size:0.82rem; color:var(--text-3); margin-bottom:14px; line-height:1.5;">${y.desc}</p>
              <div>
                <strong style="font-size:0.8rem; text-transform:uppercase; color:var(--text-2); display:block; margin-bottom:6px;">Daftar Ruangan Terkait:</strong>
                <ul style="margin:0 0 0 16px; padding:0; font-size:0.85rem; color:var(--text-1); line-height:1.7;">
                  ${y.rooms.map(k=>`<li>${k}</li>`).join("")}
                </ul>
              </div>
            </div>
          `).join("")}
        </div>
      `:d.innerHTML=`
        <div style="display:flex; flex-direction:column; gap:16px;">
          ${_a.hq.map(y=>`
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; border-bottom:2px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1); display:flex; align-items:center; gap:8px;">
                  <span>\u{1F3E2}</span> Head Quarter \u2014 ${y.floor}
                </h3>
              </div>
              <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:14px;">
                <!-- Zona Hijau -->
                <div style="background:#ECFDF5; border:1px solid #A7F3D0; border-radius:8px; padding:12px;">
                  <strong style="color:#065F46; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E2} ZONA HIJAU (Publik):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#064E3B; line-height:1.6;">
                    ${y.hijau.map(k=>`<li>${k}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Kuning -->
                <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:8px; padding:12px;">
                  <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E1} ZONA KUNING (Terbatas):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#78350F; line-height:1.6;">
                    ${y.kuning.map(k=>`<li>${k}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Merah -->
                <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:8px; padding:12px;">
                  <strong style="color:#991B1B; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F534} ZONA MERAH (Restricted):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#7F1D1D; line-height:1.6;">
                    ${y.merah.map(k=>`<li>${k}</li>`).join("")}
                  </ul>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,a.querySelector("#view-cabang-btn").addEventListener("click",()=>{o="cabang",A(a)}),a.querySelector("#view-hq-btn").addEventListener("click",()=>{o="hq",A(a)})}function s(a){a.innerHTML=`
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
          <input type="text" id="sec-asset-search" class="form-control" placeholder="Cari nama aset / tipe / PIC..." value="${u}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
            ${n?`
          <button id="btn-add-sec-asset" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
            \u2795 Tambah Aset Baru
          </button>
            `:""}
        </div>
      </div>

      <div id="sec-asset-table-container"></div>
    `;let d=a.querySelector("#sec-asset-table-container"),y=a.querySelector("#sec-asset-search");function k(){let I=u.toLowerCase().trim(),K=$.filter(B=>!I||B.item.toLowerCase().includes(I)||B.type.toLowerCase().includes(I)||B.description.toLowerCase().includes(I)||B.pic&&B.pic.toLowerCase().includes(I));d.innerHTML=`
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
            <strong style="color:var(--text-1); font-size:0.95rem;">Rekap Klasifikasi Aset Perusahaan</strong>
            <span class="badge badge-primary">${K.length} Aset</span>
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
                ${K.map((B,P)=>`
                  <tr style="border-bottom:1px solid #F1F5F9; background:${P%2===0?"#fff":"#FAFAFA"};">
                    <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${B.no||P+1}</td>
                    <td style="padding:10px 14px;">
                      <span class="badge" style="${M(B.type)}">${w(B.type)}</span>
                    </td>
                    <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${w(B.item)}</td>
                    <td style="padding:10px 14px; color:var(--text-2);">${w(B.description)}</td>
                    <td style="padding:10px 14px; font-weight:600; color:#1E40AF;">${w(B.pic||"-")}</td>
                    <td style="padding:10px 14px; text-align:center;">
                    ${n?`
                      <button class="btn btn-sm btn-ghost btn-edit-asset" data-no="${B.no}" title="Edit Aset" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                    `:""}
                    ${r?`
                      <button class="btn btn-sm btn-ghost btn-delete-asset" data-no="${B.no}" title="Hapus Aset" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    `:""}
                    </td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </div>
      `,d.querySelectorAll(".btn-edit-asset").forEach(B=>{B.addEventListener("click",()=>{let P=parseInt(B.dataset.no,10),b=$.find(h=>h.no===P);b&&l(b)})}),d.querySelectorAll(".btn-delete-asset").forEach(B=>{B.addEventListener("click",()=>{let P=parseInt(B.dataset.no,10),b=$.find(h=>h.no===P);b&&re(`Apakah Anda yakin ingin menghapus aset: "${b.item}"?`,()=>{$=$.filter(h=>h.no!==P),$.forEach((h,S)=>{h.no=S+1}),Dt($),H("Aset berhasil dihapus"),F()},"Hapus Aset")})})}function M(I){return I==="Aset Tetap"?"background:#EFF6FF; color:#1E40AF; font-weight:700; padding:3px 8px; border-radius:6px;":I==="Aset Bergerak"?"background:#ECFDF5; color:#047857; font-weight:700; padding:3px 8px; border-radius:6px;":"background:#FEF3C7; color:#92400E; font-weight:700; padding:3px 8px; border-radius:6px;"}k(),y&&y.addEventListener("input",I=>{u=I.target.value,k()}),a.querySelector("#btn-add-sec-asset").addEventListener("click",()=>{l()})}function l(a=null){if(!n)return;let d=!!a,y=`
      <form id="asset-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Tipe / Klasifikasi Aset *</label>
          <select id="form-asset-type" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Aset Tetap" ${a?.type==="Aset Tetap"?"selected":""}>Aset Tetap</option>
            <option value="Aset Bergerak" ${a?.type==="Aset Bergerak"?"selected":""}>Aset Bergerak</option>
            <option value="Low Value Asset" ${a?.type==="Low Value Asset"?"selected":""}>Low Value Asset</option>
          </select>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Aset / Kelompok *</label>
          <input type="text" id="form-asset-item" class="form-control" required value="${a?w(a.item):""}" placeholder="Contoh: Genset & Elevator" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Deskripsi / Contoh Item *</label>
          <textarea id="form-asset-desc" class="form-control" rows="3" required placeholder="Contoh: Generator daya darurat dan panel listrik utama" style="width:100%; border-radius:6px; font-size:0.85rem;">${a?w(a.description):""}</textarea>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">PIC Pengawas / Penanggung Jawab *</label>
          <input type="text" id="form-asset-pic" class="form-control" required value="${a?w(a.pic):""}" placeholder="Contoh: Teknisi & Security" style="width:100%; border-radius:6px; height:36px;">
        </div>
      </form>
    `;Z({title:d?`\u270F\uFE0F Edit Aset: ${a.item}`:"\u2795 Tambah Klasifikasi Aset Baru",content:y,confirmText:"Simpan",onConfirm:(k,M)=>{let I=k.querySelector("#form-asset-type").value,K=k.querySelector("#form-asset-item").value.trim(),B=k.querySelector("#form-asset-desc").value.trim(),P=k.querySelector("#form-asset-pic").value.trim();if(!K||!B||!P){N("Harap lengkapi semua kolom");return}if(d)a.type=I,a.item=K,a.description=B,a.pic=P,Dt($),H(`Aset "${K}" berhasil diperbarui`);else{let b=$.length>0?Math.max(...$.map(h=>h.no))+1:1;$.push({no:b,type:I,item:K,description:B,pic:P}),Dt($),H(`Aset baru "${K}" berhasil ditambahkan`)}M(),F()}})}function c(a){a.innerHTML=`
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
    `}function v(a){a.innerHTML=`
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
    `}async function x(){try{let{downloadExcel:a}=await Promise.resolve().then(()=>(pe(),fe)),d=p.map(k=>({"Kode IK":`IK-SEC-${String(k.id).padStart(2,"0")}`,"Judul Prosedur":k.title,Kategori:k.category,"Sasaran Target":k.target,"Alat & Perlengkapan":k.tools.join(", "),"Langkah Pelaksanaan":k.steps.map((M,I)=>`${I+1}. ${M}`).join(" | "),"Catatan Khusus":k.notes||"-"})),y=new Date().toISOString().slice(0,10);a(d,`SOP_Security_Instruksi_Kerja_${y}`),H("Data Excel SOP Security berhasil diunduh")}catch(a){console.error(a),N("Gagal mendownload data Excel: "+a.message)}}function w(a){return a?String(a).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}F()}ge();var Da=[{id:1,name:"Berita Acara Kerusakan Alat",category:"Berita Acara",document_link:"https://docs.google.com/document/d/1example_ba_kerusakan",description:"Format berita acara resmi pencatatan dan eskalasi kerusakan peralatan klinik dan aset fasilitas."},{id:2,name:"Checklist Apar",category:"Apar & K3",document_link:"https://docs.google.com/spreadsheets/d/1example_apar",description:"Lembar inspeksi berkala tabung pemadam api ringan (APAR), tekanan manometer, dan masa kadaluarsa."},{id:3,name:"Checklist Deep Cleaning",category:"Kegiatan Deep Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_deep_cleaning",description:"Panduan dan verifikasi checklist pembersihan mendalam (deep cleaning) area dental unit dan klinik."},{id:4,name:"Checklist Fulldraine",category:"Drain Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_fulldraine",description:"Checklist monitoring pembersihan saluran pembuangan air kotor, floor drain, dan sedotan suction."},{id:5,name:"Checklist General Cleaning",category:"Kegiatan General Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_general_cleaning",description:"Standar checklist kebersihan harian/mingguan seluruh ruangan, lobby, toilet, dan koridor."},{id:6,name:"Form Aktivity Reliefer",category:"Form Operasional",document_link:"https://forms.gle/1example_reliefer",description:"Formulir pencatatan absensi dan penugasan harian reliefer Facility Care di cabang pengganti."},{id:7,name:"Form Aktivity Security HQ",category:"Form Operasional",document_link:"https://forms.gle/1example_security_hq",description:"Formulir monitoring shift dan pengawasan keamanan harian di Kantor Pusat (Head Quarter)."}];function Pa(e=""){let t=(e||"").toLowerCase();return t.includes("berita acara")?{label:e,bg:"#F3E8FF",color:"#7E22CE",border:"#D8B4FE",icon:"\u2696\uFE0F"}:t.includes("apar")||t.includes("k3")||t.includes("safety")?{label:e,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F9EF}"}:t.includes("deep cleaning")?{label:e,bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2728"}:t.includes("drain")?{label:e,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u{1F6BF}"}:t.includes("general cleaning")||t.includes("kebersihan")?{label:e,bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F9F9}"}:t.includes("form")?{label:e,bg:"#E0E7FF",color:"#4338CA",border:"#C7D2FE",icon:"\u{1F4DD}"}:{label:e||"Dokumen",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4C1}"}}function La(e=""){if(!e)return{type:"Tidak Ada Tautan",icon:"\u{1F4CE}",badgeClass:"bg-slate-100 text-slate-500"};let t=e.toLowerCase();return t.includes("forms.gle")||t.includes("google.com/forms")?{type:"Google Form",icon:"\u{1F7E3}",badgeClass:"badge-form",color:"#673AB7"}:t.includes("docs.google.com/spreadsheets")||t.includes("sheet")?{type:"Google Sheets",icon:"\u{1F7E2}",badgeClass:"badge-sheet",color:"#0F9D58"}:t.includes("docs.google.com/document")||t.includes("doc")?{type:"Google Docs",icon:"\u{1F535}",badgeClass:"badge-doc",color:"#4285F4"}:t.includes("drive.google.com")?{type:"Google Drive",icon:"\u{1F7E1}",badgeClass:"badge-drive",color:"#F4B400"}:t.endsWith(".pdf")?{type:"PDF Document",icon:"\u{1F534}",badgeClass:"badge-pdf",color:"#EA4335"}:{type:"Web Link",icon:"\u{1F310}",badgeClass:"badge-link",color:"#2563EB"}}async function Ia(e){let t=ae(),n=X(t,"checklist","write"),r=X(t,"checklist","delete"),i=[],m="all",f="",u="fcms-master-doc-styles";if(!document.getElementById(u)){let _=document.createElement("style");_.id=u,_.textContent=`
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
        font-size: 1.45rem;
        font-weight: 800;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 10px;
        letter-spacing: -0.02em;
      }
      .doc-hero-actions {
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
      }
      .doc-hero-actions .btn {
        height: 38px;
        box-sizing: border-box;
        transition: all 0.2s ease;
      }
      #doc-btn-add:hover {
        background: #1D4ED8 !important;
        transform: translateY(-1px);
      }
      #doc-btn-excel:hover, #doc-btn-refresh:hover {
        background: rgba(255, 255, 255, 0.16) !important;
        border-color: rgba(255, 255, 255, 0.3) !important;
        transform: translateY(-1px);
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
        .doc-hero-actions {
          width: 100% !important;
          display: flex !important;
          align-items: center !important;
          gap: 6px !important;
          flex-wrap: nowrap !important;
          overflow-x: auto !important;
          scrollbar-width: none !important;
          padding-bottom: 2px !important;
        }
        .doc-hero-actions::-webkit-scrollbar {
          display: none;
        }
        #doc-btn-add {
          flex: 1 1 auto !important;
          justify-content: center !important;
          font-size: 0.82rem !important;
          padding: 8px 10px !important;
          white-space: nowrap !important;
        }
        #doc-btn-excel {
          flex: 1 1 auto !important;
          justify-content: center !important;
          font-size: 0.82rem !important;
          padding: 8px 10px !important;
          white-space: nowrap !important;
        }
        #doc-btn-refresh {
          width: 38px !important;
          height: 38px !important;
          padding: 0 !important;
          flex-shrink: 0 !important;
          justify-content: center !important;
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
          grid-template-columns: minmax(0, 1fr) !important;
          width: 100% !important;
          max-width: 100% !important;
          min-width: 0 !important;
          box-sizing: border-box !important;
          gap: 12px !important;
        }
        .doc-card {
          padding: 16px !important;
          border-radius: 14px !important;
          width: 100% !important;
          max-width: 100% !important;
          min-width: 0 !important;
          box-sizing: border-box !important;
          overflow: hidden;
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
    `,document.head.appendChild(_)}async function o(){try{let _=await q("/api/checklist?limit=1000");if(_.ok&&_.data?.data&&Array.isArray(_.data.data)&&_.data.data.length>0)i=_.data.data,localStorage.setItem("fcms_master_doc_cache",JSON.stringify(i));else{let g=localStorage.getItem("fcms_master_doc_cache");g?i=JSON.parse(g):(i=[...Da],localStorage.setItem("fcms_master_doc_cache",JSON.stringify(i)))}}catch(_){console.warn("Fallback loading documents from cache/seed:",_);let g=localStorage.getItem("fcms_master_doc_cache");i=g?JSON.parse(g):[...Da]}T()}function p(){return i.filter(_=>{let g=m==="all"||_.category===m,A=f.toLowerCase().trim(),s=!A||(_.name||"").toLowerCase().includes(A)||(_.category||"").toLowerCase().includes(A)||(_.description||"").toLowerCase().includes(A);return g&&s})}function $(){let _={all:i.length};return i.forEach(g=>{let A=g.category||"Lainnya";_[A]=(_[A]||0)+1}),_}function F(_=null){if(!n)return;let g=!!_,A=Array.from(new Set(i.map(l=>l.category).filter(Boolean))),s=document.createElement("div");s.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Nama Dokumen / Checklist <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Checklist Pembersihan Wastafel & Toilet" value="${_?.name?String(_.name).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Kategori Dokumen <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-category" list="doc-categories-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Pilih atau ketik kategori baru..." value="${_?.category?String(_.category).replace(/"/g,"&quot;"):""}" required />
          <datalist id="doc-categories-list">
            ${A.map(l=>`<option value="${l}">${l}</option>`).join("")}
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
            placeholder="https://docs.google.com/..." value="${_?.document_link?String(_.document_link).replace(/"/g,"&quot;"):""}" />
          <span style="font-size:0.75rem; color:#64748B; margin-top:4px; display:block;">
            Bisa berupa tautan Google Forms untuk inspeksi harian, atau link Google Sheets / Docs sebagai acuan.
          </span>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Deskripsi / Petunjuk Penggunaan
          </label>
          <textarea id="doc-modal-desc" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan keterangan singkat mengenai tujuan dan instruksi pengisian dokumen ini...">${_?.description||""}</textarea>
        </div>
      </div>
    `,Z({title:g?"\u270F\uFE0F Edit Master Dokumen":"\u2795 Tambah Master Dokumen Baru",content:s,size:"md",confirmText:g?"Simpan Perubahan":"Tambahkan Dokumen",confirmClass:"btn-primary",onConfirm:async(l,c)=>{let v=s.querySelector("#doc-modal-name"),x=s.querySelector("#doc-modal-category"),w=s.querySelector("#doc-modal-link"),a=s.querySelector("#doc-modal-desc"),d=(v.value||"").trim(),y=(x.value||"").trim(),k=(w.value||"").trim(),M=(a.value||"").trim();if(!d){N("Nama Dokumen wajib diisi"),v.focus();return}if(!y){N("Kategori Dokumen wajib diisi"),x.focus();return}let I={name:d,category:y,document_link:k,description:M};try{if(g)if((await q(`/api/checklist/${_.id}`,{method:"PUT",body:JSON.stringify(I)})).ok){let B=i.findIndex(P=>P.id===_.id);B!==-1&&(i[B]={...i[B],...I}),H("Dokumen berhasil diperbarui")}else{let B=i.findIndex(P=>P.id===_.id);B!==-1&&(i[B]={...i[B],...I}),H("Dokumen disimpan secara lokal")}else{let K=await q("/api/checklist",{method:"POST",body:JSON.stringify(I)});if(K.ok&&K.data?.id)i.unshift({id:K.data.id,...I}),H("Dokumen baru berhasil ditambahkan");else{let B=Date.now();i.unshift({id:B,...I}),H("Dokumen baru ditambahkan secara lokal")}}localStorage.setItem("fcms_master_doc_cache",JSON.stringify(i)),c(),T()}catch(K){console.error("Error saving document:",K),N("Gagal menyimpan dokumen: "+K.message)}}})}function E(_){re(`Apakah Anda yakin ingin menghapus dokumen <strong>"${_.name}"</strong>?`,async()=>{try{let g=await q(`/api/checklist/${_.id}`,{method:"DELETE"});g.ok||console.warn("API delete warning:",g),i=i.filter(A=>A.id!==_.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(i)),H(`Dokumen "${_.name}" berhasil dihapus`),T()}catch{i=i.filter(A=>A.id!==_.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(i)),H(`Dokumen "${_.name}" berhasil dihapus`),T()}},"Hapus Dokumen")}function C(_,g){if(!_){xt("Dokumen belum memiliki tautan digital.");return}navigator.clipboard.writeText(_).then(()=>{H(`Tautan "${g}" disalin ke clipboard!`)}).catch(()=>{xt(`Tautan: ${_}`)})}async function j(){if(!X(t,"export")){N("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:_}=await Promise.resolve().then(()=>(pe(),fe)),g=i.map((A,s)=>({No:s+1,"Nama Dokumen":A.name||"",Kategori:A.category||"","Platform / Format":La(A.document_link).type,"Link Dokumen":A.document_link||"","Deskripsi / Keterangan":A.description||""}));_(g,`Master_Dokumen_FCMS_${new Date().toISOString().slice(0,10)}`),H("Berhasil mengunduh data dokumen ke Excel.")}catch(_){N("Gagal mengekspor data: "+_.message)}}function T(){let _=p(),g=$(),A=Object.keys(g).filter(l=>l!=="all"),s=i.filter(l=>!!l.document_link).length;e.innerHTML=`
      <div class="doc-page-container">
        
        <!-- Hero Header -->
        <div class="doc-hero-card">
          <div class="doc-hero-header">
            <div>
              <h1 class="doc-hero-title">
                <span>\u{1F4CB}</span> Master Document & Template Standar
              </h1>
            </div>
            <div class="doc-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${n?`
              <button id="doc-btn-add" class="btn btn-primary" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap; border:none;">
                <span>\u2795</span>
                <span>Tambah Dokumen Baru</span>
              </button>
              `:""}
              ${X(t,"export")?`<button id="doc-btn-excel" class="btn btn-ghost" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}
              <button id="doc-btn-refresh" class="btn btn-ghost" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <span>\u{1F504}</span>
              </button>
            </div>
          </div>

          <!-- Quick Metrics -->
          <div class="doc-stats-grid">
            <div class="doc-stat-box">
              <div class="doc-stat-icon">\u{1F4C4}</div>
              <div>
                <div class="doc-stat-value">${i.length}</div>
                <div class="doc-stat-label">Total Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon">\u{1F5C2}\uFE0F</div>
              <div>
                <div class="doc-stat-value">${A.length}</div>
                <div class="doc-stat-label">Kategori Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon" style="background:rgba(16,185,129,0.18); color:#34D399;">\u{1F310}</div>
              <div>
                <div class="doc-stat-value">${s} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">/ ${i.length}</span></div>
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
              <input type="text" id="doc-search-input" class="doc-search-input" placeholder="Cari nama dokumen, kategori, atau deskripsi..." value="${f?f.replace(/"/g,"&quot;"):""}" />
            </div>
          </div>

          <!-- Category Pills -->
          <div class="doc-filter-pills">
            <button class="doc-pill ${m==="all"?"active":""}" data-cat="all">
              <span>Semua Dokumen</span>
              <span class="doc-pill-count">${g.all||0}</span>
            </button>
            ${A.map(l=>{let c=Pa(l);return`
                <button class="doc-pill ${m===l?"active":""}" data-cat="${l}">
                  <span>${c.icon}</span>
                  <span>${l}</span>
                  <span class="doc-pill-count">${g[l]||0}</span>
                </button>
              `}).join("")}
          </div>
        </div>

        <!-- Documents View Container -->
        <div id="doc-list-view">
          ${_.length===0?`
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
              ${_.map(l=>{let c=Pa(l.category),v=La(l.document_link);return`
                  <div class="doc-card" data-id="${l.id}">
                    <div>
                      <div class="doc-card-top">
                        <span class="doc-category-badge" style="background:${c.bg}; color:${c.color}; border:1px solid ${c.border};">
                          <span>${c.icon}</span>
                          <span>${c.label}</span>
                        </span>
                        <span class="doc-platform-badge">
                          <span>${v.icon}</span>
                          <span>${v.type}</span>
                        </span>
                      </div>
                      
                      <h3 class="doc-card-title">${l.name}</h3>
                      <p class="doc-card-desc">${l.description||"Formulir acuan standar operasional dan administrasi fasilitas."}</p>
                    </div>

                    <div class="doc-card-footer">
                      ${l.document_link?`
                        <a href="${l.document_link}" target="_blank" rel="noopener noreferrer" class="doc-btn-open">
                          <span>Buka Dokumen</span>
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                        </a>
                      `:`
                        <button class="doc-btn-open disabled" title="Belum ada tautan yang dikaitkan">
                          <span>Tidak Ada Link</span>
                        </button>
                      `}

                      <div class="doc-card-actions">
                        ${l.document_link?`
                          <button class="doc-icon-btn copy-link-btn" title="Salin Tautan Dokumen" data-link="${l.document_link}" data-title="${l.name}">
                            <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                          </button>
                        `:""}
                        ${n?`
                        <button class="doc-icon-btn edit-doc-btn" title="Edit Dokumen" data-id="${l.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        `:""}
                        ${r?`
                        <button class="doc-icon-btn delete delete-doc-btn" title="Hapus Dokumen" data-id="${l.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                        </button>
                        `:""}
                      </div>
                    </div>
                  </div>
                `}).join("")}
            </div>
          `}
        </div>

      </div>
    `,D()}function D(){let _=e.querySelector("#doc-btn-add");_&&(_.onclick=()=>F());let g=e.querySelector("#doc-btn-excel");g&&(g.onclick=()=>j());let A=e.querySelector("#doc-btn-refresh");A&&(A.onclick=()=>o());let s=e.querySelector("#doc-btn-reset-filter");s&&(s.onclick=()=>{f="",m="all",T()});let l=e.querySelector("#doc-search-input");l&&(l.oninput=c=>{f=c.target.value;let v=p();T();let x=e.querySelector("#doc-search-input");x&&(x.focus(),x.setSelectionRange(x.value.length,x.value.length))}),e.querySelectorAll(".doc-pill").forEach(c=>{c.onclick=()=>{m=c.getAttribute("data-cat")||"all",T()}}),e.querySelectorAll(".copy-link-btn").forEach(c=>{c.onclick=v=>{v.stopPropagation();let x=c.getAttribute("data-link"),w=c.getAttribute("data-title");C(x,w)}}),e.querySelectorAll(".edit-doc-btn").forEach(c=>{c.onclick=v=>{v.stopPropagation();let x=Number(c.getAttribute("data-id")),w=i.find(a=>a.id===x);w&&F(w)}}),e.querySelectorAll(".delete-doc-btn").forEach(c=>{c.onclick=v=>{v.stopPropagation();let x=Number(c.getAttribute("data-id")),w=i.find(a=>a.id===x);w&&E(w)}})}await o()}ge();pe();var qe={"Standar Umum":{icon:"\u{1F4CC}",bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE",rooms:["Lantai & Sudut Ruangan","Dinding, Partisi & Plint","Kaca, Cermin & Partisi Kaca","Sanitair & Area Basah Toilet","Plafon, Lampu & Kisi Ventilasi","Mebel, Meja & Kursi Kerja/Pasien","Peralatan Khusus & Medis Non-Steril","Wadah Sampah & Utilitas"]},"Area Depan & Pelayanan":{icon:"\u{1F6CB}\uFE0F",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",rooms:["Teras","Lobby 1","Receptionist","Snack Corner"]},"Area Klinis & Medis":{icon:"\u{1FA7A}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD",rooms:["Ruang Tindakan","Ronsen","Ruang Dokter","Ruang Steril"]},"Toilet & Sanitasi":{icon:"\u{1F6BF}",bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4",rooms:["Toilet 1","Janitor","Ruang Limbah"]},Sirkulasi:{icon:"\u{1F6B6}",bg:"#F1F5F9",color:"#334155",border:"#CBD5E1",rooms:["Area Tangga","Lift","Area Tangga Exit"]},"Area Anak & Fasilitas":{icon:"\u{1F9F8}",bg:"#FEF3C7",color:"#92400E",border:"#FDE68A",rooms:["Lobby 2","Playground","Musholla","Tempat Wudhu","Ruang Team"]},"Gudang & Teknis":{icon:"\u2699\uFE0F",bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA",rooms:["Gudang Farmasi & Bahan","RoofTop","Ruang Kompressor","Area Genset"]},Eksterior:{icon:"\u{1F333}",bg:"#DCFCE7",color:"#15803D",border:"#86EFAC",rooms:["Taman","Pos Security","Toilet Security","Halaman Parkir"]}};function It(e){for(let[t,n]of Object.entries(qe))if(n.rooms.includes(e))return{zoneName:t,...n};return{zoneName:"Lainnya",icon:"\u{1F4CD}",bg:"#F1F5F9",color:"#475569",border:"#E2E8F0"}}async function za(e){let t=ae(),n=t&&(t.role==="superadmin"||t.role==="admin"),r=X(t,"export");e.innerHTML=`
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
        ${r?`
        <button id="btn-export-hygiene" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          \u{1F4E5} Export Excel
        </button>
        `:""}
        ${n?`
        <button id="btn-add-hygiene" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          + Tambah Parameter
        </button>
        `:""}
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
        ${Object.entries(qe).map(([s,l])=>`
          <button class="zone-tab-btn" data-zone="${s}" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
            <span>${l.icon}</span> ${s}
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
  `;let i="all",m="table",f=[];async function u(){try{let s=await q("/api/hygiene-standards?all=1");if(!s.ok)throw new Error("Gagal memuat standar hygiene");f=Array.isArray(s.data)?s.data:s.data?.data||[],p(),o(),F()}catch(s){console.error(s),N(s.message||"Terjadi kesalahan memuat data");let l=e.querySelector("#hygiene-view-container");l&&(l.innerHTML=`
          <div class="empty-state" style="text-align:center; padding:40px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:12px;">\u26A0\uFE0F</div>
            <h3 style="margin:0 0 8px 0; color:var(--text-1);">Gagal Memuat Data</h3>
            <p style="color:var(--text-3); margin-bottom:16px;">${s.message}</p>
            <button id="btn-retry-hygiene" class="btn btn-primary" style="padding:8px 16px; border-radius:8px;">Coba Lagi</button>
          </div>
        `,l.querySelector("#btn-retry-hygiene")?.addEventListener("click",u))}}function o(){let s=e.querySelector("#badge-count-all");s&&(s.textContent=f.length),Object.entries(qe).forEach(([v,x])=>{let w=v.replace(/[\s&]+/g,"-"),a=e.querySelector(`#badge-count-${w}`);if(a){let d=f.filter(y=>x.rooms.includes(y.room_name)).length;a.textContent=d}});let l=new Set(f.map(v=>v.room_name)),c=e.querySelector("#card-total-rooms");c&&(c.textContent=l.size)}function p(){let s=e.querySelector("#hygiene-room-filter");if(!s)return;let l=s.value,c=[];if(i==="all")c=Array.from(new Set(f.map(v=>v.room_name))).filter(Boolean).sort();else{let v=qe[i];v&&(c=v.rooms.filter(x=>f.some(w=>w.room_name===x)))}s.innerHTML=`<option value="">Semua Ruangan (${c.length})</option>`+c.map(v=>`<option value="${v}" ${v===l?"selected":""}>${v}</option>`).join("")}function $(){let s=(e.querySelector("#hygiene-search")?.value||"").toLowerCase().trim(),l=e.querySelector("#hygiene-room-filter")?.value||"";return f.filter(c=>{if(i!=="all"){let v=qe[i];if(!v||!v.rooms.includes(c.room_name))return!1}if(l&&c.room_name!==l)return!1;if(s){let v=(c.room_name||"").toLowerCase(),x=(c.item_name||"").toLowerCase(),w=(c.cleanliness_standard||"").toLowerCase();if(!v.includes(s)&&!x.includes(s)&&!w.includes(s))return!1}return!0})}function F(){let s=$(),l=e.querySelector("#hygiene-stats-text");l&&(l.innerHTML=`Menampilkan <strong>${s.length}</strong> dari <strong>${f.length}</strong> parameter`);let c=e.querySelector("#hygiene-view-container");if(c){if(s.length===0){c.innerHTML=`
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
      `,c.querySelector("#btn-clear-search-empty")?.addEventListener("click",()=>{let v=e.querySelector("#hygiene-search"),x=e.querySelector("#hygiene-room-filter");v&&(v.value=""),x&&(x.value=""),i="all",e.querySelectorAll(".zone-tab-btn").forEach(w=>w.classList.toggle("active",w.dataset.zone==="all")),e.querySelectorAll(".zone-tab-btn").forEach(w=>{w.dataset.zone==="all"?(w.style.background="#2563EB",w.style.color="#fff"):(w.style.background="#F1F5F9",w.style.color="var(--text-2)")}),p(),F()});return}m==="table"?E(c,s):C(c,s)}}function E(s,l){s.innerHTML=`
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse:collapse; text-align:left; font-size:0.9rem;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:2px solid #E2E8F0; color:#475569; font-size:0.78rem; text-transform:uppercase; letter-spacing:0.6px;">
                <th style="padding:14px 16px; width:50px; text-align:center;">No</th>
                <th style="padding:14px 16px; width:220px;">Ruangan & Zona</th>
                <th style="padding:14px 16px; width:280px;">Item / Objek Inspeksi</th>
                <th style="padding:14px 16px;">Standar Kebersihan Fisik</th>
                ${n?'<th style="padding:14px 16px; width:100px; text-align:center;">Aksi</th>':""}
              </tr>
            </thead>
            <tbody>
              ${l.map((c,v)=>{let x=It(c.room_name);return`
                  <tr style="border-bottom:1px solid #F1F5F9; transition:background-color 0.15s ease;" onmouseover="this.style.backgroundColor='#F8FAFC'" onmouseout="this.style.backgroundColor='transparent'">
                    <td style="padding:14px 16px; text-align:center; color:var(--text-3); font-weight:600; font-size:0.85rem;">
                      ${v+1}
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="margin-bottom:4px;">
                        <span style="display:inline-flex; align-items:center; gap:6px; background:${x.bg}; color:${x.color}; border:1px solid ${x.border}; padding:4px 10px; border-radius:6px; font-weight:700; font-size:0.83rem;">
                          <span>${x.icon}</span> ${c.room_name}
                        </span>
                      </div>
                      <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">
                        ${x.zoneName}
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
                    ${n?`
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
    `,j(s)}function C(s,l){let c={};l.forEach(v=>{let x=It(v.room_name);c[x.zoneName]||(c[x.zoneName]={info:x,rooms:{}}),c[x.zoneName].rooms[v.room_name]||(c[x.zoneName].rooms[v.room_name]=[]),c[x.zoneName].rooms[v.room_name].push(v)}),s.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:24px;">
        ${Object.entries(c).map(([v,x])=>{let w=Object.values(x.rooms).reduce((a,d)=>a+d.length,0);return`
            <div class="card" style="background:#fff; border:1px solid ${x.info.border}; border-radius:14px; overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.03);">
              <!-- Zone Card Header -->
              <div style="background:${x.info.bg}; border-bottom:1px solid ${x.info.border}; padding:14px 20px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="font-size:1.4rem;">${x.info.icon}</span>
                  <div>
                    <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:${x.info.color};">
                      ${v}
                    </h3>
                    <div style="font-size:0.78rem; color:${x.info.color}; opacity:0.85;">
                      ${Object.keys(x.rooms).length} Ruangan \xB7 ${w} Parameter Standar
                    </div>
                  </div>
                </div>
                <span class="badge" style="background:#fff; color:${x.info.color}; border:1px solid ${x.info.border}; font-weight:700; padding:4px 10px; border-radius:12px; font-size:0.8rem;">
                  ${w} Parameter
                </span>
              </div>

              <!-- Zone Rooms List -->
              <div style="padding:16px 20px;">
                ${Object.entries(x.rooms).map(([a,d])=>`
                  <div style="margin-bottom:20px; last-child:margin-bottom:0;">
                    <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                      <h4 style="margin:0; font-size:0.95rem; font-weight:700; color:var(--text-1);">
                        \u{1F6AA} ${a}
                      </h4>
                      <span class="badge badge-sm" style="background:#F1F5F9; color:var(--text-3); font-size:0.75rem;">
                        ${d.length} item
                      </span>
                    </div>

                    <div style="border:1px solid #E2E8F0; border-radius:8px; overflow:hidden;">
                      <table style="width:100%; border-collapse:collapse; font-size:0.875rem;">
                        <thead>
                          <tr style="background:#F8FAFC; border-bottom:1px solid #E2E8F0; color:#475569; font-size:0.75rem; text-transform:uppercase;">
                            <th style="padding:8px 12px; width:40px; text-align:center;">No</th>
                            <th style="padding:8px 12px; width:260px;">Item / Objek Inspeksi</th>
                            <th style="padding:8px 12px;">Standar Kebersihan Fisik</th>
                            ${n?'<th style="padding:8px 12px; width:80px; text-align:center;">Aksi</th>':""}
                          </tr>
                        </thead>
                        <tbody>
                          ${d.map((y,k)=>`
                            <tr style="border-bottom:1px solid #F1F5F9;">
                              <td style="padding:10px 12px; text-align:center; color:var(--text-3); font-size:0.8rem;">
                                ${k+1}
                              </td>
                              <td style="padding:10px 12px; font-weight:600; color:var(--text-1); vertical-align:top;">
                                ${y.item_name}
                              </td>
                              <td style="padding:10px 12px; vertical-align:top; color:#334155;">
                                <div style="display:flex; align-items:flex-start; gap:6px;">
                                  <span style="color:#10B981; font-weight:700;">\u2713</span>
                                  <span>${y.cleanliness_standard}</span>
                                </div>
                              </td>
                              ${n?`
                                <td style="padding:10px 12px; text-align:center; vertical-align:middle;">
                                  <div style="display:inline-flex; gap:4px;">
                                    <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${y.id}" title="Edit" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                                    <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${y.id}" data-name="${y.item_name}" title="Hapus" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
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
    `,j(s)}function j(s){s.querySelectorAll(".btn-edit-item").forEach(l=>{l.addEventListener("click",()=>{let c=l.dataset.id,v=f.find(x=>String(x.id)===String(c));v&&T(v)})}),s.querySelectorAll(".btn-delete-item").forEach(l=>{l.addEventListener("click",()=>{let c=l.dataset.id,v=l.dataset.name;re(`Yakin ingin menghapus parameter standar <strong>"${v}"</strong>?`,async()=>{try{let x=await q(`/api/hygiene-standards/${c}`,{method:"DELETE"});if(!x.ok)throw new Error(x.data?.error||"Gagal menghapus parameter");H("Parameter kebersihan berhasil dihapus"),u()}catch(x){N(x.message)}},"Hapus Parameter Standar")})})}function T(s=null){if(!n){N("Anda tidak memiliki izin untuk menambah atau mengedit parameter kebersihan.");return}let l=!!s,c=Array.from(new Set(f.map(x=>x.room_name))).filter(Boolean).sort(),v=document.createElement("div");v.innerHTML=`
      <form id="form-hygiene-item" style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Ruangan / Lokasi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-room-name" list="room-suggestions" class="form-control" placeholder="Contoh: Ruang Tindakan, Lobby 1, Toilet 1..." value="${s?.room_name||""}" required style="width:100%; height:40px; border-radius:8px;">
          <datalist id="room-suggestions">
            ${c.map(x=>`<option value="${x}">`).join("")}
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
    `,Z({title:l?"\u270F\uFE0F Edit Parameter Kebersihan":"\u2795 Tambah Parameter Kebersihan",content:v,confirmText:l?"Simpan Perubahan":"Tambah Parameter",size:"md",onConfirm:async(x,w)=>{let a=v.querySelector("#modal-room-name").value.trim(),d=v.querySelector("#modal-item-name").value.trim(),y=v.querySelector("#modal-standard").value.trim();if(!a||!d||!y){N("Semua kolom wajib diisi");return}try{let k={room_name:a,item_name:d,cleanliness_standard:y},M=l?`/api/hygiene-standards/${s.id}`:"/api/hygiene-standards",K=await q(M,{method:l?"PUT":"POST",body:JSON.stringify(k)});if(!K.ok)throw new Error(K.data?.error||"Gagal menyimpan parameter");H(l?"Parameter berhasil diperbarui":"Parameter berhasil ditambahkan"),w(),u()}catch(k){N(k.message)}}})}function D(){let s=window.open("","_blank");if(!s){N("Pop-up printer diblokir oleh browser. Izinkan pop-up untuk mencetak SOP.");return}let c=`
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
            <div>Total: <strong>${f.length} Parameter Fisik</strong></div>
          </div>
        </div>

        <div class="title">
          <h1>STANDAR ACUAN KEBERSIHAN FISIK SELURUH CABANG</h1>
          <p>Buku Panduan Kriteria Kebersihan Fisik 8 Zona Area (Cleaning Service & Facility Care)</p>
        </div>

        ${Object.entries(qe).map(([v,x])=>{let w=f.filter(a=>x.rooms.includes(a.room_name));return w.length===0?"":`
            <div class="zone-title">${x.icon} ${v.toUpperCase()} (${w.length} PARAMETER)</div>
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
                ${w.map((a,d)=>`
                  <tr>
                    <td class="td-no">${d+1}</td>
                    <td class="td-room">${a.room_name}</td>
                    <td class="td-obj">${a.item_name}</td>
                    <td>${a.cleanliness_standard}</td>
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
    `;s.document.open(),s.document.write(c),s.document.close()}function _(){let s=f.map((c,v)=>{let x=It(c.room_name);return{No:v+1,"Kategori Zona":x.zoneName,"Ruangan / Lokasi":c.room_name,"Item / Objek Inspeksi":c.item_name,"Standar Kebersihan Fisik":c.cleanliness_standard}}),l=new Date().toISOString().slice(0,10);ne(s,`Master_Standar_Kebersihan_FCMS_${l}`),H("Data master kebersihan berhasil diekspor ke Excel")}e.querySelector("#hygiene-search")?.addEventListener("input",()=>F()),e.querySelector("#hygiene-room-filter")?.addEventListener("change",()=>F()),e.querySelector("#btn-reset-filters")?.addEventListener("click",()=>{let s=e.querySelector("#hygiene-search"),l=e.querySelector("#hygiene-room-filter");s&&(s.value=""),l&&(l.value=""),F()}),e.querySelectorAll(".zone-tab-btn").forEach(s=>{s.addEventListener("click",()=>{i=s.dataset.zone,e.querySelectorAll(".zone-tab-btn").forEach(l=>{let c=l.dataset.zone===i;if(l.classList.toggle("active",c),c){l.style.background="#2563EB",l.style.color="#fff";let v=l.querySelector(".badge");v&&(v.style.background="rgba(255,255,255,0.25)",v.style.color="#fff")}else{l.style.background="#F1F5F9",l.style.color="var(--text-2)";let v=l.querySelector(".badge");v&&(v.style.background="#CBD5E1",v.style.color="#334155")}}),p(),F()})});let g=e.querySelector("#view-table-btn"),A=e.querySelector("#view-zone-btn");g?.addEventListener("click",()=>{m="table",g.style.background="#fff",g.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",A.style.background="transparent",A.style.boxShadow="none",F()}),A?.addEventListener("click",()=>{m="zone",A.style.background="#fff",A.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",g.style.background="transparent",g.style.boxShadow="none",F()}),e.querySelector("#btn-print-hygiene")?.addEventListener("click",D),e.querySelector("#btn-export-hygiene")?.addEventListener("click",_),e.querySelector("#btn-add-hygiene")?.addEventListener("click",()=>T(null)),u()}ge();pe();async function Ma(e,t="supply"){return Un(e)}async function Un(e){let n=((await q("/api/branches?all=1")).data?.data||[]).map(r=>({value:r.id,label:r.full_name}));me({container:e,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:r=>r?new Date(r).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(r,i)=>i.branch_name_ref||i.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:r=>{try{let i=JSON.parse(r);return Array.isArray(i)?i.join(", "):r}catch{return r||"-"}}},{key:"chemical_items",label:"Chemical",render:r=>{try{let i=JSON.parse(r);return Array.isArray(i)?i.join(", "):r}catch{return r||"-"}}},{key:"additional_notes",label:"Catatan",render:r=>r?.length>40?r.slice(0,40)+"\u2026":r||"-"},{key:"status",label:"Status",render:r=>he(r)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:r=>{let i=r?.tools_items;try{i=Array.isArray(JSON.parse(i))?JSON.parse(i).join(", "):i}catch{}let m=r?.chemical_items;try{m=Array.isArray(JSON.parse(m))?JSON.parse(m).join(", "):m}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:r?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:r?.branch_id&&!n.find(f=>f.value==r.branch_id)?[...n,{value:r.branch_id,label:r.branch_name||r.branch_id}]:n,createApi:{path:"/api/branches",field:"full_name"},value:r?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:i},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:r?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:m},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:r?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:r?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:r?.status||""},{name:"processed_by",label:"Diproses Oleh",value:r?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async r=>{let i=new URLSearchParams(r||{}).toString(),m=await q(`/api/reports/supply?limit=10000&${i}`);if(m.ok){let f=m.data.data.map(u=>{let o=u.tools_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}let p=u.chemical_items;try{p=Array.isArray(JSON.parse(p))?JSON.parse(p).join(", "):p}catch{}return{Waktu:u.submitted_at||"",Pengirim:u.submitter_name||"",Cabang:u.branch_name_ref||u.branch_name||"","Alat/Barang":o||"",Chemical:p||"",Catatan:u.additional_notes||"",Status:u.status||"","Diproses Oleh":u.processed_by||""}});ne(f,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async r=>{let m=(await q("/api/branches?all=1")).data?.data||[],f=$=>{if(!$)return null;let F=String($||"").toLowerCase(),E=m.find(C=>String(C.full_name||"").toLowerCase()===F||String(C.code||"").toLowerCase()===F||String(C.name||"").toLowerCase()===F);return E?E.id:null},u=$=>{if($==null||$==="")return"";if($ instanceof Date&&!isNaN($.getTime()))return $.toISOString().slice(0,10);let F=String($).trim();if(F===""||F==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(F))return F.slice(0,10);if(/^\d{4,5}$/.test(F)){let C=Number(F);if(C>2e4&&C<99999){let j=new Date(Date.UTC(1899,11,30)+C*864e5);return isNaN(j.getTime())?"":j.toISOString().slice(0,10)}}let E=F.split(/[\/\-\.]/);if(E.length===3){let[C,j,T]=E.map(D=>D.trim());if(C.length===4&&j.length<=2&&T.length<=2)return`${C}-${j.padStart(2,"0")}-${T.padStart(2,"0")}`;if(T.length===4&&j.length<=2&&C.length<=2)return`${T}-${j.padStart(2,"0")}-${C.padStart(2,"0")}`}return F},o=r.map($=>({submitted_at:u($.Waktu||$.Tanggal),submitter_name:String($.Pengirim||"").trim(),branch_id:f(String($.Cabang||"").trim()),tools_items:String($["Alat/Barang"]||$.Alat||"").trim(),chemical_items:String($.Chemical||"").trim(),additional_notes:String($.Catatan||$.Keterangan||"").trim(),status:String($.Status||"").trim(),processed_by:String($["Diproses Oleh"]||$.PIC||"").trim()})).filter($=>$.submitted_at&&$.submitter_name&&$.branch_id),p=await q("/api/reports/supply/import",{method:"POST",body:JSON.stringify(o)});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(r,i)=>{let m=Z({title:"Update Status Permintaan",content:`
              <div class="form-group">
                <label class="form-label">Status</label>
                <select class="form-control" id="supply-status">
                  <option value="Pending" ${r.status==="Pending"?"selected":""}>Pending</option>
                  <option value="Diproses" ${r.status==="Diproses"?"selected":""}>Diproses</option>
                  <option value="Selesai" ${r.status==="Selesai"?"selected":""}>Selesai</option>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label">Diproses Oleh</label>
                <input type="text" class="form-control" id="supply-processed-by" value="${r.processed_by||""}" placeholder="Nama">
              </div>
            `,onConfirm:async(f,u)=>{let o=f.querySelector("#supply-status").value,p=f.querySelector("#supply-processed-by").value;(await q(`/api/reports/supply/${r.id}`,{method:"PUT",body:JSON.stringify({status:o,processed_by:p})})).ok?(H("Status diperbarui."),u(),i()):N("Gagal update status.")}})}}]})}pe();async function ja(e){let t=ae();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}me({container:e,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:n=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",editor_khusus:"badge-warning",input_lembur:"badge-info",viewer:"badge-neutral"}[n]||"badge-neutral"}">${n}</span>`},{key:"is_active",label:"Status",render:n=>n?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:n=>n?new Date(n).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:n=>{let r=!!n;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:n?.full_name},{name:"username",label:"Username",required:!0,placeholder:"username",value:n?.username,autocomplete:"off"}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!0,placeholder:"email@contoh.com",value:n?.email,autocomplete:"off"},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"editor_khusus",label:"Editor (Hanya Lembur & Masalah)"},{value:"input_lembur",label:"Input Lembur (Hanya Isi Lembur)"},{value:"viewer",label:"Viewer"}],value:n?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:r?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!r,placeholder:"Min. 6 karakter",autocomplete:"new-password"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:r?n?.is_active:1}]}]},onBeforeSubmit:(n,r)=>(n.full_name&&(n.full_name=String(n.full_name).trim()),n.username&&(n.username=String(n.username).trim()),n.email&&(n.email=String(n.email).trim().toLowerCase()),r&&(!n.password||String(n.password).trim()==="")&&delete n.password,n),exportOptions:{moduleName:"users",onExport:async()=>{let n=await q(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let r=n.data.data.map(i=>({"Nama Lengkap":i.full_name||"",Username:i.username||"",Email:i.email||"",Role:i.role||"",Status:i.is_active?"Aktif":"Nonaktif"}));ne(r,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async n=>{let r=n.map(m=>({full_name:String(m["Nama Lengkap"]||"").trim(),username:String(m.Username||"").trim(),email:String(m.Email||"").trim().toLowerCase(),role:String(m.Role||"").trim()||"viewer",password:String(m.Password||"").trim()})).filter(m=>m.username&&m.password&&m.email&&m.full_name),i=await q("/api/users/import",{method:"POST",body:JSON.stringify(r)});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}pe();async function Ka(e){me({container:e,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!1,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:t=>t?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:t=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:t?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:t?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:t?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:t?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:t?.is_active!==void 0?t.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let t=await q(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(t.ok)ne(t.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async t=>{let n=t.map(i=>({code:String(i["Kode Cabang"]||"").trim(),name:String(i["Nama Pendek"]||"").trim(),full_name:String(i["Nama Lengkap"]||"").trim(),city:String(i.Kota||"").trim()})).filter(i=>i.code&&i.name),r=await q("/api/branches/import",{method:"POST",body:JSON.stringify(n)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}async function Na(e){let t=new Date,n=[];e.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{t.setMonth(t.getMonth()-1),i()}),document.getElementById("cal-next").addEventListener("click",()=>{t.setMonth(t.getMonth()+1),i()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(f=>f.addEventListener("change",i));async function r(){try{let f=`${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,"0")}`;n=(await q(`/api/dashboard/calendar?month=${f}`)).data?.data||[]}catch(f){console.warn("[Calendar] Failed to load events, rendering empty grid:",f),n=[]}}async function i(){let f=document.getElementById("calendar-grid");if(f){f.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await r();try{let u=t.getFullYear(),o=t.getMonth(),p=t.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),$=document.getElementById("cal-month-label");$&&($.textContent=p);let F=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(l=>l.value)),E=n.filter(l=>F.has(l.type)),C={};E.forEach(l=>{let c=Gn(l.event_date);c&&(C[c]||(C[c]=[]),C[c].push(l))});let j=new Date(u,o,1).getDay(),T=new Date(u,o+1,0).getDate(),D=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],_=new Date().toISOString().slice(0,10),g='<div class="calendar-grid">';D.forEach(l=>{g+=`<div class="cal-day-header">${l}</div>`});for(let l=0;l<j;l++)g+='<div class="cal-cell cal-cell-empty"></div>';for(let l=1;l<=T;l++){let c=`${u}-${String(o+1).padStart(2,"0")}-${String(l).padStart(2,"0")}`,v=C[c]||[],x=c===_;g+=`
          <div class="cal-cell ${x?"cal-today":""} ${v.length?"cal-has-events":""}"
               data-date="${c}" tabindex="0" role="button" aria-label="${c}">
            <div class="cal-day-num ${x?"today-num":""}">${l}</div>
            <div class="cal-events-preview">
              ${v.slice(0,3).map(w=>`
                <div class="cal-event-dot cal-color-${w.color||"gray"}" title="${Ee(w.title||w.type)}">
                  <span class="cal-event-dot-label">${Jn(w.title||w.branch_name||w.type,18)}</span>
                </div>
              `).join("")}
              ${v.length>3?`<div class="cal-more">+${v.length-3} lagi</div>`:""}
            </div>
          </div>`}let s=(j+T)%7;if(s!==0)for(let l=0;l<7-s;l++)g+='<div class="cal-cell cal-cell-empty"></div>';g+="</div>",f.innerHTML=g,f.querySelectorAll(".cal-cell[data-date]").forEach(l=>{l.addEventListener("click",()=>{let c=l.dataset.date,v=C[c]||[];if(!v.length)return;let x=document.getElementById("cal-event-list"),w=new Date(c+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=w,document.getElementById("cal-event-items").innerHTML=v.map(a=>`
            <div class="cal-event-item cal-color-border-${a.color||"gray"}">
              <div class="cal-event-type">
                ${Yn(a.type)}
                ${a.date_type?` <span class="badge" style="font-size:0.75rem;padding:2px 6px;background:rgba(0,0,0,0.06);margin-left:4px;">${Ee(a.date_type)}</span>`:""}
              </div>
              <div class="cal-event-title" style="font-weight:600;font-size:0.95rem;margin-top:2px;">${Ee(a.title||"-")}</div>
              ${a.branch_name?`<div class="cal-event-branch" style="color:var(--text-2);font-size:0.85rem;margin-top:2px;">\u{1F4CD} ${Ee(a.branch_name)}</div>`:""}
              ${a.pic||a.period||a.opening_date||a.target_date||a.completion_date||a.notes?`
                <div class="cal-event-meta" style="margin-top:6px;padding-top:6px;border-top:1px dashed var(--border);font-size:0.8rem;color:var(--text-2);display:flex;flex-direction:column;gap:3px;">
                  ${a.pic?`<div>\u{1F464} <b>PIC:</b> ${Ee(a.pic)}</div>`:""}
                  ${a.period?`<div>\u{1F4C5} <b>Periode:</b> ${Ee(a.period)}</div>`:""}
                  ${a.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${Ee(zt(a.opening_date))}</div>`:""}
                  ${a.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${Ee(zt(a.target_date))}</div>`:""}
                  ${a.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${Ee(zt(a.completion_date))}</div>`:""}
                  ${a.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${Ee(a.notes)}</div>`:""}
                </div>
              `:""}
              ${a.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${a.status==="Done"?"badge-success":"badge-warning"}">${Ee(a.status)}</span></div>`:""}
              ${a.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${a.days_remaining} hari</div>`:""}
            </div>
          `).join(""),x.style.display="block"})})}catch(u){console.error("[Calendar] Render error:",u),f&&(f.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let m=()=>{r().then(()=>i())};st.on("data:changed",m),i()}function Gn(e){if(!e||e==="-"||String(e).trim()==="")return"";let t=String(e).trim();if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);let n=t.split(/[\/\-\.]/);if(n.length===3){let[r,i,m]=n.map(f=>f.trim());if(m.length===4)return`${m}-${i.padStart(2,"0")}-${r.padStart(2,"0")}`;if(r.length===4)return`${r}-${i.padStart(2,"0")}-${m.padStart(2,"0")}`}return t.slice(0,10)}function zt(e){if(!e||e==="-"||String(e).trim()==="")return"";let n=String(e).trim().split(/[\/\-\.]/);if(n.length===3){let[r,i,m]=n.map(f=>f.trim());if(r.length===4)return`${m.padStart(2,"0")}-${i.padStart(2,"0")}-${r}`;if(m.length===4)return`${r.padStart(2,"0")}-${i.padStart(2,"0")}-${m}`}return e}function Jn(e,t){return e?e.length>t?e.slice(0,t)+"\u2026":e:""}function Ee(e){return e?String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Yn(e){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging"}[e]||e}async function qa(e){let t=ae(),n=(t?.full_name||t?.username||"U")[0].toUpperCase(),i={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[t?.role]||"#64748B";e.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${i},${i}99)">
            ${n}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${t?.full_name||"\u2014"}</div>
            <div class="profile-username">@${t?.username||"\u2014"}</div>
            <span class="badge badge-info" style="background:${i}18;color:${i};margin-top:6px">
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
            <span class="info-value" style="color:${i};font-weight:700">${t?.role||"\u2014"}</span>
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
  `;let m=localStorage.getItem("fm_token"),f=document.getElementById("session-info");if(m&&f)try{let u=JSON.parse(atob(m.split(".")[1])),o=new Date(u.exp*1e3);f.textContent=`Berakhir: ${o.toLocaleString("id-ID")}`}catch{f.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async u=>{u.preventDefault();let o=document.getElementById("pwd-error"),p=document.getElementById("pwd-success"),$=document.getElementById("btn-save-pwd");o.style.display="none",p.style.display="none";let F=u.target,E=F.current_password.value,C=F.new_password.value,j=F.confirm_password.value;if(C!==j){o.textContent="\u274C Konfirmasi password tidak cocok.",o.style.display="block";return}if(C.length<6){o.textContent="\u274C Password baru minimal 6 karakter.",o.style.display="block";return}$.disabled=!0,$.textContent="\u23F3 Menyimpan...";let T=await q("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:E,new_password:C})});$.disabled=!1,$.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',T.ok?(p.textContent="\u2705 Password berhasil diubah.",p.style.display="block",F.reset(),H("Password berhasil diubah.")):(o.textContent=T.data?.error||"Gagal mengubah password.",o.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}var ct={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function ue(e){if(e==null||e==="")return null;if(e instanceof Date)return isNaN(e.getTime())?null:e.toISOString().slice(0,10);let t=String(e).trim();if(t===""||t==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);if(/^\d{4,5}$/.test(t)){let i=Number(t);if(i>2e4&&i<99999){let m=new Date(Date.UTC(1899,11,30)+i*864e5);return isNaN(m.getTime())?null:m.toISOString().slice(0,10)}}let n=t.split(/[\/\-\.]/);if(n.length===3){let[i,m,f]=n.map($=>$.trim()),u=Number(i),o=Number(m),p=Number(f);if(i.length===4&&u>1900)return`${i}-${m.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&p>1900)return u>12?`${f}-${m.padStart(2,"0")}-${i.padStart(2,"0")}`:o>12?`${f}-${i.padStart(2,"0")}-${m.padStart(2,"0")}`:`${f}-${m.padStart(2,"0")}-${i.padStart(2,"0")}`;if(f.length===2&&!isNaN(p)){let $=p>=50?`19${f}`:`20${f}`;return u>12?`${$}-${m.padStart(2,"0")}-${i.padStart(2,"0")}`:`${$}-${m.padStart(2,"0")}-${i.padStart(2,"0")}`}}let r=new Date(t);return isNaN(r.getTime())?null:r.toISOString().slice(0,10)}function Ra(e){return Object.values(e).every(t=>t==null||String(t).trim()==="")}var Vn={validation:{required:[],map:e=>({cabang:e.CABANG,pic:e.PIC,kegiatan:e.KEGIATAN,quartal:e.QUARTAL,masa_pkwt:e["MASA PKWT"],pic_pelapor:e["PIC PELAPOR"],kontrak:e.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:e=>({name:e["Nama SOP"],category:e.Kategori||"Umum",document_link:e["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({full_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||e.Divisi||"FACILITY CARE",phone:e["No. Hp"]||e["No. HP"],join_date:ue(e["Tanggal Masuk"]||e["Tgl Masuk"]),status:e.Status||"",target_pindah_os:ue(e["Target Pindah OS"]),target_selesai_os:ue(e["Target Selesai OS"]),notes:e.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({employee_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||"FACILITY CARE",start_date:ue(e["Tanggal Mulai"]),end_date:ue(e["Tanggal Selesai"]),contract_type:e["Tipe Kontrak"]||"",pkwt_number:e.PKWT||"",status:e.Status||"",notes:e.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:e=>({report_date:ue(e["Tanggal Info"]),branch_name:e.Cabang,category:e.Kategori,source:e["Sumber Laporan"],complaint:e.Keluhan,employee_name:e["Nama FC"],fc_specialist:e["FC Spesialis"],solution:e.Solusi,status:e.Status||"",completion_date:ue(e["Tanggal Selesai"])})},one_on_one:{required:[],map:e=>({meeting_date:ue(e.Tanggal),branch_name:e.Cabang,employee_name:e["Nama Karyawan"],pic:e.Pic,problem:e.Masalah,solution:e.Solusi,status:e.Status||"",completion_date:ue(e["Tanggal Selesai"]),document_link:e["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:e=>({branch_name:e.Cabang,activity_type:e.Kegiatan,period:e.Periode,pic:e.Pic||e.PIC,opening_date:ue(e["Tanggal Opening"]||e["Tgl Opening"]),target_date:ue(e["Tanggal Target"]||e["Tgl Target"]),completion_date:ue(e["Tanggal Selesai"]||e["Tgl Selesai"]),status:e.Status||"",notes:e.Keterangan||e.Catatan})},inspection:{required:[],map:e=>({inspection_date:ue(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",fc_score:e["Point FC SP"]!==void 0&&e["Point FC SP"]!==null?parseFloat(String(e["Point FC SP"]).replace(",",".")):null,spv_score:e["Point SPV"]!==void 0&&e["Point SPV"]!==null?parseFloat(String(e["Point SPV"]).replace(",",".")):null,document_link:e.Link,notes:""})},cleaning:{required:[],map:e=>({activity_date:ue(e.Tanggal),branch_name:e.Cabang,activity_type:e["Jenis Kegiatan"]||"General Cleaning",period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},fogging:{required:[],map:e=>({activity_date:ue(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:e=>({info_date:ue(e["Tgl Info"]),branch_name:e.Cabang,problem:e.Permasalahan,pic:e.PIC,done_date:ue(e["Tgl Done"]),status:e.Status||"",notes:e.Ket})},relievers:{required:[],map:e=>({branch_name:e.Cabang,original_fc_name:e["Nama Facility care"],period:e.Periode,reliever_name:e.Relifer,backup_date:ue(e["Tanggal Back Up"]),completion_date:ue(e["Tanggal Selesai"]),reason:e.Keterangan,shift:e.Shift,status:e.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:e=>({training_date:ue(e.Tanggal),batch:e.Batch,subject:e.Materi,participants:e.Peserta,branch_name:e.Cabang,trainer:e.Trainer,score:e.Nilai!==void 0&&e.Nilai!==null?parseFloat(String(e.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:e=>({name:e["Master Checklist"],category:"Umum",document_link:e["Link Document"],description:""})},supply:{required:[],map:e=>({submitted_at:ue(e.Timestamp),submitter_name:e["Nama Lengkap"],branch_name:e["Kebutuhan Untuk Cabang"],tools_items:e["Alat - Alat / Barang"],tools_quantity:e["Jumlah Permintaan Alat / Barang"],chemical_items:e.Chemical,chemical_quantity:e["Jumlah Permintaan Chemical"],additional_notes:e["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:e.Status||""})}};function Wn(e,t){let n=ct[e];if(!n)return{valid:[],errors:[],mapped:[],skipped:!0};let r=Vn[n.module];if(!r)return{valid:[],errors:[],mapped:[],skipped:!0};let i=[],m=[],f=[];return t.filter(o=>!Ra(o)).forEach((o,p)=>{let $=t.indexOf(o)+2,F=[];r.required.forEach(({key:C,label:j})=>{let T=o[C];if(T==null||String(T).trim()===""){let D=Object.keys(o).filter(_=>_.trim()).join(", ");F.push({column:j,originalValue:T||"",reason:`Kolom "${j}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${D.slice(0,120)}`})}});let E=r.map(o);F.length>0?m.push({row:$,data:E,raw:o,errors:F}):(i.push(o),f.push(E))}),{valid:i,errors:m,mapped:f}}function Oa(e){let t=[];return e.SheetNames.forEach(n=>{let r=ct[n];if(!r)return;let i=e.Sheets[n],m=window.XLSX.utils.sheet_to_json(i,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),f=Wn(n,m),u=m.filter(o=>!Ra(o));t.push({sheetName:n,module:r.module,label:r.label,total:u.length,valid:f.mapped.length,errorCount:f.errors.length,errors:f.errors,mapped:f.mapped,skipped:!1})}),t}function Ha(){let e=window.XLSX,t=e.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([r,i])=>{e.utils.book_append_sheet(t,e.utils.json_to_sheet(i),r)}),e.writeFile(t,"Template_Import_Data_Awal_FCMS.xlsx")}function Ua(e){let t=window.XLSX,n=t.utils.book_new(),r=!1;return e.forEach(i=>{if(!i.errors||i.errors.length===0)return;r=!0;let m=i.errors.map(u=>({"No. Baris":u.row,"Kolom Gagal":(u.errors||[]).map(o=>o.column||o).join("; "),"Alasan Error":(u.errors||[]).map(o=>o.reason||o).join("; "),...Object.fromEntries(Object.entries(u.data||{}).map(([o,p])=>[o,p??""]))})),f=t.utils.json_to_sheet(m);t.utils.book_append_sheet(n,f,i.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),r?(t.writeFile(n,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Qn=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","sop","inspection","cleaning","fogging","basecamp","supply"];function Ga(e){e.innerHTML=`
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
              ${Object.entries(ct).map(([T,{label:D}])=>`<span class="import-sheet-tag">\u{1F4C4} ${T} \u2192 ${D}</span>`).join("")}
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
  `;let t=null,n=null,r=0,i={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function m(T){Object.entries(i).forEach(([D,_])=>{_.style.display=D===T?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let T=document.getElementById("btn-backup-db");T.disabled=!0,T.textContent="\u23F3 Memproses Backup...";try{let D=await q("/api/import/backup");if(D.ok){if(!window.XLSX){N("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let _=window.XLSX,g=_.utils.book_new();Object.entries(D.data.database).forEach(([A,s])=>{let l=s.length>0?s:[{}],c=_.utils.json_to_sheet(l);_.utils.book_append_sheet(g,c,A.substring(0,31))}),_.writeFile(g,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),H("Backup berhasil diunduh!")}else N("Gagal memproses backup: "+(D.data?.error||"Unknown error"))}catch(D){N("Gagal memproses backup: "+D.message)}finally{T.disabled=!1,T.textContent="\u{1F4E6} Backup Database"}});let f=document.getElementById("btn-sync-google");f&&f.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let T=f.innerHTML;f.innerHTML='<span class="spinner"></span> Menyinkronkan...',f.disabled=!0;try{let D=await q("/api/sync/google-sheets",{method:"POST"});D.ok?alert("Sinkronisasi Berhasil: "+(D.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(D.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{f.innerHTML=T,f.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Ha(),H("Template Excel berhasil didownload!")});let u=document.getElementById("file-input"),o=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",T=>{T.stopPropagation(),u.click()}),u.addEventListener("change",T=>{T.target.files[0]&&p(T.target.files[0])}),o.addEventListener("dragover",T=>{T.preventDefault(),o.classList.add("drag-over")}),o.addEventListener("dragleave",()=>o.classList.remove("drag-over")),o.addEventListener("drop",T=>{T.preventDefault(),o.classList.remove("drag-over");let D=T.dataTransfer.files[0];D&&D.name.match(/\.xlsx?$/i)?p(D):N("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{t=null,u.value="",document.getElementById("file-info").style.display="none",o.style.display="",m("upload")});async function p(T){t=T,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${T.name} (${(T.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",o.style.display="none",await $(T)}async function $(T){m("validating");let D=document.getElementById("validation-status"),_=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");D.textContent="Membaca file Excel...",_.style.width="20%",await We(200);let g=await T.arrayBuffer(),A=window.XLSX.read(g,{type:"array",cellDates:!0});D.textContent=`Memvalidasi ${A.SheetNames.length} sheet...`,_.style.width="50%",await We(100),n=Oa(A),_.style.width="100%",D.textContent="Validasi selesai!",await We(300),F()}catch(g){m("upload"),N("Gagal memproses file: "+g.message),document.getElementById("file-info").style.display="flex",o.style.display="none"}}function F(){m("preview");let T=n.filter(x=>!x.skipped).length,D=n.reduce((x,w)=>x+w.total,0),_=n.reduce((x,w)=>x+w.valid,0),g=n.reduce((x,w)=>x+w.errorCount,0),A=D>0?Math.round(_/D*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${T} sheet</span>
      <span class="badge badge-secondary">${D} baris</span>
      <span class="badge badge-success">${_} valid (${A}%)</span>
      ${g>0?`<span class="badge badge-danger">${g} error</span>`:""}
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
          ${n.map((x,w)=>`
            <tr class="${x.errorCount>0?"row-error":x.skipped?"row-skipped":"row-ok"}">
              <td><strong>${x.sheetName}</strong></td>
              <td>${x.label}</td>
              <td style="text-align:center">${x.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${x.valid}</span></td>
              <td style="text-align:center">${x.errorCount>0?`<span class="badge badge-danger">${x.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${x.skipped?'<span class="badge badge-neutral">Dilewati</span>':x.errorCount>0&&x.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':x.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':x.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${x.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${w}">\u{1F50D} ${x.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,s.querySelectorAll(".btn-detail-error").forEach(x=>{x.addEventListener("click",()=>{let w=n[Number(x.dataset.idx)];E(w)})});let l=document.getElementById("error-detail-section"),c=document.getElementById("error-detail-container");c.innerHTML="",l.style.display="none";let v=document.getElementById("btn-start-import");_===0?(v.disabled=!0,v.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(v.disabled=!1,g>0?(v.innerHTML=`\u{1F680} Import ${_} Data Valid (${g} dilewati)`,v.title="Baris error akan dilewati, baris valid tetap diimport"):v.innerHTML=`\u{1F680} Mulai Import ${_} Data`)}function E(T){let D=document.getElementById("error-detail-section"),_=document.getElementById("error-detail-container");D.style.display="";let g=T.errors.slice(0,100).map(A=>(Array.isArray(A.errors)?A.errors:[]).map(l=>{let c=typeof l=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${A.row}</span></td>
            <td><strong>${c?l.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${c&&l.originalValue!==void 0?l.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${c?l.reason:l}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${c&&l.aliases?`Gunakan salah satu nama kolom:<br><em>${l.aliases}</em>`:c&&l.hint?l.hint:""}
            </td>
          </tr>
        `}).join("")).join("");_.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${T.sheetName} \u2014 ${T.errorCount} baris error dari ${T.total} total
          ${T.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
        ${T.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,D.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{m("upload"),document.getElementById("file-info").style.display="none",o.style.display="",t=null,u.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!n)return;Ua(n)?H("Log error berhasil didownload."):H("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let T=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";C(T)});async function C(T){m("importing"),r=Date.now();let D=[];Qn.forEach(l=>{let c=n?.find(v=>v.module===l&&v.mapped?.length>0);c&&D.push(c)});let _=document.getElementById("import-steps-list");_.innerHTML=D.map(l=>`
      <div class="import-step-item" id="step-item-${l.module}">
        <span class="step-item-icon" id="step-icon-${l.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${l.label} <span class="step-item-count">(${l.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${l.module}"></span>
      </div>
    `).join("");let g=document.getElementById("import-bar"),A=document.getElementById("import-current-status"),s={totalSheets:D.length,totalRows:D.reduce((l,c)=>l+c.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let l=0;l<D.length;l++){let c=D[l],v=document.getElementById(`step-icon-${c.module}`),x=document.getElementById(`step-status-${c.module}`);v.textContent="\u{1F504}",x.textContent="Mengimport...",A.textContent=`Mengimport ${c.label}...`,g.style.width=`${Math.round(l/D.length*100)}%`;try{let w=await q(`/api/import/${c.module}`,{method:"POST",body:JSON.stringify({rows:c.mapped,onDuplicate:T})});if(w.ok){let a=w.data;s.inserted+=a.inserted||0,s.skipped+=a.skipped||0,s.moduleResults.push({label:c.label,inserted:a.inserted||0,skipped:a.skipped||0,status:"ok"}),v.textContent="\u2705",x.innerHTML=`<span class="badge badge-success">${a.inserted||0} berhasil</span>${a.skipped>0?` <span class="badge badge-neutral">${a.skipped} skip</span>`:""}`}else s.failed++,s.moduleResults.push({label:c.label,inserted:0,skipped:0,status:"error",error:w.data?.error}),v.textContent="\u274C",x.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(w){s.failed++,s.moduleResults.push({label:c.label,inserted:0,skipped:0,status:"error",error:w.message}),v.textContent="\u274C",x.innerHTML='<span class="badge badge-danger">Gagal</span>'}await We(150)}g.style.width="100%",A.textContent="Selesai!",await We(400),j(s)}function j(T){m("summary");let D=((Date.now()-r)/1e3).toFixed(1),_=T.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${_?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${_?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${T.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${T.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${T.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${T.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${T.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${T.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
      <div class="summary-stat-card info">
        <div class="stat-value">${D}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${T.moduleResults.map(g=>`
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{t=null,n=null,u.value="",document.getElementById("file-info").style.display="none",o.style.display="",m("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function We(e){return new Promise(t=>setTimeout(t,e))}ge();var Ja=[{id:1,employee_name:"Teguh Ariyandi",division:"FACILITY CARE",branch_id:18,branch_name:"018. Harapan Indah",tanggal:"2025-08-19",akhir_sp:"2026-02-18",sp_type:"SP 1",status:"Aktif",document_link:"",notes:"Keterlambatan berulang dan ketidaklengkapan checklist harian lantai dasar."},{id:2,employee_name:"Andri",division:"FACILITY CARE",branch_id:1,branch_name:"001. Pondok Bambu",tanggal:"2025-04-10",akhir_sp:"2025-10-09",sp_type:"SP 1",status:"Selesai",document_link:"",notes:"Pelanggaran grooming dan standar pembersihan area tindakan tanpa APD lengkap."},{id:3,employee_name:"M. FIRMANSYAH",division:"FACILITY CARE",branch_id:25,branch_name:"025. Bogor Baru",tanggal:"2024-03-21",akhir_sp:"2024-09-20",sp_type:"SP 2",status:"Selesai",document_link:"",notes:"Meninggalkan area tugas tanpa konfirmasi supervisor selama jam operasional klinik."}];function De(e){if(!e)return{status:"Tidak Ditentukan",label:"Tidak Ada Tanggal",bg:"#F1F5F9",color:"#64748B",border:"#CBD5E1",icon:"\u26AA",daysLeft:0,isExpired:!1,isExpiringSoon:!1};let t=new Date;t.setHours(0,0,0,0);let n=new Date(e);n.setHours(0,0,0,0);let r=n.getTime()-t.getTime(),i=Math.ceil(r/(1e3*60*60*24));return i<0?{status:"Selesai",label:`Masa Sanksi Selesai (${Math.abs(i)} hari lalu)`,shortLabel:"Selesai",bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2705",daysLeft:i,isExpired:!0,isExpiringSoon:!1}:i<=30?{status:"Mendekati Berakhir",label:`Segera Berakhir (Sisa ${i} hari)`,shortLabel:`Sisa ${i} hari`,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F",daysLeft:i,isExpired:!1,isExpiringSoon:!0}:{status:"Aktif",label:`Aktif Berlaku (Sisa ${i} hari)`,shortLabel:`Aktif (${i} hari)`,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F534}",daysLeft:i,isExpired:!1,isExpiringSoon:!1}}function Mt(e=""){let t=(e||"").toUpperCase().trim();return t==="SP 3"?{label:"SP 3 (Terakhir)",bg:"#7F1D1D",color:"#FFFFFF",border:"#991B1B",icon:"\u{1F6A8}"}:t==="SP 2"?{label:"SP 2",bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F6D1}"}:t==="SP 1"?{label:"SP 1",bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F"}:t.includes("TEGURAN")?{label:"Teguran Lisan",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F4AC}"}:{label:e||"SP",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4DC}"}}function Re(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Ya(e){let t=ae(),n=X(t,"sp","write"),r=X(t,"sp","write"),i=X(t,"sp","delete"),m=[],f=[],u=[],o="all",p="all",$="all",F="",E=localStorage.getItem("fcms_sp_view_mode")||"card",C="fcms-sp-page-styles",j=document.getElementById(C);j||(j=document.createElement("style"),j.id=C,document.head.appendChild(j)),j.textContent=`
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
        font-size: 1.45rem;
        font-weight: 800;
        margin: 0;
        display: flex;
        align-items: center;
        gap: 10px;
        letter-spacing: -0.02em;
      }
      .sp-hero-actions {
        display: flex;
        align-items: center;
        gap: 8px;
        flex-wrap: wrap;
      }
      .sp-hero-actions .btn {
        height: 38px;
        box-sizing: border-box;
        transition: all 0.2s ease;
      }
      .sp-btn-add:hover {
        background: #1D4ED8 !important;
        transform: translateY(-1px);
      }
      .sp-btn-excel:hover, .sp-btn-refresh:hover {
        background: rgba(255, 255, 255, 0.16) !important;
        border-color: rgba(255, 255, 255, 0.3) !important;
        transform: translateY(-1px);
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
      /* View Mode Toggle */
      .sp-view-toggle {
        display: inline-flex;
        align-items: center;
        background: #F1F5F9;
        padding: 3px;
        border-radius: 10px;
        border: 1px solid #CBD5E1;
        gap: 3px;
        flex-shrink: 0;
      }
      .sp-view-btn {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 6px 12px;
        border-radius: 7px;
        font-size: 0.82rem;
        font-weight: 600;
        border: none;
        background: transparent;
        color: #64748B;
        cursor: pointer;
        transition: all 0.15s ease;
        height: 34px;
        box-sizing: border-box;
        user-select: none;
      }
      .sp-view-btn:hover {
        color: #0F172A;
        background: rgba(255, 255, 255, 0.5);
      }
      .sp-view-btn.active {
        background: #FFFFFF;
        color: #2563EB;
        font-weight: 700;
        box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
      }
      /* Modern Table View */
      .sp-table-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        overflow: hidden;
      }
      .sp-table-responsive {
        width: 100%;
        overflow-x: auto;
        -webkit-overflow-scrolling: touch;
      }
      .sp-table {
        width: 100%;
        border-collapse: collapse;
        text-align: left;
        min-width: 900px;
      }
      .sp-table th {
        background: #F8FAFC;
        padding: 13px 16px;
        font-size: 0.76rem;
        font-weight: 700;
        color: #475569;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        border-bottom: 1.5px solid #E2E8F0;
        white-space: nowrap;
      }
      .sp-table td {
        padding: 13px 16px;
        font-size: 0.88rem;
        color: #1E293B;
        border-bottom: 1px solid #F1F5F9;
        vertical-align: middle;
      }
      .sp-table tbody tr:hover td {
        background: #F8FAFC;
      }
      .sp-table tbody tr:last-child td {
        border-bottom: none;
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
          gap: 6px !important;
          flex-wrap: nowrap !important;
          overflow-x: auto !important;
          scrollbar-width: none !important;
          padding-bottom: 2px !important;
        }
        .sp-hero-actions::-webkit-scrollbar {
          display: none;
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
          grid-template-columns: minmax(0, 1fr) !important;
          width: 100% !important;
          max-width: 100% !important;
          min-width: 0 !important;
          box-sizing: border-box !important;
          gap: 12px !important;
        }
        .sp-card {
          padding: 16px !important;
          border-radius: 14px !important;
          width: 100% !important;
          max-width: 100% !important;
          min-width: 0 !important;
          box-sizing: border-box !important;
          overflow: hidden;
        }
        .sp-card-header {
          flex-direction: row !important;
          align-items: flex-start !important;
        }
        .sp-view-toggle {
          width: 100% !important;
          display: flex !important;
        }
        .sp-view-btn {
          flex: 1 !important;
          justify-content: center !important;
        }
        .sp-table-card {
          border-radius: 14px !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
        }
        .sp-table {
          min-width: 850px !important;
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
    `;async function T(){try{let[x,w]=await Promise.all([se(),be()]);f=x||[],u=w||[];let a=await q("/api/sp?limit=1000");if(a.ok&&a.data?.data&&Array.isArray(a.data.data)&&a.data.data.length>0)m=a.data.data,localStorage.setItem("fcms_sp_data_cache",JSON.stringify(m));else{let d=localStorage.getItem("fcms_sp_data_cache");d?m=JSON.parse(d):(m=[...Ja],localStorage.setItem("fcms_sp_data_cache",JSON.stringify(m)))}}catch(x){console.warn("Fallback loading SP data from cache/seed:",x);let w=localStorage.getItem("fcms_sp_data_cache");m=w?JSON.parse(w):[...Ja]}c()}function D(){return m.filter(x=>{let w=F.toLowerCase().trim(),a=!w||(x.employee_name||"").toLowerCase().includes(w)||(x.branch_name||"").toLowerCase().includes(w)||(x.division||"").toLowerCase().includes(w)||(x.sp_type||"").toLowerCase().includes(w),d=p==="all"||String(x.branch_id)===String(p),y=$==="all"||x.division===$,k=De(x.akhir_sp),M=!0;return o==="active"?M=!k.isExpired:o==="expiring"?M=k.isExpiringSoon:o==="expired"?M=k.isExpired:o==="sp1"?M=(x.sp_type||"").toUpperCase()==="SP 1":o==="sp2"?M=(x.sp_type||"").toUpperCase()==="SP 2":o==="sp3"?M=(x.sp_type||"").toUpperCase()==="SP 3":o==="teguran"&&(M=(x.sp_type||"").toUpperCase().includes("TEGURAN")),a&&d&&y&&M})}function _(){let x=0,w=0,a=0,d=0,y=0,k=0,M=0,I=new Set;return m.forEach(K=>{K.employee_name&&I.add(K.employee_name.toLowerCase().trim());let B=De(K.akhir_sp);B.isExpired?a++:x++,B.isExpiringSoon&&w++;let P=(K.sp_type||"").toUpperCase();P==="SP 1"?d++:P==="SP 2"?y++:P==="SP 3"?k++:P.includes("TEGURAN")&&M++}),{total:m.length,active:x,expiring:w,expired:a,sp1:d,sp2:y,sp3:k,teguran:M,uniqueEmployees:I.size}}function g(x=null){let w=!!x;if(w&&!r){N("Anda tidak memiliki izin untuk mengedit data SP.");return}if(!w&&!n){N("Anda tidak memiliki izin untuk menambah data SP.");return}let a=!!x,d=x?.tanggal?x.tanggal.slice(0,10):new Date().toISOString().slice(0,10),y=x?.akhir_sp?x.akhir_sp.slice(0,10):"";if(!y){let K=new Date(d);K.setMonth(K.getMonth()+6),y=K.toISOString().slice(0,10)}let k=document.createElement("div");k.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <input type="text" id="sp-modal-emp" list="sp-emp-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih atau ketik nama..." value="${x?.employee_name?String(x.employee_name).replace(/"/g,"&quot;"):""}" required />
            <datalist id="sp-emp-list">
              ${u.map(K=>`<option value="${K.label}">${K.label}</option>`).join("")}
            </datalist>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-div" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="FACILITY CARE" ${x?.division==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${x?.division==="SECURITY"?"selected":""}>SECURITY</option>
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
              ${f.map(K=>`<option value="${K.value}" ${String(x?.branch_id)===String(K.value)?"selected":""}>${K.label}</option>`).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tingkat / Jenis SP <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="SP 1" ${x?.sp_type==="SP 1"?"selected":""}>SP 1 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 2" ${x?.sp_type==="SP 2"?"selected":""}>SP 2 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 3" ${x?.sp_type==="SP 3"?"selected":""}>SP 3 (Masa Berlaku 6 Bulan / Terakhir)</option>
              <option value="Teguran Lisan" ${x?.sp_type==="Teguran Lisan"?"selected":""}>Teguran Lisan (Pembinaan)</option>
            </select>
          </div>
        </div>

        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Terbit SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${d}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Akhir Masa Berlaku SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-akhir" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${y}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Link Berkas / Foto BAP / Scan Surat SP
          </label>
          <input type="url" id="sp-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://drive.google.com/..." value="${x?.document_link?String(x.document_link).replace(/"/g,"&quot;"):""}" />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Kronologi Pelanggaran / Catatan Pembinaan
          </label>
          <textarea id="sp-modal-notes" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan ringkasan pelanggaran disiplin kerja yang terjadi...">${x?.notes||x?.description||""}</textarea>
        </div>
      </div>
    `;let M=k.querySelector("#sp-modal-tanggal"),I=k.querySelector("#sp-modal-akhir");M.addEventListener("change",()=>{if(M.value){let K=new Date(M.value);K.setMonth(K.getMonth()+6),I.value=K.toISOString().slice(0,10)}}),Z({title:a?"\u270F\uFE0F Edit Data Surat Peringatan (SP)":"\u2795 Terbitkan Surat Peringatan (SP) Baru",content:k,size:"md",confirmText:a?"Simpan Perubahan":"Terbitkan SP",confirmClass:"btn-primary",onConfirm:async(K,B)=>{let P=(k.querySelector("#sp-modal-emp").value||"").trim(),b=k.querySelector("#sp-modal-div").value,h=k.querySelector("#sp-modal-branch").value,S=k.querySelector("#sp-modal-type").value,O=k.querySelector("#sp-modal-tanggal").value,z=k.querySelector("#sp-modal-akhir").value,G=(k.querySelector("#sp-modal-link").value||"").trim(),L=(k.querySelector("#sp-modal-notes").value||"").trim();if(!P){N("Nama Karyawan wajib diisi");return}if(!h){N("Cabang penempatan wajib dipilih");return}if(!O||!z){N("Tanggal SP dan Akhir SP wajib diisi");return}let R=f.find(Q=>String(Q.value)===String(h)),V=R?R.label:"Cabang",J={employee_name:P,division:b,branch_id:Number(h),branch_name:V,sp_type:S,tanggal:O,akhir_sp:z,document_link:G,notes:L,status:De(z).status};try{if(a){let Q=await q(`/api/sp/${x.id}`,{method:"PUT",body:JSON.stringify(J)}),U=m.findIndex(Y=>Y.id===x.id);U!==-1&&(m[U]={...m[U],...J}),H("Data SP berhasil diperbarui")}else{let Q=await q("/api/sp",{method:"POST",body:JSON.stringify(J)}),U=Q.ok&&Q.data?.id?Q.data.id:Date.now();m.unshift({id:U,...J}),H("Surat Peringatan baru berhasil diterbitkan")}localStorage.setItem("fcms_sp_data_cache",JSON.stringify(m)),B(),c()}catch(Q){console.error("Error saving SP:",Q),N("Gagal menyimpan SP: "+Q.message)}}})}function A(x){re(`Hapus data Surat Peringatan untuk karyawan <strong>${x.employee_name}</strong> (${x.sp_type})?`,async()=>{try{await q(`/api/sp/${x.id}`,{method:"DELETE"})}catch(w){console.warn("API delete error fallback to local:",w)}m=m.filter(w=>w.id!==x.id),localStorage.setItem("fcms_sp_data_cache",JSON.stringify(m)),H(`Data SP ${x.employee_name} berhasil dihapus`),c()},"Hapus Data SP")}function s(x){let w=De(x.akhir_sp),a=Mt(x.sp_type),d=document.createElement("div");d.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div style="display:flex; align-items:center; gap:14px; padding-bottom:14px; border-bottom:1px solid #E2E8F0;">
          <div class="sp-avatar-box" style="width:54px; height:54px; font-size:1.2rem;">
            ${x.employee_name.slice(0,2).toUpperCase()}
          </div>
          <div>
            <h3 style="margin:0 0 4px 0; font-size:1.15rem; font-weight:700; color:#0F172A;">
              ${x.employee_name}
            </h3>
            <div style="display:flex; align-items:center; gap:8px;">
              <span class="badge" style="background:#EFF6FF; color:#1D4ED8; font-weight:700;">${x.division||"FACILITY CARE"}</span>
              <span style="color:#64748B; font-size:0.85rem;">\u{1F4CD} ${x.branch_name||"Cabang"}</span>
            </div>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Tingkat Sanksi</div>
            <span class="sp-type-badge" style="background:${a.bg}; color:${a.color}; border:1px solid ${a.border}; font-size:0.85rem;">
              <span>${a.icon}</span>
              <span>${a.label}</span>
            </span>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Status Masa Berlaku</div>
            <span class="sp-status-badge" style="background:${w.bg}; color:${w.color}; border:1px solid ${w.border}; font-size:0.85rem;">
              <span>${w.icon}</span>
              <span>${w.label}</span>
            </span>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Tanggal Terbit SP</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${Re(x.tanggal)}</div>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Akhir Masa Berlaku</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${Re(x.akhir_sp)}</div>
          </div>
        </div>

        ${x.notes||x.description?`
          <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:10px; padding:14px;">
            <div style="font-size:0.78rem; font-weight:700; color:#92400E; margin-bottom:4px;">\u{1F4DD} Catatan Pelanggaran / Pembinaan:</div>
            <div style="font-size:0.88rem; color:#78350F; line-height:1.5;">${x.notes||x.description}</div>
          </div>
        `:""}

        ${x.document_link?`
          <div>
            <a href="${x.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-primary" style="width:100%; justify-content:center; padding:10px; font-weight:600; display:flex; align-items:center; gap:8px;">
              <span>\u{1F4C4}</span> Buka Dokumen BAP / Scan SP Fisik
            </a>
          </div>
        `:`
          <div style="text-align:center; padding:8px; font-size:0.82rem; color:#94A3B8;">
            (Belum ada scan dokumen / BAP yang diunggah)
          </div>
        `}
      </div>
    `,Z({title:"\u{1F4CB} Lembar Riwayat Surat Peringatan (SP)",content:d,size:"md",cancelText:"Tutup"})}async function l(){if(!X(t,"export")){N("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:x}=await Promise.resolve().then(()=>(pe(),fe)),w=m.map((a,d)=>{let y=De(a.akhir_sp);return{No:d+1,"Nama Karyawan":a.employee_name||"",Divisi:a.division||"",Cabang:a.branch_name||"","Jenis SP":a.sp_type||"","Tanggal SP":a.tanggal||"","Akhir SP":a.akhir_sp||"","Status Masa Berlaku":y.status,"Sisa Hari":y.daysLeft>0?y.daysLeft:0,"Link BAP / Foto":a.document_link||"","Catatan / Kronologi":a.notes||a.description||""}});x(w,`Data_Surat_Peringatan_FCMS_${new Date().toISOString().slice(0,10)}`),H("Berhasil mengunduh data SP ke format Excel.")}catch(x){N("Gagal ekspor: "+x.message)}}function c(){let x=D(),w=_();e.innerHTML=`
      <div class="sp-page-container">
        
        <!-- Hero Header -->
        <div class="sp-hero-card">
          <div class="sp-hero-header">
            <div>
              <h1 class="sp-hero-title">
                <span>\u{1F4DC}</span> Monitoring Data SP (Surat Peringatan)
              </h1>
            </div>
            <div class="sp-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${n?`
                <button id="sp-btn-add" class="btn btn-primary sp-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Terbitkan SP</span>
                </button>
              `:""}
              
              ${X(t,"export")?`<button id="sp-btn-excel" class="btn btn-ghost sp-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="sp-btn-refresh" class="btn btn-ghost sp-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- Quick Metrics -->
          <div class="sp-stats-grid">
            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(239,68,68,0.2); color:#F87171;">\u{1F534}</div>
              <div>
                <div class="sp-stat-value">${w.active}</div>
                <div class="sp-stat-label">SP Aktif Berlaku</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">\u26A0\uFE0F</div>
              <div>
                <div class="sp-stat-value">${w.expiring}</div>
                <div class="sp-stat-label">Segera Berakhir (&lt;30 Hari)</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">\u2705</div>
              <div>
                <div class="sp-stat-value">${w.expired}</div>
                <div class="sp-stat-label">Masa Sanksi Selesai</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">\u{1F465}</div>
              <div>
                <div class="sp-stat-value">${w.uniqueEmployees} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">karyawan</span></div>
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
              <input type="text" id="sp-search-input" class="sp-search-input" placeholder="Cari nama karyawan, cabang, atau divisi..." value="${F?F.replace(/"/g,"&quot;"):""}" />
            </div>

            <!-- Branch Dropdown -->
            <select id="sp-filter-branch" class="sp-filter-select">
              <option value="all">Semua Cabang (${f.length})</option>
              ${f.map(a=>`<option value="${a.value}" ${String(p)===String(a.value)?"selected":""}>${a.label}</option>`).join("")}
            </select>

            <!-- Division Dropdown -->
            <select id="sp-filter-division" class="sp-filter-select">
              <option value="all" ${$==="all"?"selected":""}>Semua Divisi</option>
              <option value="FACILITY CARE" ${$==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${$==="SECURITY"?"selected":""}>SECURITY</option>
            </select>

            <div class="sp-view-toggle">
              <button type="button" class="sp-view-btn ${E==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="sp-view-btn ${E==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Filter Pills -->
          <div class="sp-filter-pills">
            <button class="sp-pill ${o==="all"?"active":""}" data-filter="all">
              <span>Semua SP</span>
              <span class="sp-pill-count">${w.total}</span>
            </button>
            <button class="sp-pill ${o==="active"?"active":""}" data-filter="active">
              <span>\u{1F534} SP Aktif</span>
              <span class="sp-pill-count">${w.active}</span>
            </button>
            <button class="sp-pill ${o==="expiring"?"active":""}" data-filter="expiring">
              <span>\u26A0\uFE0F &lt;30 Hari</span>
              <span class="sp-pill-count">${w.expiring}</span>
            </button>
            <button class="sp-pill ${o==="expired"?"active":""}" data-filter="expired">
              <span>\u2705 Selesai</span>
              <span class="sp-pill-count">${w.expired}</span>
            </button>
            <button class="sp-pill ${o==="sp1"?"active":""}" data-filter="sp1">
              <span>SP 1</span>
              <span class="sp-pill-count">${w.sp1}</span>
            </button>
            <button class="sp-pill ${o==="sp2"?"active":""}" data-filter="sp2">
              <span>SP 2</span>
              <span class="sp-pill-count">${w.sp2}</span>
            </button>
            <button class="sp-pill ${o==="sp3"?"active":""}" data-filter="sp3">
              <span>SP 3</span>
              <span class="sp-pill-count">${w.sp3}</span>
            </button>
          </div>
        </div>

        <!-- Content Area -->
        <div id="sp-list-container">
          ${x.length===0?`
            <div class="sp-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data SP Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="sp-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:E==="table"?`
            <!-- Table View -->
            <div class="sp-table-card">
              <div class="sp-table-responsive">
                <table class="sp-table">
                  <thead>
                    <tr>
                      <th style="width:45px; text-align:center;">No</th>
                      <th>Nama Karyawan</th>
                      <th>Cabang & Divisi</th>
                      <th style="text-align:center;">Jenis SP</th>
                      <th>Tanggal Terbit</th>
                      <th>Akhir SP</th>
                      <th style="text-align:center;">Status</th>
                      <th>Keterangan / BAP</th>
                      <th style="width:110px; text-align:center;">Aksi</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${x.map((a,d)=>{let y=De(a.akhir_sp),k=Mt(a.sp_type);return`
                        <tr data-id="${a.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${d+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="sp-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${(a.employee_name||"SP").slice(0,2).toUpperCase()}
                              </div>
                              <div style="font-weight:700; color:#0F172A; white-space:nowrap;">
                                ${a.employee_name}
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#334155; font-size:0.84rem;">\u{1F4CD} ${a.branch_name||"Cabang"}</div>
                            <div style="font-size:0.74rem; color:#2563EB; font-weight:600; margin-top:2px;">${a.division||"FACILITY CARE"}</div>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="sp-type-badge" style="background:${k.bg}; color:${k.color}; border:1px solid ${k.border};">
                              <span>${k.icon}</span>
                              <span>${k.label}</span>
                            </span>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            ${Re(a.tanggal)}
                          </td>
                          <td style="white-space:nowrap; font-weight:700; color:#0F172A;">
                            ${Re(a.akhir_sp)}
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="sp-status-badge" style="background:${y.bg}; color:${y.color}; border:1px solid ${y.border};">
                              <span>${y.icon}</span>
                              <span>${y.shortLabel}</span>
                            </span>
                          </td>
                          <td style="max-width:240px;">
                            <div style="font-size:0.82rem; color:#64748B; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; margin-bottom:4px;" title="${String(a.notes||a.description||"").replace(/"/g,"&quot;")}">
                              ${a.notes||a.description||"Pencatatan sanksi peringatan disiplin operasional."}
                            </div>
                            <button class="btn btn-sm btn-ghost sp-detail-btn" style="border:1px solid #CBD5E1; border-radius:6px; font-weight:600; font-size:0.74rem; padding:2px 8px;" data-id="${a.id}">
                              \u{1F441}\uFE0F Rincian BAP
                            </button>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <div style="display:inline-flex; align-items:center; gap:6px;">
                              ${a.document_link?`
                                <a href="${a.document_link}" target="_blank" rel="noopener noreferrer" class="sp-icon-btn" title="Buka Dokumen">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                                </a>
                              `:""}
                              ${r?`
                                <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${a.id}">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              `:""}
                              ${i?`
                                <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${a.id}">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                                </button>
                              `:""}
                            </div>
                          </td>
                        </tr>
                      `}).join("")}
                  </tbody>
                </table>
              </div>
            </div>
          `:`
            <!-- Grid Card View -->
            <div class="sp-cards-grid">
              ${x.map(a=>{let d=De(a.akhir_sp),y=Mt(a.sp_type);return`
                  <div class="sp-card" data-id="${a.id}">
                    <div>
                      <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px;">
                        <span class="sp-type-badge" style="background:${y.bg}; color:${y.color}; border:1px solid ${y.border};">
                          <span>${y.icon}</span>
                          <span>${y.label}</span>
                        </span>
                        <span class="sp-status-badge" style="background:${d.bg}; color:${d.color}; border:1px solid ${d.border};">
                          <span>${d.icon}</span>
                          <span>${d.shortLabel}</span>
                        </span>
                      </div>

                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="sp-avatar-box" style="width:48px; height:48px; font-size:1.05rem;">
                          ${a.employee_name.slice(0,2).toUpperCase()}
                        </div>
                        <div>
                          <h3 style="margin:0 0 3px 0; font-size:1.1rem; font-weight:700; color:#0F172A;">
                            ${a.employee_name}
                          </h3>
                          <div style="font-size:0.8rem; color:#64748B;">
                            \u{1F4CD} ${a.branch_name||"Cabang"} \u2022 <span style="color:#2563EB; font-weight:600;">${a.division||"FC"}</span>
                          </div>
                        </div>
                      </div>

                      <div style="background:#F8FAFC; border-radius:10px; padding:10px 12px; margin-bottom:14px; font-size:0.82rem; display:flex; justify-content:space-between;">
                        <div>
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Tanggal Terbit</div>
                          <div style="font-weight:700; color:#1E293B;">${Re(a.tanggal)}</div>
                        </div>
                        <div style="text-align:right;">
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Akhir SP</div>
                          <div style="font-weight:700; color:#0F172A;">${Re(a.akhir_sp)}</div>
                        </div>
                      </div>

                      <p style="font-size:0.83rem; color:#64748B; line-height:1.45; margin:0 0 16px 0; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">
                        ${a.notes||a.description||"Pencatatan sanksi peringatan disiplin operasional."}
                      </p>
                    </div>

                    <div style="padding-top:12px; border-top:1px solid #F1F5F9; display:flex; align-items:center; justify-content:space-between;">
                      <button class="btn btn-sm btn-ghost sp-detail-btn" style="border:1px solid #CBD5E1; border-radius:8px; font-weight:600; font-size:0.82rem; padding:6px 12px;" data-id="${a.id}">
                        \u{1F441}\uFE0F Rincian BAP
                      </button>

                      <div style="display:flex; align-items:center; gap:6px;">
                        ${a.document_link?`
                          <a href="${a.document_link}" target="_blank" rel="noopener noreferrer" class="sp-icon-btn" title="Buka Dokumen">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                          </a>
                        `:""}
                        ${r?`
                          <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${a.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                        `:""}
                        ${i?`
                          <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${a.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                          </button>
                        `:""}
                      </div>
                    </div>
                  </div>
                `}).join("")}
            </div>
          `}
        </div>

      </div>
    `,v()}function v(){let x=e.querySelector("#sp-btn-add");x&&(x.onclick=()=>g());let w=e.querySelector("#sp-btn-excel");w&&(w.onclick=()=>l());let a=e.querySelector("#sp-btn-refresh");a&&(a.onclick=()=>T());let d=e.querySelector("#sp-btn-reset-filter");d&&(d.onclick=()=>{F="",p="all",$="all",o="all",c()});let y=e.querySelector("#sp-search-input");y&&(y.oninput=I=>{F=I.target.value,c();let K=e.querySelector("#sp-search-input");K&&(K.focus(),K.setSelectionRange(K.value.length,K.value.length))});let k=e.querySelector("#sp-filter-branch");k&&(k.onchange=I=>{p=I.target.value,c()});let M=e.querySelector("#sp-filter-division");M&&(M.onchange=I=>{$=I.target.value,c()}),e.querySelectorAll(".sp-pill").forEach(I=>{I.onclick=()=>{o=I.getAttribute("data-filter")||"all",c()}}),e.querySelectorAll(".sp-view-btn").forEach(I=>{I.onclick=()=>{let K=I.getAttribute("data-view");K&&K!==E&&(E=K,localStorage.setItem("fcms_sp_view_mode",K),c())}}),e.querySelectorAll(".sp-detail-btn, .sp-view-detail-link").forEach(I=>{I.onclick=K=>{K.stopPropagation();let B=Number(I.getAttribute("data-id")),P=m.find(b=>b.id===B);P&&s(P)}}),e.querySelectorAll(".sp-edit-btn").forEach(I=>{I.onclick=K=>{K.stopPropagation();let B=Number(I.getAttribute("data-id")),P=m.find(b=>b.id===B);P&&g(P)}}),e.querySelectorAll(".sp-delete-btn").forEach(I=>{I.onclick=K=>{K.stopPropagation();let B=Number(I.getAttribute("data-id")),P=m.find(b=>b.id===B);P&&A(P)}})}await T()}ge();var jt=[{id:1,tanggal:"2026-10-02",employee_name:"TAUFIK HIDAYAT",from_branch_id:12,from_branch_name:"012. BSD",to_branch_id:0,to_branch_name:"000. Management",status:"Selesai",document_link:""},{id:2,tanggal:"2026-09-30",employee_name:"ELVAN JANGGA PRAMANA",from_branch_id:55,from_branch_name:"055. Gajah Mada",to_branch_id:0,to_branch_name:"000. Management",status:"Proses",document_link:""},{id:3,tanggal:"2026-06-23",employee_name:"Widya Astuti",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:7,to_branch_name:"007. Bekasi",status:"Selesai",document_link:""},{id:4,tanggal:"2026-06-22",employee_name:"Gema Azan Saputra",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:21,to_branch_name:"021. Kelapa Gading Boulevard",status:"Selesai",document_link:""},{id:5,tanggal:"2026-06-22",employee_name:"Ahmad Farhan",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:62,to_branch_name:"062. Summarecon Bekasi",status:"Selesai",document_link:""}];function Va(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Wa(e){let t=[],n=[],r=[],i="all",m="all",f="all",u="",o=!1,p=localStorage.getItem("fcms_mutasi_view_mode")||"card",$=ae(),F=X($,"mutasi","write"),E=X($,"mutasi","write"),C=X($,"mutasi","delete");try{[n,r]=await Promise.all([se(),be()])}catch(l){console.error("Error fetching cache:",l)}async function j(){o=!0,A();try{let l=await q("/api/mutasi?limit=10000");if(l.ok){let c=l.data?.data||l.data||[];t=Array.isArray(c)&&c.length>0?c:jt}else t=jt}catch(l){console.warn("Gagal load mutasi, menggunakan data lokal:",l),t=jt}finally{o=!1,A()}}async function T(){if(!X($,"export")){N("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:l}=await Promise.resolve().then(()=>(pe(),fe)),c=t.map((v,x)=>({No:x+1,Tanggal:v.tanggal||"","Nama Karyawan":v.employee_name||"","Cabang Asal":v.from_branch_name||"","Cabang Tujuan":v.to_branch_name||"",Status:v.status||"","Link Dokumen":v.document_link||""}));l(c,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`),H("Data mutasi berhasil diekspor ke Excel!")}catch(l){N("Gagal ekspor Excel: "+l.message)}}async function D(l){if(!E)return;let c=t.find(v=>v.id===l);re({title:"Hapus Data Mutasi",message:`Apakah Anda yakin ingin menghapus data mutasi untuk "${c?.employee_name||"karyawan ini"}"?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let v=await q(`/api/mutasi/${l}`,{method:"DELETE"});v.ok?(H("Data mutasi berhasil dihapus"),t=t.filter(x=>x.id!==l),A()):N(v.data?.error||"Gagal menghapus mutasi")}catch(v){N("Gagal menghapus data mutasi: "+v.message)}}})}function _(l=null){let c=!!l;if(c&&!E){N("Anda tidak memiliki izin untuk mengedit data mutasi.");return}if(!c&&!F){N("Anda tidak memiliki izin untuk menambah data mutasi.");return}if(!E)return;let v=!!l,x=l?.tanggal?l.tanggal.slice(0,10):new Date().toISOString().slice(0,10),w=document.createElement("div");w.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="mutasi-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${x}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Proses" ${l?.status==="Proses"?"selected":""}>\u23F3 Proses</option>
              <option value="Selesai" ${l?.status==="Selesai"||!l?"selected":""}>\u2705 Selesai</option>
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Karyawan <span style="color:#DC2626;">*</span>
          </label>
          <input list="emp-list-options" id="mutasi-modal-emp" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Ketik atau pilih nama karyawan..." value="${l?.employee_name?String(l.employee_name).replace(/"/g,"&quot;"):""}" required />
          <datalist id="emp-list-options">
            ${r.map(a=>`<option value="${typeof a=="string"?a:a.label||a.name||""}">`).join("")}
          </datalist>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Asal <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-from" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang Asal --</option>
              ${n.map(a=>`
                <option value="${a.value}" ${String(l?.from_branch_id)===String(a.value)?"selected":""}>
                  ${a.label}
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
              ${n.map(a=>`
                <option value="${a.value}" ${String(l?.to_branch_id)===String(a.value)?"selected":""}>
                  ${a.label}
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
            placeholder="https://drive.google.com/file/d/..." value="${l?.document_link?String(l.document_link).replace(/"/g,"&quot;"):""}" />
        </div>
      </div>
    `,Z({title:v?"\u270F\uFE0F Edit Data Mutasi Karyawan":"\u2795 Tambah Data Mutasi Karyawan",content:w,size:"md",confirmText:v?"Simpan Perubahan":"Simpan Mutasi",confirmClass:"btn-primary",onConfirm:async(a,d)=>{let y=w.querySelector("#mutasi-modal-tanggal").value,k=(w.querySelector("#mutasi-modal-emp").value||"").trim(),M=w.querySelector("#mutasi-modal-from").value,I=w.querySelector("#mutasi-modal-to").value,K=w.querySelector("#mutasi-modal-status").value,B=(w.querySelector("#mutasi-modal-link").value||"").trim();if(!y){N("Tanggal mutasi wajib diisi");return}if(!k){N("Nama karyawan wajib diisi");return}if(!M){N("Cabang asal wajib dipilih");return}if(!I){N("Cabang tujuan wajib dipilih");return}let P=n.find(S=>String(S.value)===String(M)),b=n.find(S=>String(S.value)===String(I)),h={tanggal:y,employee_name:k,from_branch_id:Number(M),from_branch_name:P?P.label:"Cabang Asal",to_branch_id:Number(I),to_branch_name:b?b.label:"Cabang Tujuan",status:K,document_link:B};try{if(v){let S=await q(`/api/mutasi/${l.id}`,{method:"PUT",body:JSON.stringify(h)});if(S.ok){H("Data mutasi berhasil diperbarui");let O=t.findIndex(z=>z.id===l.id);O!==-1&&(t[O]={...t[O],...h}),d(),A()}else N(S.data?.error||"Gagal memperbarui data mutasi")}else{let S=await q("/api/mutasi",{method:"POST",body:JSON.stringify(h)});if(S.ok){H("Data mutasi baru berhasil ditambahkan");let O=S.data?.id||Date.now();t.unshift({id:O,...h}),d(),A()}else N(S.data?.error||"Gagal menambahkan mutasi")}}catch(S){N("Terjadi kesalahan: "+S.message)}}})}function g(){return t.filter(l=>{if(u){let c=u.toLowerCase(),v=(l.employee_name||"").toLowerCase().includes(c),x=(l.from_branch_name||"").toLowerCase().includes(c),w=(l.to_branch_name||"").toLowerCase().includes(c);if(!v&&!x&&!w)return!1}return!(m!=="all"&&String(l.from_branch_id)!==String(m)||f!=="all"&&String(l.to_branch_id)!==String(f)||i!=="all"&&l.status!==i)})}function A(){let l=g(),c=t.length,v=t.filter(k=>k.status==="Proses").length,x=t.filter(k=>k.status==="Selesai").length,w=new Date,a=w.getFullYear(),d=w.getMonth(),y=t.filter(k=>{if(!k.tanggal)return!1;let M=new Date(k.tanggal);return!isNaN(M.getTime())&&M.getFullYear()===a&&M.getMonth()===d}).length;e.innerHTML=`
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
        .mutasi-hero-actions {
          display: flex;
          align-items: center;
          gap: 8px;
          flex-wrap: wrap;
        }
        .mutasi-hero-actions .btn {
          height: 38px;
          box-sizing: border-box;
          transition: all 0.2s ease;
        }
        .mutasi-btn-add:hover {
          background: #1D4ED8 !important;
          transform: translateY(-1px);
        }
        .mutasi-btn-excel:hover, .mutasi-btn-refresh:hover {
          background: rgba(255, 255, 255, 0.16) !important;
          border-color: rgba(255, 255, 255, 0.3) !important;
          transform: translateY(-1px);
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
          align-items: center;
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

        /* View Mode Toggle */
        .mutasi-view-toggle {
          display: inline-flex;
          align-items: center;
          background: #F1F5F9;
          padding: 3px;
          border-radius: 10px;
          border: 1px solid #CBD5E1;
          gap: 3px;
          flex-shrink: 0;
        }
        .mutasi-view-btn {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 6px 12px;
          border-radius: 7px;
          font-size: 0.82rem;
          font-weight: 600;
          border: none;
          background: transparent;
          color: #64748B;
          cursor: pointer;
          transition: all 0.15s ease;
          height: 34px;
          box-sizing: border-box;
          user-select: none;
        }
        .mutasi-view-btn:hover {
          color: #0F172A;
          background: rgba(255, 255, 255, 0.5);
        }
        .mutasi-view-btn.active {
          background: #FFFFFF;
          color: #2563EB;
          font-weight: 700;
          box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
        }

        /* Modern Table View */
        .mutasi-table-card {
          background: #FFFFFF;
          border: 1px solid #E2E8F0;
          border-radius: 14px;
          box-shadow: 0 2px 6px rgba(0, 0, 0, 0.03);
          overflow: hidden;
        }
        .mutasi-table-responsive {
          width: 100%;
          overflow-x: auto;
          -webkit-overflow-scrolling: touch;
        }
        .mutasi-table {
          width: 100%;
          border-collapse: collapse;
          text-align: left;
          font-size: 0.88rem;
          min-width: 860px;
        }
        .mutasi-table th {
          background: #F8FAFC;
          color: #475569;
          font-size: 0.76rem;
          font-weight: 700;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          padding: 13px 16px;
          border-bottom: 1.5px solid #E2E8F0;
          white-space: nowrap;
        }
        .mutasi-table td {
          padding: 13px 16px;
          border-bottom: 1px solid #F1F5F9;
          color: #1E293B;
          vertical-align: middle;
        }
        .mutasi-table tbody tr:hover td {
          background: #F8FAFC;
        }
        .mutasi-table tbody tr:last-child td {
          border-bottom: none;
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
            gap: 6px !important;
            flex-wrap: nowrap !important;
            overflow-x: auto !important;
            scrollbar-width: none !important;
            padding-bottom: 2px !important;
          }
          .mutasi-hero-actions::-webkit-scrollbar {
            display: none;
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
          .mutasi-view-toggle {
            width: 100%;
            display: flex;
          }
          .mutasi-view-btn {
            flex: 1;
            justify-content: center;
          }
          .mutasi-table-card {
            border-radius: 12px;
          }
          .mutasi-cards-grid {
            grid-template-columns: minmax(0, 1fr) !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .mutasi-card {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
            overflow: hidden;
            padding: 14px !important;
          }
        }
      </style>

      <div class="mutasi-page">

        <!-- Hero Header Card -->
        <div class="mutasi-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">\u{1F501}</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Data Mutasi Karyawan
                </h1>
              </div>
            </div>

            <div class="mutasi-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${F?`
                <button id="mutasi-btn-add" class="btn btn-primary mutasi-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Mutasi</span>
                </button>
              `:""}
              
              ${X($,"export")?`<button id="mutasi-btn-excel" class="btn btn-ghost mutasi-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="mutasi-btn-refresh" class="btn btn-ghost mutasi-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
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
                <div class="mutasi-stat-val">${c}</div>
                <div class="mutasi-stat-lbl">Total Mutasi</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u23F3
              </div>
              <div>
                <div class="mutasi-stat-val">${v}</div>
                <div class="mutasi-stat-lbl">Dalam Proses</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(34,197,94,0.2); color:#4ADE80;">
                \u2705
              </div>
              <div>
                <div class="mutasi-stat-val">${x}</div>
                <div class="mutasi-stat-lbl">Mutasi Selesai</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(168,85,247,0.2); color:#C084FC;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="mutasi-stat-val">${y}</div>
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
              <option value="all">Semua Cabang Asal (${n.length})</option>
              ${n.map(k=>`<option value="${k.value}" ${String(m)===String(k.value)?"selected":""}>${k.label}</option>`).join("")}
            </select>

            <select id="mutasi-filter-to" class="mutasi-select">
              <option value="all">Semua Cabang Tujuan (${n.length})</option>
              ${n.map(k=>`<option value="${k.value}" ${String(f)===String(k.value)?"selected":""}>${k.label}</option>`).join("")}
            </select>

            <div class="mutasi-view-toggle">
              <button type="button" class="mutasi-view-btn ${p==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="mutasi-view-btn ${p==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Filter Pills (Swipeable on Mobile) -->
          <div class="mutasi-filter-pills">
            <div class="mutasi-pill ${i==="all"?"active":""}" data-status="all">
              <span>Semua Mutasi</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${c}</span>
            </div>
            <div class="mutasi-pill ${i==="Proses"?"active":""}" data-status="Proses">
              <span style="color:#D97706;">\u23F3</span>
              <span>Proses</span>
              <span class="badge" style="background:rgba(217,119,6,0.15); color:#B45309; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${v}</span>
            </div>
            <div class="mutasi-pill ${i==="Selesai"?"active":""}" data-status="Selesai">
              <span style="color:#16A34A;">\u2705</span>
              <span>Selesai</span>
              <span class="badge" style="background:rgba(22,163,74,0.15); color:#15803D; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${x}</span>
            </div>
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="mutasi-list-container">
          ${l.length===0?`
            <div class="mutasi-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data Mutasi Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="mutasi-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:p==="table"?`
            <div class="mutasi-table-card">
              <div class="mutasi-table-responsive">
                <table class="mutasi-table">
                  <thead>
                    <tr>
                      <th style="width:45px; text-align:center;">No</th>
                      <th>Nama Karyawan</th>
                      <th>Tanggal Mutasi</th>
                      <th>Cabang Asal</th>
                      <th>Cabang Tujuan</th>
                      <th style="text-align:center;">Status</th>
                      <th style="text-align:center;">Dokumen SK</th>
                      ${E||C?'<th style="width:90px; text-align:center;">Aksi</th>':""}
                    </tr>
                  </thead>
                  <tbody>
                    ${l.map((k,M)=>{let I=k.status==="Selesai",K=(k.employee_name||"Karyawan").split(" ").map(B=>B[0]).filter(Boolean).slice(0,2).join("").toUpperCase();return`
                        <tr data-id="${k.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${M+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="mutasi-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${K}
                              </div>
                              <div>
                                <div style="font-weight:700; color:#0F172A; white-space:nowrap;">${k.employee_name}</div>
                                <div style="font-size:0.72rem; color:#64748B;">ID Mutasi: #${k.id}</div>
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            \u{1F5D3}\uFE0F ${Va(k.tanggal)}
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#334155; font-size:0.85rem; display:inline-flex; align-items:center; gap:4px;">
                              <span>\u{1F4CD}</span>
                              <span>${k.from_branch_name||"-"}</span>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:700; color:#2563EB; font-size:0.85rem; display:inline-flex; align-items:center; gap:5px;">
                              <span>\u2794</span>
                              <span>${k.to_branch_name||"-"}</span>
                            </div>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="mutasi-status-badge ${I?"selesai":"proses"}">
                              <span>${I?"\u2705":"\u23F3"}</span>
                              <span>${k.status||"Proses"}</span>
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            ${k.document_link?`
                              <a href="${k.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.76rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                                <span>\u{1F4C4} Buka SK</span>
                              </a>
                            `:`
                              <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                            `}
                          </td>
                          ${E||C?`
                            <td style="text-align:center; white-space:nowrap;">
                              <div style="display:inline-flex; align-items:center; gap:6px;">
                                ${E?`
                                  <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${k.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                  </button>
                                `:""}
                                ${C?`
                                  <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${k.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                                  </button>
                                `:""}
                              </div>
                            </td>
                          `:""}
                        </tr>
                      `}).join("")}
                  </tbody>
                </table>
              </div>
            </div>
          `:`
            <div class="mutasi-cards-grid">
              ${l.map(k=>{let M=k.status==="Selesai",I=(k.employee_name||"Karyawan").split(" ").map(K=>K[0]).filter(Boolean).slice(0,2).join("").toUpperCase();return`
                  <div class="mutasi-card" data-id="${k.id}">
                    <div>
                      <!-- Card Header: Status & Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="mutasi-status-badge ${M?"selesai":"proses"}">
                          <span>${M?"\u2705":"\u23F3"}</span>
                          <span>${k.status||"Proses"}</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B; display:inline-flex; align-items:center; gap:4px;">
                          <span>\u{1F5D3}\uFE0F</span>
                          <span>${Va(k.tanggal)}</span>
                        </span>
                      </div>

                      <!-- Employee Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="mutasi-avatar-box">
                          ${I}
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

                      ${E||C?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${E?`
                            <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${k.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                          `:""}
                          ${C?`
                            <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${k.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                            </button>
                          `:""}
                        </div>
                      `:""}
                    </div>

                  </div>
                `}).join("")}
            </div>
          `}
        </div>

      </div>
    `,s()}function s(){let l=e.querySelector("#mutasi-btn-add");l&&(l.onclick=()=>_());let c=e.querySelector("#mutasi-btn-excel");c&&(c.onclick=()=>T());let v=e.querySelector("#mutasi-btn-refresh");v&&(v.onclick=()=>j());let x=e.querySelector("#mutasi-btn-reset-filter");x&&(x.onclick=()=>{u="",m="all",f="all",i="all",A()});let w=e.querySelector("#mutasi-search-input");w&&(w.oninput=y=>{u=y.target.value,A();let k=e.querySelector("#mutasi-search-input");k&&(k.focus(),k.setSelectionRange(k.value.length,k.value.length))});let a=e.querySelector("#mutasi-filter-from");a&&(a.onchange=y=>{m=y.target.value,A()});let d=e.querySelector("#mutasi-filter-to");d&&(d.onchange=y=>{f=y.target.value,A()}),e.querySelectorAll(".mutasi-pill").forEach(y=>{y.onclick=()=>{i=y.getAttribute("data-status")||"all",A()}}),e.querySelectorAll(".mutasi-view-btn").forEach(y=>{y.onclick=()=>{let k=y.getAttribute("data-view");k&&k!==p&&(p=k,localStorage.setItem("fcms_mutasi_view_mode",k),A())}}),e.querySelectorAll(".mutasi-edit-btn").forEach(y=>{y.onclick=k=>{k.stopPropagation();let M=Number(y.getAttribute("data-id")),I=t.find(K=>K.id===M);I&&_(I)}}),e.querySelectorAll(".mutasi-delete-btn").forEach(y=>{y.onclick=k=>{k.stopPropagation();let M=Number(y.getAttribute("data-id"));D(M)}})}await j()}ge();var Kt=[{id:1,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:1,employee_name:"ADE SURAHMAN",start_time:"23:00",end_time:"08:00",break_hours:1,total_hours:8,reason:"General Cleaning"},{id:2,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:2,employee_name:"TAUFIK HIDAYAT",start_time:"00:00",end_time:"08:00",break_hours:1,total_hours:7,reason:"General Cleaning"},{id:3,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:3,employee_name:"YUDA WILDAN WIGUNA",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"},{id:4,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:4,employee_name:"REZA RIZKI",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"},{id:5,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:5,employee_name:"ADITYA PRATAMA",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"}],Qa=["General Cleaning","Deep Cleaning","Cover manpower","Longshift","Pekerjaan urgent"];function Xa(e,t,n){if(!e||!t)return 0;try{let[r,i]=e.split(":").map(Number),[m,f]=t.split(":").map(Number),u=m*60+f-(r*60+i);u<0&&(u+=24*60);let o=u/60-(parseFloat(n)||0);return o<0&&(o=0),Number.isInteger(o)?o:parseFloat(o.toFixed(2))}catch{return 0}}function Nt(e=""){let t=(e||"").toLowerCase();return t.includes("general")?{bg:"#EFF6FF",color:"#1D4ED8",border:"#BFDBFE",icon:"\u{1F9F9}"}:t.includes("deep")?{bg:"#F5F3FF",color:"#6D28D9",border:"#DDD6FE",icon:"\u2728"}:t.includes("cover")?{bg:"#ECFDF5",color:"#047857",border:"#A7F3D0",icon:"\u{1F465}"}:t.includes("longshift")?{bg:"#FFFBEB",color:"#B45309",border:"#FDE68A",icon:"\u23F3"}:t.includes("urgent")?{bg:"#FFF1F2",color:"#BE123C",border:"#FECDD3",icon:"\u{1F6A8}"}:{bg:"#F8FAFC",color:"#475569",border:"#CBD5E1",icon:"\u23F1\uFE0F"}}function qt(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Za(e){let t=[],n=[],r=[],i="all",m="all",f="",u=!1,o=localStorage.getItem("fcms_overtime_view_mode")||"card",p=ae(),$=X(p,"overtime","write"),F=X(p,"overtime","write"),E=X(p,"overtime","delete");try{[n,r]=await Promise.all([se(),Be()])}catch(s){console.error("Error fetching cache:",s)}async function C(){u=!0,g();try{let s=await q("/api/overtime?limit=10000");if(s.ok){let l=s.data?.data||s.data||[];t=Array.isArray(l)&&l.length>0?l:Kt}else t=Kt}catch(s){console.warn("Gagal load lembur, menggunakan data lokal:",s),t=Kt}finally{u=!1,g()}}async function j(){if(!X(p,"export")){N("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:s}=await Promise.resolve().then(()=>(pe(),fe)),l=t.map((c,v)=>({No:v+1,Tanggal:c.date||"","Cabang Lembur":c.branch_name||"","Nama Karyawan":c.employee_name||"","Jam Mulai":c.start_time||"","Jam Selesai":c.end_time||"","Istirahat (Jam)":c.break_hours!=null?c.break_hours:0,"Total Jam":c.total_hours!=null?c.total_hours:0,"Alasan / Keterangan":c.reason||""}));s(l,`Data_Lembur_${new Date().toISOString().slice(0,10)}`),H("Data lembur berhasil diekspor ke Excel!")}catch(s){N("Gagal ekspor Excel: "+s.message)}}async function T(s){if(!E)return;let l=t.find(c=>c.id===s);re({title:"Hapus Catatan Lembur",message:`Apakah Anda yakin ingin menghapus catatan lembur untuk "${l?.employee_name||"karyawan ini"}" pada tanggal ${qt(l?.date)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let c=await q(`/api/overtime/${s}`,{method:"DELETE"});c.ok?(H("Catatan lembur berhasil dihapus"),t=t.filter(v=>v.id!==s),g()):N(c.data?.error||"Gagal menghapus lembur")}catch(c){N("Gagal menghapus catatan lembur: "+c.message)}}})}function D(s=null){if(!F)return;let l=!!s,c=s?.date?s.date.slice(0,10):new Date().toISOString().slice(0,10),v=s?.start_time||"22:00",x=s?.end_time||"08:00",w=s?.break_hours!=null?s.break_hours:1,a=s?.total_hours!=null?s.total_hours:Xa(v,x,w),d=document.createElement("div");d.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="ot-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${c}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penugasan <span style="color:#DC2626;">*</span>
            </label>
            <select id="ot-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${n.map(B=>`
                <option value="${B.value}" ${String(s?.branch_id)===String(B.value)?"selected":""}>
                  ${B.label}
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
            ${r.map(B=>{let P=typeof B=="string"?B:B.label||B.full_name||B.name||"",b=typeof B=="object"&&B!==null?B.value!=null?B.value:B.id!=null?B.id:P:B,h=String(s?.employee_id)===String(b)||s?.employee_name===P;return`<option value="${b}" ${h?"selected":""}>${P}</option>`}).join("")}
          </select>
        </div>

        <!-- Shift & Waktu Lembur -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Mulai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${v}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Selesai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${x}" required />
          </div>
        </div>

        <!-- Durasi & Istirahat -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Istirahat (Jam) <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.5" min="0" max="10" id="ot-modal-break" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${w}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Total Jam Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.1" min="0" id="ot-modal-total" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; background:#F8FAFC; font-weight:700; color:#2563EB;" 
              value="${a}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Alasan / Keterangan Penugasan Lembur <span style="color:#DC2626;">*</span>
          </label>
          <select id="ot-modal-reason" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
            ${Qa.map(B=>`
              <option value="${B}" ${s?.reason===B?"selected":""}>${B}</option>
            `).join("")}
          </select>
        </div>
      </div>
    `;let y=d.querySelector("#ot-modal-start"),k=d.querySelector("#ot-modal-end"),M=d.querySelector("#ot-modal-break"),I=d.querySelector("#ot-modal-total");function K(){let B=y.value,P=k.value,b=M.value;I.value=Xa(B,P,b)}y.addEventListener("input",K),k.addEventListener("input",K),M.addEventListener("input",K),Z({title:l?"\u270F\uFE0F Edit Data Lembur Karyawan":"\u2795 Catat Penugasan Lembur Baru",content:d,size:"md",confirmText:l?"Simpan Perubahan":"Catat Lembur",confirmClass:"btn-primary",onConfirm:async(B,P)=>{let b=d.querySelector("#ot-modal-date").value,h=d.querySelector("#ot-modal-branch").value,S=d.querySelector("#ot-modal-emp").value,O=d.querySelector("#ot-modal-start").value,z=d.querySelector("#ot-modal-end").value,G=parseFloat(d.querySelector("#ot-modal-break").value)||0,L=parseFloat(d.querySelector("#ot-modal-total").value)||0,R=d.querySelector("#ot-modal-reason").value;if(!b){N("Tanggal lembur wajib diisi");return}if(!h){N("Cabang penugasan wajib dipilih");return}if(!S){N("Nama karyawan wajib dipilih");return}if(!O||!z){N("Jam mulai dan jam selesai wajib diisi");return}let V=n.find(Y=>String(Y.value)===String(h)),J=r.find(Y=>String(Y.value)===String(S)||String(Y.id)===String(S)),Q=J?J.label||J.full_name||J.name||"Karyawan":typeof S=="string"&&isNaN(Number(S))?S:"Karyawan",U={date:b,branch_id:Number(h),branch_name:V?V.label:"Cabang Lembur",employee_id:Number(S),employee_name:J?J.label:"Karyawan",start_time:O,end_time:z,break_hours:G,total_hours:L,reason:R};try{if(l){let Y=await q(`/api/overtime/${s.id}`,{method:"PUT",body:JSON.stringify(U)});if(Y.ok){H("Data lembur berhasil diperbarui");let ee=t.findIndex(W=>W.id===s.id);ee!==-1&&(t[ee]={...t[ee],...U}),P(),g()}else N(Y.data?.error||"Gagal memperbarui data lembur")}else{let Y=await q("/api/overtime",{method:"POST",body:JSON.stringify(U)});if(Y.ok){H("Data lembur berhasil ditambahkan");let ee=Y.data?.id||Date.now();t.unshift({id:ee,...U}),P(),g()}else N(Y.data?.error||"Gagal menambahkan lembur")}}catch(Y){N("Terjadi kesalahan: "+Y.message)}}})}function _(){return t.filter(s=>{if(f){let l=f.toLowerCase(),c=(s.employee_name||"").toLowerCase().includes(l),v=(s.branch_name||"").toLowerCase().includes(l),x=(s.reason||"").toLowerCase().includes(l);if(!c&&!v&&!x)return!1}return!(m!=="all"&&String(s.branch_id)!==String(m)||i!=="all"&&s.reason!==i)})}function g(){let s=_(),l=t.length,c=t.reduce((y,k)=>y+(parseFloat(k.total_hours)||0),0),v=l>0?(c/l).toFixed(1):0,x=new Date,w=x.getFullYear(),a=x.getMonth(),d=t.filter(y=>{if(!y.date)return!1;let k=new Date(y.date);return!isNaN(k.getTime())&&k.getFullYear()===w&&k.getMonth()===a}).length;e.innerHTML=`
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
        .lembur-hero-actions {
          display: flex;
          align-items: center;
          gap: 8px;
          flex-wrap: wrap;
        }
        .lembur-hero-actions .btn {
          height: 38px;
          box-sizing: border-box;
          transition: all 0.2s ease;
        }
        .lembur-btn-add:hover {
          background: #1D4ED8 !important;
          transform: translateY(-1px);
        }
        .lembur-btn-excel:hover, .lembur-btn-refresh:hover {
          background: rgba(255, 255, 255, 0.16) !important;
          border-color: rgba(255, 255, 255, 0.3) !important;
          transform: translateY(-1px);
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

        /* View Mode Toggle */
        .lembur-view-toggle {
          display: inline-flex;
          align-items: center;
          background: #F1F5F9;
          padding: 3px;
          border-radius: 10px;
          border: 1px solid #CBD5E1;
          gap: 3px;
          flex-shrink: 0;
        }
        .lembur-view-btn {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 6px 12px;
          border-radius: 7px;
          font-size: 0.82rem;
          font-weight: 600;
          border: none;
          background: transparent;
          color: #64748B;
          cursor: pointer;
          transition: all 0.15s ease;
          height: 34px;
          box-sizing: border-box;
          user-select: none;
        }
        .lembur-view-btn:hover {
          color: #0F172A;
          background: rgba(255, 255, 255, 0.5);
        }
        .lembur-view-btn.active {
          background: #FFFFFF;
          color: #2563EB;
          font-weight: 700;
          box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
        }

        /* Modern Table View */
        .lembur-table-card {
          background: #FFFFFF;
          border: 1px solid #E2E8F0;
          border-radius: 14px;
          box-shadow: 0 2px 6px rgba(0, 0, 0, 0.03);
          overflow: hidden;
        }
        .lembur-table-responsive {
          width: 100%;
          overflow-x: auto;
          -webkit-overflow-scrolling: touch;
        }
        .lembur-table {
          width: 100%;
          border-collapse: collapse;
          text-align: left;
          font-size: 0.88rem;
          min-width: 860px;
        }
        .lembur-table th {
          background: #F8FAFC;
          color: #475569;
          font-size: 0.76rem;
          font-weight: 700;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          padding: 13px 16px;
          border-bottom: 1.5px solid #E2E8F0;
          white-space: nowrap;
        }
        .lembur-table td {
          padding: 13px 16px;
          border-bottom: 1px solid #F1F5F9;
          color: #1E293B;
          vertical-align: middle;
        }
        .lembur-table tbody tr:hover td {
          background: #F8FAFC;
        }
        .lembur-table tbody tr:last-child td {
          border-bottom: none;
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
            gap: 6px !important;
            flex-wrap: nowrap !important;
            overflow-x: auto !important;
            scrollbar-width: none !important;
            padding-bottom: 2px !important;
          }
          .lembur-hero-actions::-webkit-scrollbar {
            display: none;
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
          .lembur-view-toggle {
            width: 100%;
            display: flex;
          }
          .lembur-view-btn {
            flex: 1;
            justify-content: center;
          }
          .lembur-table-card {
            border-radius: 12px;
          }
          .lembur-cards-grid {
            grid-template-columns: minmax(0, 1fr) !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .lembur-card {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
            overflow: hidden;
            padding: 14px !important;
          }
        }
      </style>

      <div class="lembur-page">

        <!-- Hero Header Card -->
        <div class="lembur-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">\u23F1\uFE0F</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Data Lembur Karyawan
                </h1>
              </div>
            </div>

            <div class="lembur-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${$?`
                <button id="lembur-btn-add" class="btn btn-primary lembur-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Lembur</span>
                </button>
              `:""}
              
              ${X(p,"export")?`<button id="lembur-btn-excel" class="btn btn-ghost lembur-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="lembur-btn-refresh" class="btn btn-ghost lembur-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
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
                <div class="lembur-stat-val">${c} Jam</div>
                <div class="lembur-stat-lbl">Total Jam Lembur</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F4DD}
              </div>
              <div>
                <div class="lembur-stat-val">${l}</div>
                <div class="lembur-stat-lbl">Catatan Penugasan</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4CA}
              </div>
              <div>
                <div class="lembur-stat-val">${v} Jam</div>
                <div class="lembur-stat-lbl">Rata-rata Durasi</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="lembur-stat-val">${d}</div>
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
              <input type="text" id="lembur-search-input" class="lembur-search-input" placeholder="Cari nama karyawan, cabang, atau alasan lembur..." value="${f}" />
            </div>

            <select id="lembur-filter-branch" class="lembur-select">
              <option value="all">Semua Cabang (${n.length})</option>
              ${n.map(y=>`<option value="${y.value}" ${String(m)===String(y.value)?"selected":""}>${y.label}</option>`).join("")}
            </select>

            <div class="lembur-view-toggle">
              <button type="button" class="lembur-view-btn ${o==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="lembur-view-btn ${o==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Reason Filter Pills (Swipeable on Mobile) -->
          <div class="lembur-filter-pills">
            <div class="lembur-pill ${i==="all"?"active":""}" data-reason="all">
              <span>Semua Lembur</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${l}</span>
            </div>
            ${Qa.map(y=>{let k=t.filter(I=>I.reason===y).length,M=Nt(y);return`
                <div class="lembur-pill ${i===y?"active":""}" data-reason="${y}">
                  <span>${M.icon}</span>
                  <span>${y}</span>
                  ${k>0?`<span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${k}</span>`:""}
                </div>
              `}).join("")}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="lembur-list-container">
          ${s.length===0?`
            <div class="lembur-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data Lembur Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="lembur-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:o==="table"?`
            <div class="lembur-table-card">
              <div class="lembur-table-responsive">
                <table class="lembur-table">
                  <thead>
                    <tr>
                      <th style="width:45px; text-align:center;">No</th>
                      <th>Nama Karyawan</th>
                      <th>Cabang Penugasan</th>
                      <th>Tanggal</th>
                      <th>Waktu Shift</th>
                      <th style="text-align:center;">Total Jam</th>
                      <th style="text-align:center;">Istirahat</th>
                      <th>Alasan Lembur</th>
                      ${F||E?'<th style="width:90px; text-align:center;">Aksi</th>':""}
                    </tr>
                  </thead>
                  <tbody>
                    ${s.map((y,k)=>{let M=(y.employee_name||"Karyawan").split(" ").map(K=>K[0]).filter(Boolean).slice(0,2).join("").toUpperCase(),I=Nt(y.reason);return`
                        <tr data-id="${y.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${k+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="lembur-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${M}
                              </div>
                              <div style="font-weight:700; color:#0F172A; white-space:nowrap;">
                                ${y.employee_name}
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#334155; font-size:0.85rem; display:inline-flex; align-items:center; gap:4px;">
                              <span>\u{1F4CD}</span>
                              <span>${y.branch_name||"Cabang"}</span>
                            </div>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            \u{1F4C5} ${qt(y.date)}
                          </td>
                          <td style="white-space:nowrap;">
                            <span style="font-weight:700; color:#2563EB;">\u23F0 ${y.start_time||"00:00"} - ${y.end_time||"00:00"}</span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="lembur-hours-badge" style="padding:3px 8px; font-size:0.8rem;">
                              <span>\u23F1\uFE0F</span>
                              <span>${y.total_hours!=null?y.total_hours:0} Jam</span>
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap; color:#64748B; font-weight:600; font-size:0.82rem;">
                            \u2615 ${y.break_hours!=null?y.break_hours:0} Jam
                          </td>
                          <td style="white-space:nowrap;">
                            <span class="lembur-reason-badge" style="background:${I.bg}; color:${I.color}; border:1px solid ${I.border};">
                              <span>${I.icon}</span>
                              <span>${y.reason||"Lembur Operasional"}</span>
                            </span>
                          </td>
                          ${F||E?`
                            <td style="text-align:center; white-space:nowrap;">
                              <div style="display:inline-flex; align-items:center; gap:6px;">
                                ${F?`
                                  <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${y.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                  </button>
                                `:""}
                                ${E?`
                                  <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${y.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                                  </button>
                                `:""}
                              </div>
                            </td>
                          `:""}
                        </tr>
                      `}).join("")}
                  </tbody>
                </table>
              </div>
            </div>
          `:`
            <div class="lembur-cards-grid">
              ${s.map(y=>{let k=(y.employee_name||"Karyawan").split(" ").map(I=>I[0]).filter(Boolean).slice(0,2).join("").toUpperCase(),M=Nt(y.reason);return`
                  <div class="lembur-card" data-id="${y.id}">
                    <div>
                      <!-- Card Header: Jam & Istirahat -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="lembur-hours-badge">
                          <span>\u23F1\uFE0F</span>
                          <span>${y.total_hours!=null?y.total_hours:0} Jam Lembur</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                          \u2615 Istirahat: ${y.break_hours!=null?y.break_hours:0} Jam
                        </span>
                      </div>

                      <!-- Employee & Branch -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="lembur-avatar-box">
                          ${k}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${y.employee_name}
                          </h4>
                          <div style="font-size:0.8rem; color:#64748B; display:flex; align-items:center; gap:4px;">
                            <span>\u{1F4CD}</span>
                            <span style="font-weight:600; color:#334155;">${y.branch_name||"Cabang"}</span>
                          </div>
                        </div>
                      </div>

                      <!-- Shift Times & Date Box -->
                      <div class="lembur-time-box">
                        <div>
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Tanggal</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#1E293B;">
                            \u{1F5D3}\uFE0F ${qt(y.date)}
                          </div>
                        </div>

                        <div style="text-align:right;">
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Waktu Shift</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#2563EB;">
                            \u23F0 ${y.start_time||"00:00"} - ${y.end_time||"00:00"}
                          </div>
                        </div>
                      </div>

                      <!-- Reason Badge -->
                      <div>
                        <span class="lembur-reason-badge" style="background:${M.bg}; color:${M.color}; border:1px solid ${M.border};">
                          <span>${M.icon}</span>
                          <span>${y.reason||"Lembur Operasional"}</span>
                        </span>
                      </div>
                    </div>

                    <!-- Card Footer: Aksi -->
                    ${F||E?`
                      <div style="display:flex; justify-content:flex-end; align-items:center; gap:6px; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:12px;">
                        ${F?`
                          <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${y.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                        `:""}
                        ${E?`
                          <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${y.id}">
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
    `,A()}function A(){let s=e.querySelector("#lembur-btn-add");s&&(s.onclick=()=>D());let l=e.querySelector("#lembur-btn-excel");l&&(l.onclick=()=>j());let c=e.querySelector("#lembur-btn-refresh");c&&(c.onclick=()=>C());let v=e.querySelector("#lembur-btn-reset-filter");v&&(v.onclick=()=>{f="",m="all",i="all",g()});let x=e.querySelector("#lembur-search-input");x&&(x.oninput=a=>{f=a.target.value,g();let d=e.querySelector("#lembur-search-input");d&&(d.focus(),d.setSelectionRange(d.value.length,d.value.length))});let w=e.querySelector("#lembur-filter-branch");w&&(w.onchange=a=>{m=a.target.value,g()}),e.querySelectorAll(".lembur-pill").forEach(a=>{a.onclick=()=>{i=a.getAttribute("data-reason")||"all",g()}}),e.querySelectorAll(".lembur-view-btn").forEach(a=>{a.onclick=()=>{let d=a.getAttribute("data-view");d&&d!==o&&(o=d,localStorage.setItem("fcms_overtime_view_mode",d),g())}}),e.querySelectorAll(".lembur-edit-btn").forEach(a=>{a.onclick=d=>{d.stopPropagation();let y=Number(a.getAttribute("data-id")),k=t.find(M=>M.id===y);k&&D(k)}}),e.querySelectorAll(".lembur-delete-btn").forEach(a=>{a.onclick=d=>{d.stopPropagation();let y=Number(a.getAttribute("data-id"));T(y)}})}await C()}async function en(e){let t=ae();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}me({container:e,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:n=>new Date(n).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(n,r)=>`<strong>${n||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${r.user_role||""}</small>`},{key:"action",label:"Aksi",render:n=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[n]||"badge-neutral"}">${n}</span>`},{key:"module",label:"Modul",render:n=>`<span style="text-transform:capitalize">${(n||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(n,r)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${n}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async n=>{try{let m=(((await q(`/api/audit-logs?search=${n}`)).data||{}).data||[]).find(p=>String(p.id)===String(n));if(!m)return alert("Data tidak ditemukan");let f=p=>{if(!p)return"Tidak ada data";try{return JSON.stringify(JSON.parse(p),null,2)}catch{return p}},u=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${f(m.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${f(m.new_data)}</pre>
           </div>
         </div>
       `,{createModal:o}=await Promise.resolve().then(()=>(ge(),Vt));o({title:`Detail Audit Log #${n}`,content:u,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}ge();pe();var Rt={"2026-09":"September 2026","2026-10":"Oktober 2026","2026-11":"November 2026","2026-12":"Desember 2026"},Qe={scheduled:{label:"Terjadwal",badgeClass:"badge-warning",icon:"\u{1F7E1}",bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},in_progress:{label:"Sedang Proses",badgeClass:"badge-info",icon:"\u{1F535}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},completed:{label:"Selesai",badgeClass:"badge-success",icon:"\u{1F7E2}",bg:"#DCFCE7",color:"#166534",border:"#86EFAC"}};async function tn(e){let t=ae(),n=t&&(t.role==="superadmin"||t.role==="admin"),r=X(t,"export");e.innerHTML=`
    <div class="page-header" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
      <div>
        <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0; font-size:1.45rem; font-weight:800; color:var(--text-1); letter-spacing:-0.02em;">
          <span>\u{1F504}</span> Time Line Migrasi OS
        </h1>
      </div>
      <div class="os-header-actions" style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
        <div class="btn-group" style="background:#F1F5F9; padding:3px; border-radius:10px; display:inline-flex; align-items:center; border:1px solid var(--border); height:38px; box-sizing:border-box;">
          <button id="view-board-btn" class="btn btn-sm" style="background:#fff; font-weight:600; border-radius:7px; box-shadow:0 1px 2px rgba(0,0,0,0.05); padding:0 12px; height:100%; display:inline-flex; align-items:center; cursor:pointer;">
            \u{1F5C2}\uFE0F Matrix Board
          </button>
          <button id="view-table-btn" class="btn btn-sm btn-ghost" style="font-weight:600; border-radius:7px; padding:0 12px; height:100%; display:inline-flex; align-items:center; cursor:pointer;">
            \u{1F4CB} Tabel Data
          </button>
        </div>
        ${r?`
        <button id="btn-export-migrations" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; height:38px; padding:0 14px; border-radius:10px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; box-sizing:border-box;">
          \u{1F4E5} Export Excel
        </button>
        `:""}
        ${n?`
        <button id="btn-add-migration" class="btn btn-primary" style="font-weight:700; height:38px; padding:0 16px; border-radius:10px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; box-sizing:border-box; background:#2563EB; box-shadow:0 4px 12px rgba(37,99,235,0.3);">
          + Tambah Cabang
        </button>
        `:""}
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
  `;let i="board",m=null,f=[];q("/api/branches?all=1").then(g=>{g.ok&&g.data?.data&&(f=g.data.data)}).catch(()=>{});async function u(){try{let g=await q("/api/os-migrations/summary");if(!g.ok)throw new Error("Gagal memuat data migrasi");m=g.data?.data||g.data,o(m),$()}catch(g){document.getElementById("migration-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${g.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function o(g){let A=document.getElementById("migration-summary-cards");if(!A||!g)return;let s=g.total_branches||0,l=g.vendors?.AEON?.total||0,c=g.vendors?.["Triple S"]?.total||0,v=g.by_status?.completed||0,x=g.by_status?.in_progress||0,w=g.by_status?.scheduled||0,a=s>0?Math.round(v/s*100):0;A.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${s} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.8rem; font-weight:800; color:#1E40AF; line-height:1.2;">${l} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
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
          <span style="font-size:0.85rem; font-weight:700; color:#15803D;">${a}%</span>
        </div>
        <div style="width:100%; height:8px; background:#E2E8F0; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${a}%; height:100%; background:#16A34A; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.75rem; color:var(--text-2); font-weight:500;">
          <span>\u{1F7E2} Selesai: <strong>${v}</strong></span>
          <span>\u{1F535} Proses: <strong>${x}</strong></span>
          <span>\u{1F7E1} Terjadwal: <strong>${w}</strong></span>
        </div>
      </div>
    `}function p(){if(!m||!m.matrix)return[];let g=(document.getElementById("migration-search")?.value||"").toLowerCase().trim(),A=document.getElementById("filter-vendor")?.value||"",s=document.getElementById("filter-month")?.value||"",l=document.getElementById("filter-status")?.value||"";return m.matrix.filter(c=>!(A&&c.vendor!==A||s&&c.target_month!==s||l&&c.status!==l||g&&!`${c.branch_code} ${c.branch_name} ${c.branch_full_name} ${c.vendor} ${c.notes||""} ${c.pic_vendor||""}`.toLowerCase().includes(g)))}function $(){let g=p(),A=document.getElementById("filter-stats-text");A&&(A.innerHTML=`Menampilkan <strong>${g.length}</strong> dari <strong>${m?.total_branches||0}</strong> cabang`);let s=document.getElementById("migration-view-container");i==="board"?F(s,g):E(s,g)}function F(g,A){if(!A.length){g.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;return}let s=["AEON","Triple S"],l=["2026-09","2026-10","2026-11","2026-12"],c="";s.forEach(v=>{let x=A.filter(M=>M.vendor===v);if(document.getElementById("filter-vendor")?.value&&document.getElementById("filter-vendor").value!==v||x.length===0)return;let w=x.length,a=v==="AEON"?"#1E40AF":"#0F766E",d=v==="AEON"?"#EFF6FF":"#F0FDFA",y=v==="AEON"?"#BFDBFE":"#99F6E4",k=l.filter(M=>x.some(I=>I.target_month===M));c+=`
        <div class="vendor-board-section" style="margin-bottom:32px; background:#fff; border:1px solid var(--border); border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <!-- Vendor Section Header -->
          <div style="padding:14px 20px; background:${d}; border-bottom:1px solid ${y}; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.1rem; font-weight:800; color:${a};">${v==="AEON"?"\u{1F3E2} AEON":"\u{1F6E1}\uFE0F Triple S"}</span>
              <span class="badge" style="background:#fff; color:${a}; border:1px solid ${y}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px;">
                ${w} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
              Format Timeline Berdasarkan Dokumen Migrasi
            </div>
          </div>

          <!-- Months Grid (Hanya menampilkan bulan yang memiliki jadwal) -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px; padding:18px;">
            ${k.map(M=>{let I=x.filter(P=>P.target_month===M);return`
                <div class="month-column" style="background:#F8FAFC; border:1px solid var(--border); border-radius:10px; display:flex; flex-direction:column; min-height:220px;">
                  <!-- Month Header -->
                  <div style="padding:12px 14px; background:#F1F5F9; border-bottom:1px solid var(--border); border-radius:9px 9px 0 0; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">
                      \u{1F4C5} ${(Rt[M]||M).split(" ")[0]}
                    </div>
                    <span class="badge" style="background:#E2E8F0; color:var(--text-2); font-weight:700; font-size:0.75rem; padding:2px 7px; border-radius:4px;">
                      ${I.length}
                    </span>
                  </div>

                  <!-- Branch Cards Container -->
                  <div style="padding:10px; display:flex; flex-direction:column; gap:8px; flex:1; max-height:480px; overflow-y:auto;">
                    ${I.map(P=>{let b=Qe[P.status]||Qe.scheduled;return`
                        <div class="branch-card" style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 12px; box-shadow:0 1px 2px rgba(0,0,0,0.03); transition:transform 0.15s, box-shadow 0.15s;">
                          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px; margin-bottom:6px;">
                            <div style="font-weight:700; font-size:0.9rem; color:var(--text-1); line-height:1.3;">
                              <span style="color:#2563EB; font-family:monospace; margin-right:4px;">${P.branch_code}.</span>${P.branch_name}
                            </div>
                            ${n?`
                              <div style="display:flex; gap:2px;">
                                <button class="btn-icon-mini edit-card-btn" data-id="${P.id}" title="Edit" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--text-3);">\u270F\uFE0F</button>
                                <button class="btn-icon-mini delete-card-btn" data-id="${P.id}" data-name="${P.branch_name}" title="Hapus" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--danger);">\u{1F5D1}\uFE0F</button>
                              </div>
                            `:""}
                          </div>

                          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px; margin-top:8px;">
                            <!-- Status Selector Dropdown -->
                            <div style="position:relative;">
                              <button class="status-toggle-btn" data-id="${P.id}" data-status="${P.status}" style="background:${b.bg}; color:${b.color}; border:1px solid ${b.border}; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; display:inline-flex; align-items:center; gap:4px; cursor:pointer;" title="Klik untuk ubah status">
                                ${b.icon} ${b.label} \u25BE
                              </button>
                            </div>

                            ${P.employee_count>0?`
                              <span style="font-size:0.75rem; color:var(--text-3); font-weight:600; display:inline-flex; align-items:center; gap:3px;" title="${P.employee_count} personil Facility Care aktif di cabang ini">
                                \u{1F464} ${P.employee_count} FC
                              </span>
                            `:""}
                          </div>

                          ${P.actual_date?`
                            <div style="font-size:0.7rem; color:var(--text-3); margin-top:6px;">
                              Realisasi: <strong>${P.actual_date}</strong>
                            </div>
                          `:""}
                        </div>
                      `}).join("")}
                  </div>
                </div>
              `}).join("")}
          </div>
        </div>
      `}),g.innerHTML=c,C(g)}function E(g,A){if(!A.length){g.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
        </div>
      `;return}g.innerHTML=`
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
                ${n?'<th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${A.map((s,l)=>{let c=Qe[s.status]||Qe.scheduled;return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3);">${l+1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">${s.branch_full_name}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${s.vendor==="AEON"?"#DBEAFE":"#CCFBF1"}; color:${s.vendor==="AEON"?"#1E40AF":"#0F766E"}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${s.vendor}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:var(--text-2);">
                      ${Rt[s.target_month]||s.target_month}
                    </td>
                    <td style="padding:12px 16px;">
                      <button class="status-toggle-btn" data-id="${s.id}" data-status="${s.status}" style="background:${c.bg}; color:${c.color}; border:1px solid ${c.border}; padding:4px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; cursor:pointer;" title="Klik untuk ubah status">
                        ${c.icon} ${c.label} \u25BE
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
                    ${n?`
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
    `,C(g)}function C(g){g.querySelectorAll(".status-toggle-btn").forEach(A=>{A.addEventListener("click",s=>{s.stopPropagation();let l=A.dataset.id,c=A.dataset.status;j(A,l,c)})}),g.querySelectorAll(".edit-card-btn").forEach(A=>{A.addEventListener("click",()=>{let s=A.dataset.id,l=m?.matrix?.find(c=>String(c.id)===String(s));l&&T(l)})}),g.querySelectorAll(".delete-card-btn").forEach(A=>{A.addEventListener("click",()=>{let s=A.dataset.id,l=A.dataset.name||"Cabang";re(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${l}</strong>?`,async()=>{try{if(!(await q(`/api/os-migrations/${s}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus");H(`Jadwal migrasi ${l} berhasil dihapus`),u()}catch(c){N(c.message)}})})})}function j(g,A,s){document.querySelectorAll(".status-popup-menu").forEach(w=>w.remove());let l=g.getBoundingClientRect(),c=document.createElement("div");c.className="status-popup-menu",c.style.cssText=`
      position: fixed;
      top: ${l.bottom+window.scrollY+4}px;
      left: ${l.left+window.scrollX}px;
      background: #fff;
      border: 1px solid var(--border);
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 9999;
      padding: 6px 0;
      min-width: 160px;
    `;let v=[{key:"scheduled",label:"\u{1F7E1} Terjadwal"},{key:"in_progress",label:"\u{1F535} Sedang Proses"},{key:"completed",label:"\u{1F7E2} Selesai"}];c.innerHTML=v.map(w=>`
      <button class="status-option-item" data-key="${w.key}" style="
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
        color: ${w.key===s?"#2563EB":"var(--text-1)"};
        font-weight: ${w.key===s?"700":"500"};
      ">
        ${w.label} ${w.key===s?"\u2713":""}
      </button>
    `).join("");let x=w=>{!c.contains(w.target)&&w.target!==g&&(c.remove(),document.removeEventListener("click",x))};setTimeout(()=>document.addEventListener("click",x),50),c.querySelectorAll(".status-option-item").forEach(w=>{w.addEventListener("click",async()=>{let a=w.dataset.key;if(c.remove(),document.removeEventListener("click",x),a!==s)try{if(!(await q(`/api/os-migrations/${A}/status`,{method:"PATCH",body:JSON.stringify({status:a})})).ok)throw new Error("Gagal memperbarui status");H("Status berhasil diubah"),u()}catch(d){N(d.message)}})}),document.body.appendChild(c)}function T(g=null){if(!n){N("Anda tidak memiliki izin untuk menambah atau mengedit migrasi OS.");return}let A=!!g,s=A?"Edit Jadwal Migrasi OS":"Tambah Jadwal Migrasi OS",c=`
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${f.map(v=>`
      <option value="${v.id}" ${g&&g.branch_id===v.id?"selected":""}>
        ${v.full_name}
      </option>
    `).join("")}
          </select>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Vendor Outsourcing *</label>
            <select name="vendor" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="AEON" ${g?.vendor==="AEON"?"selected":""}>AEON</option>
              <option value="Triple S" ${g?.vendor==="Triple S"?"selected":""}>Triple S</option>
              <option value="Lainnya" ${g&&!["AEON","Triple S"].includes(g.vendor)?"selected":""}>Lainnya</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Target Bulan *</label>
            <select name="target_month" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="2026-09" ${g?.target_month==="2026-09"?"selected":""}>September 2026</option>
              <option value="2026-10" ${g?.target_month==="2026-10"?"selected":""}>Oktober 2026</option>
              <option value="2026-11" ${g?.target_month==="2026-11"?"selected":""}>November 2026</option>
              <option value="2026-12" ${g?.target_month==="2026-12"?"selected":""}>Desember 2026</option>
            </select>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Status Pelaksanaan</label>
            <select name="status" class="form-control" style="width:100%; height:38px; border-radius:8px;">
              <option value="scheduled" ${g?.status==="scheduled"?"selected":""}>\u{1F7E1} Terjadwal</option>
              <option value="in_progress" ${g?.status==="in_progress"?"selected":""}>\u{1F535} Sedang Proses</option>
              <option value="completed" ${g?.status==="completed"?"selected":""}>\u{1F7E2} Selesai</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Tanggal Realisasi (Opsional)</label>
            <input type="date" name="actual_date" class="form-control" value="${g?.actual_date||""}" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">PIC / Koordinator Vendor (Opsional)</label>
          <input type="text" name="pic_vendor" class="form-control" placeholder="Nama PIC vendor / no HP" value="${g?.pic_vendor||""}" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Catatan / Keterangan</label>
          <textarea name="notes" class="form-control" rows="2" placeholder="Catatan serah terima / info tambahan..." style="width:100%; border-radius:8px;">${g?.notes||""}</textarea>
        </div>
      </form>
    `;Z({title:s,content:c,confirmText:A?"Simpan Perubahan":"Tambah Cabang",onConfirm:async(v,x)=>{let w=v.querySelector("#migration-form");if(!w.checkValidity()){w.reportValidity();return}let a=new FormData(w),d={branch_id:parseInt(a.get("branch_id"),10),vendor:a.get("vendor"),target_month:a.get("target_month"),status:a.get("status"),actual_date:a.get("actual_date")||null,pic_vendor:a.get("pic_vendor")||null,notes:a.get("notes")||null};try{let y=A?`/api/os-migrations/${g.id}`:"/api/os-migrations",M=await q(y,{method:A?"PUT":"POST",body:JSON.stringify(d)});if(!M.ok){let I=M.data?.error||M.data?.message||"Gagal menyimpan data";throw new Error(I)}H(A?"Jadwal migrasi berhasil diupdate":"Jadwal migrasi berhasil ditambahkan"),x(),u()}catch(y){N(y.message)}}})}let D=document.getElementById("view-board-btn"),_=document.getElementById("view-table-btn");D.addEventListener("click",()=>{i="board",D.style.background="#fff",D.classList.remove("btn-ghost"),_.style.background="transparent",_.classList.add("btn-ghost"),$()}),_.addEventListener("click",()=>{i="table",_.style.background="#fff",_.classList.remove("btn-ghost"),D.style.background="transparent",D.classList.add("btn-ghost"),$()}),["migration-search","filter-vendor","filter-month","filter-status"].forEach(g=>{let A=document.getElementById(g);A&&(A.addEventListener("input",$),A.addEventListener("change",$))}),document.getElementById("btn-add-migration")?.addEventListener("click",()=>{T()}),document.getElementById("btn-export-migrations")?.addEventListener("click",()=>{let g=p();if(!g.length){N("Tidak ada data untuk diexport");return}let A=g.map((s,l)=>({No:l+1,"Kode Cabang":s.branch_code,"Nama Cabang":s.branch_name,"Cabang Lengkap":s.branch_full_name,"Vendor OS":s.vendor,"Target Bulan":Rt[s.target_month]||s.target_month,Status:Qe[s.status]?.label||s.status,"Tanggal Realisasi":s.actual_date||"","Jml Personil FC":s.employee_count||0,"PIC Vendor":s.pic_vendor||"",Catatan:s.notes||""}));ne(A,`Timeline_Migrasi_OS_${new Date().toISOString().slice(0,10)}`),H("File Excel berhasil diunduh")}),u()}ge();pe();var an={Persiapan:{bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE"},"Area Umum":{bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD"},"Ruang Tindakan":{bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},"Ruang Penunjang":{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},Toilet:{bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4"},"Area Luar":{bg:"#FEF3C7",color:"#92400E",border:"#FDE68A"},"Kontrol & Check":{bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA"},Istirahat:{bg:"#E2E8F0",color:"#475569",border:"#CBD5E1"},"Layanan Pasien":{bg:"#DCFCE7",color:"#15803D",border:"#86EFAC"},"Sampah Umum":{bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},"Sampah Medis":{bg:"#FEE2E2",color:"#991B1B",border:"#FCA5A5"},Penutup:{bg:"#E0E7FF",color:"#3730A3",border:"#C7D2FE"}};async function nn(e){let t=ae(),n=t&&(t.role==="superadmin"||t.role==="admin"),r=X(t,"export");e.innerHTML=`
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
        ${r?`
        <button id="btn-export-dac" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          \u{1F4E5} Export Excel
        </button>
        `:""}
        ${n?`
        <button id="btn-add-dac" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          + Tambah Aktivitas
        </button>
        `:""}
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
  `;let i="pagi",m="timeline",f=[];async function u(){try{let g=await q("/api/dac?all=1");if(!g.ok)throw new Error("Gagal memuat standar DAC");f=g.data?.data||g.data||[];let A=f.filter(v=>v.shift==="pagi").length,s=f.filter(v=>v.shift==="siang").length;document.getElementById("badge-count-pagi").textContent=A,document.getElementById("badge-count-siang").textContent=s,document.getElementById("badge-count-all").textContent=f.length;let l=document.getElementById("dac-category-filter"),c=Array.from(new Set(f.map(v=>v.category))).sort();l.innerHTML='<option value="">Semua Kategori</option>'+c.map(v=>`<option value="${v}">${v}</option>`).join(""),p(),$()}catch(g){document.getElementById("dac-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${g.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function o(){let g=(document.getElementById("dac-search")?.value||"").toLowerCase().trim(),A=document.getElementById("dac-category-filter")?.value||"",s=document.getElementById("dac-period-filter")?.value||"";return f.filter(l=>{if(i!=="all"&&l.shift!==i||A&&l.category!==A)return!1;if(s==="morning"){let c=parseInt(l.start_time.split(":")[0],10);if(l.shift==="pagi"&&c>=12||l.shift==="siang"&&c>=17)return!1}else if(s==="afternoon"){let c=parseInt(l.start_time.split(":")[0],10);if(l.shift==="pagi"&&c<12||l.shift==="siang"&&c<17)return!1}return!(g&&!`${l.start_time} ${l.end_time} ${l.category} ${l.location} ${l.activity}`.toLowerCase().includes(g))})}function p(){let g=document.getElementById("dac-summary-cards");if(!g)return;let A=f.filter(x=>x.shift==="pagi"),s=f.filter(x=>x.shift==="siang"),l=i==="pagi"?A:i==="siang"?s:f,c=i==="pagi"?"08:00 \u2013 17:00":i==="siang"?"13:00 \u2013 22:00":"08:00 \u2013 22:00",v=i==="pagi"?"Shift Pagi":i==="siang"?"Shift Siang":"Semua Shift";g.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Aktivitas (${v})</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${l.length} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Aktivitas</span></div>
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
    `}function $(){let g=o(),A=document.getElementById("dac-stats-text");A&&(A.innerHTML=`Menampilkan <strong>${g.length}</strong> aktivitas`);let s=document.getElementById("dac-view-container");m==="timeline"?F(s,g):E(s,g)}function F(g,A){if(!A.length){g.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;return}g.innerHTML=`
      <div class="timeline-container" style="display:flex; flex-direction:column; gap:12px; max-width:960px; margin:0 auto;">
        ${A.map(s=>{let l=an[s.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},c=s.shift==="pagi"?"\u{1F305} Pagi":"\u{1F307} Siang",v=s.shift==="pagi"?"#FEF3C7":"#E0E7FF",x=s.shift==="pagi"?"#92400E":"#3730A3";return`
            <div class="dac-timeline-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:16px 20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; gap:16px; align-items:flex-start; position:relative;">
              <!-- Time Column -->
              <div style="min-width:130px; text-align:center; background:#F8FAFC; border:1px solid var(--border); border-radius:8px; padding:10px 8px;">
                <div style="font-size:1.1rem; font-weight:800; color:var(--text-1); font-family:monospace;">
                  ${s.start_time} - ${s.end_time}
                </div>
                <div style="font-size:0.75rem; color:var(--text-3); font-weight:600; margin-top:4px;">
                  Item #${s.item_no}
                </div>
                ${i==="all"?`
                  <div style="margin-top:6px;">
                    <span class="badge" style="background:${v}; color:${x}; font-size:0.7rem; font-weight:700; padding:2px 6px; border-radius:4px;">
                      ${c}
                    </span>
                  </div>
                `:""}
              </div>

              <!-- Content Column -->
              <div style="flex:1;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                  <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span class="badge" style="background:${l.bg}; color:${l.color}; border:1px solid ${l.border}; font-weight:700; font-size:0.75rem; padding:3px 10px; border-radius:6px;">
                      ${s.category}
                    </span>
                    <span style="font-weight:700; font-size:0.95rem; color:var(--text-1);">
                      \u{1F4CD} ${s.location}
                    </span>
                  </div>
                  ${n?`
                    <div style="display:flex; gap:4px;">
                      <button class="btn btn-ghost btn-sm edit-dac-btn" data-id="${s.id}" title="Edit" style="padding:4px 8px; font-size:0.8rem; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-ghost btn-sm delete-dac-btn" data-id="${s.id}" data-name="${s.location}" title="Hapus" style="padding:4px 8px; font-size:0.8rem; color:var(--danger); cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </div>
                  `:""}
                </div>

                <div style="font-size:0.9rem; color:var(--text-2); line-height:1.5; background:#FAFBFD; padding:10px 14px; border-radius:8px; border-left:3px solid ${l.color}; margin-top:8px;">
                  ${s.activity}
                </div>
              </div>
            </div>
          `}).join("")}
      </div>
    `,C(g)}function E(g,A){if(!A.length){g.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
        </div>
      `;return}g.innerHTML=`
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div class="table-responsive">
          <table class="table" style="width:100%; border-collapse:collapse;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:1px solid var(--border); text-align:left;">
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:50px;">NO</th>
                ${i==="all"?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">SHIFT</th>':""}
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:110px;">JAM</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:140px;">KATEGORI</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:200px;">LOKASI / AREA</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">AKTIVITAS</th>
                ${n?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right; width:100px;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${A.map(s=>{let l=an[s.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"};return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 14px; font-size:0.85rem; color:var(--text-3); font-weight:600;">${s.item_no}</td>
                    ${i==="all"?`
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
                      <span class="badge" style="background:${l.bg}; color:${l.color}; border:1px solid ${l.border}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px; white-space:nowrap;">
                        ${s.category}
                      </span>
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; font-weight:600; color:var(--text-1);">
                      ${s.location}
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; color:var(--text-2); line-height:1.4;">
                      ${s.activity}
                    </td>
                    ${n?`
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
    `,C(g)}function C(g){g.querySelectorAll(".edit-dac-btn").forEach(A=>{A.addEventListener("click",()=>{let s=A.dataset.id,l=f.find(c=>String(c.id)===String(s));l&&j(l)})}),g.querySelectorAll(".delete-dac-btn").forEach(A=>{A.addEventListener("click",()=>{let s=A.dataset.id,l=A.dataset.name||"item";re(`Apakah Anda yakin ingin menghapus aktivitas pada lokasi <strong>${l}</strong>?`,async()=>{try{if(!(await q(`/api/dac/${s}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus item");H("Aktivitas DAC berhasil dihapus"),u()}catch(c){N(c.message)}})})})}function j(g=null){if(!n){N("Anda tidak memiliki izin untuk menambah atau mengedit aktivitas.");return}let A=!!g,s=A?"Edit Aktivitas DAC":"Tambah Aktivitas DAC",l=["Persiapan","Area Umum","Ruang Tindakan","Ruang Penunjang","Toilet","Area Luar","Kontrol & Check","Istirahat","Layanan Pasien","Sampah Umum","Sampah Medis","Penutup"],c=`
      <form id="dac-form" style="display:flex; flex-direction:column; gap:14px;">
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Shift *</label>
            <select name="shift" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="pagi" ${g?.shift==="pagi"||!g&&i==="pagi"?"selected":""}>\u{1F305} Shift Pagi (08:00 - 17:00)</option>
              <option value="siang" ${g?.shift==="siang"||!g&&i==="siang"?"selected":""}>\u{1F307} Shift Siang (13:00 - 22:00)</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Nomor Urut *</label>
            <input type="number" name="item_no" class="form-control" value="${g?.item_no||""}" placeholder="Otomatis jika kosong" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Mulai (HH:MM) *</label>
            <input type="time" name="start_time" class="form-control" value="${g?.start_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Selesai (HH:MM) *</label>
            <input type="time" name="end_time" class="form-control" value="${g?.end_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Kategori *</label>
          <input list="cat-suggestions" name="category" class="form-control" value="${g?.category||""}" required placeholder="Pilih atau ketik kategori..." style="width:100%; height:38px; border-radius:8px;">
          <datalist id="cat-suggestions">
            ${l.map(v=>`<option value="${v}">`).join("")}
          </datalist>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Lokasi / Area *</label>
          <input type="text" name="location" class="form-control" value="${g?.location||""}" required placeholder="Contoh: Ruang tindakan Du 101, Lobby, Toilet GF" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Aktivitas Detail *</label>
          <textarea name="activity" class="form-control" rows="3" required placeholder="Deskripsi aktivitas pembersihan..." style="width:100%; border-radius:8px;">${g?.activity||""}</textarea>
        </div>
      </form>
    `;Z({title:s,content:c,confirmText:A?"Simpan Perubahan":"Tambah Aktivitas",onConfirm:async(v,x)=>{let w=v.querySelector("#dac-form");if(!w.checkValidity()){w.reportValidity();return}let a=new FormData(w),d={shift:a.get("shift"),item_no:a.get("item_no")?parseInt(a.get("item_no"),10):void 0,start_time:a.get("start_time"),end_time:a.get("end_time"),category:a.get("category"),location:a.get("location"),activity:a.get("activity")};try{let y=A?`/api/dac/${g.id}`:"/api/dac",M=await q(y,{method:A?"PUT":"POST",body:JSON.stringify(d)});if(!M.ok){let I=M.data?.error||M.data?.message||"Gagal menyimpan aktivitas";throw new Error(I)}H(A?"Aktivitas berhasil diperbarui":"Aktivitas berhasil ditambahkan"),x(),u()}catch(y){N(y.message)}}})}function T(){let g=o();if(!g.length){N("Tidak ada data untuk dicetak");return}let A=i==="pagi"?"SHIFT PAGI (08:00 - 17:00)":i==="siang"?"SHIFT SIANG (13:00 - 22:00)":"SELURUH SHIFT (PAGI & SIANG)",s=window.open("","_blank");if(!s){N("Popup diblokir browser, mohon izinkan popup untuk mencetak");return}s.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>STANDAR ACUAN DAC - ${A}</title>
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
          <div class="subtitle">STANDAR RUTINITAS DAILY ACTIVITY CLEANING (DAC) \u2014 ${A}</div>
        </div>
        <table>
          <thead>
            <tr>
              <th style="width:35px; text-align:center;">NO</th>
              ${i==="all"?'<th style="width:55px;">SHIFT</th>':""}
              <th style="width:105px;">JAM</th>
              <th style="width:130px;">KATEGORI</th>
              <th style="width:180px;">LOKASI / AREA</th>
              <th>RINCIAN AKTIVITAS</th>
            </tr>
          </thead>
          <tbody>
            ${g.map(l=>`
              <tr>
                <td style="text-align:center; font-weight:bold;">${l.item_no}</td>
                ${i==="all"?`<td>${l.shift==="pagi"?"Pagi":"Siang"}</td>`:""}
                <td class="jam">${l.start_time} - ${l.end_time}</td>
                <td><strong>${l.category}</strong></td>
                <td><strong>${l.location}</strong></td>
                <td>${l.activity}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
        <div style="margin-top:20px; text-align:right; font-size:9pt; color:#64748b;">
          Dicetak pada: ${new Date().toLocaleString("id-ID")} | Standar Resmi Operasional Facility Care Audy Dental
        </div>
      </body>
      </html>
    `),s.document.close(),s.focus(),setTimeout(()=>{s.print()},400)}document.querySelectorAll(".shift-tab-btn").forEach(g=>{g.addEventListener("click",()=>{document.querySelectorAll(".shift-tab-btn").forEach(A=>{A.style.background="#F1F5F9",A.style.color="var(--text-2)"}),g.style.background="#2563EB",g.style.color="#fff",i=g.dataset.shift,p(),$()})});let D=document.getElementById("view-timeline-btn"),_=document.getElementById("view-table-btn");D.addEventListener("click",()=>{m="timeline",D.style.background="#fff",D.classList.remove("btn-ghost"),_.style.background="transparent",_.classList.add("btn-ghost"),$()}),_.addEventListener("click",()=>{m="table",_.style.background="#fff",_.classList.remove("btn-ghost"),D.style.background="transparent",D.classList.add("btn-ghost"),$()}),["dac-search","dac-category-filter","dac-period-filter"].forEach(g=>{let A=document.getElementById(g);A&&(A.addEventListener("input",$),A.addEventListener("change",$))}),document.getElementById("btn-add-dac")?.addEventListener("click",()=>{j()}),document.getElementById("btn-print-dac").addEventListener("click",()=>{T()}),document.getElementById("btn-export-dac")?.addEventListener("click",()=>{let g=o();if(!g.length){N("Tidak ada data untuk diexport");return}let A=g.map(s=>({No:s.item_no,Shift:s.shift==="pagi"?"Shift Pagi":"Shift Siang","Jam Mulai":s.start_time,"Jam Selesai":s.end_time,Kategori:s.category,"Lokasi / Area":s.location,"Rincian Aktivitas":s.activity}));ne(A,`DAC_Daily_Activity_Cleaning_${i}_${new Date().toISOString().slice(0,10)}`),H("File Excel berhasil diunduh")}),u()}window.parseFlexibleDate=e=>{if(!e||e==="-")return"";if(e=String(e).trim(),/^\d{5}$/.test(e)){let n=Math.floor(Number(e)-25569);return new Date(n*86400*1e3).toISOString().split("T")[0]}if(e.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let n=e.split(/[\/\-]/);return`${n[2]}-${n[1]}-${n[0]}`}let t=e.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(t){let n=t[1],r=parseInt(t[2],10),i=parseInt(t[3],10);if(r>12&&i<=12)return`${n}-${t[3]}-${t[2]}`}return e.split("T")[0]};window.formatDate=e=>{let t=window.parseFlexibleDate(e);if(!t)return"";let n=t.split("-");if(n.length===3&&n[0].length===4){let r=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],i=parseInt(n[2],10),m=r[parseInt(n[1],10)-1];return`${i} ${m} ${n[0]}`}return t};function le(e){return async t=>{if(!Ze()){$e("/login");return}return e(t)}}var Xe=null;function Xn(){Xe&&clearInterval(Xe);let e=()=>{let t=new Date,n=t.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),r=t.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),i=document.getElementById("header-clock-time"),m=document.getElementById("header-clock-date");i&&(i.textContent=n),m&&(m.textContent=r)};e(),Xe=setInterval(e,1e3)}async function Zn(){try{let e=await q("/api/dashboard/kpi");if(!e.ok)return;let t=e.data?.data||e.data||{},n=(r,i)=>{let m=document.getElementById(r);m&&(m.textContent=i>0?i:"",m.style.display=i>0?"inline-flex":"none")};n("badge-issues",t.issues?.current||0),n("badge-contracts",t.expiring30?.current||0),n("badge-oo1",t.one_on_one?.current||0),n("badge-schedule",t.schedule?.current||0),n("badge-supply",t.supply?.current||0)}catch{}}var Oe=[];async function ei(){try{let e=await q("/api/dashboard/notifications");if(!e.ok)return;Oe=e.data?.data||e.data||[];let t=document.getElementById("notif-dot");t&&(t.style.display=Oe.length>0?"block":"none",t.textContent=Oe.length)}catch{}}function ti(){if(!Oe.length){Z({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(t,n)=>n()});return}let e=`
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
  `;Z({title:`Notifikasi (${Oe.length})`,content:e,confirmText:"Tutup",onConfirm:(t,n)=>n()})}function rn(){let e=ae(),t=(e?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
                <span class="topbar-greeting-time">${(()=>{let p=new Date().getHours();return p>=4&&p<11?"Selamat Pagi":p>=11&&p<15?"Selamat Siang":p>=15&&p<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${e?.full_name||e?.username||"Guest"}</span> \u{1F44B}
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
  `;let n=document.getElementById("sidebar"),r=document.getElementById("sidebar-overlay"),i=document.getElementById("topbar-menu-btn"),m=document.getElementById("sidebar-close"),f=()=>{n.classList.add("open"),r.classList.add("show")},u=()=>{n.classList.remove("open"),r.classList.remove("show")};i?.addEventListener("click",f),m?.addEventListener("click",u),r?.addEventListener("click",u),document.querySelectorAll(".nav-item").forEach(p=>p.addEventListener("click",u)),e&&e.role==="input_lembur"&&(document.querySelectorAll(".nav-item").forEach(p=>{p.dataset.route!=="/overtime"&&(p.style.display="none")}),document.querySelectorAll(".nav-section").forEach(p=>{Array.from(p.querySelectorAll(".nav-item")).some(F=>F.style.display!=="none")||(p.style.display="none")}));function o(){let p=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(E=>{let C=E.dataset.route;E.classList.toggle("active",p===C||C!=="/dashboard"&&p.startsWith(C))});let $=document.getElementById("topbar-title"),F=document.querySelector(".nav-item.active .nav-label");$&&F&&($.textContent=F.textContent)}window.addEventListener("hashchange",o),o(),Xn(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await q("/api/auth/logout",{method:"POST"}),et(),Xe&&clearInterval(Xe),$e("/login")}),Zn(),ei(),document.getElementById("btn-notif")?.addEventListener("click",p=>{p.preventDefault(),ti()})}async function ai(){oe("/login",({main:t})=>ma(t)),oe("/dashboard",le(({main:t})=>da(t))),oe("/calendar",le(({main:t})=>Na(t))),oe("/employees",le(({main:t,params:n})=>ea(t,n))),oe("/contracts",le(({main:t,params:n})=>aa(t,n))),oe("/sp",le(({main:t})=>Ya(t))),oe("/mutasi",le(({main:t})=>Wa(t))),oe("/overtime",le(({main:t})=>Za(t))),oe("/timeline",le(({main:t,params:n})=>Qt(t,n))),oe("/os-migrations",le(({main:t})=>tn(t))),oe("/issues",le(({main:t,params:n})=>ra(t,n))),oe("/one-on-one",le(({main:t,params:n})=>oa(t,n))),oe("/training",le(({main:t})=>fa(t))),oe("/relievers",le(({main:t,params:n})=>ya(t,n))),oe("/reports/inspection",le(({main:t})=>va(t))),oe("/reports/cleaning",le(({main:t})=>ka(t))),oe("/reports/fogging",le(({main:t})=>wa(t))),oe("/reports/basecamp",le(({main:t})=>Ea(t))),oe("/reports/supply",le(({main:t})=>Ma(t,"supply"))),oe("/sop",()=>{window.location.hash="#/sop-fc"}),oe("/sop-fc",le(({main:t})=>$a(t))),oe("/sop-security",le(({main:t})=>Ta(t))),oe("/checklist",le(({main:t})=>Ia(t))),oe("/hygiene-standards",le(({main:t})=>za(t))),oe("/dac",le(({main:t})=>nn(t))),oe("/users",le(({main:t})=>ja(t))),oe("/branches",le(({main:t})=>Ka(t))),oe("/profile",le(({main:t})=>qa(t))),oe("/settings/import",le(({main:t})=>{let n=ae();if(!n||n.role!=="superadmin"&&n.role!=="admin"){$e("/dashboard");return}Ga(t)})),oe("/audit-logs",le(({main:t})=>en(t)));let e=Ze();if(!e&&window.location.hash!=="#/login"&&$e("/login"),e){let t=await q("/api/auth/me");if(t.ok){tt(t.data.data),rn();let n=ae(),r=window.location.hash;n&&n.role==="input_lembur"&&(r==="#/dashboard"||r===""||r==="#/")&&$e("/overtime")}else et(),$e("/login")}window.addEventListener("fm:login",()=>{rn();let t=ae();t&&t.role==="input_lembur"?$e("/overtime"):$e("/dashboard")}),Yt()}ai();
