var yn=Object.defineProperty;var Vt=(e,t)=>()=>(e&&(t=e(e=0)),t);var Wt=(e,t)=>{for(var n in t)yn(e,n,{get:t[n],enumerable:!0})};var Zt={};Wt(Zt,{confirmDialog:()=>ie,createModal:()=>te});function te({title:e,content:t,onConfirm:n,onCancel:i,confirmText:r="Simpan",cancelText:m="Batal",size:b="md",confirmClass:g="btn-primary"}){let s={sm:"400px",md:"560px",lg:"720px",xl:"900px"},d=document.createElement("div");d.className="modal-overlay",d.innerHTML=`
    <div class="modal" style="max-width:${s[b]||s.md}">
      <div class="modal-header">
        <h3 class="modal-title">${e}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof t=="string"?t:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${m}</button>
        ${n?`<button class="btn ${g} modal-confirm">${r}</button>`:""}
      </div>
    </div>
  `,t instanceof HTMLElement&&d.querySelector(".modal-body").appendChild(t);let F=()=>{d.classList.remove("show"),setTimeout(()=>d.remove(),250)};return d.querySelector(".modal-close").addEventListener("click",()=>{i&&i(),F()}),d.querySelector(".modal-cancel").addEventListener("click",()=>{i&&i(),F()}),n&&d.querySelector(".modal-confirm").addEventListener("click",()=>n(d,F)),d.addEventListener("click",S=>{S.target===d&&(i&&i(),F())}),document.body.appendChild(d),requestAnimationFrame(()=>d.classList.add("show")),{overlay:d,close:F}}function ie(e,t,n="Konfirmasi"){if(typeof e=="object"&&e!==null){let i=e;return te({title:i.title||"Konfirmasi",content:typeof i.message=="string"?`<p>${i.message}</p>`:i.message,onConfirm:async(r,m)=>{i.onConfirm&&await i.onConfirm(),m()},onCancel:i.onCancel,confirmText:i.confirmText||"Ya, Lanjutkan",cancelText:i.cancelText||"Batal",confirmClass:i.confirmClass||"btn-danger"})}return te({title:n,content:typeof e=="string"?`<p>${e}</p>`:e,onConfirm:async(i,r)=>{t&&await t(),r()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var me=Vt(()=>{});var xe={};Wt(xe,{downloadExcel:()=>ne,parseExcel:()=>Je,renderExcelButtons:()=>En});function Je(e){return new Promise((t,n)=>{let i=new FileReader;i.onload=r=>{try{let m=new Uint8Array(r.target.result),b=XLSX.read(m,{type:"array"}),g=b.SheetNames[0],s=b.Sheets[g];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${e.name}`),console.log(`File Size: ${(e.size/1024).toFixed(2)} KB`),console.log(`File Type: ${e.type||"unknown"}`),console.log(`Sheets Found: ${b.SheetNames.join(", ")}`),console.log(`Sheet Used: ${g}`);let d=XLSX.utils.decode_range(s["!ref"]||"A1:A1"),F=d.e.r-d.s.r+1,S=d.e.c-d.s.c+1;console.log(`Total Rows (including empty): ${F}`),console.log(`Total Columns: ${S}`);let w=[];for(let M=d.s.c;M<=d.e.c;++M){let A=s[XLSX.utils.encode_cell({c:M,r:d.s.r})];A&&A.v&&w.push(A.v)}console.log(`Headers Found: ${w.join(", ")}`),console.log("---------------------------");let C=XLSX.utils.sheet_to_json(s,{defval:""});Object.defineProperty(C,"__worksheet",{value:s,enumerable:!1}),Object.defineProperty(C,"__headers",{value:w,enumerable:!1}),t(C)}catch(m){n(m)}},i.onerror=r=>n(r),i.readAsArrayBuffer(e)})}function ne(e,t){try{let n=XLSX.utils.json_to_sheet(e),i=XLSX.utils.book_new();XLSX.utils.book_append_sheet(i,n,"Data"),XLSX.writeFile(i,`${t}.xlsx`)}catch(n){throw console.error("Error generating Excel file:",n),n}}function En(e){return`
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
  `}var de=Vt(()=>{});var vn="https://fm-operations-api.facilitycare-audydental.workers.dev",kn=vn,Ie=1e4;function tt(){return localStorage.getItem("fm_token")}function Qt(e){localStorage.setItem("fm_token",e)}function at(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ae(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function nt(e){localStorage.setItem("fm_user",JSON.stringify(e))}function Q(e,t="read",n=null){if(!e||!e.role)return!1;let i=e.role;if(i==="superadmin"||i==="admin")return!0;let r=null,m="read";return n!==null?(r=t,m=n):m=t||"read",m==="export"||m==="import"||m==="admin"?!1:i==="editor_khusus"?r==="issues"||r==="overtime"?m==="read"||m==="write"||m==="delete":m==="read":i==="input_lembur"&&r==="overtime"?m==="read"||m==="write":m==="read"}async function R(e,t={}){let n=tt(),i={"Content-Type":"application/json",...n?{Authorization:`Bearer ${n}`}:{},...t.headers||{}};try{let r=`cb=${Date.now()}`,m=e.includes("?")?"&":"?",b=`${kn}${e}${m}${r}`,g=await fetch(b,{...t,headers:i}),s;try{let d=await g.text();try{s=JSON.parse(d)}catch{s={error:`Server Error (${g.status}): ${d.substring(0,80)}...`}}}catch{s={error:"Gagal membaca respon dari server"}}return g.status===401&&(at(),window.location.hash="#/login"),{ok:g.ok,status:g.status,data:s}}catch(r){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${r.message})`}}}}var vt={},it=null;function oe(e,t){vt[e]=t}function _e(e){window.location.hash=e}function Xt(){async function e(){let t=window.location.hash.replace("#","")||"/dashboard",[n,...i]=t.split("?"),r=vt[n];if(!r){for(let[b,g]of Object.entries(vt))if(b.endsWith("/*")&&n.startsWith(b.slice(0,-2))){r=g;break}}it&&(it(),it=null);let m=document.getElementById("main-content");if(m&&(m.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),r){let b=new URLSearchParams(i.join("?")),g=n.split("/").filter(Boolean),s=await r({path:n,params:b,segments:g,main:m});s&&(it=s)}else{let b=m||document.getElementById("app");b&&(b.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",e),e()}var Ue;function wn(){return Ue||(Ue=document.createElement("div"),Ue.id="toast-container",document.body.appendChild(Ue)),Ue}function kt(e,t="info",n=3500){let i=wn(),r=document.createElement("div");r.className=`toast toast-${t}`;let m={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};r.innerHTML=`<span class="toast-icon">${m[t]||"\u2139"}</span><span class="toast-msg">${e}</span>`,i.appendChild(r),requestAnimationFrame(()=>r.classList.add("show")),setTimeout(()=>{r.classList.remove("show"),setTimeout(()=>r.remove(),350)},n)}var G=e=>kt(e,"success"),q=e=>kt(e,"error");var wt=e=>kt(e,"info");me();me();function rt({columns:e,data:t,onEdit:n,onDelete:i,onView:r,actions:m=[],emptyText:b="Tidak ada data",bulkSelect:g=null}){let s=document.createElement("div");if(s.className="table-wrapper",!t||t.length===0)return s.innerHTML=`<div class="empty-state"><p>${b}</p></div>`,s;let d=document.createElement("table");d.className="data-table";let F=document.createElement("thead"),S=document.createElement("tr");if(g){let C=document.createElement("th");C.style.width="40px",C.style.textAlign="center";let M=document.createElement("input");M.type="checkbox",M.id="select-all-checkbox",M.title="Pilih semua",M.addEventListener("change",()=>{t.forEach(A=>{M.checked?g.selectedIds.add(A.id):g.selectedIds.delete(A.id)}),s.querySelectorAll(".row-checkbox").forEach(A=>A.checked=M.checked),g.onToggle()}),C.appendChild(M),S.appendChild(C)}if(e.forEach(C=>{let M=document.createElement("th");M.textContent=C.label,C.width&&(M.style.width=C.width),S.appendChild(M)}),n||i||r||m.length>0){let C=document.createElement("th");C.textContent="Aksi",C.style.width="120px",S.appendChild(C)}F.appendChild(S),d.appendChild(F);let w=document.createElement("tbody");return t.forEach(C=>{let M=document.createElement("tr");if(g){let A=document.createElement("td");A.style.textAlign="center",A.style.width="40px";let B=document.createElement("input");B.type="checkbox",B.className="row-checkbox",B.checked=g.selectedIds.has(C.id),B.addEventListener("change",()=>{if(B.checked)g.selectedIds.add(C.id);else{g.selectedIds.delete(C.id);let T=document.getElementById("select-all-checkbox");T&&(T.checked=!1)}g.onToggle()}),A.appendChild(B),M.appendChild(A)}if(e.forEach(A=>{let B=document.createElement("td");if(A.render){let T=A.render(C[A.key],C);T instanceof HTMLElement?B.appendChild(T):B.innerHTML=T||""}else B.textContent=C[A.key]!==null&&C[A.key]!==void 0&&C[A.key]!==""?C[A.key]:"";A.nowrap&&(B.style.whiteSpace="nowrap"),M.appendChild(B)}),n||i||r||m.length>0){let A=document.createElement("td");A.className="actions-cell";let B=document.createElement("div");if(B.className="btn-group",r){let T=document.createElement("button");T.className="btn btn-xs btn-ghost",T.innerHTML="\u{1F441}",T.title="Lihat",T.addEventListener("click",()=>r(C)),B.appendChild(T)}if(n){let T=document.createElement("button");T.className="btn btn-xs btn-secondary",T.innerHTML="\u270F\uFE0F",T.title="Edit",T.addEventListener("click",()=>n(C)),B.appendChild(T)}m.forEach(T=>{let u=document.createElement("button");u.className=`btn btn-xs ${T.class||"btn-ghost"}`,u.innerHTML=T.icon||T.label,u.title=T.label,u.addEventListener("click",()=>T.handler(C)),B.appendChild(u)}),A.appendChild(B),M.appendChild(A)}w.appendChild(M)}),d.appendChild(w),s.appendChild(d),s}function ot({page:e,pages:t,total:n,limit:i,onPage:r}){if(t<=1)return null;let m=document.createElement("div");m.className="pagination";let b=document.createElement("span");b.className="pagination-info",b.textContent=`Total: ${n} data`,m.appendChild(b);let g=document.createElement("div");g.className="pagination-btns";let s=(S,w,C=!1,M=!1)=>{let A=document.createElement("button");A.className=`btn btn-sm ${M?"btn-primary":"btn-ghost"} pagination-btn`,A.textContent=S,A.disabled=C,A.addEventListener("click",()=>r(w)),g.appendChild(A)};s("\xAB",1,e===1),s("\u2039",e-1,e===1);let d=Math.max(1,e-2),F=Math.min(t,e+2);for(let S=d;S<=F;S++)s(S,S,!1,S===e);return s("\u203A",e+1,e===t),s("\xBB",t,e===t),m.appendChild(g),m}me();function Ge(e){return e.map(t=>{if(t.type==="hidden")return`<input type="hidden" name="${t.name}" value="${t.value||""}">`;if(t.type==="html")return t.html||"";if(t.type==="row")return`<div class="form-row">${Ge(t.fields)}</div>`;let n=t.required?"required":"",i=t.label?`<label class="form-label">${t.label}${t.required?' <span class="required">*</span>':""}</label>`:"",r="";switch(t.type){case"textarea":r=`<textarea name="${t.name}" class="form-control" placeholder="${t.placeholder||""}" ${n} rows="${t.rows||3}">${t.value||""}</textarea>`;break;case"select":let b=(t.options||[]).map(C=>{let M=typeof C=="object"?C.value:C,A=typeof C=="object"?C.label:C,B=t.value==M?"selected":"";return`<option value="${M}" ${B}>${A}</option>`}).join("");r=`<select name="${t.name}" class="form-control" ${n}><option value="">-- Pilih ${t.label||""} --</option>${b}</select>`;break;case"combobox":let g=`dl-${t.name}-${Math.random().toString(36).substring(7)}`,s=(t.options||[]).map(C=>{let M=typeof C=="object"?C.value:C,A=typeof C=="object"?C.label||C.value||"":C||"";return(A==="undefined"||A==="[object Object]"||A==="null")&&(A=""),A?`<option value="${A}"></option>`:""}).join(""),d=t.value||"";if(t.value){let C=(t.options||[]).find(M=>(typeof M=="object"?M.value:M)==t.value);if(C){let M=typeof C=="object"?C.label||C.value||"":C||"";M&&M!=="undefined"&&M!=="[object Object]"&&M!=="null"&&(d=M)}}r=`
          <input type="text" name="${t.name}" list="${g}" class="form-control" value="${d}" placeholder="Pilih atau ketik baru..." ${n} autocomplete="off">
          <datalist id="${g}">${s}</datalist>
        `;break;case"checkbox":r=`<label class="checkbox-label"><input type="checkbox" name="${t.name}" value="1" ${t.value?"checked":""}> ${t.checkLabel||t.label}</label>`;break;case"date":let F=window.parseFlexibleDate&&t.value?window.parseFlexibleDate(t.value):t.value||"";r=`<input type="date" name="${t.name}" class="form-control" value="${F}" ${n}>`;break;case"month":let S=t.value?String(t.value).slice(0,7):"";r=`<input type="month" name="${t.name}" class="form-control" value="${S}" ${n}>`;break;case"number":let w=t.readonly?"readonly":"";r=`<input type="number" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" min="${t.min||""}" max="${t.max||""}" step="${t.step||"1"}" ${n} ${w}>`;break;case"email":r=`<input type="email" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${n} autocomplete="${t.autocomplete||"off"}">`;break;case"password":r=`<input type="password" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${n} autocomplete="${t.autocomplete||"new-password"}">`;break;case"url":r=`<input type="url" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||"https://..."}" ${n}>`;break;default:r=`<input type="${t.type||"text"}" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${n} autocomplete="${t.autocomplete||"off"}">`}let m=t.hint?`<div class="form-hint">${t.hint}</div>`:"";return`<div class="form-group ${t.class||""}">${i}${r}${m}</div>`}).join("")}function st(e){let t={},n=new FormData(e);for(let[i,r]of n.entries())t[i]=r===""?null:r;return e.querySelectorAll("input[type=checkbox]").forEach(i=>{i.checked||(t[i.name]=null)}),t}function lt(e,t){t&&Object.entries(t).forEach(([n,i])=>{let r=e.querySelector(`[name="${n}"]`);r&&(r.hasAttribute("list")||(r.type==="checkbox"?r.checked=!!i:r.type==="date"&&i&&window.parseFlexibleDate?r.value=window.parseFlexibleDate(i):r.type==="month"&&i?r.value=String(i).slice(0,7):r.value=i??""))})}de();var Ae={},dt={on(e,t){Ae[e]||(Ae[e]=new Set),Ae[e].add(t)},off(e,t){Ae[e]&&Ae[e].delete(t)},emit(e,t){Ae[e]&&Ae[e].forEach(n=>{try{n(t)}catch(i){console.warn("[calendarBus] Handler error:",i)}})},clear(){Object.keys(Ae).forEach(e=>delete Ae[e])}},Sn=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function ye(e){if(!e){dt.emit("data:changed",{module:"unknown"});return}let t=String(e).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");dt.emit("data:changed",{module:t,relevant:Sn.has(t)})}function be({container:e,title:t,icon:n,apiPath:i,columns:r,formFields:m,filterFields:b,defaultFilters:g={},defaultSort:s="",itemLabel:d="Data",canCreate:F=!0,canEdit:S=!0,canDelete:w=!0,onBeforeSubmit:C,onAfterLoad:M,onDataLoaded:A,extraActions:B=[],initialSearch:T="",exportOptions:u=null,bulkDelete:$=!1,paginationMode:p="server"}){let o=ae();o&&(o.role==="superadmin"||o.role==="admin")?Q(o,"delete")||(w=!1,$=!1):(u=null,$=!1,o&&o.role==="editor_khusus"?i!=="/api/issues"&&i!=="/api/overtime"&&(F=!1,S=!1,w=!1):o&&o.role==="input_lembur"?(i==="/api/overtime"||(F=!1,S=!1),w=!1):(F=!1,S=!1,w=!1));let y=1,x={...g};T&&(x.search=T);let k=new Set;e.innerHTML=`
    ${$?`
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
        ${F?`<button class="btn btn-primary" id="btn-create">+ Tambah ${d}</button>`:""}
        ${u?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${u.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${u.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${u.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${u.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
    

    ${b&&b.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${b.filter(l=>l.type==="search").map(l=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${l.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${b.filter(l=>l.type!=="search").map(l=>{if(l.type==="select"||l.type==="combobox"){let E=(l.label||"").startsWith("Pilih")?l.label:`Pilih ${l.label||""}`;return`<select class="filter-select" name="${l.name}" id="filter-${l.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${E}</option>${(l.options||[]).map(D=>`<option value="${typeof D=="object"?D.value:D}" ${x[l.name]===(typeof D=="object"?D.value:D)?"selected":""}>${typeof D=="object"?D.label:D}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${b.filter(l=>l.type!=="search").map(l=>{if(l.type==="select"||l.type==="combobox"){let E=(l.label||"").startsWith("Pilih")?l.label:`Pilih ${l.label||""}`;return`<select class="filter-select filter-select-sheet" name="${l.name}-sheet" id="filter-sheet-${l.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${E}</option>${(l.options||[]).map(D=>`<option value="${typeof D=="object"?D.value:D}" ${x[l.name]===(typeof D=="object"?D.value:D)?"selected":""}>${typeof D=="object"?D.label:D}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function a(){let l=document.getElementById("bulk-toolbar");if(!l)return;let E=document.getElementById("bulk-count"),D=document.getElementById("btn-bulk-delete"),z=document.getElementById("btn-bulk-cancel");E.textContent=`${k.size} item dipilih`,k.size>0?(l.style.display="flex",D.disabled=!1,z.disabled=!1):(l.style.display="none",D.disabled=!0,z.disabled=!0);let J=document.getElementById("select-all-checkbox");if(J){let N=document.querySelectorAll(".row-checkbox");if(N.length>0){let O=[...N].every(H=>H.checked),U=[...N].some(H=>H.checked);J.checked=O,J.indeterminate=U&&!O}else J.checked=!1,J.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{k.clear(),document.querySelectorAll(".row-checkbox").forEach(E=>E.checked=!1);let l=document.getElementById("select-all-checkbox");l&&(l.checked=!1),a()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(k.size===0)return;let l=[...k],E=document.createElement("div");E.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",E.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${l.length} ${d}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${l.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(E),E.querySelector("#bulk-cancel-btn").addEventListener("click",()=>E.remove()),E.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let D=E.querySelector("#bulk-confirm-btn");D.disabled=!0,D.textContent="Menghapus...";let z=await R(`${i}/bulk`,{method:"DELETE",body:JSON.stringify({ids:l})});E.remove(),z.ok?(G(`${l.length} ${d} berhasil dihapus.`),k.clear(),a(),ye(i),L()):q(z.data?.error||"Gagal menghapus data.")})});let c=document.getElementById("filter-search"),h;c?.addEventListener("input",l=>{clearTimeout(h),h=setTimeout(()=>{x.search=l.target.value,y=1,k.clear(),a(),L()},400)}),b?.forEach(l=>{(l.type==="select"||l.type==="combobox")&&(document.getElementById(`filter-${l.name}`)?.addEventListener("change",E=>{x[l.name]=E.target.value;let D=document.getElementById(`filter-sheet-${l.name}`);D&&(D.value=E.target.value),y=1,k.clear(),a(),L()}),document.getElementById(`filter-sheet-${l.name}`)?.addEventListener("change",E=>{x[l.name]=E.target.value;let D=document.getElementById(`filter-${l.name}`);D&&(D.value=E.target.value),y=1,k.clear(),a(),L(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{x={},c&&(c.value=""),b?.forEach(l=>{let E=document.getElementById(`filter-${l.name}`);E&&(E.value="");let D=document.getElementById(`filter-sheet-${l.name}`);D&&(D.value="")}),y=1,k.clear(),a(),L()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{x={},c&&(c.value=""),b?.forEach(l=>{let E=document.getElementById(`filter-${l.name}`);E&&(E.value="");let D=document.getElementById(`filter-sheet-${l.name}`);D&&(D.value="")}),y=1,k.clear(),a(),L(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>P(null)),u&&document.addEventListener("click",function(l){let E=document.getElementById("aksi-menu-main"),D=document.getElementById("btn-aksi-main");E&&D&&!D.contains(l.target)&&!E.contains(l.target)&&E.classList.remove("show-aksi-menu")});let v=document.getElementById("btn-mobile-filter"),I=document.getElementById("filter-options-wrapper"),j=document.getElementById("btn-close-filter-sheet");if(v&&I&&(v.addEventListener("click",l=>{l.preventDefault(),I.classList.add("sheet-open")}),j&&j.addEventListener("click",l=>{l.preventDefault(),I.classList.remove("sheet-open")})),u){document.getElementById(`btn-export-${u.moduleName}`)?.addEventListener("click",async E=>{let D=E.target,z=D.innerHTML;D.innerHTML="\u23F3 Loading...",D.disabled=!0;try{await u.onExport()}catch{q("Gagal export data")}finally{D.innerHTML=z,D.disabled=!1}}),document.getElementById(`btn-template-${u.moduleName}`)?.addEventListener("click",()=>{u.onTemplate()});let l=document.getElementById(`input-import-${u.moduleName}`);l?.addEventListener("change",async E=>{let D=E.target.files[0];if(!D)return;let z=document.getElementById(`label-import-${u.moduleName}`),J=z?z.querySelector(".import-text"):null,N=J?J.innerText:"";J&&(J.innerText="\u231B Memproses..."),z&&(z.style.pointerEvents="none"),l.disabled=!0;try{let O=await Je(D);if(O.length===0)throw new Error("File kosong atau format salah");await u.onImport(O),G("Import berhasil!"),ye(i),L()}catch(O){q(O.message||"Gagal import data")}finally{J&&(J.innerText=N),z&&(z.style.pointerEvents="auto"),l.disabled=!1,l.value=""}})}async function L(){a();let l=document.getElementById("table-container");if(!l)return;l.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let E=p==="client",D=E?1:y,z=E?Ie:20,J={...s?{sort:s}:{},...Object.fromEntries(Object.entries(x).filter(([,ee])=>ee))},N=new URLSearchParams({page:D,limit:z,...J}),O=await R(`${i}?${N}`);if(!O.ok){l.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${O.data?.error||"Error"}</p></div>`;return}let U=O.data?.data||O.data||[],H=O.data?.pagination,X=U.length;if(E){typeof A=="function"&&(U=A(U));let ee=U.length,Z=20,Y=Math.ceil(ee/Z);y>Y&&Y>0&&(y=Y);let re=(y-1)*Z,pe=y*Z;U=U.slice(re,pe),H={page:y,limit:Z,total:ee,pages:Y}}!1,M&&M(U);let W=rt({columns:r,data:U,onEdit:S?ee=>P(ee):null,actions:B.map(ee=>({...ee,handler:Z=>ee.handler(Z,L)})),emptyText:`Tidak ada ${String(d||"").toLowerCase()}`,bulkSelect:$?{selectedIds:k,onToggle:a}:null});l.innerHTML="",l.appendChild(W);let V=document.getElementById("pagination-container");if(V&&(V.innerHTML="",H&&H.pages>1)){let ee=ot({page:H.page,pages:H.pages,total:H.total,limit:H.limit,onPage:Z=>{y=Z,L()}});ee&&V.appendChild(ee)}}function _(l){let E=typeof m=="function"?m(l):m;return Ge(E)}function P(l){let E=!!l,D=document.createElement("form");if(D.noValidate=!0,D.innerHTML=_(l),E){let J=typeof m=="function"?m(l):m;lt(D,l)}let{close:z}=te({title:E?`Edit ${d}`:`Tambah ${d}`,content:D,size:"lg",confirmText:E?"Simpan Perubahan":`Tambah ${d}`,onConfirm:async(J,N)=>{if(!D.reportValidity())return;let O=J.querySelector(".modal-confirm");O.disabled=!0,O.textContent="Menyimpan...";let U=st(D),H=typeof m=="function"?m(l):m,X=async Z=>{for(let Y of Z)if(Y.type==="row")await X(Y.fields);else if(Y.type==="combobox"&&U[Y.name]){let re=U[Y.name],pe=(Y.options||[]).find(ce=>{let ge=String(typeof ce=="object"?ce.value:ce),Le=String(typeof ce=="object"?ce.label:ce);return ge===re||Le===re});if(pe)U[Y.name]=typeof pe=="object"?pe.value:pe;else if(Y.createApi){let ce={};ce[Y.createApi.field]=re,Y.createApi.extra&&Object.assign(ce,Y.createApi.extra);let ge=await R(Y.createApi.path,{method:"POST",body:JSON.stringify(ce)});if(ge.ok&&ge.data?.id)U[Y.name]=ge.data.id;else if(ge.ok&&!ge.data?.id)U[Y.name]=re;else throw new Error(`Gagal membuat master data: ${ge.data?.error||"Unknown error"}`)}}};try{await X(H)}catch(Z){q(Z.message),O.disabled=!1,O.textContent=E?"Simpan Perubahan":`Tambah ${d}`;return}C&&(U=await C(U,l));let W=E?"PUT":"POST",V=E?`${i}/${l.id}`:i,ee=await R(V,{method:W,body:JSON.stringify(U)});ee.ok?(G(E?`${d} berhasil diperbarui.`:`${d} berhasil ditambahkan.`),N(),ye(i),L()):(q(ee.data?.error||"Gagal menyimpan data."),O.disabled=!1,O.textContent=E?"Simpan Perubahan":`Tambah ${d}`)}})}function K(l){ie(`Hapus ${d} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let E=await R(`${i}/${l.id}`,{method:"DELETE"});E.ok?(G(`${d} berhasil dihapus.`),ye(i),L()):q(E.data?.error||"Gagal menghapus.")},`Hapus ${d}`)}return L(),L}var ze=null,pt=null;async function Te(e=!1){if(ze&&!e)return console.log("Employees Raw (Cache Hit)",ze.slice(0,5)),ze;let t=await R(`/api/employees?limit=${Ie}&status=Aktif`);return ze=(t.data?.data||[]).map(n=>({value:n.id,label:n.full_name})),console.log("Employees Raw",t.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",ze.slice(0,5)),ze}async function he(e=!1){let n=(await Te(e)).map(i=>({value:i.label,label:i.label}));return console.log("Employee Options",n.slice(0,5)),n}async function se(e=!1){return pt&&!e||(pt=((await R("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name}))),pt}function fe(e){let t={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!e||e==="-"||String(e).trim()===""?"":`<span class="badge ${t[e]||"badge-neutral"}">${e}</span>`}function ea(e){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[e]||"badge-neutral"}">${e||"-"}</span>`}function Fe(e){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[e]||"badge-neutral"}">${e||"-"}</span>`}de();var Fn=/^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$/,Cn=/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})$/,$n=/^(\d{4})-(\d{2})$/,_n=/^\d{5}$/;function St(e,t,n){if(t<1||t>12||n<1||n>31)return null;let i=new Date(Date.UTC(e,t-1,n));return i.getUTCFullYear()!==e||i.getUTCMonth()!==t-1||i.getUTCDate()!==n?null:i.getTime()}function Ye(e,{allowYearMonth:t=!0}={}){if(e==null)return null;let n=String(e).trim();if(!n||n==="-")return null;let i=n.match(Fn);if(i)return St(Number(i[1]),Number(i[2]),Number(i[3]));if(i=n.match(Cn),i)return St(Number(i[3]),Number(i[2]),Number(i[1]));if(_n.test(n)){let r=Number(n);return r>=1e4?Date.UTC(1899,11,30)+r*864e5:null}return t&&(i=n.match($n),i)?St(Number(i[1]),Number(i[2]),1):null}function ct(e){if(e==null)return null;let t=String(e).trim();if(!t||t==="-")return null;let n=t.includes(" ")&&!t.includes("T")?`${t.replace(" ","T")}Z`:t,i=Date.parse(n);return Number.isFinite(i)?i:Ye(t)}function ut(e,t,n){let i=Ye(e?.[n]),r=Ye(t?.[n]);if(i===null&&r!==null)return 1;if(i!==null&&r===null)return-1;if(i!==null&&r!==null&&r!==i)return r-i;let m=ct(e?.updated_at),b=ct(t?.updated_at);if(m===null&&b!==null)return 1;if(m!==null&&b===null)return-1;if(m!==null&&b!==null&&b!==m)return b-m;let g=ct(e?.created_at),s=ct(t?.created_at);return g===null&&s!==null?1:g!==null&&s===null?-1:g!==null&&s!==null&&s!==g?s-g:Number(t?.id||0)-Number(e?.id||0)}function ta(e){return[...e].sort((t,n)=>{let i=Ye(t,{allowYearMonth:!0}),r=Ye(n,{allowYearMonth:!0});return i===null&&r!==null?1:i!==null&&r===null?-1:i!==null&&r!==null&&r!==i?r-i:String(n||"").localeCompare(String(t||""))})}var Ft=[],Ve=[];function An(e){if(!Array.isArray(e))return"Q3";let t=["Q4","Q3","Q2","Q1"];for(let n of t)if(e.some(i=>i.period===n))return n;return"Q3"}async function aa(e,t){Ft=await se();let n=await he();Ve=["Berlin Ariansyah","Ade Surahman"];let i=B=>B&&!Ve.find(T=>String(typeof T=="object"?T.value:T).toLowerCase()===String(B).toLowerCase())?[...Ve,B]:Ve,r=await R(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),m=B=>{if(!B||B==="-"||String(B).trim()==="")return"";let T=String(B).split("-");return T.length===3&&T[0].length===4?`${T[2]}-${T[1]}-${T[0]}`:B},b=r.data?.data||[],g=An(b),s=t?t.get("dash_filter"):null,d=new Date,F=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`,S={},w=t&&t.get("month")?t.get("month"):null;s==="inspeksi"?S={status:"Done",activity_type:"Inspeksi Hygiene",month:w}:s==="gcdc"?S={status:"Done",activity_type:"GCDC",month:w}:s&&s.startsWith("period_")&&(S={period:s.replace("period_","").toUpperCase()});let C=new Date().getFullYear(),A=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((B,T)=>{let u=String(T+1).padStart(2,"0");return{value:`${C}-${u}`,label:`${B} ${C}`}});be({container:e,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:S,defaultSort:t&&t.get("sort")||new URLSearchParams(window.location.search).get("sort")||"opening_date_desc",onDataLoaded:B=>{let T=t&&t.get("sort")||new URLSearchParams(window.location.search).get("sort")||"opening_date_desc";return T!=="opening_date_desc"&&T!=="date_desc"?[...B||[]]:[...B||[]].sort((u,$)=>ut(u,$,"opening_date"))},columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:B=>ea(B)},{key:"period",label:"Periode",render:B=>Fe(B)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:B=>m(B)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:B=>m(B)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:B=>m(B)},{key:"status",label:"Status",render:B=>fe(B)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:Ft},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:A},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Ve}],formFields:B=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Ft,value:B?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:B?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:B?.period},{name:"pic",label:"PIC",type:"combobox",options:i(B?.pic),value:B?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:B?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:B?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:B?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:B?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:B?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let B=await R(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(B.ok){let T=B.data.data.map(u=>({Cabang:u.branch_name||"",Kegiatan:u.activity_type||"",Periode:u.period||"",PIC:u.pic||"","Tgl Opening":u.opening_date||"","Tgl Target":u.target_date||"","Tgl Selesai":u.completion_date||"",Status:u.status||""}));ne(T,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async B=>{let u=(await R("/api/branches?all=1")).data?.data||[],$=y=>{if(!y)return null;let x=String(y||"").toLowerCase().trim(),k=x.match(/^(\d{3})/);if(k){let c=u.find(h=>String(h.code||"")===k[1]);if(c)return c.id}let a=u.find(c=>String(c.full_name||"").toLowerCase()===x||String(c.code||"").toLowerCase()===x||String(c.name||"").toLowerCase()===x);return a?a.id:null},p=y=>{if(y==null||y==="")return"";if(y instanceof Date&&!isNaN(y.getTime()))return y.toISOString().slice(0,10);let x=String(y).trim();if(x===""||x==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(x))return x.slice(0,10);if(/^\d{4,5}$/.test(x)){let a=Number(x);if(a>2e4&&a<99999){let c=new Date(Date.UTC(1899,11,30)+a*864e5);return isNaN(c.getTime())?"":c.toISOString().slice(0,10)}}let k=x.split(/[\/\-\.]/);if(k.length===3){let[a,c,h]=k.map(v=>v.trim());if(a.length===4&&c.length<=2&&h.length<=2)return`${a}-${c.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&c.length<=2&&a.length<=2)return`${h}-${c.padStart(2,"0")}-${a.padStart(2,"0")}`}return x},o=B.map(y=>({branch_id:$(String(y.Cabang||"").trim()),branch_name:String(y.Cabang||"").trim(),activity_type:String(y.Kegiatan||"").trim(),period:String(y.Periode||"").trim(),pic:String(y.PIC||y.Pic||"").trim(),opening_date:p(y["Tgl Opening"]||y["Tanggal Opening"]||y["Tgl Openir"]),target_date:p(y["Tgl Target"]||y["Tanggal Target"]),completion_date:p(y["Tgl Selesai"]||y["Tanggal Selesai"]),status:String(y.Status||"").trim(),notes:String(y.Catatan||y.Keterangan||"").trim()})).filter(y=>y.activity_type&&y.period),f=await R("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!f.ok)throw new Error(f.data?.error||"Import gagal");return ye("schedule"),f.data}}})}me();var Ct=[{id:1,full_name:"ACHMAD NURYAINI",branch_id:7,branch_name:"007. Bekasi",division:"FACILITY CARE",phone:"081298765431",join_date:"2023-10-09",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:2,full_name:"ADAM JHORDY",branch_id:32,branch_name:"032. Cawang",division:"FACILITY CARE",phone:"085712345678",join_date:"2024-03-18",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:3,full_name:"ADE MUSTIKA REZDAFITRA",branch_id:16,branch_name:"016. Cibubur",division:"FACILITY CARE",phone:"081345678901",join_date:"2022-06-21",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:4,full_name:"ADE SURAHMAN",branch_id:0,branch_name:"000. Management",division:"FACILITY CARE",phone:"081223344556",join_date:"2021-01-15",status:"Aktif",notes:"Trainer & Supervisor Operasional",target_pindah_os:"",target_selesai_os:""},{id:5,full_name:"ADITYA",branch_id:26,branch_name:"026. Tangerang City",division:"FACILITY CARE",phone:"087811223344",join_date:"2024-10-19",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:6,full_name:"ADITYA PRATAMA",branch_id:54,branch_name:"054. Bogor Gunung Batu",division:"FACILITY CARE",phone:"089699887766",join_date:"2025-09-04",status:"Aktif",notes:"Lulusan Training Batch 5 nilai tertinggi",target_pindah_os:"",target_selesai_os:""},{id:7,full_name:"AGUNG SEPTIADI",branch_id:0,branch_name:"000. Management",division:"FC - RELIEFER",phone:"082155667788",join_date:"2026-02-12",status:"Aktif",notes:"Personil cover cabang mobile",target_pindah_os:"",target_selesai_os:""}],na=["FACILITY CARE","SECURITY","FC - RELIEFER"],Bn=["Aktif","Tidak Aktif","Resign","Cut"];function Me(e){if(!e)return"-";try{let t=String(e).trim();if(/^\d{5}(\.\d+)?$/.test(t)){let i=new Date(Math.round((Number(t)-25569)*86400*1e3));if(!isNaN(i.getTime()))return i.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}let n=new Date(t);return isNaN(n.getTime())?t:n.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function ia(e){if(!e||e==="-")return"\u{1F464}";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function ra(e,t){let n=[],i=[],r="all",m="all",b="all",g="",s=!1,d=localStorage.getItem("fcms_employees_view_mode")||"card",F=t?t.get("dash_filter"):null;F==="active"?r="Aktif":F==="reliefer"&&(m="FC - RELIEFER");let S=ae(),w=Q(S,"employees","write"),C=Q(S,"employees","write"),M=Q(S,"employees","delete");try{i=await se()}catch(o){console.error("Error fetching branches:",o)}async function A(){s=!0,p();try{let o=await R("/api/employees?limit=10000");if(o.ok){let f=o.data?.data||o.data||[];n=Array.isArray(f)&&f.length>0?f:Ct}else n=Ct}catch(o){console.warn("Gagal load employees, menggunakan data lokal:",o),n=Ct}finally{s=!1,p()}}async function B(){if(!Q(S,"export")){q("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:o}=await Promise.resolve().then(()=>(de(),xe)),f=n.map((y,x)=>({No:x+1,"Nama Lengkap":y.full_name||"",Cabang:y.branch_name||"",Divisi:y.division||"","No. HP":y.phone||"","Tgl Masuk":y.join_date||"",Status:y.status||"","Target Pindah OS":y.target_pindah_os||"","Target Selesai OS":y.target_selesai_os||"",Catatan:y.notes||""}));o(f,`Data_Karyawan_FCMS_${new Date().toISOString().split("T")[0]}`),G("Berhasil mengekspor data karyawan ke Excel")}catch(o){q("Gagal ekspor Excel: "+o.message)}}function T(o=null){let f=!!o;if(f&&!C){q("Anda tidak memiliki izin untuk mengedit data karyawan.");return}if(!f&&!w){q("Anda tidak memiliki izin untuk menambah data karyawan.");return}let y=!!o,x=o?.division||"FACILITY CARE",k=o?.status||"Aktif",a=document.createElement("div");a.className="emp-modal-body",a.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Lengkap Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <input type="text" id="emp-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: Achmad Nuryaini" value="${o?.full_name?String(o.full_name).replace(/"/g,"&quot;"):""}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              No. Handphone / WhatsApp
            </label>
            <input type="tel" id="emp-modal-phone" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="08xx-xxxx-xxxx" value="${o?.phone?String(o.phone).replace(/"/g,"&quot;"):""}" />
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${i.map(c=>`
                <option value="${c.value}" ${String(o?.branch_id)===String(c.value)?"selected":""}>
                  ${c.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi / Bagian <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-division" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${na.map(c=>`
                <option value="${c}" ${x===c?"selected":""}>${c}</option>
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
              value="${o?.join_date||""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${Bn.map(c=>`
                <option value="${c}" ${k===c?"selected":""}>${c}</option>
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
              value="${o?.target_pindah_os?String(o.target_pindah_os).slice(0,7):""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Target Selesai OS (Opsional)
            </label>
            <input type="month" id="emp-modal-target-selesai" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${o?.target_selesai_os?String(o.target_selesai_os).slice(0,7):""}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan / Keterangan Karyawan
          </label>
          <textarea id="emp-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Catatan keahlian, tugas khusus, riwayat cover...">${o?.notes||""}</textarea>
        </div>
      </div>
    `,te({title:y?"\u270F\uFE0F Edit Data Karyawan":"\u2795 Tambah Data Karyawan",content:a,size:"md",confirmText:y?"Simpan Perubahan":"Simpan Karyawan",confirmClass:"btn-primary",onConfirm:async(c,h)=>{let v=(a.querySelector("#emp-modal-name").value||"").trim(),I=(a.querySelector("#emp-modal-phone").value||"").trim(),j=a.querySelector("#emp-modal-branch").value,L=a.querySelector("#emp-modal-division").value,_=a.querySelector("#emp-modal-joindate").value||null,P=a.querySelector("#emp-modal-status").value,K=a.querySelector("#emp-modal-target-pindah").value||null,l=a.querySelector("#emp-modal-target-selesai").value||null,E=(a.querySelector("#emp-modal-notes").value||"").trim();if(!v||!j||!L||!P){q("Harap lengkapi Nama Lengkap, Cabang, Divisi, dan Status!");return}let D={full_name:v,phone:I,branch_id:Number(j),division:L,join_date:_,status:P,target_pindah_os:K,target_selesai_os:l,notes:E};try{let z;y?z=await R(`/api/employees/${o.id}`,{method:"PUT",body:JSON.stringify(D)}):z=await R("/api/employees",{method:"POST",body:JSON.stringify(D)}),z.ok?(G(y?"Data karyawan berhasil diperbarui!":"Karyawan baru berhasil ditambahkan!"),h(),await A()):q(z.data?.error||"Gagal menyimpan data karyawan")}catch(z){q("Terjadi kesalahan: "+z.message)}}})}async function u(o){ie({title:"Nonaktifkan Karyawan?",message:`Apakah Anda yakin ingin menonaktifkan personil <strong>${o.full_name}</strong> dari cabang <em>${o.branch_name||"cabang"}</em>?`,confirmText:"Ya, Nonaktifkan",confirmClass:"btn-danger",onConfirm:async()=>{try{let f=await R(`/api/employees/${o.id}`,{method:"DELETE"});f.ok?(G("Status karyawan berhasil diubah menjadi Tidak Aktif"),await A()):q(f.data?.error||"Gagal menonaktifkan data")}catch(f){q("Terjadi kesalahan: "+f.message)}}})}function $(){return n.filter(o=>{if(g){let f=g.toLowerCase(),y=(o.full_name||"").toLowerCase().includes(f),x=(o.branch_name||"").toLowerCase().includes(f),k=(o.phone||"").toLowerCase().includes(f),a=(o.division||"").toLowerCase().includes(f),c=(o.notes||"").toLowerCase().includes(f);if(!y&&!x&&!k&&!a&&!c)return!1}return!(r!=="all"&&(r==="Aktif"&&o.status!=="Aktif"||r==="Nonaktif"&&o.status==="Aktif")||m!=="all"&&o.division!==m||b!=="all"&&String(o.branch_id)!==String(b))})}function p(){let o=$(),f=n.length,y=n.filter(l=>l.status==="Aktif").length,x=n.filter(l=>l.division==="FACILITY CARE"&&l.status==="Aktif").length,k=n.filter(l=>(l.division==="SECURITY"||l.division==="FC - RELIEFER")&&l.status==="Aktif").length,a=n.length,c=n.filter(l=>l.status==="Aktif").length,h=n.filter(l=>l.status!=="Aktif").length;e.innerHTML=`
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
              ${w?`
                <button id="emp-btn-add" class="btn btn-primary emp-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Karyawan</span>
                </button>
              `:""}
              
              ${Q(S,"export")?`<button id="emp-btn-excel" class="btn btn-ghost emp-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="emp-btn-refresh" class="btn btn-ghost emp-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${s?"spin-icon":""}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="emp-stats-row">
            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(37, 99, 235, 0.2); color:#93C5FD;">\u{1F465}</div>
              <div>
                <div class="emp-stat-val">${f}</div>
                <div class="emp-stat-lbl">Total Karyawan</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">\u{1F7E2}</div>
              <div>
                <div class="emp-stat-val">${y}</div>
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
                <div class="emp-stat-val">${k}</div>
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
                placeholder="Cari nama karyawan, no HP, cabang..." value="${g?String(g).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="emp-branch-filter" class="emp-select">
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(l=>`
                <option value="${l.value}" ${String(b)===String(l.value)?"selected":""}>
                  ${l.label}
                </option>
              `).join("")}
            </select>

            <select id="emp-division-filter" class="emp-select">
              <option value="all">Semua Divisi</option>
              ${na.map(l=>`
                <option value="${l}" ${m===l?"selected":""}>
                  ${l}
                </option>
              `).join("")}
            </select>

            <div class="emp-view-toggle" title="Pilih Tampilan">
              <button type="button" class="emp-view-btn ${d==="card"?"active":""}" data-view="card">
                <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="emp-view-btn ${d==="table"?"active":""}" data-view="table">
                <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="emp-filter-pills">
            <div class="emp-pill ${r==="all"&&m==="all"?"active":""}" data-status="all" data-div="all">
              <span>Semua Karyawan</span>
              <span class="emp-pill-count">${a}</span>
            </div>
            <div class="emp-pill ${r==="Aktif"?"active":""}" data-status="Aktif" data-div="all">
              <span class="emp-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="emp-pill-count">${c}</span>
            </div>
            <div class="emp-pill ${m==="FC - RELIEFER"?"active":""}" data-status="all" data-div="FC - RELIEFER">
              <span class="emp-pill-dot" style="background:#8B5CF6;"></span>
              <span>FC - Reliefer</span>
              <span class="emp-pill-count">${n.filter(l=>l.division==="FC - RELIEFER").length}</span>
            </div>
            <div class="emp-pill ${r==="Nonaktif"?"active":""}" data-status="Nonaktif" data-div="all">
              <span class="emp-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="emp-pill-count">${h}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive List Container (Card / Table) -->
        <div id="emp-list-container">
          ${o.length===0?`
            <div class="emp-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada karyawan yang cocok</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian, filter cabang, atau filter divisi.</p>
              <button id="emp-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
          `:d==="table"?`
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
                    ${o.map((l,E)=>{let D=l.division||"FACILITY CARE",z="emp-div-fc",J="";D==="SECURITY"?(z="emp-div-sec",J="sec"):D==="FC - RELIEFER"&&(z="emp-div-rel",J="rel");let N=l.status==="Aktif",O=N?"emp-status-aktif":"emp-status-nonaktif",U=N?"\u{1F7E2} Aktif":`\u26AA ${l.status||"Nonaktif"}`;return`
                        <tr data-id="${l.id}">
                          <td style="text-align:center; font-weight:700; color:#64748B;">${E+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="emp-avatar ${J}" style="width:36px; height:36px; border-radius:9px; font-size:0.85rem;">
                                ${ia(l.full_name)}
                              </div>
                              <div>
                                <div style="font-weight:700; color:#0F172A; white-space:nowrap;">${l.full_name||"-"}</div>
                                <div style="font-size:0.75rem; color:#64748B;">ID #${l.id}</div>
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#334155;">
                            \u{1F4CD} ${l.branch_name||"Cabang Tidak Ditentukan"}
                          </td>
                          <td style="white-space:nowrap;">
                            <span class="emp-div-badge ${z}">
                              ${D==="SECURITY"?"\u{1F6E1}\uFE0F":D==="FC - RELIEFER"?"\u{1F504}":"\u{1F9F9}"} ${D}
                            </span>
                          </td>
                          <td style="white-space:nowrap;">
                            ${l.phone?`
                              <a href="tel:${l.phone}" class="emp-phone-link">
                                \u{1F4DE} ${l.phone}
                              </a>
                            `:'<span style="color:#94A3B8;">-</span>'}
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            ${Me(l.join_date)}
                          </td>
                          <td style="max-width:260px;">
                            ${l.target_pindah_os||l.target_selesai_os?`
                              <div style="font-size:0.78rem; font-weight:600; color:#4F46E5; margin-bottom:2px;">
                                \u{1F504} ${l.target_pindah_os?"Pindah: "+Me(l.target_pindah_os):""}
                                ${l.target_selesai_os?(l.target_pindah_os?" \u2192 ":"")+"Selesai: "+Me(l.target_selesai_os):""}
                              </div>
                            `:""}
                            ${l.notes?`
                              <div style="font-size:0.78rem; color:#475569; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${String(l.notes).replace(/"/g,"&quot;")}">
                                \u{1F4DD} ${l.notes}
                              </div>
                            `:!l.target_pindah_os&&!l.target_selesai_os?'<span style="color:#94A3B8;">-</span>':""}
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="emp-status-badge ${O}">
                              ${U}
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <div style="display:inline-flex; align-items:center; gap:6px;">
                              ${l.phone?`
                                <a href="https://wa.me/${l.phone.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="emp-icon-btn" title="Chat WhatsApp" style="color:#16A34A;">
                                  <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                                </a>
                              `:""}
                              ${C?`
                                <button class="emp-icon-btn edit" data-action="edit" data-id="${l.id}" title="Edit Data Karyawan">
                                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              `:""}
                              ${M?`
                                <button class="emp-icon-btn delete" data-action="delete" data-id="${l.id}" title="Nonaktifkan Karyawan">
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
              ${o.map(l=>{let E=l.division||"FACILITY CARE",D="emp-div-fc",z="";E==="SECURITY"?(D="emp-div-sec",z="sec"):E==="FC - RELIEFER"&&(D="emp-div-rel",z="rel");let J=l.status==="Aktif",N=J?"emp-status-aktif":"emp-status-nonaktif",O=J?"\u{1F7E2} Aktif":`\u26AA ${l.status||"Nonaktif"}`;return`
              <div class="emp-card" data-id="${l.id}">
                <!-- Top Badges Row -->
                <div class="emp-card-top">
                  <span class="emp-div-badge ${D}">
                    ${E==="SECURITY"?"\u{1F6E1}\uFE0F":E==="FC - RELIEFER"?"\u{1F504}":"\u{1F9F9}"} ${E}
                  </span>
                  <div style="display:flex; align-items:center; gap:8px;">
                    <span class="emp-status-badge ${N}">
                      ${O}
                    </span>
                  </div>
                </div>

                <!-- Profile Box -->
                <div class="emp-profile-box">
                  <div class="emp-avatar ${z}">
                    ${ia(l.full_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="emp-name" style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${l.full_name||"-"}
                    </div>
                    <div class="emp-subtitle">
                      Personil ${E}
                    </div>
                  </div>
                </div>

                <!-- Information Grid -->
                <div class="emp-info-card">
                  <div class="emp-info-row">
                    <span class="emp-info-label">\u{1F4CD} Cabang:</span>
                    <span class="emp-info-val">${l.branch_name||"Cabang Tidak Ditentukan"}</span>
                  </div>
                  <div class="emp-info-row">
                    <span class="emp-info-label">\u{1F4DE} No. HP:</span>
                    <span class="emp-info-val">
                      ${l.phone?`
                        <a href="tel:${l.phone}" class="emp-phone-link">
                          <span>${l.phone}</span>
                        </a>
                      `:'<span style="color:#94A3B8;">-</span>'}
                    </span>
                  </div>
                  <div class="emp-info-row">
                    <span class="emp-info-label">\u{1F5D3}\uFE0F Tgl Masuk:</span>
                    <span class="emp-info-val">${Me(l.join_date)}</span>
                  </div>
                  ${l.target_pindah_os||l.target_selesai_os?`
                    <div class="emp-info-row target-os" style="border-top:1px dashed #E2E8F0; padding-top:6px; margin-top:2px; flex-direction:column; align-items:flex-start; gap:3px; width:100%;">
                      <span class="emp-info-label" style="font-weight:600; color:#475569;">\u{1F504} Target Migrasi OS:</span>
                      <span class="emp-info-val" style="font-size:0.78rem; color:#4F46E5; text-align:left; width:100%; word-break:break-word; overflow-wrap:anywhere;">
                        ${l.target_pindah_os?"Pindah: "+Me(l.target_pindah_os):""}
                        ${l.target_selesai_os?(l.target_pindah_os?" \u2192 ":"")+"Selesai: "+Me(l.target_selesai_os):""}
                      </span>
                    </div>
                  `:""}
                </div>

                <!-- Notes box (if any) -->
                ${l.notes?`
                  <div class="emp-notes-box">
                    \u{1F4DD} <strong>Catatan:</strong> ${l.notes}
                  </div>
                `:""}

                <!-- Footer Actions -->
                <div class="emp-card-footer">
                  <div style="font-size:0.75rem; color:#94A3B8;">
                    ID #${l.id}
                  </div>

                  <div class="emp-action-btns">
                    ${l.phone?`
                      <a href="https://wa.me/${l.phone.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="emp-icon-btn" title="Chat WhatsApp" style="color:#16A34A;">
                        <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                      </a>
                    `:""}

                    ${C?`
                      <button class="emp-icon-btn edit" data-action="edit" data-id="${l.id}" title="Edit Data Karyawan">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                    `:""}
                    ${M?`
                      <button class="emp-icon-btn delete" data-action="delete" data-id="${l.id}" title="Nonaktifkan Karyawan">
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
    `,e.querySelectorAll(".emp-view-btn").forEach(l=>{l.addEventListener("click",()=>{let E=l.getAttribute("data-view");E&&E!==d&&(d=E,localStorage.setItem("fcms_employees_view_mode",E),p())})});let v=e.querySelector("#emp-search-input");v&&v.addEventListener("input",l=>{g=l.target.value,p();let E=e.querySelector("#emp-search-input");E&&(E.focus(),E.setSelectionRange(E.value.length,E.value.length))});let I=e.querySelector("#emp-branch-filter");I&&I.addEventListener("change",l=>{b=l.target.value,p()});let j=e.querySelector("#emp-division-filter");j&&j.addEventListener("change",l=>{m=l.target.value,p()}),e.querySelectorAll(".emp-pill").forEach(l=>{l.addEventListener("click",()=>{r=l.dataset.status;let E=l.dataset.div;E&&E!=="all"?m=E:l.dataset.status==="all"&&(m="all"),p()})});let L=e.querySelector("#emp-reset-filter");L&&L.addEventListener("click",()=>{g="",b="all",m="all",r="all",p()});let _=e.querySelector("#emp-btn-add");_&&_.addEventListener("click",()=>T());let P=e.querySelector("#emp-btn-excel");P&&P.addEventListener("click",()=>B());let K=e.querySelector("#emp-btn-refresh");K&&K.addEventListener("click",()=>A()),e.querySelectorAll(".emp-icon-btn.edit").forEach(l=>{l.addEventListener("click",E=>{let D=Number(E.currentTarget.getAttribute("data-id")),z=n.find(J=>J.id===D);z&&T(z)})}),e.querySelectorAll(".emp-icon-btn.delete").forEach(l=>{l.addEventListener("click",E=>{let D=Number(E.currentTarget.getAttribute("data-id")),z=n.find(J=>J.id===D);z&&u(z)})})}await A()}me();de();function We(e){if(!e||e==="-")return"-";if(String(e).startsWith("2099"))return"Tetap / PKWTT";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Ce(e){return e?String(e).startsWith("2099"):!1}function oa(e){if(!e||e==="-")return"\u{1F464}";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function sa(e,t){let n=[],i=[],r=[],m={},b="",g="all",s="all",d="all",F="all",S=!1,w=localStorage.getItem("fcms_contracts_view_mode")||"card",C=t?t.get("dash_filter"):null;C==="expiring30"?(F="30",d="Aktif"):C==="active"&&(d="Aktif");let M=ae(),A=Q(M,"contracts","write"),B=Q(M,"contracts","write"),T=Q(M,"contracts","delete");try{let[a,c]=await Promise.all([se(),Te()]);i=Array.isArray(a)?a:[],r=Array.isArray(c)?c:[],r.forEach(h=>{let v=h.value!==void 0?h.value:h.id,I=h.label||h.full_name;v&&h.phone&&(m[v]=h.phone),I&&h.phone&&(m[I.toLowerCase()]=h.phone)})}catch(a){console.error("Error fetching cache options for contracts:",a)}async function u(){S=!0,k();try{let a=await R("/api/contracts?limit=10000");if(a.ok){let c=a.data?.data||a.data||[];n=Array.isArray(c)?c:[]}else n=[],q("Gagal memuat data kontrak dari server.")}catch(a){console.error("Error load contracts:",a),n=[],q("Terjadi kesalahan koneksi saat memuat kontrak.")}finally{S=!1,k()}}function $(){return n.filter(a=>{if(b.trim()){let c=b.toLowerCase().trim(),h=String(a.employee_name||"").toLowerCase(),v=String(a.branch_name||"").toLowerCase(),I=String(a.pkwt_number||"").toLowerCase(),j=String(a.contract_type||"").toLowerCase(),L=String(a.notes||"").toLowerCase();if(!(h.includes(c)||v.includes(c)||I.includes(c)||j.includes(c)||L.includes(c)))return!1}if(g!=="all"&&String(a.branch_id)!==String(g)||s!=="all"&&a.division!==s||d==="Aktif"&&a.status!=="Aktif")return!1;if(d==="Expiring30"){if(!!!(a.end_date&&String(a.end_date).trim()&&a.end_date!=="-"))return!1;let h=Ce(a.end_date),v=a.days_remaining!==null&&a.days_remaining!==void 0&&a.days_remaining!==""?Number(a.days_remaining):null;if(a.status!=="Aktif"||h||v===null||isNaN(v)||v<0||v>30)return!1}if(d==="Tetap"&&(!Ce(a.end_date)||a.status!=="Aktif")||d==="Nonaktif"&&a.status==="Aktif")return!1;if(F!=="all"){if(!!!(a.end_date&&String(a.end_date).trim()&&a.end_date!=="-"))return!1;let h=parseInt(F,10),v=Ce(a.end_date),I=a.days_remaining!==null&&a.days_remaining!==void 0&&a.days_remaining!==""?Number(a.days_remaining):null;if(v||a.status!=="Aktif"||I===null||isNaN(I)||I<0||I>h)return!1}return!0})}async function p(){te({title:"\u{1F50D} Analisis Selisih Kontrak Karyawan",content:`
        <div style="text-align:center; padding:24px 10px;">
          <div style="font-size:2rem; margin-bottom:10px;">\u23F3</div>
          <p style="color:#64748B; font-size:0.9rem;">Sedang menyinkronkan data karyawan aktif dengan kontrak aktif di database...</p>
        </div>
      `,cancelText:"Tutup"});try{let[a,c]=await Promise.all([R("/api/employees?limit=10000&status=Aktif"),R("/api/contracts?limit=10000")]),h=a.ok?a.data?.data||a.data||[]:[],v=c.ok?c.data?.data||c.data||[]:[],I=new Set(v.filter(_=>_.status==="Aktif").map(_=>_.employee_id)),j=h.filter(_=>!I.has(_.id)),L="";j.length===0?L=`
          <div style="text-align:center; padding:20px 10px;">
            <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F389}</div>
            <h4 style="font-size:1.1rem; font-weight:800; color:#10B981; margin:0 0 8px 0;">Semua Karyawan Aktif Memiliki Kontrak!</h4>
            <p style="font-size:0.86rem; color:#64748B; margin:0;">
              Total <strong>${h.length}</strong> karyawan aktif tercatat memiliki perikatan kontrak aktif yang sah di sistem.
            </p>
          </div>
        `:L=`
          <div style="display:flex; flex-direction:column; gap:12px;">
            <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:10px; padding:12px; font-size:0.84rem; color:#991B1B;">
              Ditemukan <strong>${j.length}</strong> dari <strong>${h.length}</strong> karyawan aktif yang <strong>belum memiliki Kontrak Aktif</strong>.
            </div>
            <div style="max-height:360px; overflow-y:auto; display:flex; flex-direction:column; gap:8px; padding-right:4px;">
              ${j.map(_=>{let P=v.filter(l=>l.employee_id===_.id),K='<span style="color:#D97706; font-weight:600;">Belum pernah ada data kontrak</span>';if(P.length>0){let l=P[0];K=`Riwayat: Status <strong>${l.status}</strong> (Selesai: ${We(l.end_date)})`}return`
                  <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:8px; padding:10px 12px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">
                    <div>
                      <div style="font-weight:700; font-size:0.88rem; color:#0F172A;">${_.full_name}</div>
                      <div style="font-size:0.76rem; color:#64748B;">
                        Cabang: ${_.branch_name||"-"} \u2022 Divisi: ${_.division||"FACILITY CARE"}
                      </div>
                      <div style="font-size:0.75rem; color:#475569; margin-top:2px;">
                        ${K}
                      </div>
                    </div>
                    ${B?`
                      <button class="btn btn-primary btn-add-contract-for-emp" data-empid="${_.id}" data-empname="${String(_.full_name).replace(/"/g,"&quot;")}" data-branchid="${_.branch_id||""}" data-division="${_.division||"FACILITY CARE"}" style="font-size:0.76rem; padding:5px 10px; border-radius:6px; background:#2563EB;">
                        + Buat Kontrak
                      </button>
                    `:""}
                  </div>
                `}).join("")}
            </div>
          </div>
        `,te({title:"Analisis Selisih Kontrak Karyawan",content:L,cancelText:"Tutup"}),setTimeout(()=>{document.querySelectorAll(".btn-add-contract-for-emp").forEach(_=>{_.addEventListener("click",P=>{let K=P.currentTarget,l=K.getAttribute("data-empid"),E=K.getAttribute("data-empname"),D=K.getAttribute("data-branchid"),z=K.getAttribute("data-division"),J=document.querySelector(".modal-close");J&&J.click(),f({employee_id:l,employee_name:E,branch_id:D,division:z,status:"Aktif"})})})},100)}catch(a){console.error(a),q("Gagal melakukan analisis selisih data kontrak.")}}function o(a,c){if(!a)return"";let h=String(a).split("-").map(Number);if(h.length!==3||!h[0]||!h[1]||!h[2])return"";let v=new Date(h[0],h[1]-1,h[2]);if(isNaN(v.getTime()))return"";if(c==="KONTRAK 6 BULAN")v.setMonth(v.getMonth()+6),v.setDate(v.getDate()-1);else if(c==="KONTRAK 1 TAHUN")v.setFullYear(v.getFullYear()+1),v.setDate(v.getDate()-1);else if(c==="KONTRAK 2 TAHUN")v.setFullYear(v.getFullYear()+2),v.setDate(v.getDate()-1);else return"";let I=v.getFullYear(),j=String(v.getMonth()+1).padStart(2,"0"),L=String(v.getDate()).padStart(2,"0");return`${I}-${j}-${L}`}function f(a=null){let c=!!(a&&a.id);if(c&&!A){q("Anda tidak memiliki izin untuk mengedit kontrak.");return}if(!c&&!B){q("Anda tidak memiliki izin untuk menambah kontrak.");return}let h=!!(a&&a.id),v=h?`\u270F\uFE0F Edit Kontrak: ${a.employee_name||"Karyawan"}`:"\u2795 Tambah Kontrak Karyawan Baru",I=a?.employee_id||"",j=a?.branch_id||"",L=a?.division||"FACILITY CARE",_=a?.status||"Aktif",P=a?.contract_type||"KONTRAK 1 TAHUN",K=a?.pkwt_number||"PKWT 1",l=a?.notes||"",E=!1;a&&(a.contract_type==="PKWTT (TETAP)"||a.pkwt_number==="TETAP"?E=!0:a.contract_type&&(a.contract_type.startsWith("KONTRAK")||a.contract_type.includes("BULAN")||a.contract_type.includes("TAHUN"))||a.pkwt_number&&a.pkwt_number.startsWith("PKWT")?E=!1:E=Ce(a.end_date));let D=a?.start_date?String(a.start_date).slice(0,10):"",z=a?.end_date&&!Ce(a.end_date)?String(a.end_date).slice(0,10):"";!E&&D&&!z&&(z=o(D,P));let J=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-employee" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Karyawan --</option>
              ${r.map(O=>{let U=O.value!==void 0?O.value:O.id,H=O.label||O.full_name||"Karyawan",X=O.division||O.branch_name?` (${O.division||"FC"} - ${O.branch_name||"Cabang -"})`:"";return`
                <option value="${U}" ${String(I)===String(U)?"selected":""}>
                  ${H}${X}
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
              ${i.map(O=>`
                <option value="${O.value}" ${String(j)===String(O.value)?"selected":""}>
                  ${O.label}
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
              <option value="FACILITY CARE" ${L==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${L==="SECURITY"?"selected":""}>SECURITY</option>
              <option value="FC - RELIEFER" ${L==="FC - RELIEFER"?"selected":""}>FC - RELIEFER</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Kontrak <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Aktif" ${_==="Aktif"?"selected":""}>\u{1F7E2} Aktif</option>
              <option value="Tidak Aktif" ${_==="Tidak Aktif"?"selected":""}>\u26AA Tidak Aktif</option>
              <option value="Resign" ${_==="Resign"?"selected":""}>\u{1F534} Resign</option>
              <option value="Cut" ${_==="Cut"?"selected":""}>\u2702\uFE0F Cut / Pemutusan</option>
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
              <option value="PKWT 1" ${K==="PKWT 1"?"selected":""}>PKWT 1</option>
              <option value="PKWT 2" ${K==="PKWT 2"?"selected":""}>PKWT 2</option>
              <option value="PKWT 3" ${K==="PKWT 3"?"selected":""}>PKWT 3</option>
              <option value="PKWT 4" ${K==="PKWT 4"?"selected":""}>PKWT 4</option>
              <option value="PKWT 5" ${K==="PKWT 5"?"selected":""}>PKWT 5</option>
              <option value="PKWT 6" ${K==="PKWT 6"?"selected":""}>PKWT 6</option>
              <option value="TETAP" ${K==="TETAP"?"selected":""}>PKWTT / TETAP</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Mulai Kontrak
            </label>
            <input type="date" id="con-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${D}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Selesai Kontrak
            </label>
            <input type="date" id="con-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; ${E?"opacity:0.5; background-color:#F8FAFC; cursor:not-allowed;":"background-color:#FFFFFF;"}" 
              value="${z}" ${E?"disabled":""} />
            <div style="margin-top:6px; display:flex; align-items:center; gap:6px;">
              <input type="checkbox" id="con-modal-is-tetap" ${E?"checked":""} style="cursor:pointer;" />
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
            placeholder="Keterangan perpanjangan, klausul khusus, atau nomor SK...">${l}</textarea>
        </div>
      </div>
    `,N=te({title:v,content:J,confirmText:h?"\u{1F4BE} Simpan Perubahan":"\u2795 Tambah Kontrak",cancelText:"Batal",onConfirm:async(O,U)=>{let H=O.querySelector("#con-modal-employee"),X=O.querySelector("#con-modal-branch"),W=O.querySelector("#con-modal-division"),V=O.querySelector("#con-modal-status"),ee=O.querySelector("#con-modal-start"),Z=O.querySelector("#con-modal-end"),Y=O.querySelector("#con-modal-is-tetap"),re=O.querySelector("#con-modal-type"),pe=O.querySelector("#con-modal-pkwt"),ce=O.querySelector("#con-modal-notes"),ge=H?.value,Le=X?.value,cn=W?.value||"FACILITY CARE",un=V?.value||"Aktif",ht=ee?.value?.trim()||null,Yt=Y?.checked,gn=Z?.value?.trim(),ft=Yt?"2099-12-31":gn||null,mn=re?.value,bn=pe?.value,hn=ce?.value?.trim()||null;if(!ge){q("Harap pilih nama karyawan!"),H?.focus();return}if(!Le){q("Harap pilih cabang penempatan!"),X?.focus();return}if(!Yt&&ht&&ft&&ft<ht){q("Tanggal selesai kontrak tidak boleh lebih awal dari tanggal mulai kontrak!"),Z?.focus();return}let fn={employee_id:parseInt(ge,10),branch_id:parseInt(Le,10),division:cn,status:un,start_date:ht,end_date:ft,contract_type:mn,pkwt_number:bn,notes:hn};try{let xt=h?`/api/contracts/${a.id}`:"/api/contracts",yt=await R(xt,{method:h?"PUT":"POST",body:JSON.stringify(fn)});if(yt.ok)G(h?"Data kontrak berhasil diperbarui!":"Kontrak baru berhasil ditambahkan!"),U(),await u();else{let xn=yt.data?.error||yt.data?.message||"Gagal menyimpan data kontrak.";q(xn)}}catch(xt){console.error(xt),q("Terjadi kesalahan koneksi ke server.")}}});setTimeout(()=>{let O=N?.overlay||document.querySelector(".modal-overlay:last-child");if(!O)return;let U=O.querySelector("#con-modal-is-tetap"),H=O.querySelector("#con-modal-end"),X=O.querySelector("#con-modal-start"),W=O.querySelector("#con-modal-type"),V=O.querySelector("#con-modal-pkwt");if(!U||!H)return;function ee(Z,Y=!0){Z?(U.checked=!0,H.disabled=!0,H.style.opacity="0.5",H.style.backgroundColor="#F8FAFC",H.style.cursor="not-allowed",H.value="",Y&&(W&&(W.value="PKWTT (TETAP)"),V&&(V.value="TETAP"))):(U.checked=!1,H.disabled=!1,H.style.opacity="1",H.style.backgroundColor="#FFFFFF",H.style.cursor="default",Y&&(W&&W.value==="PKWTT (TETAP)"&&(W.value="KONTRAK 1 TAHUN"),V&&V.value==="TETAP"&&(V.value="PKWT 1")))}U.addEventListener("change",Z=>{ee(Z.target.checked,!0),Z.target.checked||H.focus()}),W&&W.addEventListener("change",Z=>{let Y=Z.target.value;Y==="PKWTT (TETAP)"?(ee(!0,!1),V&&(V.value="TETAP")):(ee(!1,!1),V&&V.value==="TETAP"&&(V.value="PKWT 1"),X&&X.value&&(H.value=o(X.value,Y)))}),V&&V.addEventListener("change",Z=>{Z.target.value==="TETAP"?(W&&(W.value="PKWTT (TETAP)"),ee(!0,!1)):(W&&W.value==="PKWTT (TETAP)"&&(W.value="KONTRAK 1 TAHUN"),ee(!1,!1),X&&X.value&&W&&(H.value=o(X.value,W.value)))}),X&&X.addEventListener("change",Z=>{!U.checked&&Z.target.value&&W&&(H.value=o(Z.target.value,W.value))})},20)}function y(a){te({title:"\u{1F5D1}\uFE0F Hapus Data Kontrak",content:`
        <div style="font-size:0.92rem; color:#334155; line-height:1.5;">
          Apakah Anda yakin ingin menghapus catatan kontrak untuk <strong>${a.employee_name||"karyawan ini"}</strong> (ID #${a.id})?
          <div style="margin-top:10px; padding:10px; background:#FEF2F2; border-radius:8px; border:1px solid #FECACA; font-size:0.8rem; color:#991B1B;">
            \u26A0\uFE0F Tindakan ini akan menghapus riwayat masa berlaku kontrak terkait dari sistem.
          </div>
        </div>
      `,confirmText:"Ya, Hapus",cancelText:"Batal",onConfirm:async(c,h)=>{try{let v=await R(`/api/contracts/${a.id}`,{method:"DELETE"});v.ok?(G("Data kontrak berhasil dihapus."),h(),await u()):q("Gagal menghapus kontrak: "+(v.data?.message||"Terjadi kesalahan"))}catch(v){console.error(v),q("Terjadi kendala koneksi.")}}})}function x(){if(!Q(M,"export")){q("Anda tidak memiliki izin untuk mengekspor data.");return}if(!n.length){q("Belum ada data kontrak untuk diekspor.");return}let a=n.map(c=>{let h=Ce(c.end_date),v=!!(c.end_date&&String(c.end_date).trim()&&c.end_date!=="-"),I=c.days_remaining!==null&&c.days_remaining!==void 0&&c.days_remaining!==""?Number(c.days_remaining):null;return{"Nama Lengkap":c.employee_name||"",Cabang:c.branch_name||"",Divisi:c.division||"","No. PKWT":c.pkwt_number||"","Tipe Kontrak":c.contract_type||"","Tanggal Mulai":c.start_date||"","Tanggal Selesai":h?"Tetap / PKWTT":v?c.end_date:"-","Sisa Hari":h?"Tetap":v&&I!==null&&!isNaN(I)?I+" Hari":"-",Status:c.status||"",Catatan:c.notes||""}});ne(a,`Data_Kontrak_Karyawan_${new Date().toISOString().slice(0,10)}`),G("File Excel Data Kontrak berhasil diunduh!")}function k(){let a=$(),c=n.length,h=n.filter(N=>N.status==="Aktif").length,v=n.filter(N=>{if(!!!(N.end_date&&String(N.end_date).trim()&&N.end_date!=="-"))return!1;let U=Ce(N.end_date),H=N.days_remaining!==null&&N.days_remaining!==void 0&&N.days_remaining!==""?Number(N.days_remaining):null;return N.status==="Aktif"&&!U&&H!==null&&!isNaN(H)&&H>=0&&H<=30}).length,I=n.filter(N=>Ce(N.end_date)&&N.status==="Aktif").length,j=n.filter(N=>N.status!=="Aktif").length;e.innerHTML=`
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
              ${B?`
                <button id="con-btn-add" class="btn btn-primary con-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Kontrak</span>
                </button>
              `:""}

              ${Q(M,"export")?`
              <button id="con-btn-excel" class="btn btn-ghost con-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>
              `:""}

              ${A?`
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
                <div class="con-stat-val">${c}</div>
                <div class="con-stat-lbl">Total Kontrak</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#34D399;">
                \u{1F7E2}
              </div>
              <div>
                <div class="con-stat-val" style="color:#34D399;">${h}</div>
                <div class="con-stat-lbl">Kontrak Aktif</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FBBF24;">
                \u26A0\uFE0F
              </div>
              <div>
                <div class="con-stat-val" style="color:#FBBF24;">${v}</div>
                <div class="con-stat-lbl">Habis &lt; 30 Hari</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(99, 102, 241, 0.2); color:#A5B4FC;">
                \u{1F3DB}\uFE0F
              </div>
              <div>
                <div class="con-stat-val" style="color:#A5B4FC;">${I}</div>
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
                placeholder="Cari nama karyawan, cabang, no PKWT, catatan..." value="${b?String(b).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="con-branch-filter" class="con-select">
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(N=>`
                <option value="${N.value}" ${String(g)===String(N.value)?"selected":""}>
                  ${N.label}
                </option>
              `).join("")}
            </select>

            <select id="con-division-filter" class="con-select">
              <option value="all">Semua Divisi</option>
              <option value="FACILITY CARE" ${s==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${s==="SECURITY"?"selected":""}>SECURITY</option>
              <option value="FC - RELIEFER" ${s==="FC - RELIEFER"?"selected":""}>FC - RELIEFER</option>
            </select>

            <select id="con-expiring-filter" class="con-select">
              <option value="all" ${F==="all"?"selected":""}>Semua Durasi Masa Habis</option>
              <option value="7" ${F==="7"?"selected":""}>Akan Habis \u2264 7 Hari</option>
              <option value="14" ${F==="14"?"selected":""}>Akan Habis \u2264 14 Hari</option>
              <option value="30" ${F==="30"?"selected":""}>Akan Habis \u2264 30 Hari</option>
              <option value="60" ${F==="60"?"selected":""}>Akan Habis \u2264 60 Hari</option>
            </select>

            <div class="con-view-toggle">
              <button type="button" class="con-view-btn ${w==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="con-view-btn ${w==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="con-filter-pills">
            <div class="con-pill ${d==="all"&&F==="all"?"active":""}" data-status="all">
              <span>Semua Kontrak</span>
              <span class="con-pill-count">${c}</span>
            </div>
            <div class="con-pill ${d==="Aktif"&&F==="all"?"active":""}" data-status="Aktif">
              <span class="con-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="con-pill-count">${h}</span>
            </div>
            <div class="con-pill ${d==="Expiring30"?"active":""}" data-status="Expiring30">
              <span class="con-pill-dot" style="background:#F59E0B;"></span>
              <span>Akan Habis (&lt;30 Hari)</span>
              <span class="con-pill-count">${v}</span>
            </div>
            <div class="con-pill ${d==="Tetap"?"active":""}" data-status="Tetap">
              <span class="con-pill-dot" style="background:#6366F1;"></span>
              <span>PKWTT (Tetap)</span>
              <span class="con-pill-count">${I}</span>
            </div>
            <div class="con-pill ${d==="Nonaktif"?"active":""}" data-status="Nonaktif">
              <span class="con-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="con-pill-count">${j}</span>
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
          `:w==="table"?`
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
                    ${a.map((N,O)=>{let U=N.division||"FACILITY CARE",H="con-div-fc",X="";U==="SECURITY"?(H="con-div-sec",X="sec"):U==="FC - RELIEFER"&&(H="con-div-rel",X="rel");let W=!!(N.end_date&&String(N.end_date).trim()&&N.end_date!=="-"),V=Ce(N.end_date),ee=N.status==="Aktif",Z=N.days_remaining!==null&&N.days_remaining!==void 0&&N.days_remaining!==""?Number(N.days_remaining):null,Y=W&&!V&&Z!==null&&!isNaN(Z)?Z:null,re="con-status-aktif",pe="\u{1F7E2} Aktif";V?(re="con-status-tetap",pe="\u{1F3DB}\uFE0F PKWTT (Tetap)"):ee?Y!==null?Y<=0?(re="con-status-danger",pe="\u26A0\uFE0F Habis Hari Ini"):Y<=30?(re="con-status-warning",pe=`\u26A0\uFE0F Sisa ${Y} Hari`):(re="con-status-aktif",pe=`\u{1F7E2} Aktif (${Y} Hari)`):(re="con-status-aktif",pe="\u{1F7E2} Aktif"):(re="con-status-danger",pe=`\u26AA ${N.status||"Nonaktif"}`);let ce=m[N.employee_id]||m[String(N.employee_name||"").toLowerCase()]||"";return`
                        <tr data-id="${N.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${O+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="con-avatar ${X}" style="width:36px; height:36px; font-size:0.8rem; border-radius:9px;">
                                ${oa(N.employee_name)}
                              </div>
                              <div>
                                <div style="font-weight:700; color:#0F172A; white-space:nowrap;">${N.employee_name||"Karyawan Tanpa Nama"}</div>
                                <div style="font-size:0.74rem; color:#64748B;">${N.contract_type||"Kontrak Kerja"} \u2022 ID #${N.id}</div>
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#0F172A; font-size:0.84rem; margin-bottom:4px;">\u{1F4CD} ${N.branch_name||"Cabang Tidak Ditentukan"}</div>
                            <span class="con-div-badge ${H}" style="padding:2px 8px; font-size:0.7rem;">
                              ${U==="SECURITY"?"\u{1F6E1}\uFE0F":U==="FC - RELIEFER"?"\u{1F504}":"\u{1F9F9}"} ${U}
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            ${N.pkwt_number?`
                              <span class="con-pkwt-badge">
                                \u{1F4DC} ${N.pkwt_number}
                              </span>
                            `:'<span style="color:#94A3B8;">-</span>'}
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            ${We(N.start_date)}
                          </td>
                          <td style="white-space:nowrap; font-weight:600; ${V?"color:#4338CA; font-weight:700;":"color:#1E293B;"}">
                            ${V?"Tetap / PKWTT":W?We(N.end_date):"-"}
                          </td>
                          <td style="white-space:nowrap; font-size:0.82rem; font-weight:600; ${V?"color:#15803D;":Y!==null&&Y<=30&&ee?"color:#D97706; font-weight:700;":"color:#475569;"}">
                            ${V?"Karyawan Tetap (Permanen)":Y!==null?Y<=0?"Masa Kontrak Telah Habis":`${Y} Hari Tersisa`:"-"}
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="con-status-badge ${re}">
                              ${pe}
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <div style="display:inline-flex; align-items:center; gap:6px;">
                              ${ce?`
                                <a href="https://wa.me/${ce.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="con-icon-btn" title="Chat WhatsApp Karyawan" style="color:#16A34A;">
                                  <svg width="15" height="15" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                                </a>
                              `:""}
                              ${A?`
                                <button class="con-icon-btn edit" data-action="edit" data-id="${N.id}" title="Edit Data Kontrak">
                                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              `:""}
                              ${T?`
                                <button class="con-icon-btn delete" data-action="delete" data-id="${N.id}" title="Hapus Kontrak">
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
              ${a.map(N=>{let O=N.division||"FACILITY CARE",U="con-div-fc",H="";O==="SECURITY"?(U="con-div-sec",H="sec"):O==="FC - RELIEFER"&&(U="con-div-rel",H="rel");let X=!!(N.end_date&&String(N.end_date).trim()&&N.end_date!=="-"),W=Ce(N.end_date),V=N.status==="Aktif",ee=N.days_remaining!==null&&N.days_remaining!==void 0&&N.days_remaining!==""?Number(N.days_remaining):null,Z=X&&!W&&ee!==null&&!isNaN(ee)?ee:null,Y="con-status-aktif",re="\u{1F7E2} Aktif";W?(Y="con-status-tetap",re="\u{1F3DB}\uFE0F PKWTT (Tetap)"):V?Z!==null?Z<=0?(Y="con-status-danger",re="\u26A0\uFE0F Habis Hari Ini"):Z<=30?(Y="con-status-warning",re=`\u26A0\uFE0F Sisa ${Z} Hari`):(Y="con-status-aktif",re=`\u{1F7E2} Aktif (${Z} Hari)`):(Y="con-status-aktif",re="\u{1F7E2} Aktif"):(Y="con-status-danger",re=`\u26AA ${N.status||"Nonaktif"}`);let pe=m[N.employee_id]||m[String(N.employee_name||"").toLowerCase()]||"";return`
              <div class="con-card" data-id="${N.id}">
                <!-- Top Badges Row -->
                <div class="con-card-top">
                  <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
                    <span class="con-div-badge ${U}">
                      ${O==="SECURITY"?"\u{1F6E1}\uFE0F":O==="FC - RELIEFER"?"\u{1F504}":"\u{1F9F9}"} ${O}
                    </span>
                    ${N.pkwt_number?`
                      <span class="con-pkwt-badge">
                        \u{1F4DC} ${N.pkwt_number}
                      </span>
                    `:""}
                  </div>
                  <span class="con-status-badge ${Y}">
                    ${re}
                  </span>
                </div>

                <!-- Profile Box -->
                <div class="con-profile-box">
                  <div class="con-avatar ${H}">
                    ${oa(N.employee_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="con-name">
                      ${N.employee_name||"Karyawan Tanpa Nama"}
                    </div>
                    <div class="con-subtitle">
                      Personil ${O} \u2022 ${N.contract_type||"Kontrak Kerja"}
                    </div>
                  </div>
                </div>

                <!-- Information Grid -->
                <div class="con-info-card">
                  <div class="con-info-row">
                    <span class="con-info-label">\u{1F4CD} Cabang:</span>
                    <span class="con-info-val">${N.branch_name||"Cabang Tidak Ditentukan"}</span>
                  </div>
                  <div class="con-info-row">
                    <span class="con-info-label">\u{1F5D3}\uFE0F Mulai Kontrak:</span>
                    <span class="con-info-val">${We(N.start_date)}</span>
                  </div>
                  <div class="con-info-row">
                    <span class="con-info-label">\u{1F3C1} Selesai Kontrak:</span>
                    <span class="con-info-val" style="${W?"color:#4338CA; font-weight:700;":""}">
                      ${W?"Tetap / PKWTT":X?We(N.end_date):"-"}
                    </span>
                  </div>
                  <div class="con-info-row">
                    <span class="con-info-label">\u23F3 Masa Berlaku:</span>
                    <span class="con-info-val" style="${W?"color:#15803D;":Z!==null&&Z<=30&&V?"color:#D97706; font-weight:700;":""}">
                      ${W?"Karyawan Tetap (Permanen)":Z!==null?Z<=0?"Masa Kontrak Telah Habis":`${Z} Hari Tersisa`:"-"}
                    </span>
                  </div>
                </div>

                <!-- Notes box (if any) -->
                ${N.notes?`
                  <div class="con-notes-box">
                    \u{1F4DD} <strong>Catatan:</strong> ${N.notes}
                  </div>
                `:""}

                <!-- Footer Actions -->
                <div class="con-card-footer">
                  <div style="font-size:0.75rem; color:#94A3B8;">
                    ID #${N.id} \u2022 ${N.pkwt_number||"PKWT"}
                  </div>

                  <div class="con-action-btns">
                    ${pe?`
                      <a href="https://wa.me/${pe.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="con-icon-btn" title="Chat WhatsApp Karyawan" style="color:#16A34A;">
                        <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                      </a>
                    `:""}

                    ${A?`
                      <button class="con-icon-btn edit" data-action="edit" data-id="${N.id}" title="Edit Data Kontrak">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                    `:""}
                    ${T?`
                      <button class="con-icon-btn delete" data-action="delete" data-id="${N.id}" title="Hapus Kontrak">
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
    `,e.querySelectorAll(".con-view-btn").forEach(N=>{N.addEventListener("click",()=>{let O=N.getAttribute("data-view");O&&O!==w&&(w=O,localStorage.setItem("fcms_contracts_view_mode",O),k())})});let L=e.querySelector("#con-search-input");L&&L.addEventListener("input",N=>{b=N.target.value,k()});let _=e.querySelector("#con-branch-filter");_&&_.addEventListener("change",N=>{g=N.target.value,k()});let P=e.querySelector("#con-division-filter");P&&P.addEventListener("change",N=>{s=N.target.value,k()});let K=e.querySelector("#con-expiring-filter");K&&K.addEventListener("change",N=>{F=N.target.value,k()}),e.querySelectorAll(".con-pill").forEach(N=>{N.addEventListener("click",O=>{let U=O.currentTarget.getAttribute("data-status");d=U,U==="Expiring30"?F="30":F="all",k()})});let l=e.querySelector("#con-reset-filter");l&&l.addEventListener("click",()=>{b="",g="all",s="all",d="all",F="all",k()});let E=e.querySelector("#con-btn-add");E&&E.addEventListener("click",()=>f());let D=e.querySelector("#con-btn-excel");D&&D.addEventListener("click",x);let z=e.querySelector("#con-btn-missing");z&&z.addEventListener("click",p);let J=e.querySelector("#con-btn-refresh");J&&J.addEventListener("click",()=>u()),e.querySelectorAll(".con-icon-btn.edit").forEach(N=>{N.addEventListener("click",O=>{let U=parseInt(O.currentTarget.getAttribute("data-id"),10),H=n.find(X=>X.id===U);H&&f(H)})}),e.querySelectorAll(".con-icon-btn.delete").forEach(N=>{N.addEventListener("click",O=>{let U=parseInt(O.currentTarget.getAttribute("data-id"),10),H=n.find(X=>X.id===U);H&&y(H)})})}await u()}me();var $t=[{id:1,report_date:"2026-08-27",branch_id:19,branch_name:"019. Cinere",category:"SDM",source:"AM",complaint:"Keluhan AM fc atas nama Malik bau badan, dan sudah dibilangin berulang kali terkait kebersihan seragam dan deodoran.",employee_name:"MALIK",fc_specialist:"Mizwar",solution:"Diberikan teguran lisan, edukasi kebersihan diri, dan pendampingan seragam cadangan di klinik.",status:"Done",completion_date:"2026-08-28",day_count:1},{id:2,report_date:"2026-08-26",branch_id:0,branch_name:"000. Management",category:"SDM",source:"Berlin",complaint:"Security atas nama Reza suhelmi meninggalkan area tugas jaga tanpa personil pengganti saat jam operasional klinik.",employee_name:"REZA SUHELMI",fc_specialist:"Ade Surahman",solution:"Pemanggilan investigasi dan penerbitan Surat Peringatan (SP) sesuai SOP security.",status:"Done",completion_date:"2026-08-27",day_count:1},{id:3,report_date:"2026-07-27",branch_id:30,branch_name:"030. Ciputat",category:"SDM",source:"AM",complaint:"Terinfo kembali permasalahan FC Ciputat, terkait standar pembersihan area dental unit yang belum memenuhi checklist hygiene.",employee_name:"HENDRA",fc_specialist:"Berlin Ariansyah",solution:"Re-training modul pembersihan dental unit dan monitoring harian oleh supervisor.",status:"In Progress",completion_date:null,day_count:null},{id:4,report_date:"2026-06-24",branch_id:30,branch_name:"030. Ciputat",category:"SDM",source:"AM",complaint:"FC atas nama Hendra mengintimidasi rcp pada saat briefing pagi mengenai pembagian area kerja.",employee_name:"HENDRA",fc_specialist:"Berlin Ariansyah",solution:"Mediasi bersama AM dan HR, diberikan pembinaan konseling perilaku kerja.",status:"Done",completion_date:"2026-06-25",day_count:1},{id:5,report_date:"2026-06-05",branch_id:0,branch_name:"000. Management",category:"SDM",source:"Berlin",complaint:"Kinerja kurang dan ada komplain dari cabang mengenai ketidakhadiran berkala personil pengganti.",employee_name:"Indra",fc_specialist:"Ade",solution:"Evaluasi penempatan dan pembinaan kedisiplinan kerja.",status:"Done",completion_date:"2026-06-08",day_count:3},{id:6,report_date:"2026-06-05",branch_id:0,branch_name:"000. Management",category:"SDM",source:"Ade",complaint:"Sering tidak masuk kerja dan memalsukan surat dokter saat izin sakit.",employee_name:"Kristian",fc_specialist:"Ade Surahman",solution:"Investigasi ke klinik dokter terkait dan eskalasi ke HRD untuk tindakan tegas.",status:"Done",completion_date:"2026-06-07",day_count:2},{id:7,report_date:"2026-05-25",branch_id:19,branch_name:"019. Cinere",category:"SDM",source:"SPV",complaint:"Tidak masuk kerja tanpa kabar (alpa), dan sudah di-one on one oleh SPV terkait absensi.",employee_name:"Malik",fc_specialist:"Mizwar",solution:"Pemberian SP 1 dan penandatanganan komitmen kehadiran kerja.",status:"Done",completion_date:"2026-05-26",day_count:1}],la=["SDM","Cleaning","Aset","K3","Lainnya"],Tn=["Open","In Progress","Done"],Dn=["SPV","AM","RCP","Perawat","FC","Security","Berlin","Ade","Pattrel","Dentrel"];function da(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Pn(e){if(!e||e==="-")return"\u26A0\uFE0F";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function pa(e,t){let n=[],i=[],r=[],m="all",b="all",g="all",s="",d=!1;(t?t.get("dash_filter"):null)==="open"&&(m="Open");let S=ae(),w=Q(S,"issues","write");try{[i,r]=await Promise.all([se(),he()])}catch(p){console.error("Error fetching cache:",p)}async function C(){d=!0,$();try{let p=await R("/api/issues?limit=10000");if(p.ok){let o=p.data?.data||p.data||[];n=Array.isArray(o)&&o.length>0?o:$t}else n=$t}catch(p){console.warn("Gagal load issues, menggunakan data lokal:",p),n=$t}finally{d=!1,$()}}async function M(){if(!Q(S,"export")){q("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:p}=await Promise.resolve().then(()=>(de(),xe)),o=n.map((f,y)=>({No:y+1,"Tanggal Info":f.report_date||"",Cabang:f.branch_name||"",Kategori:f.category||"","Sumber Laporan":f.source||"",Keluhan:f.complaint||"","Nama FC / Security":f.employee_name||"","FC Spesialis":f.fc_specialist||"","Solusi / Tindak Lanjut":f.solution||"",Status:f.status||"","Tanggal Selesai":f.completion_date||"","Durasi Hari":f.day_count!=null?f.day_count:""}));p(o,`Data_Permasalahan_FCMS_${new Date().toISOString().split("T")[0]}`),G("Berhasil mengekspor data permasalahan ke Excel")}catch(p){q("Gagal ekspor Excel: "+p.message)}}async function A(p,o){try{let f=o==="Done"?p.completion_date||new Date().toISOString().split("T")[0]:o==="Open"?null:p.completion_date,y=p.day_count;if(p.report_date&&f){let a=new Date(p.report_date),c=new Date(f);y=Math.max(0,Math.floor((c-a)/(1e3*60*60*24)))}else o==="Open"&&(y=null);let x={...p,status:o,completion_date:f,day_count:y},k=await R(`/api/issues/${p.id}`,{method:"PUT",body:JSON.stringify(x)});k.ok?(G(`Status masalah diubah ke "${o}"`),await C()):q(k.data?.error||"Gagal memperbarui status")}catch(f){q("Terjadi kesalahan: "+f.message)}}function B(p=null){let o=!!p,f=new Date().toISOString().split("T")[0],y=p?.report_date||f,x=p?.category||"SDM",k=p?.source||"AM",a=p?.status||"Open",c=document.createElement("div");c.className="issue-modal-body",c.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Info <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="iss-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${y}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${i.map(_=>`
                <option value="${_.value}" ${String(p?.branch_id)===String(_.value)?"selected":""}>
                  ${_.label}
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
              ${la.map(_=>`
                <option value="${_}" ${x===_?"selected":""}>${_}</option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Sumber Laporan
            </label>
            <input list="iss-source-options" id="iss-modal-source" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: AM, SPV, Berlin..." value="${k}" />
            <datalist id="iss-source-options">
              ${Dn.map(_=>`<option value="${_}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Detail Keluhan / Permasalahan <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="iss-modal-complaint" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Jelaskan kendala, kronologi, atau komplain..." required>${p?.complaint||""}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Personil FC / Security
            </label>
            <input list="iss-emp-options" id="iss-modal-employee" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Ketik nama karyawan terkait..." value="${p?.employee_name?String(p.employee_name).replace(/"/g,"&quot;"):""}" />
            <datalist id="iss-emp-options">
              ${r.map(_=>`<option value="${_.label}">`).join("")}
            </datalist>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              FC Spesialis / Penanggung Jawab
            </label>
            <input list="iss-spec-options" id="iss-modal-spec" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: Mizwar, Ade, Berlin..." value="${p?.fc_specialist?String(p.fc_specialist).replace(/"/g,"&quot;"):""}" />
            <datalist id="iss-spec-options">
              ${r.map(_=>`<option value="${_.label}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Solusi / Tindak Lanjut Penanganan
          </label>
          <textarea id="iss-modal-solution" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tindakan yang telah atau akan diambil untuk menyelesaikan masalah...">${p?.solution||""}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Masalah <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${Tn.map(_=>`
                <option value="${_}" ${a===_?"selected":""}>${_}</option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Selesai
            </label>
            <input type="date" id="iss-modal-completion" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${p?.completion_date||""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Durasi (Hari)
            </label>
            <input type="number" id="iss-modal-daycount" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; background:#F8FAFC;" 
              value="${p?.day_count!=null?p.day_count:""}" readonly placeholder="Otomatis" />
          </div>
        </div>
      </div>
    `;let h=c.querySelector("#iss-modal-date"),v=c.querySelector("#iss-modal-completion"),I=c.querySelector("#iss-modal-daycount"),j=c.querySelector("#iss-modal-status");function L(){if(h.value&&v.value){let _=new Date(h.value),P=new Date(v.value),K=Math.floor((P-_)/(1e3*60*60*24));I.value=isNaN(K)?"":Math.max(0,K)}else I.value=""}h.addEventListener("change",L),v.addEventListener("change",L),j.addEventListener("change",()=>{j.value==="Done"&&!v.value?(v.value=new Date().toISOString().split("T")[0],L()):j.value==="Open"&&(v.value="",I.value="")}),te({title:o?"\u270F\uFE0F Edit Data Permasalahan":"\u2795 Tambah Data Permasalahan",content:c,size:"md",confirmText:o?"Simpan Perubahan":"Simpan Permasalahan",confirmClass:"btn-primary",onConfirm:async(_,P)=>{let K=c.querySelector("#iss-modal-date").value,l=c.querySelector("#iss-modal-branch").value,E=c.querySelector("#iss-modal-category").value,D=(c.querySelector("#iss-modal-source").value||"").trim(),z=(c.querySelector("#iss-modal-complaint").value||"").trim(),J=(c.querySelector("#iss-modal-employee").value||"").trim(),N=(c.querySelector("#iss-modal-spec").value||"").trim(),O=(c.querySelector("#iss-modal-solution").value||"").trim(),U=c.querySelector("#iss-modal-status").value,H=c.querySelector("#iss-modal-completion").value||null;if(!K||!l||!z||!E){q("Harap lengkapi Tanggal, Cabang, Kategori, dan Keluhan!");return}let X=null;if(K&&H){let V=new Date(K),ee=new Date(H);X=Math.max(0,Math.floor((ee-V)/(1e3*60*60*24)))}let W={report_date:K,branch_id:Number(l),category:E,source:D,complaint:z,employee_name:J,fc_specialist:N,solution:O,status:U,completion_date:H,day_count:X};try{let V;o?V=await R(`/api/issues/${p.id}`,{method:"PUT",body:JSON.stringify(W)}):V=await R("/api/issues",{method:"POST",body:JSON.stringify(W)}),V.ok?(G(o?"Permasalahan berhasil diperbarui!":"Permasalahan baru berhasil ditambahkan!"),P(),await C()):q(V.data?.error||"Gagal menyimpan permasalahan")}catch(V){q("Terjadi kesalahan: "+V.message)}}})}async function T(p){ie({title:"Hapus Permasalahan?",message:`Apakah Anda yakin ingin menghapus tiket permasalahan dari <strong>${p.branch_name||"cabang"}</strong>:<br><br><em>"${p.complaint}"</em>?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let o=await R(`/api/issues/${p.id}`,{method:"DELETE"});o.ok?(G("Permasalahan berhasil dihapus"),await C()):q(o.data?.error||"Gagal menghapus data")}catch(o){q("Terjadi kesalahan: "+o.message)}}})}function u(){return n.filter(p=>{if(s){let o=s.toLowerCase(),f=(p.complaint||"").toLowerCase().includes(o),y=(p.employee_name||"").toLowerCase().includes(o),x=(p.branch_name||"").toLowerCase().includes(o),k=(p.source||"").toLowerCase().includes(o),a=(p.category||"").toLowerCase().includes(o),c=(p.solution||"").toLowerCase().includes(o);if(!f&&!y&&!x&&!k&&!a&&!c)return!1}return!(m!=="all"&&p.status!==m||b!=="all"&&p.category!==b||g!=="all"&&String(p.branch_id)!==String(g))})}function $(){let p=u(),o=n.length,f=n.filter(z=>z.status==="Open"||z.status==="In Progress").length,y=n.filter(z=>z.status==="Done").length,x=new Date,k=x.getFullYear(),a=x.getMonth(),c=n.filter(z=>{if(!z.report_date)return!1;let J=new Date(z.report_date);return!isNaN(J.getTime())&&J.getFullYear()===k&&J.getMonth()===a}).length,h=n.length,v=n.filter(z=>z.status==="Open").length,I=n.filter(z=>z.status==="In Progress").length,j=n.filter(z=>z.status==="Done").length;e.innerHTML=`
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
              ${w?`
                <button id="issues-btn-add" class="btn btn-primary issues-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Permasalahan</span>
                </button>
              `:""}
              
              ${Q(S,"export")?`<button id="issues-btn-excel" class="btn btn-ghost issues-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>`:""}

              <button id="issues-btn-refresh" class="btn btn-ghost issues-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${d?"spin-icon":""}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="issues-stats-row">
            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(239, 68, 68, 0.2); color:#FCA5A5;">\u26A0\uFE0F</div>
              <div>
                <div class="issues-stat-val">${o}</div>
                <div class="issues-stat-lbl">Total Permasalahan</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FCD34D;">\u23F3</div>
              <div>
                <div class="issues-stat-val">${f}</div>
                <div class="issues-stat-lbl">Open & In Progress</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">\u2705</div>
              <div>
                <div class="issues-stat-val">${y}</div>
                <div class="issues-stat-lbl">Tuntas (Done)</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(59, 130, 246, 0.2); color:#93C5FD;">\u{1F4C5}</div>
              <div>
                <div class="issues-stat-val">${c}</div>
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
                placeholder="Cari keluhan, nama personil, sumber, cabang..." value="${s?String(s).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="issues-branch-filter" class="issues-select">
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(z=>`
                <option value="${z.value}" ${String(g)===String(z.value)?"selected":""}>
                  ${z.label}
                </option>
              `).join("")}
            </select>

            <select id="issues-category-filter" class="issues-select">
              <option value="all">Semua Kategori</option>
              ${la.map(z=>`
                <option value="${z}" ${b===z?"selected":""}>
                  Kategori: ${z}
                </option>
              `).join("")}
            </select>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="issues-filter-pills">
            <div class="issues-pill ${m==="all"?"active":""}" data-status="all">
              <span>Semua Status</span>
              <span class="issues-pill-count">${h}</span>
            </div>
            <div class="issues-pill ${m==="Open"?"active":""}" data-status="Open">
              <span class="issues-pill-dot" style="background:#F59E0B;"></span>
              <span>Open</span>
              <span class="issues-pill-count">${v}</span>
            </div>
            <div class="issues-pill ${m==="In Progress"?"active":""}" data-status="In Progress">
              <span class="issues-pill-dot" style="background:#3B82F6;"></span>
              <span>In Progress</span>
              <span class="issues-pill-count">${I}</span>
            </div>
            <div class="issues-pill ${m==="Done"?"active":""}" data-status="Done">
              <span class="issues-pill-dot" style="background:#10B981;"></span>
              <span>Done</span>
              <span class="issues-pill-count">${j}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive Cards Grid -->
        <div class="issues-cards-grid">
          ${p.length===0?`
            <div class="issues-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada permasalahan yang cocok</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian atau filter status yang dipilih.</p>
              <button id="issues-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
          `:p.map(z=>{let J=z.category||"Lainnya",N=`issues-category-${J.toLowerCase().replace(/[^a-z0-9]/g,"")}`||"issues-category-lainnya",O=z.status||"Open",U="issues-status-open",H="\u26A0\uFE0F Open";O==="In Progress"?(U="issues-status-inprogress",H="\u{1F504} In Progress"):O==="Done"&&(U="issues-status-done",H="\u2705 Done");let X=O==="Done"?"status-done":O==="In Progress"?"status-inprogress":"";return`
              <div class="issues-card" data-id="${z.id}">
                <!-- Top Badges Row -->
                <div class="issues-card-top">
                  <span class="issues-category-badge ${N}">
                    \u{1F3F7}\uFE0F ${J}
                  </span>
                  <div style="display:flex; align-items:center; gap:8px;">
                    <span class="issues-status-badge ${U}">
                      ${H}
                    </span>
                    <span style="font-size:0.8rem; color:#64748B; display:flex; align-items:center; gap:4px;">
                      \u{1F5D3}\uFE0F ${da(z.report_date)}
                    </span>
                  </div>
                </div>

                <!-- Complaint Box -->
                <div class="issues-complaint-box ${X}">
                  "${z.complaint||"-"}"
                </div>

                <!-- Person & Location Box -->
                <div class="issues-person-box">
                  <div class="issues-avatar-box">
                    ${Pn(z.employee_name)}
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
                        <span>\u23F1\uFE0F Selesai: ${da(z.completion_date)}</span>
                        ${z.day_count!=null?`<span>(Durasi: ${z.day_count} Hari)</span>`:""}
                      </div>
                    `:""}
                  </div>
                `:""}

                <!-- Footer Actions -->
                <div class="issues-card-footer">
                  <div>
                    ${w?O==="Open"?`
                        <button class="issues-quick-btn start" data-action="quick-progress" data-id="${z.id}" title="Ubah status ke In Progress">
                          <span>\u{1F504} Mulai Tangani</span>
                        </button>
                      `:O==="In Progress"?`
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
                    ${w?`
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
    `;let L=e.querySelector("#issues-search-input");L&&L.addEventListener("input",z=>{s=z.target.value,$();let J=e.querySelector("#issues-search-input");J&&(J.focus(),J.setSelectionRange(J.value.length,J.value.length))});let _=e.querySelector("#issues-branch-filter");_&&_.addEventListener("change",z=>{g=z.target.value,$()});let P=e.querySelector("#issues-category-filter");P&&P.addEventListener("change",z=>{b=z.target.value,$()}),e.querySelectorAll(".issues-pill").forEach(z=>{z.addEventListener("click",()=>{m=z.dataset.status,$()})});let K=e.querySelector("#issues-reset-filter");K&&K.addEventListener("click",()=>{s="",g="all",b="all",m="all",$()});let l=e.querySelector("#issues-btn-add");l&&l.addEventListener("click",()=>B());let E=e.querySelector("#issues-btn-excel");E&&E.addEventListener("click",()=>M());let D=e.querySelector("#issues-btn-refresh");D&&D.addEventListener("click",()=>C()),e.querySelectorAll(".issues-card").forEach(z=>{let J=Number(z.dataset.id),N=n.find(V=>V.id===J);if(!N)return;let O=z.querySelector('[data-action="edit"]');O&&O.addEventListener("click",()=>B(N));let U=z.querySelector('[data-action="delete"]');U&&U.addEventListener("click",()=>T(N));let H=z.querySelector('[data-action="quick-progress"]');H&&H.addEventListener("click",()=>A(N,"In Progress"));let X=z.querySelector('[data-action="quick-done"]');X&&X.addEventListener("click",()=>A(N,"Done"));let W=z.querySelector('[data-action="quick-reopen"]');W&&W.addEventListener("click",()=>A(N,"Open"))})}await C()}var je=[];function Ln(e,t){let n=String(e.status||"").toLowerCase();return t==="pending"?n==="pending":!1}async function ca(e,t){let n=t?t.get("dash_filter"):null;je=await se();let i=await he(),r=["Ade","Berlin"],m=g=>g&&!i.find(s=>s.value===g)?[...i,{value:g,label:g}]:i,b=g=>g&&!r.find(s=>(typeof s=="object"?s.value:s)===g)?[...r,g]:r;be({container:e,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:g=>n?g.filter(s=>Ln(s,n)):g,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:g=>window.formatDate(g)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:g=>`<span title="${g||""}">${g?.length>50?g.slice(0,50)+"\u2026":g||"-"}</span>`},{key:"solution",label:"Solusi",render:g=>`<span title="${g||""}">${g?.length>40?g.slice(0,40)+"\u2026":g||"-"}</span>`},{key:"status",label:"Status",render:g=>fe(g)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:g=>window.formatDate(g)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:g=>g?`<a href="${g}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:je},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async g=>{let s=new URLSearchParams(g||{}).toString(),d=await R(`/api/one-on-one?limit=10000&${s}`);if(d.ok){let F=d.data.data.map(w=>({Tanggal:w.meeting_date||"",Cabang:w.branch_name||"","Nama Karyawan":w.employee_name||"",PIC:w.pic||"",Masalah:w.problem||"",Solusi:w.solution||"",Status:w.status||"","Tgl Selesai":w.completion_date||"",Dokumen:w.document_link||""})),{downloadExcel:S}=await Promise.resolve().then(()=>(de(),xe));S(F,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let g=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:s}=await Promise.resolve().then(()=>(de(),xe));s(g,"Template_Import_OneOnOne")},onImport:async g=>{let s=w=>{if(!w)return null;let C=String(w||"").toLowerCase(),M=je.find(A=>String(A.label||"").toLowerCase()===C);return M?M.value:null},d=w=>{if(!w)return"";if(w instanceof Date&&!isNaN(w.getTime()))return w.toISOString().slice(0,10);let C=String(w).trim();if(/^\d{4,5}$/.test(C)){let A=Number(C);if(A>2e4&&A<99999){let B=new Date(Date.UTC(1899,11,30)+A*864e5);return isNaN(B.getTime())?"":B.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(C))return C.slice(0,10);let M=C.split(/[\/\-\.]/);if(M.length===3){let[A,B,T]=M.map(u=>u.trim());if(A.length===4&&B.length<=2&&T.length<=2)return`${A}-${B.padStart(2,"0")}-${T.padStart(2,"0")}`;if(T.length===4&&B.length<=2&&A.length<=2)return`${T}-${B.padStart(2,"0")}-${A.padStart(2,"0")}`}return C},F=g.map(w=>({meeting_date:d(w.Tanggal),employee_name:String(w["Nama Karyawan"]||"").trim(),branch_id:s(String(w.Cabang||"").trim()),pic:String(w.PIC||"").trim(),problem:String(w.Masalah||"").trim(),solution:String(w.Solusi||"").trim(),status:String(w.Status||"").trim(),completion_date:d(w["Tgl Selesai"]),document_link:String(w.Dokumen||"").trim()})).filter(w=>w.meeting_date&&w.employee_name&&w.branch_id),S=await R("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:F,onDuplicate:"update"})});if(!S.ok)throw new Error(S.data?.error||"Import gagal");return S.data}},formFields:g=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:g?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:g?.branch_id&&!je.find(s=>s.value==g.branch_id)?[...je,{value:g.branch_id,label:g.branch_name||g.branch_id}]:je,createApi:{path:"/api/branches",field:"full_name"},value:g?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:m(g?.employee_name),value:g?.employee_name},{name:"pic",label:"PIC",type:"select",options:b(g?.pic),createApi:{path:"/api/pic",field:"name"},value:g?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:g?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:g?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:g?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:g?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:g?.document_link}]})}var Be={};function qe(e){if(Be[e]){try{Be[e].destroy()}catch{}delete Be[e]}}function In(){Object.keys(Be).forEach(qe)}var ve=(e,t=0)=>{let n=Number(e);return isNaN(n)||e===null||e===void 0?t:n},Ke=(e,t="\u2014")=>{if(e==null||e==="")return t;let n=String(e).trim();return n===""||n==="[object Object]"?t:n};function gt(e,t,n=900){if(!e)return;e._animFrame&&cancelAnimationFrame(e._animFrame);let i=Math.max(0,Math.round(ve(t))),r=e.textContent?e.textContent.replace(/[^\d]/g,""):"",m=r!==""&&parseInt(r)||0;if(i===m){e.textContent=i.toLocaleString("id-ID"),e._animFrame=null;return}let b=Date.now(),g=()=>{let s=Math.min((Date.now()-b)/n,1),d=1-Math.pow(1-s,3);e.textContent=Math.round(m+d*(i-m)).toLocaleString("id-ID"),s<1?e._animFrame=requestAnimationFrame(g):(e.textContent=i.toLocaleString("id-ID"),e._animFrame=null)};e._animFrame=requestAnimationFrame(g)}var zn={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},Mn=e=>{let t=Ke(e,"\u2014");return`<span class="status-pill ${zn[t]||"pill-neutral"}">${t}</span>`};var we={family:"Inter",size:11},$e="#94A3B8",Ne="#F1F5F9",_t=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],jn=()=>window.innerWidth<768;function mt(e={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:jn()?"bottom":"top",labels:{font:we,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:we,titleFont:{...we,weight:"700"}}},scales:{x:{grid:{color:Ne},ticks:{font:we,color:$e,maxRotation:0}},y:{grid:{color:Ne},ticks:{font:we,color:$e},beginAtZero:!0}},...e}}var Kn=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),Nn=()=>Array(5).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function ua(e=3){return Array(e).fill(0).map((t,n)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${n<e-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function ke(e,t,n=8e3){try{let i=new AbortController,r=setTimeout(()=>i.abort(),n),m=await R(e,{signal:i.signal}).catch(()=>null);if(clearTimeout(r),!m||!m.ok)return t;let b=m.data;return b?b.data!==void 0?b.data??t:b:t}catch{return t}}function qn(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(i=>{let r=document.getElementById(i);r&&(r.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(i=>{let r=document.getElementById(i);if(r&&r.style.display==="none"){r.style.display="block";let m=r.parentElement;if(m&&!m.querySelector(".chart-empty")){let b=document.createElement("div");b.className="chart-empty",b.textContent="Belum ada data",r.style.display="none",m.appendChild(b)}}});let e=document.getElementById("kpi-row");e&&e.querySelector(".skeleton")&&ba({});let t=document.getElementById("mini-stats-row");t&&t.querySelector(".skeleton")&&ha({}),["table-contracts","table-issues"].forEach(i=>{let r=document.getElementById(i);r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada data</div>')});let n=document.getElementById("activity-log");n&&n.querySelector(".skeleton")&&(n.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function ma(e){In(),e._dashRefresh&&clearInterval(e._dashRefresh),e._skelTimeout&&clearTimeout(e._skelTimeout);let t=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});e.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">
      
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h2 style="margin:0; font-size:1.2rem; color:var(--text-1);">Dashboard Overview</h2>
        <button id="btn-dash-refresh" class="btn-primary" style="padding:6px 12px; font-size:0.8rem; display:flex; align-items:center; gap:6px; cursor:pointer; background:var(--primary); color:#fff; border:none; border-radius:6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
          Refresh Data
        </button>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${Kn()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${Nn()}</div>

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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${ua(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${ua(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>ga(e)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async n=>{let i=n.target.value,r=document.getElementById("jadwal-year-label");r&&(r.textContent=i);let m=document.getElementById("skel-jadwal"),b=document.getElementById("chart-jadwal");m&&(m.style.display="block",m.style.position="absolute"),b&&(b.style.display="none");let g=await ke(`/api/dashboard/schedule-chart?year=${i}`,{},8e3);try{fa(g)}catch(s){console.warn("ScheduleChart render:",s),Ee("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async n=>{let i=n.target.value,r=i?`/api/dashboard/inspection-bar?month=${i}`:"/api/dashboard/inspection-bar",m=document.getElementById("skel-insp"),b=document.getElementById("chart-insp");m&&(m.style.display="block",m.style.position="absolute"),b&&(b.style.display="none");let g=await ke(r,{},8e3);try{xa(g)}catch(s){console.warn("InspBar render:",s),Ee("skel-insp","chart-insp")}}),e._skelTimeout=setTimeout(()=>qn(),5e3),await ga(e)}async function ga(e){e._skelTimeout&&(clearTimeout(e._skelTimeout),e._skelTimeout=null);let[t,n,i,r,m,b,g,s]=await Promise.all([ke("/api/dashboard/kpi",{},8e3),ke("/api/dashboard/issues-trend",{},8e3),ke("/api/dashboard/issues-summary",{},8e3),ke("/api/dashboard/stats",{},8e3),ke("/api/dashboard/calendar",[],8e3),ke("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),ke(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),ke("/api/dashboard/dropdown-aggregates",{},8e3)]),d=document.getElementById("filter-insp-month"),F=d?d.value:"",S=F?`/api/dashboard/inspection-bar?month=${F}`:"/api/dashboard/inspection-bar",w=await ke(S,{},8e3);window.dashboardAggregates=s||{};try{ba(t)}catch(C){console.warn("KPI render:",C)}try{ha(t)}catch(C){console.warn("MiniStats render:",C)}try{fa(g)}catch(C){console.warn("ScheduleChart render:",C),Ee("skel-jadwal","chart-jadwal")}try{Rn(Array.isArray(i?.by_category)?i.by_category:[])}catch(C){console.warn("Donut render:",C),Ee("skel-donut","chart-donut")}try{On(n)}catch(C){console.warn("Trend render:",C),Ee("skel-trend","chart-trend")}try{xa(w)}catch(C){console.warn("InspBar render:",C),Ee("skel-insp","chart-insp")}try{let C=Array.isArray(r)?r:Array.isArray(r?.recent_issues)?r.recent_issues:[];Un(C)}catch(C){console.warn("IssuesTable render:",C)}try{let C=Array.isArray(r?.expiring_contracts)?r.expiring_contracts:[];Hn(b)}catch(C){console.warn("ContractsTable render:",C)}try{Gn(Array.isArray(m)?m:[])}catch(C){console.warn("Agenda render:",C)}try{Jn()}catch(C){console.warn("Quick Actions render:",C)}}function ba(e){let t=document.getElementById("kpi-row");if(!t)return;e=e||{};let n=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,i=window.dashboardAggregates||{},r=s=>{let d=window.dashboardAggregates?.schedule_quarter_details?.[s];return d?`Insp: ${d.insp} \u2022 GC: ${d.gc} \u2022 DC: ${d.dc} \u2022 Fog: ${d.fog}`:""},m=s=>{let d=window.dashboardAggregates?.schedule_quarter_details?.[s];return d?`Rincian ${s}: Inspeksi: ${d.insp} | General Cleaning: ${d.gc} | Deep Cleaning: ${d.dc} | Fogging: ${d.fog}`:`Total Kegiatan ${s}`},b=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
          <option value="Q1" ${n==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${n==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${n==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${n==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:i.schedule_by_quarter?.[n]??e.schedule?.current??0,href:`#/timeline?dash_filter=period_${n.toLowerCase()}`,color:"kpi-blue",sub:r(n),subId:"mini-jadwal-detail",tooltip:m(n)}];t.innerHTML=b.map(s=>{let d=s.val!==void 0?ve(s.val,0):ve(e[s.key]?.current,0);return s.dropdown?`
        <div class="kpi-card ${s.color}" ${s.id?`id="${s.id}"`:""} ${s.tooltip?`title="${s.tooltip}"`:""} data-href="${s.href}" tabindex="0" role="link" style="cursor:pointer;padding:10px 12px">
          <div style="display:flex; gap:10px; align-items:center;">
            <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${s.icon}</span></div>
            <div style="flex:1;min-width:0;">
              <div style="display:flex; align-items:baseline; gap:3px;">
                <div class="kpi-value" data-target="${d}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${d}</div>
                ${s.dropdown}
              </div>
              <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${s.label}</div>
              <div class="kpi-subtitle" ${s.subId?`id="${s.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${s.tooltip||""}">${s.sub}</div>
            </div>
          </div>
        </div>`:`
      <a href="${s.href}" class="kpi-card ${s.color}" ${s.id?`id="${s.id}"`:""} ${s.tooltip?`title="${s.tooltip}"`:""} style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${s.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="kpi-value" data-target="${d}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${d}</div>
              ${s.dropdown?s.dropdown:""}
            </div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${s.label}</div>
            <div class="kpi-subtitle" ${s.subId?`id="${s.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${s.tooltip||""}">${s.sub}</div>
          </div>
        </div>
      </a>`}).join(""),t.querySelectorAll(".kpi-card[data-href]").forEach(s=>{s.addEventListener("click",d=>{d.target.closest("select")||s.dataset.href&&(window.location.hash=s.dataset.href)}),s.addEventListener("keydown",d=>{d.key==="Enter"&&!d.target.closest("select")&&s.dataset.href&&(window.location.hash=s.dataset.href)})}),t.querySelectorAll(".kpi-value").forEach(s=>{gt(s,parseInt(s.dataset.target)||0)});let g=document.getElementById("dash-jadwal-period");if(g){let s=d=>{let S=(window.dashboardAggregates?.schedule_by_quarter||{})[d]||0,w=document.querySelector("#mini-jadwal .kpi-value");w&&(w.dataset.target=S,gt(w,S,400));let C=document.getElementById("mini-jadwal-detail");C&&(C.textContent=r(d),C.title=m(d));let M=document.getElementById("mini-jadwal");if(M){let A=`#/timeline?dash_filter=period_${d.toLowerCase()}`;M.tagName==="A"?M.href=A:M.dataset.href=A,M.title=m(d)}};g.addEventListener("change",d=>s(d.target.value))}}function ha(e){let t=document.getElementById("mini-stats-row");if(!t)return;e=e||{};let n=new Date().getFullYear(),i=String(new Date().getMonth()+1).padStart(2,"0"),r=`${n}-${i}`,m=d=>`
    <select id="${d}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
      <option value="${n}-01" ${r===`${n}-01`?"selected":""}>Jan</option>
      <option value="${n}-02" ${r===`${n}-02`?"selected":""}>Feb</option>
      <option value="${n}-03" ${r===`${n}-03`?"selected":""}>Mar</option>
      <option value="${n}-04" ${r===`${n}-04`?"selected":""}>Apr</option>
      <option value="${n}-05" ${r===`${n}-05`?"selected":""}>Mei</option>
      <option value="${n}-06" ${r===`${n}-06`?"selected":""}>Jun</option>
      <option value="${n}-07" ${r===`${n}-07`?"selected":""}>Jul</option>
      <option value="${n}-08" ${r===`${n}-08`?"selected":""}>Agu</option>
      <option value="${n}-09" ${r===`${n}-09`?"selected":""}>Sep</option>
      <option value="${n}-10" ${r===`${n}-10`?"selected":""}>Okt</option>
      <option value="${n}-11" ${r===`${n}-11`?"selected":""}>Nov</option>
      <option value="${n}-12" ${r===`${n}-12`?"selected":""}>Des</option>
    </select>
  `,b=window.dashboardAggregates||{},g=[{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:m("dash-inspeksi-month"),val:b.inspeksi_by_month?.[r]??e.inspection_month?.current??0,href:`#/timeline?dash_filter=inspeksi&month=${r}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:m("dash-gcdc-month"),val:b.gcdc_by_month?.[r]??e.cleaning_month?.current??0,href:`#/timeline?dash_filter=gcdc&month=${r}`,color:"mini-green"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:m("dash-fogging-month"),val:b.fogging_by_month?.[r]??e.fogging_month?.current??0,href:`#/reports/fogging?dash_filter=fogging&month=${r}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:e.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:e.branches?.current,href:"#/branches",color:"mini-teal"}];t.innerHTML=g.map(d=>d.dropdown?`
        <div class="mini-stat ${d.color}" style="cursor:pointer;" id="${d.id||""}" data-href="${d.href}" tabindex="0" role="link">
          <div class="mini-stat-icon">${d.icon}</div>
          <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="mini-stat-value" data-target="${ve(d.val)}">0</div>
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
            <div class="mini-stat-value" data-target="${ve(d.val)}">0</div>
            ${d.dropdown?d.dropdown:""}
          </div>
          <div class="mini-stat-text">${d.label}</div>
          ${d.extra||""}
        </div>
      </a>`).join(""),t.querySelectorAll(".mini-stat[data-href]").forEach(d=>{d.addEventListener("click",F=>{F.target.closest("select")||d.dataset.href&&(window.location.hash=d.dataset.href)}),d.addEventListener("keydown",F=>{F.key==="Enter"&&!F.target.closest("select")&&d.dataset.href&&(window.location.hash=d.dataset.href)})}),t.querySelectorAll(".mini-stat-value").forEach(d=>gt(d,parseInt(d.dataset.target)||0,700));let s=(d,F,S,w)=>{let C=document.getElementById(d);if(C){let M=A=>{let T=(window.dashboardAggregates?.[S]||{})[A]||0,u=document.querySelector(`#${F} .mini-stat-value`);u&&(u.dataset.target=T,gt(u,T,400));let $=document.getElementById(F);if($){let p=`${w}&month=${A}`;$.tagName==="A"?$.href=p:$.dataset.href=p}};C.addEventListener("change",A=>M(A.target.value))}};s("dash-inspeksi-month","mini-inspeksi","inspeksi_by_month","#/timeline?dash_filter=inspeksi"),s("dash-gcdc-month","mini-gcdc","gcdc_by_month","#/timeline?dash_filter=gcdc"),s("dash-fogging-month","mini-fogging","fogging_by_month","#/reports/fogging?dash_filter=fogging")}function Rn(e){Ee("skel-donut","chart-donut");let t=document.getElementById("chart-donut"),n=document.getElementById("donut-legend");if(!t||!n)return;qe("donut");let i=(e||[]).filter(s=>ve(s.count)>0);if(!i.length){Qe(t,"Belum ada data permasalahan");return}let r=i.map(s=>`${Ke(s.category,"Lainnya")}`),m=i.map(s=>ve(s.count)),b=m.reduce((s,d)=>s+d,0);n.innerHTML=i.map((s,d)=>{let F=_t[d%_t.length],S=b>0?Math.round(s.count/b*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${F}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${s.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${S}%)</span></div>
          <div class="donut-legend-label">${r[d]}</div>
        </div>
      </div>
    `}).join("");let g={id:"centerText",beforeDraw:function(s){let d=s.width,F=s.height,S=s.ctx;S.restore();let w=(F/80).toFixed(2);S.font="bold "+w+"em Inter",S.textBaseline="middle",S.fillStyle="#1E293B";let C=b.toString(),M=Math.round((d-S.measureText(C).width)/2),A=F/2;S.fillText(C,M,A-4),S.font="600 "+(w*.35).toFixed(2)+"em Inter",S.fillStyle="#64748B";let B="Total",T=Math.round((d-S.measureText(B).width)/2);S.fillText(B,T,A+10),S.save()}};Be.donut=new Chart(t,{type:"doughnut",data:{labels:r,datasets:[{data:m,backgroundColor:_t,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:we,titleFont:{...we,weight:"700"},callbacks:{label:s=>` ${s.label}: ${s.parsed} kasus`}}},cutout:"75%"},plugins:[g]})}function On(e){Ee("skel-trend","chart-trend");let t=document.getElementById("chart-trend");if(!t)return;qe("trend"),e=e||{};let n=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],i=(e.labels||[]).map(b=>{if(!b||typeof b!="string")return"";try{let[g,s]=b.split("-");return(n[Number(s)-1]||s)+" "+String(g).slice(-2)}catch{return b}}),r=(e.open||[]).map(b=>ve(b)),m=(e.closed||[]).map(b=>ve(b));if(!i.length){Qe(t,"Belum ada data trend");return}Be.trend=new Chart(t,{type:"line",data:{labels:i,datasets:[{label:"Open",data:r,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:m,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:mt({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:$e,maxRotation:45,autoSkip:!0}},y:{grid:{color:Ne},ticks:{font:{family:"Inter",size:9},color:$e},beginAtZero:!0}}})})}function fa(e){Ee("skel-jadwal","chart-jadwal");let t=document.getElementById("chart-jadwal");if(!t)return;qe("jadwal"),e=e||{};let n=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(e).some(s=>Array.isArray(s)&&s.some(d=>d>0))){Qe(t,"Belum ada data jadwal");return}let r=e["Inspeksi Hygiene"]||Array(12).fill(0),m=e["General Cleaning"]||Array(12).fill(0),b=e["Deep Cleaning"]||Array(12).fill(0),g=e.Fogging||Array(12).fill(0);Be.jadwal=new Chart(t,{type:"bar",data:{labels:n,datasets:[{label:"Inspeksi",data:r,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:m,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:b,backgroundColor:"#F59E0B"},{label:"Fogging",data:g,backgroundColor:"#EF4444"}]},options:mt({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:$e,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Ne},ticks:{font:{family:"Inter",size:9},color:$e},min:0}}})})}function xa(e){Ee("skel-insp","chart-insp");let t=document.getElementById("chart-insp");if(!t)return;qe("inspBar"),e=e||{};let n=e.labels||[],i=(e.fc||[]).map(m=>ve(m)),r=(e.spv||[]).map(m=>ve(m));if(!n.length){Qe(t,"Belum ada data inspeksi");return}Be.inspBar=new Chart(t,{type:"bar",data:{labels:n,datasets:[{label:"Skor FC",data:i,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:r,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:mt({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:we,color:$e,maxRotation:45,minRotation:45}},y:{grid:{color:Ne},ticks:{font:we,color:$e},min:0,max:100}}})})}function Hn(e){Ee("skel-contract-mini","chart-contract-mini");let t=document.getElementById("chart-contract-mini");if(!t)return;qe("contractMiniBar"),e=e||{};let n={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},i=(e.labels||[]).map(b=>{let g=b.split("-")[1];return n[g]||b}),r=(e.data||[]).map(b=>ve(b));if(!i.length){Qe(t,"Belum ada data");return}let m=t.getContext("2d");Be.contractMiniBar=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Kontrak Habis",data:r,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:mt({onClick:(b,g)=>{if(g&&g.length>0){let s=g[0].index,d=(e.labels||[])[s];d&&(window.location.hash="#/contracts?month_expiry="+d)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:we,color:$e,maxRotation:0,autoSkip:!1}},y:{grid:{color:Ne,borderDash:[4,4],drawBorder:!1},ticks:{font:we,color:$e,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function Un(e){let t=document.getElementById("table-issues");if(!t)return;let n=(e||[]).slice(0,8);if(!n.length){t.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}t.innerHTML=`
    <div class="dash-list">
      ${n.map(i=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${Mn(i.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Ke(i.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Ke(i.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function Gn(e){let t=document.getElementById("widget-agenda");if(!t)return;let n=new Date,i=`${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,"0")}-${String(n.getDate()).padStart(2,"0")}`,m=(e||[]).filter(b=>(b.event_date||"").startsWith(i)).slice(0,10);if(!m.length){t.innerHTML="";return}t.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${m.map(b=>{let g="#3B82F6",s="#EFF6FF",d="Agenda",F=(b.title||"").toLowerCase();return F.includes("inspeksi")?(g="#10B981",s="#ECFDF5",d="Inspeksi"):F.includes("cleaning")||F.includes("gcdc")?(g="#3B82F6",s="#EFF6FF",d="Cleaning"):F.includes("reliefer")?(g="#F59E0B",s="#FFFBEB",d="Reliefer"):F.includes("fogging")&&(g="#8B5CF6",s="#F5F3FF",d="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(b.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${g};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Ke(b.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Ke(b.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${s};color:${g}">${d}</div>
        </div>
      `}).join("")}
    </div>
  `}function Jn(){let e=document.getElementById("quick-actions");if(!e)return;let t=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];e.innerHTML=t.map(n=>`
    <a href="${n.href}" class="action-btn">
      <div class="action-icon" style="background:${n.bg}">${n.icon}</div>
      ${n.label}
    </a>
  `).join("")}function Ee(e,t){let n=document.getElementById(e),i=document.getElementById(t);if(n&&(n.style.display="none",n.style.position=""),i){i.style.display="block";let r=i.parentElement;if(r){let m=r.querySelector(".chart-empty");m&&m.remove()}}}function Qe(e,t="Belum ada data"){if(!e)return;e.style.display="none";let n=e.parentElement;if(!n)return;if(!n.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent=t,n.appendChild(r)}}async function ya(e){document.getElementById("app").innerHTML=`
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
  `;let t=document.getElementById("login-form"),n=document.getElementById("login-error"),i=document.getElementById("login-btn"),r=document.getElementById("toggle-password"),m=document.getElementById("login-password");r?.addEventListener("click",()=>{let b=m.type==="text";m.type=b?"password":"text",r.style.color=b?"":"var(--primary)"}),t?.addEventListener("submit",async b=>{b.preventDefault(),n.style.display="none";let g=t.username.value.trim(),s=t.password.value;if(!g||!s){n.textContent="Username dan password wajib diisi.",n.style.display="block";return}i.querySelector(".btn-text").style.display="none",i.querySelector(".btn-spinner").style.display="",i.disabled=!0;try{let d=await R("/api/auth/login",{method:"POST",body:JSON.stringify({username:g,password:s})});d.ok&&d.data.success?(Qt(d.data.data.token),nt(d.data.data.user),G("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(n.textContent=d.data.error||"Username atau password salah.",n.style.display="block",i.classList.add("shake"),setTimeout(()=>i.classList.remove("shake"),600))}catch{n.textContent="Gagal terhubung ke server. Periksa koneksi internet.",n.style.display="block"}finally{i.querySelector(".btn-text").style.display="",i.querySelector(".btn-spinner").style.display="none",i.disabled=!1}})}me();var At=[{id:1,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:6,branch_name:"006. Kelapa Gading MOI",trainer:"Berlin Ariansyah",participants:"MUHAMMAD PRAKAS AGAZI",score:90,document_link:""},{id:2,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:11,branch_name:"011. Kuningan",trainer:"Berlin Ariansyah",participants:"KAMALLUDIN",score:88,document_link:""},{id:3,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:42,branch_name:"042. Muara Karang",trainer:"Berlin Ariansyah",participants:"SAHRUL ICHSANI",score:85,document_link:""},{id:4,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:51,branch_name:"051. Bintaro Sektor 9",trainer:"Berlin Ariansyah",participants:"MUHAMAD VICKRI FATHURROHMAN",score:92,document_link:""},{id:5,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:45,branch_name:"045. FX Mall",trainer:"Berlin Ariansyah",participants:"DEDE AHDORI",score:86,document_link:""},{id:6,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:54,branch_name:"054. Bogor Gunung Batu",trainer:"Berlin Ariansyah",participants:"ADITYA PRATAMA",score:95,document_link:""},{id:7,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:55,branch_name:"055. Gajah Mada",trainer:"Berlin Ariansyah",participants:"ELVAN JANGGA PRAMANA",score:89,document_link:""}],va=["Batch 5","Batch 4","Batch 3","Batch 2","Batch 1"],ka=["Berlin Ariansyah","Ade Surahman","Mizwar","Fajar","Ade","Berlin"];function Yn(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function De(e){if(!e)return"-";try{let t=JSON.parse(e);if(Array.isArray(t))return t.join(", ")}catch{}return String(e)}async function wa(e){let t=[],n=[],i=[],r="all",m="all",b="all",g="",s=!1,d=ae(),F=Q(d,"training","write");try{[n,i]=await Promise.all([se(),he()])}catch(u){console.error("Error fetching cache:",u)}async function S(){s=!0,B();try{let u=await R("/api/training?limit=10000");if(u.ok){let $=u.data?.data||u.data||[];t=Array.isArray($)&&$.length>0?$:At}else t=At}catch(u){console.warn("Gagal load training, menggunakan data lokal:",u),t=At}finally{s=!1,B()}}async function w(){if(!Q(d,"export")){q("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:u}=await Promise.resolve().then(()=>(de(),xe)),$=t.map((p,o)=>({No:o+1,Tanggal:p.training_date||"",Batch:p.batch||"",Materi:p.subject||"",Cabang:p.branch_name||"",Trainer:p.trainer||"",Peserta:De(p.participants),Nilai:p.score!=null?p.score:"",Dokumen:p.document_link||"",Catatan:p.notes||""}));u($,`Data_Training_${new Date().toISOString().slice(0,10)}`),G("Data training berhasil diekspor ke Excel!")}catch(u){q("Gagal ekspor Excel: "+u.message)}}async function C(u){if(!F)return;let $=t.find(p=>p.id===u);ie({title:"Hapus Data Training",message:`Apakah Anda yakin ingin menghapus data sesi training "${$?.subject||"ini"}" untuk ${De($?.participants)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let p=await R(`/api/training/${u}`,{method:"DELETE"});p.ok?(G("Data training berhasil dihapus"),t=t.filter(o=>o.id!==u),B()):q(p.data?.error||"Gagal menghapus data training")}catch(p){q("Gagal menghapus training: "+p.message)}}})}function M(u=null){if(!F)return;let $=!!u,p=u?.training_date?u.training_date.slice(0,10):new Date().toISOString().slice(0,10),o=u?.batch||"Batch 5",f=u?.trainer||"Berlin Ariansyah",y=De(u?.participants)!=="-"?De(u?.participants):"",x=document.createElement("div");x.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Training <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="tr-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${p}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Batch Pelatihan <span style="color:#DC2626;">*</span>
            </label>
            <input list="tr-batch-options" id="tr-modal-batch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih atau ketik batch..." value="${o}" required />
            <datalist id="tr-batch-options">
              ${va.map(k=>`<option value="${k}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Materi / Topik Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="tr-modal-subject" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Basic Knowledge & Practical Skills" value="${u?.subject?String(u.subject).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="tr-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${n.map(k=>`
                <option value="${k.value}" ${String(u?.branch_id)===String(k.value)?"selected":""}>
                  ${k.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Trainer / Instruktur <span style="color:#DC2626;">*</span>
            </label>
            <input list="tr-trainer-options" id="tr-modal-trainer" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih trainer..." value="${f}" required />
            <datalist id="tr-trainer-options">
              ${ka.map(k=>`<option value="${k}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Peserta Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="tr-modal-participants" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Ketik nama karyawan peserta training..." required>${y}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nilai / Score Evaluasi (0 - 100)
            </label>
            <input type="number" min="0" max="100" step="1" id="tr-modal-score" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: 90" value="${u?.score!=null?u.score:""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Link Dokumen / Sertifikat (Opsional)
            </label>
            <input type="url" id="tr-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="https://drive.google.com/..." value="${u?.document_link?String(u.document_link).replace(/"/g,"&quot;"):""}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan Pelatihan
          </label>
          <textarea id="tr-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Catatan hasil pelatihan, rekomendasi materi lanjutan...">${u?.notes||""}</textarea>
        </div>
      </div>
    `,te({title:$?"\u270F\uFE0F Edit Data Sesi Training":"\u2795 Tambah Data Sesi Training",content:x,size:"md",confirmText:$?"Simpan Perubahan":"Simpan Training",confirmClass:"btn-primary",onConfirm:async(k,a)=>{let c=x.querySelector("#tr-modal-date").value,h=(x.querySelector("#tr-modal-batch").value||"").trim(),v=(x.querySelector("#tr-modal-subject").value||"").trim(),I=x.querySelector("#tr-modal-branch").value,j=(x.querySelector("#tr-modal-trainer").value||"").trim(),L=(x.querySelector("#tr-modal-participants").value||"").trim(),_=x.querySelector("#tr-modal-score").value,P=_!==""?Number(_):null,K=(x.querySelector("#tr-modal-link").value||"").trim(),l=(x.querySelector("#tr-modal-notes").value||"").trim();if(!c){q("Tanggal training wajib diisi");return}if(!v){q("Materi / topik pelatihan wajib diisi");return}if(!I){q("Cabang klinik wajib dipilih");return}if(!j){q("Trainer wajib diisi");return}if(!L){q("Nama peserta training wajib diisi");return}let E=n.find(z=>String(z.value)===String(I)),D={training_date:c,batch:h,subject:v,branch_id:Number(I),branch_name:E?E.label:"Cabang",trainer:j,participants:L,score:P,document_link:K,notes:l};try{if($){let z=await R(`/api/training/${u.id}`,{method:"PUT",body:JSON.stringify(D)});if(z.ok){G("Data training berhasil diperbarui");let J=t.findIndex(N=>N.id===u.id);J!==-1&&(t[J]={...t[J],...D}),a(),B()}else q(z.data?.error||"Gagal memperbarui data training")}else{let z=await R("/api/training",{method:"POST",body:JSON.stringify(D)});if(z.ok){G("Sesi training baru berhasil ditambahkan");let J=z.data?.id||Date.now();t.unshift({id:J,...D}),a(),B()}else q(z.data?.error||"Gagal menambahkan training")}}catch(z){q("Terjadi kesalahan: "+z.message)}}})}function A(){return t.filter(u=>{if(g){let $=g.toLowerCase(),p=(u.subject||"").toLowerCase().includes($),o=(u.trainer||"").toLowerCase().includes($),f=De(u.participants).toLowerCase().includes($),y=(u.branch_name||"").toLowerCase().includes($),x=(u.batch||"").toLowerCase().includes($);if(!p&&!o&&!f&&!y&&!x)return!1}return!(r!=="all"&&u.batch!==r||m!=="all"&&String(u.branch_id)!==String(m)||b!=="all"&&u.trainer!==b)})}function B(){let u=A(),$=t.length,p=new Set;t.forEach(h=>{De(h.participants).split(",").forEach(I=>{let j=I.trim();j&&j!=="-"&&p.add(j)})});let o=p.size,y=new Set(t.map(h=>h.batch).filter(Boolean)).size,x=new Date,k=x.getFullYear(),a=x.getMonth(),c=t.filter(h=>{if(!h.training_date)return!1;let v=new Date(h.training_date);return!isNaN(v.getTime())&&v.getFullYear()===k&&v.getMonth()===a}).length;e.innerHTML=`
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
              ${F?`
                <button id="tr-btn-add" class="btn btn-primary training-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Training</span>
                </button>
              `:""}
              
              ${Q(d,"export")?`<button id="tr-btn-excel" class="btn btn-ghost training-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
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
                <div class="training-stat-val">${$}</div>
                <div class="training-stat-lbl">Sesi Training</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F465}
              </div>
              <div>
                <div class="training-stat-val">${o}</div>
                <div class="training-stat-lbl">Peserta Terlatih</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4DA}
              </div>
              <div>
                <div class="training-stat-val">${y} Batch</div>
                <div class="training-stat-lbl">Total Angkatan</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="training-stat-val">${c}</div>
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
              <input type="text" id="tr-search-input" class="training-search-input" placeholder="Cari materi, trainer, peserta, atau cabang..." value="${g}" />
            </div>

            <select id="tr-filter-branch" class="training-select">
              <option value="all">Semua Cabang (${n.length})</option>
              ${n.map(h=>`<option value="${h.value}" ${String(m)===String(h.value)?"selected":""}>${h.label}</option>`).join("")}
            </select>

            <select id="tr-filter-trainer" class="training-select">
              <option value="all">Semua Trainer</option>
              ${ka.map(h=>`<option value="${h}" ${b===h?"selected":""}>${h}</option>`).join("")}
            </select>
          </div>

          <!-- Batch Filter Pills (Swipeable on Mobile) -->
          <div class="training-filter-pills">
            <div class="training-pill ${r==="all"?"active":""}" data-batch="all">
              <span>Semua Batch</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${$}</span>
            </div>
            ${va.map(h=>{let v=t.filter(I=>I.batch===h).length;return`
                <div class="training-pill ${r===h?"active":""}" data-batch="${h}">
                  <span>\u{1F393}</span>
                  <span>${h}</span>
                  ${v>0?`<span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${v}</span>`:""}
                </div>
              `}).join("")}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="training-list-container">
          ${u.length===0?`
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
              ${u.map(h=>{let v=De(h.participants),j=v.split(",")[0].trim().split(" ").map(L=>L[0]).filter(Boolean).slice(0,2).join("").toUpperCase()||"TR";return`
                  <div class="training-card" data-id="${h.id}">
                    <div>
                      <!-- Card Header: Batch & Nilai/Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="training-batch-badge">
                          <span>\u{1F393}</span>
                          <span>${h.batch||"Reguler"}</span>
                        </span>
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${h.score!=null?`
                            <span class="training-score-badge">
                              <span>\u2B50</span>
                              <span>${h.score}</span>
                            </span>
                          `:""}
                          <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                            \u{1F5D3}\uFE0F ${Yn(h.training_date)}
                          </span>
                        </div>
                      </div>

                      <!-- Subject / Materi Title -->
                      <h4 style="margin:0 0 10px 0; font-size:1.05rem; font-weight:700; color:#0F172A; line-height:1.4;">
                        ${h.subject||"Materi Pelatihan"}
                      </h4>

                      <!-- Participant & Branch Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:10px;">
                        <div class="training-avatar-box">
                          ${j}
                        </div>
                        <div style="min-width:0;">
                          <div style="font-size:0.72rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:1px;">Peserta Training</div>
                          <div style="font-size:0.95rem; font-weight:700; color:#1E293B; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${v}">
                            ${v}
                          </div>
                        </div>
                      </div>

                      <!-- Detail Box: Cabang & Trainer -->
                      <div class="training-info-box">
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#334155;">
                          <span>\u{1F4CD}</span>
                          <span style="font-weight:600;">${h.branch_name||"Cabang Belum Ditentukan"}</span>
                        </div>
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#64748B;">
                          <span>\u{1F468}\u200D\u{1F3EB}</span>
                          <span>Trainer: <strong style="color:#0F172A;">${h.trainer||"-"}</strong></span>
                        </div>
                      </div>

                      ${h.notes?`
                        <div style="font-size:0.78rem; color:#64748B; font-style:italic; line-height:1.4; margin-bottom:8px;">
                          \u{1F4DD} ${h.notes}
                        </div>
                      `:""}
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:8px;">
                      <div>
                        ${h.document_link?`
                          <a href="${h.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>\u{1F4C4} Materi / Sertifikat</span>
                          </a>
                        `:`
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${F?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          <button class="training-icon-btn training-edit-btn" title="Edit Data Training" data-id="${h.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                          <button class="training-icon-btn delete training-delete-btn" title="Hapus Training" data-id="${h.id}">
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
    `,T()}function T(){let u=e.querySelector("#tr-btn-add");u&&(u.onclick=()=>M());let $=e.querySelector("#tr-btn-excel");$&&($.onclick=()=>w());let p=e.querySelector("#tr-btn-refresh");p&&(p.onclick=()=>S());let o=e.querySelector("#tr-btn-reset-filter");o&&(o.onclick=()=>{g="",m="all",r="all",b="all",B()});let f=e.querySelector("#tr-search-input");f&&(f.oninput=k=>{g=k.target.value,B();let a=e.querySelector("#tr-search-input");a&&(a.focus(),a.setSelectionRange(a.value.length,a.value.length))});let y=e.querySelector("#tr-filter-branch");y&&(y.onchange=k=>{m=k.target.value,B()});let x=e.querySelector("#tr-filter-trainer");x&&(x.onchange=k=>{b=k.target.value,B()}),e.querySelectorAll(".training-pill").forEach(k=>{k.onclick=()=>{r=k.getAttribute("data-batch")||"all",B()}}),e.querySelectorAll(".training-edit-btn").forEach(k=>{k.onclick=a=>{a.stopPropagation();let c=Number(k.getAttribute("data-id")),h=t.find(v=>v.id===c);h&&M(h)}}),e.querySelectorAll(".training-delete-btn").forEach(k=>{k.onclick=a=>{a.stopPropagation();let c=Number(k.getAttribute("data-id"));C(c)}})}await S()}me();de();function Ea({container:e,title:t,icon:n,apiPath:i,columns:r,formFields:m,filterFields:b,defaultFilters:g={},defaultSort:s="",itemLabel:d="Data",canCreate:F=!0,canEdit:S=!0,canDelete:w=!0,onBeforeSubmit:C,onAfterLoad:M,onDataLoaded:A,extraActions:B=[],initialSearch:T="",exportOptions:u=null,bulkDelete:$=!1,paginationMode:p="server"}){let o=ae();o&&(o.role==="superadmin"||o.role==="admin")?Q(o,"delete")||(w=!1,$=!1):(u=null,$=!1,o&&o.role==="editor_khusus"?i!=="/api/issues"&&i!=="/api/overtime"&&(F=!1,S=!1,w=!1):o&&o.role==="input_lembur"?(i==="/api/overtime"||(F=!1,S=!1),w=!1):(F=!1,S=!1,w=!1));let y=1,x={...g};T&&(x.search=T);let k=new Set;e.innerHTML=`
    ${$?`
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
        ${F?`<button class="btn btn-primary" id="btn-create">+ Tambah ${d}</button>`:""}
        ${u?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${u.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${u.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${u.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${u.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
    

    ${b&&b.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${b.filter(l=>l.type==="search").map(l=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${l.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${b.filter(l=>l.type!=="search").map(l=>{if(l.type==="select"||l.type==="combobox"){let E=(l.label||"").startsWith("Pilih")?l.label:`Pilih ${l.label||""}`;return`<select class="filter-select" name="${l.name}" id="filter-${l.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${E}</option>${(l.options||[]).map(D=>`<option value="${typeof D=="object"?D.value:D}" ${x[l.name]===(typeof D=="object"?D.value:D)?"selected":""}>${typeof D=="object"?D.label:D}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${b.filter(l=>l.type!=="search").map(l=>{if(l.type==="select"||l.type==="combobox"){let E=(l.label||"").startsWith("Pilih")?l.label:`Pilih ${l.label||""}`;return`<select class="filter-select filter-select-sheet" name="${l.name}-sheet" id="filter-sheet-${l.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${E}</option>${(l.options||[]).map(D=>`<option value="${typeof D=="object"?D.value:D}" ${x[l.name]===(typeof D=="object"?D.value:D)?"selected":""}>${typeof D=="object"?D.label:D}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function a(){let l=document.getElementById("bulk-toolbar");if(!l)return;let E=document.getElementById("bulk-count"),D=document.getElementById("btn-bulk-delete"),z=document.getElementById("btn-bulk-cancel");E.textContent=`${k.size} item dipilih`,k.size>0?(l.style.display="flex",D.disabled=!1,z.disabled=!1):(l.style.display="none",D.disabled=!0,z.disabled=!0);let J=document.getElementById("select-all-checkbox");if(J){let N=document.querySelectorAll(".row-checkbox");if(N.length>0){let O=[...N].every(H=>H.checked),U=[...N].some(H=>H.checked);J.checked=O,J.indeterminate=U&&!O}else J.checked=!1,J.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{k.clear(),document.querySelectorAll(".row-checkbox").forEach(E=>E.checked=!1);let l=document.getElementById("select-all-checkbox");l&&(l.checked=!1),a()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(k.size===0)return;let l=[...k],E=document.createElement("div");E.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",E.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${l.length} ${d}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${l.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(E),E.querySelector("#bulk-cancel-btn").addEventListener("click",()=>E.remove()),E.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let D=E.querySelector("#bulk-confirm-btn");D.disabled=!0,D.textContent="Menghapus...";let z=await R(`${i}/bulk`,{method:"DELETE",body:JSON.stringify({ids:l})});E.remove(),z.ok?(G(`${l.length} ${d} berhasil dihapus.`),k.clear(),a(),ye(i),L()):q(z.data?.error||"Gagal menghapus data.")})});let c=document.getElementById("filter-search"),h;c?.addEventListener("input",l=>{clearTimeout(h),h=setTimeout(()=>{x.search=l.target.value,y=1,k.clear(),a(),L()},400)}),b?.forEach(l=>{(l.type==="select"||l.type==="combobox")&&(document.getElementById(`filter-${l.name}`)?.addEventListener("change",E=>{x[l.name]=E.target.value;let D=document.getElementById(`filter-sheet-${l.name}`);D&&(D.value=E.target.value),y=1,k.clear(),a(),L()}),document.getElementById(`filter-sheet-${l.name}`)?.addEventListener("change",E=>{x[l.name]=E.target.value;let D=document.getElementById(`filter-${l.name}`);D&&(D.value=E.target.value),y=1,k.clear(),a(),L(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{x={},c&&(c.value=""),b?.forEach(l=>{let E=document.getElementById(`filter-${l.name}`);E&&(E.value="");let D=document.getElementById(`filter-sheet-${l.name}`);D&&(D.value="")}),y=1,k.clear(),a(),L()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{x={},c&&(c.value=""),b?.forEach(l=>{let E=document.getElementById(`filter-${l.name}`);E&&(E.value="");let D=document.getElementById(`filter-sheet-${l.name}`);D&&(D.value="")}),y=1,k.clear(),a(),L(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>P(null)),u&&document.addEventListener("click",function(l){let E=document.getElementById("aksi-menu-main"),D=document.getElementById("btn-aksi-main");E&&D&&!D.contains(l.target)&&!E.contains(l.target)&&E.classList.remove("show-aksi-menu")});let v=document.getElementById("btn-mobile-filter"),I=document.getElementById("filter-options-wrapper"),j=document.getElementById("btn-close-filter-sheet");if(v&&I&&(v.addEventListener("click",l=>{l.preventDefault(),I.classList.add("sheet-open")}),j&&j.addEventListener("click",l=>{l.preventDefault(),I.classList.remove("sheet-open")})),u){document.getElementById(`btn-export-${u.moduleName}`)?.addEventListener("click",async E=>{let D=E.target,z=D.innerHTML;D.innerHTML="\u23F3 Loading...",D.disabled=!0;try{await u.onExport()}catch{q("Gagal export data")}finally{D.innerHTML=z,D.disabled=!1}}),document.getElementById(`btn-template-${u.moduleName}`)?.addEventListener("click",()=>{u.onTemplate()});let l=document.getElementById(`input-import-${u.moduleName}`);l?.addEventListener("change",async E=>{let D=E.target.files[0];if(!D)return;let z=document.getElementById(`label-import-${u.moduleName}`),J=z?z.querySelector(".import-text"):null,N=J?J.innerText:"";J&&(J.innerText="\u231B Memproses..."),z&&(z.style.pointerEvents="none"),l.disabled=!0;try{let O=await Je(D);if(O.length===0)throw new Error("File kosong atau format salah");await u.onImport(O),G("Import berhasil!"),ye(i),L()}catch(O){q(O.message||"Gagal import data")}finally{J&&(J.innerText=N),z&&(z.style.pointerEvents="auto"),l.disabled=!1,l.value=""}})}async function L(){a();let l=document.getElementById("table-container");if(!l)return;l.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let E=p==="client",D=E?1:y,z=E?Ie:20,J={...s?{sort:s}:{},...Object.fromEntries(Object.entries(x).filter(([,ee])=>ee))},N=new URLSearchParams({page:D,limit:z,...J}),O=await R(`${i}?${N}`);if(!O.ok){l.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${O.data?.error||"Error"}</p></div>`;return}let U=O.data?.data||O.data||[],H=O.data?.pagination,X=U.length;if(E){typeof A=="function"&&(U=A(U));let ee=U.length,Z=20,Y=Math.ceil(ee/Z);y>Y&&Y>0&&(y=Y);let re=(y-1)*Z,pe=y*Z;U=U.slice(re,pe),H={page:y,limit:Z,total:ee,pages:Y}}!1,M&&M(U);let W=rt({columns:r,data:U,onEdit:S?ee=>P(ee):null,actions:B.map(ee=>({...ee,handler:Z=>ee.handler(Z,L)})),emptyText:`Tidak ada ${String(d||"").toLowerCase()}`,bulkSelect:$?{selectedIds:k,onToggle:a}:null});l.innerHTML="",l.appendChild(W);let V=document.getElementById("pagination-container");if(V&&(V.innerHTML="",H&&H.pages>1)){let ee=ot({page:H.page,pages:H.pages,total:H.total,limit:H.limit,onPage:Z=>{y=Z,L()}});ee&&V.appendChild(ee)}}function _(l){let E=typeof m=="function"?m(l):m;return Ge(E)}function P(l){let E=!!l,D=document.createElement("form");if(D.noValidate=!0,D.innerHTML=_(l),E){let J=typeof m=="function"?m(l):m;lt(D,l)}let{close:z}=te({title:E?`Edit ${d}`:`Tambah ${d}`,content:D,size:"lg",confirmText:E?"Simpan Perubahan":`Tambah ${d}`,onConfirm:async(J,N)=>{if(!D.reportValidity())return;let O=J.querySelector(".modal-confirm");O.disabled=!0,O.textContent="Menyimpan...";let U=st(D),H=typeof m=="function"?m(l):m,X=async Z=>{for(let Y of Z)if(Y.type==="row")await X(Y.fields);else if(Y.type==="combobox"&&U[Y.name]){let re=U[Y.name],pe=(Y.options||[]).find(ce=>{let ge=String(typeof ce=="object"?ce.value:ce),Le=String(typeof ce=="object"?ce.label:ce);return ge===re||Le===re});if(pe)U[Y.name]=typeof pe=="object"?pe.value:pe;else if(Y.createApi){let ce={};ce[Y.createApi.field]=re,Y.createApi.extra&&Object.assign(ce,Y.createApi.extra);let ge=await R(Y.createApi.path,{method:"POST",body:JSON.stringify(ce)});if(ge.ok&&ge.data?.id)U[Y.name]=ge.data.id;else if(ge.ok&&!ge.data?.id)U[Y.name]=re;else throw new Error(`Gagal membuat master data: ${ge.data?.error||"Unknown error"}`)}}};try{await X(H)}catch(Z){q(Z.message),O.disabled=!1,O.textContent=E?"Simpan Perubahan":`Tambah ${d}`;return}C&&(U=await C(U,l));let W=E?"PUT":"POST",V=E?`${i}/${l.id}`:i,ee=await R(V,{method:W,body:JSON.stringify(U)});ee.ok?(G(E?`${d} berhasil diperbarui.`:`${d} berhasil ditambahkan.`),N(),ye(i),L()):(q(ee.data?.error||"Gagal menyimpan data."),O.disabled=!1,O.textContent=E?"Simpan Perubahan":`Tambah ${d}`)}})}function K(l){ie(`Hapus ${d} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let E=await R(`${i}/${l.id}`,{method:"DELETE"});E.ok?(G(`${d} berhasil dihapus.`),ye(i),L()):q(E.data?.error||"Gagal menghapus.")},`Hapus ${d}`)}return L(),L}de();async function Sa(e,t){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let n=await se(),i=await he(),r=t?t.get("dash_filter"):null,m={};if(r==="reliever"){let w=new Date,C=`${w.getFullYear()}-${String(w.getMonth()+1).padStart(2,"0")}`;m={status:"Done",month:t&&t.get("month")?t.get("month"):C}}console.log("RAW",await Te()),console.log("OPTIONS",i);let b=w=>w&&!i.find(C=>C.value===w)?[...i,{value:w,label:w}]:i,g=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],s=w=>w&&!g.includes(w)?[...g,w]:g,d=new Date().getFullYear(),S=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((w,C)=>{let M=String(C+1).padStart(2,"0");return{value:`${d}-${M}`,label:`${w} ${d}`}});Ea({container:e,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:m,onDataLoaded:w=>[...w||[]].sort((C,M)=>ut(C,M,"backup_date")),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:w=>Fe(w)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:w=>window.formatDate(w)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:w=>window.formatDate(w)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:w=>w?`<span class="badge badge-info">${w}</span>`:"-"},{key:"status",label:"Status",render:w=>fe(w)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:g},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"month",label:"Bulan",options:S},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:w=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:n,value:w?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:w?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:b(w?.original_fc_name),value:w?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:s(w?.reliever_name),value:w?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:w?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:w?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:w?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:w?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:w?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let w=await R(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(w.ok){let C=w.data.data.map(M=>({Cabang:M.branch_name||"","Nama Facility care":M.original_fc_name||"",Periode:M.period||"",Relifer:M.reliever_name||"","Tanggal Back Up":M.backup_date||"","Tanggal Selesai":M.completion_date||"",Keterangan:M.reason||"",Shift:M.shift||"",Status:M.status||""}));C.length===0&&C.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),ne(C,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async w=>{let M=(await R("/api/branches?all=1")).data?.data||[],A=u=>{if(!u)return null;let $=String(u||"").toLowerCase(),p=M.find(o=>String(o.full_name||"").toLowerCase()===$||String(o.code||"").toLowerCase()===$||String(o.name||"").toLowerCase()===$);return p?p.id:null},B=w.map(u=>({branch_name:String(u.Cabang||"").trim(),backup_date:String(u["Tanggal Back Up"]||u["Tanggal Backup"]||"").trim(),original_fc_name:String(u["Nama Facility care"]||u["FC Digantikan"]||"").trim(),reliever_name:String(u.Relifer||u.Reliefer||"").trim(),period:String(u.Periode||"").trim(),reason:String(u.Keterangan||"").trim(),shift:String(u.Shift||"").trim(),completion_date:String(u["Tanggal Selesai"]||"").trim(),status:String(u.Status||"").trim()})).filter(u=>u.reliever_name&&u.backup_date),T=await R("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:B,onDuplicate:"update"})});if(!T.ok)throw new Error(T.data?.error||"Import gagal");return T.data}}})}de();async function Fa(e){let t=await se(),n=Array.from({length:4},(i,r)=>String(new Date().getFullYear()-r));be({container:e,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:i=>Fe(i)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"fc_score",label:"Point FC",render:i=>i!=null?`<strong class="${i>=80?"text-success":i>=60?"text-warning":"text-danger"}">${i}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:i=>i!=null?`<strong>${i}</strong>`:"-"},{key:"status",label:"Status",render:i=>fe(i)},{key:"document_link",label:"Dokumen",render:i=>i?`<a href="${i}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:i=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:i?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:i?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:i?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:i?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:i?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:i?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:i?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:i?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async i=>{let r=new URLSearchParams(i||{}).toString(),m=await R(`/api/reports/inspection?limit=10000&${r}`);if(m.ok){let b=m.data.data.map(g=>({Cabang:g.branch_name||"",Periode:g.period||"",Tanggal:g.inspection_date||"","Point FC":g.fc_score!==null&&g.fc_score!==void 0?g.fc_score:"","Point SPV":g.spv_score!==null&&g.spv_score!==void 0?g.spv_score:"",Status:g.status||"","Link Dokumen":g.document_link||""}));ne(b,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async i=>{let r=s=>{if(!s)return null;let d=String(s||"").toLowerCase(),F=t.find(S=>String(S.label||"").toLowerCase()===d);return F?F.value:null},m=s=>{if(s==null||s==="")return"";if(s instanceof Date&&!isNaN(s.getTime()))return s.toISOString().slice(0,10);let d=String(s).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let S=Number(d);if(S>2e4&&S<99999){let w=new Date(Date.UTC(1899,11,30)+S*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}let F=d.split(/[\/\-\.]/);if(F.length===3){let[S,w,C]=F.map(M=>M.trim());if(S.length===4&&w.length<=2&&C.length<=2)return`${S}-${w.padStart(2,"0")}-${C.padStart(2,"0")}`;if(C.length===4&&w.length<=2&&S.length<=2)return`${C}-${w.padStart(2,"0")}-${S.padStart(2,"0")}`}return d},b=i.map(s=>({branch_id:r(String(s.Cabang||"").trim()),period:String(s.Periode||"").trim(),inspection_date:m(s.Tanggal),fc_score:s["Point FC"]!==void 0&&s["Point FC"]!==""?Number(s["Point FC"]):null,spv_score:s["Point SPV"]!==void 0&&s["Point SPV"]!==""?Number(s["Point SPV"]):null,status:String(s.Status||"").trim(),document_link:String(s["Link Dokumen"]||"").trim(),notes:String(s.Catatan||s.Keterangan||"").trim()})).filter(s=>s.branch_id&&s.period&&s.inspection_date),g=await R("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:b,onDuplicate:"update"})});if(!g.ok)throw new Error(g.data?.error||"Import gagal");return g.data}}})}de();async function Ca(e){let t=await se(),n=Array.from({length:4},(i,r)=>String(new Date().getFullYear()-r));be({container:e,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:i=>`<span class="badge ${i==="Deep Cleaning"?"badge-purple":"badge-success"}">${i}</span>`},{key:"period",label:"Periode",render:i=>Fe(i)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"status",label:"Status",render:i=>fe(i)},{key:"document_link",label:"Dokumen",render:i=>i?`<a href="${i}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:i=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t,value:i?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:i?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:i?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:i?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:i?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:i?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:i?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async i=>{let r=new URLSearchParams(i||{}).toString(),m=await R(`/api/reports/cleaning?limit=10000&${r}`);if(m.ok){let b=m.data.data.map(g=>({Cabang:g.branch_name||"",Jenis:g.activity_type||"",Periode:g.period||"",Tanggal:g.activity_date||"",Status:g.status||"","Link Dokumen":g.document_link||""}));ne(b,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async i=>{let r=s=>{if(!s)return null;let d=String(s||"").toLowerCase(),F=t.find(S=>String(S.label||"").toLowerCase()===d);return F?F.value:null},m=s=>{if(s==null||s==="")return"";if(s instanceof Date&&!isNaN(s.getTime()))return s.toISOString().slice(0,10);let d=String(s).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let S=Number(d);if(S>2e4&&S<99999){let w=new Date(Date.UTC(1899,11,30)+S*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}let F=d.split(/[\/\-\.]/);if(F.length===3){let[S,w,C]=F.map(M=>M.trim());if(S.length===4&&w.length<=2&&C.length<=2)return`${S}-${w.padStart(2,"0")}-${C.padStart(2,"0")}`;if(C.length===4&&w.length<=2&&S.length<=2)return`${C}-${w.padStart(2,"0")}-${S.padStart(2,"0")}`}return d},b=i.map(s=>({branch_id:r(String(s.Cabang||"").trim()),activity_type:String(s.Jenis||s.Kegiatan||"").trim(),period:String(s.Periode||"").trim(),activity_date:m(s.Tanggal),status:String(s.Status||"").trim(),document_link:String(s["Link Dokumen"]||"").trim(),notes:String(s.Catatan||s.Keterangan||"").trim()})).filter(s=>s.branch_id&&s.activity_type&&s.period&&s.activity_date),g=await R("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:b,onDuplicate:"update"})});if(!g.ok)throw new Error(g.data?.error||"Import gagal");return g.data}}})}de();async function $a(e,t){let n=await se(),i=Array.from({length:4},(b,g)=>String(new Date().getFullYear()-g)),r=t?t.get("dash_filter"):null,m={};if(r==="fogging"){let b=new Date,g=String(b.getMonth()+1).padStart(2,"0"),s=String(b.getFullYear()),d=t?t.get("month"):null;d&&d.length===7&&(s=d.split("-")[0],g=d.split("-")[1]),m={status:"Done",month:g,year:s}}be({container:e,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:m,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:b=>`<span class="badge badge-warning">${b}</span>`},{key:"period",label:"Periode",render:b=>Fe(b)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:b=>window.formatDate(b)},{key:"status",label:"Status",render:b=>fe(b)},{key:"document_link",label:"Dokumen",render:b=>b?`<a href="${b}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:b=>b||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:b=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:n,value:b?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:b?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:b?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:b?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:b?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:b?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async b=>{let g=new URLSearchParams(b||{}).toString(),s=await R(`/api/reports/fogging?limit=10000&${g}`);if(s.ok){let d=s.data.data.map(F=>({Cabang:F.branch_name||"",Jenis:F.activity_type||"Fogging",Periode:F.period||"",Tanggal:F.activity_date||"",Status:F.status||"","Link Dokumen":F.document_link||""}));ne(d,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async b=>{let g=S=>{if(!S)return null;let w=String(S||"").toLowerCase(),C=n.find(M=>String(M.label||"").toLowerCase()===w);return C?C.value:null},s=S=>{if(S==null||S==="")return"";if(S instanceof Date&&!isNaN(S.getTime()))return S.toISOString().slice(0,10);let w=String(S).trim();if(w===""||w==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(w))return w.slice(0,10);if(/^\d{4,5}$/.test(w)){let M=Number(w);if(M>2e4&&M<99999){let A=new Date(Date.UTC(1899,11,30)+M*864e5);return isNaN(A.getTime())?"":A.toISOString().slice(0,10)}}let C=w.split(/[\/\-\.]/);if(C.length===3){let[M,A,B]=C.map(T=>T.trim());if(M.length===4&&A.length<=2&&B.length<=2)return`${M}-${A.padStart(2,"0")}-${B.padStart(2,"0")}`;if(B.length===4&&A.length<=2&&M.length<=2)return`${B}-${A.padStart(2,"0")}-${M.padStart(2,"0")}`}return w},d=b.map(S=>({branch_id:g(String(S.Cabang||"").trim()),activity_type:String(S.Jenis||S.Kegiatan||"Fogging").trim(),period:String(S.Periode||"").trim(),activity_date:s(S.Tanggal),status:String(S.Status||"").trim(),document_link:String(S["Link Dokumen"]||"").trim(),notes:String(S.Catatan||S.Keterangan||"").trim()})).filter(S=>S.branch_id&&S.period&&S.activity_date),F=await R("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(d)});if(!F.ok)throw new Error(F.data?.error||"Import gagal");return F.data}}})}de();async function _a(e){let t=await se(),n=await he(),i=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],r=Array.from({length:4},(g,s)=>String(new Date().getFullYear()-s)),m=g=>g&&!n.find(s=>s.value===g)?[...n,{value:g,label:g}]:n,b=g=>g&&!i.find(s=>s.value===g)?[...i,{value:g,label:g}]:i;be({container:e,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:g=>window.formatDate(g)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:g=>`<span title="${g||""}">${g?.length>60?g.slice(0,60)+"\u2026":g||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:g=>window.formatDate(g)},{key:"status",label:"Status",render:g=>fe(g)},{key:"notes",label:"Keterangan",render:g=>g?.length>40?g.slice(0,40)+"\u2026":g||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:g=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:g?.branch_id},{name:"pic",label:"PIC",type:"select",options:b(g?.pic),value:g?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:g?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:g?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:g?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:g?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:g?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async g=>{let s=new URLSearchParams(g||{}).toString(),d=await R(`/api/reports/basecamp?limit=10000&${s}`);if(d.ok){let F=d.data.data.map(S=>({"Tgl Info":S.info_date||"",Cabang:S.branch_name||"",Permasalahan:S.problem||"",PIC:S.pic||"","Tgl Done":S.done_date||"",Status:S.status||"",Keterangan:S.notes||""}));ne(F,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async g=>{let s=w=>{if(!w)return null;let C=String(w||"").toLowerCase(),M=t.find(A=>String(A.label||"").toLowerCase()===C);return M?M.value:null},d=w=>{if(w==null||w==="")return"";if(w instanceof Date&&!isNaN(w.getTime()))return w.toISOString().slice(0,10);let C=String(w).trim();if(C===""||C==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(C))return C.slice(0,10);if(/^\d{4,5}$/.test(C)){let A=Number(C);if(A>2e4&&A<99999){let B=new Date(Date.UTC(1899,11,30)+A*864e5);return isNaN(B.getTime())?"":B.toISOString().slice(0,10)}}let M=C.split(/[\/\-\.]/);if(M.length===3){let[A,B,T]=M.map(u=>u.trim());if(A.length===4&&B.length<=2&&T.length<=2)return`${A}-${B.padStart(2,"0")}-${T.padStart(2,"0")}`;if(T.length===4&&B.length<=2&&A.length<=2)return`${T}-${B.padStart(2,"0")}-${A.padStart(2,"0")}`}return C},F=g.map(w=>({info_date:d(w["Tgl Info"]||w["Tanggal Info"]),branch_id:s(String(w.Cabang||"").trim()),problem:String(w.Permasalahan||"").trim(),pic:String(w.PIC||"").trim(),done_date:d(w["Tgl Done"]||w["Tanggal Done"]),status:String(w.Status||"").trim(),notes:String(w.Keterangan||w.Catatan||"").trim()})).filter(w=>w.info_date&&w.branch_id&&w.problem),S=await R("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(F)});if(!S.ok)throw new Error(S.data?.error||"Import gagal");return S.data}}})}me();var Vn=[{id:1,title:"Dusting (Pembersihan Debu)",category:"Pembersihan Rutin",icon:"\u{1F9F9}",color:"#2563EB",bg:"#EFF6FF",target:"Meja resepsionis, meja pasien, meja dokter, furniture, kaca, monitor & keyboard",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Kain Lap Microfiber Biru (Area Umum)","Kanebo (Kaca & Layar Monitor)","Kemoceng","Botol Spray Glass Cleaner","Kuas 4 Inch (Keyboard & Sela Sempit)"],steps:["Siapkan larutan glass cleaner di botol spray khusus kaca dan perabot.","Lipat kain microfiber menjadi 8 bidang kerja agar efisien dan bersih merata.","Semprotkan cairan pembersih ke permukaan kain lap, bukan langsung ke monitor atau bodi elektronik.","Usap permukaan perabot dari bagian paling atas ke bagian bawah secara searah atau pola S.","Gunakan kuas cat 4 inch untuk membersihkan sela-sela keyboard dan lubang ventilasi.","Keringkan kaca atau monitor menggunakan lap microfiber kering atau kanebo bersih."],notes:"Khusus layar monitor medis & dental unit, dilarang menyemprotkan cairan langsung ke layar."},{id:2,title:"Sweeping (Lobby Duster & Sapu Nylon)",category:"Pembersihan Rutin",icon:"\u{1F33E}",color:"#D97706",bg:"#FFFBEB",target:"Lobby utama, lorong, ruang tindakan, balkon, pantry, teras",chemical_ratio:"-",tools:["Lobby Duster Set 60cm (Area Lobby & Lorong)","Lobby Duster Set 40cm (Area Ruang Tindakan)","Sapu & Pengki Nagata 742s","Kapi Scrapper (Kerak/Permen Karet)"],steps:["Pastikan kain lobby duster dalam keadaan kering, bersih, dan bebas pasir kasar.","Dorong lobby duster menyusuri tepi dinding ruangan terlebih dahulu membentuk perimeter.","Lanjutkan mendorong dari sudut terjauh menuju arah pintu keluar tanpa mengangkat kain pel lobi dari lantai.","Kumpulkan debu dan kotoran di dekat pintu keluar, lalu angkat menggunakan sapu dan pengki Nagata.","Jika ada noda kerak atau bekas stiker/permen karet menempel di lantai, kikis perlahan menggunakan kapi scrapper.","Kibaskan kain duster di luar ruangan atau sedot sisa debu pada kain duster."],notes:"Jangan pernah menggunakan lobby duster di area lantai yang masih basah atau tergenang air."},{id:3,title:"Mopping (Pengepelan Lantai)",category:"Pembersihan Rutin",icon:"\u{1FAA3}",color:"#059669",bg:"#ECFDF5",target:"Seluruh lantai ruangan, koridor, dan area tunggu klinik",chemical_ratio:"25 ml Floor Cleaner per 1 Liter air bersih",tools:["Single Bucket Mop Wringer 20L / Double Bucket 46L","Mopping Stick / Tongkat & Kain Pel","Chemical Floor Cleaner Prokleen Lavender","Warning Sign Lantai Licin"],steps:['Pasang papan peringatan "Caution Wet Floor" di titik strategis yang mudah terlihat.',"Campurkan cairan Floor Cleaner ke dalam ember air bersih dengan takaran 25 ml per 1 Liter air.","Celupkan kain pel, masukkan ke dalam alat pemeras (wringer), dan peras hingga kondisi kain pel lembap (tidak menetes).","Mulai mengepel dari sudut terjauh ruangan mundur ke arah pintu keluar dengan gerakan membentuk pola angka delapan (8).","Bilas kain pel secara berkala apabila air pada ember bilas sudah terlihat keruh.","Biarkan lantai mengering sempurna secara alami sebelum papan peringatan wet floor dilepas."],notes:'Wajib pasang papan peringatan "Caution Wet Floor" sebelum mulai mengepel untuk mencegah risiko terpeleset.'},{id:4,title:"Membersihkan Toilet & Wastafel",category:"Sanitasi Toilet",icon:"\u{1F6BD}",color:"#DC2626",bg:"#FEF2F2",target:"Toilet pasien, toilet tim, closet, wastafel, dinding keramik, floor drain",chemical_ratio:"Karbol 25 ml:1 L Air / Bowl Cleaner 100 ml spray (kerak) / Bubuk KIFA",tools:["Sarung Tangan Karet (Wajib)","Sikat Kloset Scotch-Brite 3M","Kain Lap Microfiber Merah (Khusus Toilet)","Sponge Scotch Brite & Sikat Lantai ID-51","Floor Squeegee 75cm","Chemical Karbol Pine, Bowl Cleaner, Bubuk KIFA"],steps:["Kenakan sarung tangan karet dan siapkan perlengkapan sanitasi toilet.","Flush kloset, tuangkan Toilet Bowl Cleaner ke bagian dalam dinding kloset dan leher angsa, diamkan 5\u201310 menit.","Sikat bagian dalam kloset secara menyeluruh dengan sikat kloset, lalu bilas (flush) kembali hingga bersih.","Seka dudukan, penutup, dan tabung luar kloset menggunakan kain lap microfiber merah yang dibasahi larutan disinfektan karbol.","Bersihkan wastafel dan kran stainless menggunakan sponge tapas dengan bubuk pembersih KIFA (jangan gunakan Bowl Cleaner pada kran stainless).","Sikat lantai kamar mandi dengan sikat lantai ID-51 dan cairan karbol disinfektan.","Tarik genangan air ke arah floor drain menggunakan floor squeegee hingga lantai tidak licin dan cepat kering.","Pasang kamper/urinal block, isi ulang sabun cuci tangan, dan pasang plastik sampah baru."],notes:"Kain lap warna MERAH mutlak hanya untuk toilet dan mangkuk spittoon. Dilarang mencampur cairan kimia sembarangan."},{id:5,title:"Pembersihan Area Parkir & Luar",category:"Pembersihan Rutin",icon:"\u{1F17F}\uFE0F",color:"#4F46E5",bg:"#EEF2FF",target:"Halaman parkir mobil & motor, saluran air luar, teras depan",chemical_ratio:"-",tools:["Sapu Lidi Gagang Panjang","Pengki Nagata","Plastik Sampah Hitam Besar (90x120)","Selang Air"],steps:["Sapu daun kering, puntung rokok, plastik, dan pasir dari batas trotoar/pagar menuju satu titik.","Kumpulkan dan angkat sampah ke dalam pengki, masukkan langsung ke plastik sampah hitam besar.","Periksa dan bersihkan saluran drainase air di area parkir dari sumbatan sampah padat.","Bila ada tumpahan oli atau kotoran tanah membandel, siram dengan air dan sikat.","Rapikan pembatas parkir dan pastikan area depan klinik terlihat bersih dan rapi menyambut pasien."],notes:"Lakukan sebelum jam operasional klinik dibuka (pukul 07.30 - 08.00)."},{id:6,title:"Pembersihan Kaca & Jendela",category:"Pembersihan Rutin",icon:"\u{1FA9F}",color:"#0891B2",bg:"#ECFEFF",target:"Pintu kaca utama, jendela kaca klinik, partisi kaca, cermin toilet",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Botol Spray Glass Cleaner Prokleen","Window Wiper / Window Squeegee","Kain Microfiber Kering / Kanebo","Tongkat Teleskopik (Kaca Tinggi)"],steps:["Semprotkan larutan glass cleaner merata pada permukaan kaca yang hendak dibersihkan.","Tarik wiper karet dari arah atas ke bawah secara lurus vertikal atau zigzag horizontal tanpa terputus.","Lap bilah karet wiper setiap selesai satu tarikan menggunakan kain lap bersih agar tidak meninggalkan garis bekas air.","Gunakan kain microfiber kering untuk membersihkan sisa tetesan air di sudut-sudut lis bingkai kaca.","Untuk kaca bagian luar yang tinggi, pasang wiper pada tongkat sambungan teleskopik."],notes:"Hindari membersihkan kaca saat terpapar terik sinar matahari langsung karena cairan akan cepat mengering dan membekas noda."},{id:7,title:"Mencuci Bola Mandi Bola (Playground)",category:"Pembersihan Khusus",icon:"\u{1F939}",color:"#7C3AED",bg:"#F5F3FF",target:"Area bermain anak (Kids Playground) dan bola-bola mandi bola",chemical_ratio:"30 ml Dishwashing Liquid per 1 Liter Air + Karbol Disinfektan",tools:["Baskom Plastik Besar 55cm","Dishwashing Liquid Prokleen","Karbol Disinfektan Pine","Lap Microfiber Bersih","Karung / Wadah Pengering Jaring"],steps:["Kumpulkan seluruh bola mandi bola ke dalam wadah atau karung bersih.","Siapkan baskom plastik besar berisi air hangat kuku dicampur Dishwashing Liquid dan sedikit karbol disinfektan.","Rendam bola dan cuci secara bertahap sambil digosok perlahan menggunakan sponge/lap bersih.","Bilas bola dengan air bersih mengalir hingga tidak ada sisa busa sabun yang menempel.","Tiriskan dan angin-anginkan bola di tempat bersih beralaskan terpal/kain kering hingga benar-benar kering.","Seka dan sterilkan area kolam mandi bola sebelum bola dimasukkan kembali."],notes:"Jadwal pencucian berkala dilakukan minimal 1 kali seminggu atau sesuai kebutuhan kebersihan arena anak."},{id:8,title:"Membersihkan Filter Air Purifier",category:"Elektronik & Utility",icon:"\u{1F4A8}",color:"#0D9488",bg:"#F0FDFA",target:"Unit Air Purifier di ruang tunggu, ruang dokter, dan koridor",chemical_ratio:"-",tools:["Vacuum Cleaner (Nozzle Sikat Halus)","Kain Microfiber Kering","Kuas Cat Halus"],steps:["Matikan saklar power unit air purifier dan cabut kabel daya dari stop kontak dinding.","Buka penutup bodi (casing) depan atau belakang air purifier secara hati-hati.","Keluarkan pre-filter jaring nilon dari unit.","Bersihkan debu tebal pada pre-filter menggunakan vacuum cleaner nozzle sikat halus atau cuci dengan air mengalir (jika tipe wash-able).","Untuk filter HEPA dan Karbon Aktif: HANYA sedot debu permukaannya secara perlahan dengan vacuum, JANGAN DICUCI DENGAN AIR.","Lap bodi luar dan kisi-kisi udara masuk menggunakan lap microfiber kering.","Pastikan pre-filter benar-benar kering sebelum dipasang kembali ke posisinya semula."],notes:"Jangan pernah mencuci filter HEPA dengan air karena akan merusak serat mikro penyaring bakteri/virus."},{id:9,title:"Membersihkan Filter & Bodi AC",category:"Elektronik & Utility",icon:"\u2744\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Filter udara dan penutup indoor unit AC split di seluruh ruangan",chemical_ratio:"-",tools:["Tangga Besi Lipat","Kain Lap Microfiber Lembap","Kuas Cat Halus","Kran Air Mengalir"],steps:["Matikan AC melalui remote dan matikan MCB/saklar AC bila diperlukan.","Gunakan tangga besi lipat yang stabil di lantai datar.","Buka tutup cover depan unit indoor AC ke atas hingga terkunci.","Lepaskan kedua kisi filter saringan udara plastik secara perlahan agar debu tidak berjatuhan.","Cuci saringan filter di bawah kran air mengalir dari arah bagian dalam ke bagian luar.","Keringkan filter di tempat teduh (jangan dijemur langsung di bawah terik matahari).","Lap bodi casing dan louver/swing AC dengan kain microfiber lembap.","Pasang kembali filter yang sudah kering, tutup cover AC, dan nyalakan AC untuk pengetesan."],notes:"Pembersihan rutin mandiri dilakukan 2 minggu sekali. Deep cleaning / cuci evaporator dilakukan oleh teknisi AC vendor."},{id:10,title:"Membersihkan Dental Unit",category:"Klinis & Medis",icon:"\u{1F9B7}",color:"#E11D48",bg:"#FFF1F2",target:"Kursi dental pasien, spittoon bowl, meja instrumen, lampu dental, cover bawah",chemical_ratio:"Glass Cleaner (netral) / Karbol Disinfektan (mangkuk) / Bowl Cleaner (kerak)",tools:["Sarung Tangan Karet (Wajib)","Kain Lap Microfiber Biru (Khusus Bodi Luar)","Kain Lap Microfiber Merah (Khusus Mangkuk Spittoon)","Botol Spray Glass Cleaner (25ml + Air)","Botol Spray Karbol (20ml + Air)","Botol Spray Bowl Cleaner (100ml + Air, bila berkerak)"],steps:["Pastikan saklar utama dental unit dalam posisi OFF dan kursi dinaikkan ke posisi tertinggi untuk memudahkan akses.","Lap seluruh jok kulit, sandaran kepala, sandaran tangan, dan tiang lampu menggunakan lap microfiber biru dan cairan pembersih netral/glass cleaner.","Bersihkan mangkuk kumur (spittoon bowl) menggunakan kain lap merah dan cairan karbol disinfektan.","Jika ditemukan kerak sisa darah/plak membandel di mangkuk spittoon, gunakan sedikit bowl cleaner, sikat perlahan, lalu bilas hingga bersih.","Lap permukaan meja instrumen meja dokter. Bila ada instrumen steril, geser dengan hati-hati dan jangan mengotori instrumen.","Buka penutup cover bawah unit dental, lap debu yang terjangkau dan pastikan kondisi area selang dalam keadaan kering."],notes:"Kain lap warna MERAH hanya untuk mangkuk spittoon. Dilarang keras menggosok jok kulit dental unit dengan bahan kimia asam keras."},{id:11,title:"Membersihkan Wastafel & Keran Air",category:"Sanitasi Toilet",icon:"\u{1F6B0}",color:"#0284C7",bg:"#F0F9FF",target:"Wastafel keramik ruang tindakan, wastafel lorong, kran stainless steel, afur sink",chemical_ratio:"Bubuk Pembersih KIFA / Dishwashing Liquid Prokleen",tools:["Sponge Scotch Brite Hijau/Kuning","Kain Microfiber Hijau (Area Basah)","Bubuk Pembersih KIFA","Sikat Gigi Bekas / Sikat Kecil"],steps:["Basahi seluruh mangkuk wastafel keramik dan kran air stainless steel.","Taburkan sedikit Bubuk KIFA pada tapas sponge yang basah.","Gosok perlahan noda sabun dan kerak air putih pada kran stainless steel searah guratan logam hingga mengkilap kembali.","Gunakan sikat kecil untuk membersihkan sela leher kran dan sudut sambungan afur pembuangan wastafel.","Bilas mangkuk wastafel dan kran dengan air bersih mengalir hingga seluruh sisa residu bubuk KIFA hilang.","Keringkan kran air stainless steel dan cermin wastafel menggunakan lap microfiber hijau kering."],notes:"DILARANG KERAS menggunakan Toilet Bowl Cleaner pada kran air stainless steel karena menyebabkan karat dan warna menghitam permanen."},{id:12,title:"Membersihkan Kulkas Klinik & Pantry",category:"Elektronik & Utility",icon:"\u{1F9CA}",color:"#475569",bg:"#F8FAFC",target:"Kulkas pantry staf dan kulkas penyimpan obat/material medis (eksterior & rak)",chemical_ratio:"30 ml Dishwashing Liquid : 1 Liter Air Hangat + Lap Kering",tools:["Kain Microfiber Biru & Hijau","Dishwashing Liquid","Ember Air Hangat","Sponge Scotch Brite Halus"],steps:["Keluarkan makanan/minuman yang telah kedaluwarsa atau berbau tidak sedap.","Lepas rak-rak kaca/plastik interior kulkas yang dapat dilepas, cuci dengan sabun cuci piring dan bilas air bersih.","Seka dinding interior kulkas menggunakan kain microfiber yang dicelupkan ke dalam air hangat bersabun lembut.","Lap kering seluruh permukaan dalam kulkas dengan microfiber kering bersih.","Pasang kembali rak interior dan susun kembali bahan/makanan secara tertib dan rapi.","Bersihkan karet pintu kulkas dari noda hitam/jamur menggunakan sikat halus dan lap kering.","Lap bodi luar kulkas dan pegangan pintu menggunakan cairan pembersih netral."],notes:"Untuk kulkas penyimpanan obat/material dental, koordinasikan terlebih dahulu dengan Perawat/PIC Obat sebelum membersihkan rak interior."},{id:13,title:"Membersihkan Sofa & Kursi Ruang Tunggu",category:"Pembersihan Khusus",icon:"\u{1F6CB}\uFE0F",color:"#9333EA",bg:"#FAF5FF",target:"Sofa ruang tunggu pasien, kursi reception, kursi kerja tim",chemical_ratio:"Baking Soda (penghilang bau) / Pembersih Upholstery secukupnya",tools:["Vacuum Cleaner (Nozzle Kecil & Sikat Lembut)","Kain Microfiber Bersih","Sikat Bulu Halus","Ember Air Hangat","Cairan Pembersih Khusus Sofa (Upholstery Cleaner)","Baking Soda (Penghilang Bau Tidak Sedap)","Sarung Tangan Karet"],steps:["Kenali bahan sofa (kain fabric, beludru, kulit asli, atau kulit sintetis/oscar).","Sedot debu di seluruh permukaan, lipatan jahitan, dan celah bantal sofa menggunakan vacuum cleaner nozzle kecil.","Gunakan sikat bulu halus untuk mengangkat rambut atau kotoran yang menempel kuat di serat kain.","Jika ada noda bercak (kopi, makanan, tinta), semprotkan pembersih upholstery secukupnya, gosok memutar perlahan dengan microfiber.","Celupkan kain microfiber ke air hangat suam-suam kuku, peras hingga lembap, lalu seka seluruh permukaan sofa.","Lap kembali dengan kain kering bersih untuk menyerap sisa kelembapan.","Jika sofa berbau apek: taburkan baking soda tipis di permukaan, diamkan 20 menit, lalu sedot kembali dengan vacuum cleaner.","Biarkan sofa mengering secara alami di dalam ruangan berventilasi atau ber-AC."],notes:"Jangan menjemur sofa langsung di bawah terik matahari karena dapat menyebabkan bahan kulit pecah-pecah dan warna kain memudar."}],Wn=[{no:1,name:"Sapu Lidi Gagang",qty:"1 unit",area:"Area parkir mobil & motor klinik"},{no:2,name:"Sapu & Pengki Nagata 742s",qty:"1 set",area:"Area pantry, balkon, dak lantai atas, atau teras"},{no:3,name:"Lobby Duster Set 60 cm",qty:"1 set",area:"Lobby utama dan koridor/lorong klinik"},{no:4,name:"Lobby Duster Set 40 cm / Refill",qty:"1 unit",area:"Area ruang tindakan dokter"},{no:5,name:"Single Bucket Mop Wringer 20L",qty:"1 paket",area:"Ember pel dengan pemeras + tongkat pel"},{no:6,name:"Double Bucket Wringer Trolley 46L",qty:"1 paket",area:"Khusus klinik besar / klinik dengan fasilitas lift"},{no:7,name:"Kemoceng Bulu Lembut",qty:"1 unit",area:"Membersihkan debu kasar meja dan perabot"},{no:8,name:"Lap Microfiber Warna Merah",qty:"2 pcs",area:"Khusus area toilet, sanitasi, dan mangkuk spittoon"},{no:9,name:"Lap Microfiber Warna Biru",qty:"2 pcs",area:"Area umum kering (meja, kursi, bodi peralatan)"},{no:10,name:"Lap Microfiber Warna Hijau",qty:"2 pcs",area:"Area umum basah (pantry, wastafel luar)"},{no:11,name:"Kanebo / Synthetic Chamois",qty:"1 pcs",area:"Area umum, cermin, dan pengering kaca"},{no:12,name:"Caddy Bag Cleaning",qty:"1 unit",area:"Tas perlengkapan pinggang petugas cleaning"},{no:13,name:"Sponge Scotch Brite 3M",qty:"2 pcs",area:"Pembersihan noda lantai dan kerak ringan (spotting)"},{no:14,name:"Scotch-Brite 3M Sikat Toilet",qty:"1 unit",area:"Sikat khusus kloset dan leher angsa"},{no:15,name:"Kapi Gagang Karet / Scrapper",qty:"1 unit",area:"Pembersihan kerak lantai, noda permen karet, stiker"},{no:16,name:"Kuas Cat 4 Inch",qty:"2 pcs",area:"Pembersih sela keyboard, ventilasi, dan sudut sempit"},{no:17,name:"Warning Sign (Lantai Licin)",qty:"1 unit",area:"Digunakan saat mopping dan general cleaning"},{no:18,name:"Teleskopik + Pembersih Kaca/Langit",qty:"1 set",area:"Membersihkan jendela tinggi, plafon, dan sarang laba-laba"},{no:19,name:"Sarung Tangan Karet 3M",qty:"1 pasang",area:"Perlindungan tangan area toilet dan bahan kimia"},{no:20,name:"Botol Spray 500ml",qty:"5 unit",area:"Wadah larutan chemical (diberi label masing-masing)"},{no:21,name:"Pengharum Ruangan Stella Aerosol",qty:"1 kaleng",area:"Digunakan darurat apabila timbul bau menyengat"},{no:22,name:"Kleaner Gantungan Sapu & Pel",qty:"1 set",area:"Gantungan penyimpanan alat cleaning di janitor"},{no:23,name:"Floor Squeegee 75 cm",qty:"1 unit",area:"Pendorong genangan air lantai toilet/teras"},{no:24,name:"Sikat Botol Ikea",qty:"1 unit",area:"Di setiap ruang tindakan dokter gigi"},{no:25,name:"Scotch-Brite 3M Sikat Lantai ID-51",qty:"1 unit",area:"Sikat lantai kamar mandi dan teras"},{no:26,name:"Kamper Toilet / Urinal Block",qty:"1 pack",area:"Pengharum di toilet dan urinoir"},{no:27,name:"Plastik Sampah Hitam Sedang (60x100)",qty:"1 pack",area:"Tempat sampah indoor ruangan"},{no:28,name:"Plastik Sampah Hitam Besar (90x120)",qty:"1 pack",area:"Tempat sampah luar & penampung sampah akhir"},{no:29,name:"Plastik Sampah Medis Kuning (90x120)",qty:"1 pack",area:"Limbah medis padat B3 klinik"},{no:30,name:"Plastik Sampah Medis Kuning (40x60)",qty:"1 pack",area:"Tempat sampah medis kecil di ruang tindakan"},{no:31,name:"Tangga Besi Lipat",qty:"1 unit",area:"Alat bantu membersihkan plafon, lampu, dan AC"},{no:32,name:"Baskom Plastik Besar 55cm",qty:"1 unit",area:"Mencuci bola kids playground"}],Qn=[{no:1,name:"Floor Cleaner Prokleen Lavender",size:"5 Liter",ratio:"25 ml : 1 Liter Air",function:"Pembersihan dan pengepelan rutin seluruh lantai ruangan"},{no:2,name:"Karbol Prokleen Pine Disinfektan",size:"5 Liter",ratio:"25 ml : 1 Liter Air (Mop) / 20 ml per botol spray",function:"Mengepel lantai toilet, disinfeksi area basah dan berbau tak sedap"},{no:3,name:"Glass Cleaner Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air / Siap pakai di botol spray",function:"Membersihkan pintu kaca, jendela, cermin, dan layar monitor"},{no:4,name:"Manual Dishwashing Liquid Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air",function:"Mencuci piring, gelas, sendok pantry, dan mencuci bola playground"},{no:5,name:"Hand Soap Prokleen Green Tea",size:"5 Liter",ratio:"Siap pakai (tanpa campur)",function:"Refill sabun pencuci tangan di wastafel pasien dan tim"},{no:6,name:"Liquid Detergent Prokleen",size:"5 Liter",ratio:"Sesuai kapasitas mesin",function:"Mencuci handuk, kain microfiber, dan seragam"},{no:7,name:"Toilet Bowl Cleaner Prokleen",size:"5 Liter",ratio:"100 ml per botol spray / Khusus kerak",function:"Pembersih kerak membandel pada kloset, wastafel keramik, dan nat keramik"},{no:8,name:"Bubuk Pembersih KIFA",size:"500 gram",ratio:"Taburkan tipis pada tapas basah",function:"Membersihkan kerak air pada kran air stainless steel, jet washer, dan sink"}],Pt="fcms_sop_fc_ik_v1",Lt="fcms_sop_fc_tools_v1",It="fcms_sop_fc_chem_v1";function Aa(){try{let e=localStorage.getItem(Pt);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK",e)}return JSON.parse(JSON.stringify(Vn))}function Bt(e){localStorage.setItem(Pt,JSON.stringify(e))}function Ba(){try{let e=localStorage.getItem(Lt);if(e)return JSON.parse(e)}catch(e){console.error("Error load tools",e)}return JSON.parse(JSON.stringify(Wn))}function Tt(e){localStorage.setItem(Lt,JSON.stringify(e))}function Ta(){try{let e=localStorage.getItem(It);if(e)return JSON.parse(e)}catch(e){console.error("Error load chem",e)}return JSON.parse(JSON.stringify(Qn))}function Dt(e){localStorage.setItem(It,JSON.stringify(e))}async function Da(e){let t=ae(),n=Q(t,"sop","write"),i=Q(t,"sop","delete"),r="ik",m="",b="all",g="tools",s="",d=Aa(),F=Ba(),S=Ta();function w(){e.innerHTML=`
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
          ${i?`
          <button id="btn-reset-default" class="btn btn-outline" title="Kembalikan data ke dokumen standar asli" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F504} Reset ke Standar
          </button>
          `:""}
          <button id="btn-print-sop" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F5A8}\uFE0F Cetak Ringkasan SOP
          </button>
          ${Q(t,"export")?`<button id="btn-export-tools" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F4E5} Download Data Excel
          </button>`:""}
        </div>
      </div>

      <!-- Navigation Tabs -->
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto; flex-wrap:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none;">
        <button class="sop-tab-btn ${r==="ik"?"active":""}" data-tab="ik" style="${C(r==="ik")}">
          \u{1F4CB} Instruksi Kerja (IK) <span class="badge" style="${M(r==="ik")}">${d.length}</span>
        </button>
        <button class="sop-tab-btn ${r==="tools"?"active":""}" data-tab="tools" style="${C(r==="tools")}">
          \u{1F6E0}\uFE0F Alat & Chemical <span class="badge" style="${M(r==="tools")}">${F.length+S.length}</span>
        </button>
        <button class="sop-tab-btn ${r==="policy"?"active":""}" data-tab="policy" style="${C(r==="policy")}">
          \u{1F4D6} Pedoman & Ketentuan
        </button>
        <button class="sop-tab-btn ${r==="links"?"active":""}" data-tab="links" style="${C(r==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sop-content-area"></div>
    `,A(),B()}function C(c){return c?"border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function M(c){return c?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function A(){e.querySelectorAll(".sop-tab-btn").forEach(I=>{I.addEventListener("click",()=>{r=I.dataset.tab,w()})});let c=e.querySelector("#btn-reset-default");c&&c.addEventListener("click",()=>{ie("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja, Alat, dan Chemical ke data standar dokumen asli?",()=>{localStorage.removeItem(Pt),localStorage.removeItem(Lt),localStorage.removeItem(It),d=Aa(),F=Ba(),S=Ta(),G("Data SOP Facility Care berhasil di-reset ke standar asli"),w()},"Konfirmasi Reset Standar")});let h=e.querySelector("#btn-print-sop");h&&h.addEventListener("click",()=>{window.print()});let v=e.querySelector("#btn-export-tools");v&&v.addEventListener("click",()=>{k()})}function B(){let c=e.querySelector("#sop-content-area");c&&(r==="ik"?T(c):r==="tools"?p(c):r==="policy"?y(c):r==="links"&&x(c))}function T(c){c.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="ik-search" class="form-control" placeholder="Cari pekerjaan, alat, atau chemical..." value="${m}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="cat-pills-wrap">
              <button class="cat-pill ${b==="all"?"active":""}" data-cat="all">Semua (${d.length})</button>
              <button class="cat-pill ${b==="Pembersihan Rutin"?"active":""}" data-cat="Pembersihan Rutin">Rutin (${d.filter(P=>P.category==="Pembersihan Rutin").length})</button>
              <button class="cat-pill ${b==="Sanitasi Toilet"?"active":""}" data-cat="Sanitasi Toilet">Toilet & Wastafel (${d.filter(P=>P.category==="Sanitasi Toilet").length})</button>
              <button class="cat-pill ${b==="Klinis & Medis"?"active":""}" data-cat="Klinis & Medis">Klinis Dental (${d.filter(P=>P.category==="Klinis & Medis").length})</button>
              <button class="cat-pill ${b==="Elektronik & Utility"?"active":""}" data-cat="Elektronik & Utility">Elektronik & AC (${d.filter(P=>P.category==="Elektronik & Utility").length})</button>
              <button class="cat-pill ${b==="Pembersihan Khusus"?"active":""}" data-cat="Pembersihan Khusus">Khusus (${d.filter(P=>P.category==="Pembersihan Khusus").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${d.length} dari ${d.length} Instruksi Kerja
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
    `;let h=document.createElement("style");h.innerHTML=`
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
    `,c.appendChild(h);let v=c.querySelector("#ik-cards-grid-container"),I=c.querySelector("#ik-count-label");function j(){let P=d.filter(K=>{let l=b==="all"||K.category===b,E=m.toLowerCase().trim(),D=!E||K.title.toLowerCase().includes(E)||K.target.toLowerCase().includes(E)||K.tools.some(z=>z.toLowerCase().includes(E))||K.steps.some(z=>z.toLowerCase().includes(E));return l&&D});if(I&&(I.textContent=`Menampilkan ${P.length} dari ${d.length} Instruksi Kerja`),P.length===0){v.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}v.innerHTML=`
        <div class="sop-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${P.map(K=>`
            <div class="card ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${K.bg||"#EFF6FF"}; border-radius:10px;">${K.icon||"\u{1F4CB}"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        ${K.id}. ${K.title}
                      </h3>
                      <span class="badge" style="background:${K.bg||"#EFF6FF"}; color:${K.color||"#2563EB"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${K.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    ${n?`
                    <button class="btn btn-sm btn-ghost btn-edit-ik" data-id="${K.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    `:""}
                    ${i?`
                    <button class="btn btn-sm btn-ghost btn-delete-ik" data-id="${K.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                    `:""}
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${K.color||"#2563EB"};">
                  <strong>Area:</strong> ${K.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Bahan Utama:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${K.tools.slice(0,3).map(l=>`<li style="margin-bottom:2px;">${l}</li>`).join("")}
                    ${K.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${K.tools.length-3} alat lainnya...</li>`:""}
                  </ul>
                </div>

                ${K.chemical_ratio&&K.chemical_ratio!=="-"?`
                  <div style="font-size:0.8rem; background:#FEF3C7; color:#92400E; padding:6px 10px; border-radius:6px; margin-bottom:14px; font-weight:600;">
                    \u{1F9EA} Takaran: ${K.chemical_ratio}
                  </div>
                `:""}
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${K.steps.length} Langkah Kerja
                </span>
                <div style="display:flex; gap:6px;">
                  <button class="btn btn-sm btn-primary btn-view-ik" data-id="${K.id}" style="padding:6px 14px; font-weight:600; border-radius:6px;">
                    Lihat Panduan \u2794
                  </button>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,v.querySelectorAll(".btn-view-ik").forEach(K=>{K.addEventListener("click",()=>{let l=parseInt(K.dataset.id,10),E=d.find(D=>D.id===l);E&&u(E)})}),v.querySelectorAll(".btn-edit-ik").forEach(K=>{K.addEventListener("click",()=>{let l=parseInt(K.dataset.id,10),E=d.find(D=>D.id===l);E&&$(E)})}),v.querySelectorAll(".btn-delete-ik").forEach(K=>{K.addEventListener("click",()=>{let l=parseInt(K.dataset.id,10),E=d.find(D=>D.id===l);E&&ie(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${E.title}"?`,()=>{d=d.filter(D=>D.id!==l),Bt(d),G("Instruksi Kerja berhasil dihapus"),w()},"Hapus Instruksi Kerja")})})}j();let L=c.querySelector("#ik-search");L&&L.addEventListener("input",P=>{m=P.target.value,j()}),c.querySelectorAll(".cat-pill").forEach(P=>{P.addEventListener("click",()=>{c.querySelectorAll(".cat-pill").forEach(K=>K.classList.remove("active")),P.classList.add("active"),b=P.dataset.cat,j()})});let _=c.querySelector("#btn-add-ik");_&&_.addEventListener("click",()=>{$()})}function u(c){te({title:`${c.icon||"\u{1F4CB}"} IK-${c.id}: ${c.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${c.bg||"#EFF6FF"}; color:${c.color||"#2563EB"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${c.category}</span>
            <span>ID: IK-${String(c.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${c.color||"#2563EB"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${c.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Bahan yang Dibutuhkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${c.tools.map(h=>`<li>${h}</li>`).join("")}
              </ul>
            </div>
          </div>

          ${c.chemical_ratio&&c.chemical_ratio!=="-"?`
            <div style="background:#FEF3C7; border:1px solid #FCD34D; border-radius:8px; padding:10px 14px;">
              <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:2px;">\u{1F9EA} Rasio & Takaran Chemical:</strong>
              <span style="color:#78350F; font-size:0.9rem; font-weight:600;">${c.chemical_ratio}</span>
            </div>
          `:""}

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah-Langkah Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${c.steps.map(h=>`<li style="margin-bottom:6px;">${h}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${c.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Penting & Safety:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${c.notes}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(h,v)=>{v()}})}function $(c=null){if(!n)return;let h=!!c,v=c?c.tools.join(`
`):"",I=c?c.steps.join(`
`):"",j=`
      <form id="ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-ik-title" class="form-control" required value="${c?a(c.title):""}" placeholder="Contoh: Membersihkan Dental Unit" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-ik-icon" class="form-control" value="${c?c.icon:"\u{1F4CB}"}" placeholder="\u{1F9F9}" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
            <select id="form-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
              <option value="Pembersihan Rutin" ${c?.category==="Pembersihan Rutin"?"selected":""}>Pembersihan Rutin</option>
              <option value="Sanitasi Toilet" ${c?.category==="Sanitasi Toilet"?"selected":""}>Sanitasi Toilet</option>
              <option value="Klinis & Medis" ${c?.category==="Klinis & Medis"?"selected":""}>Klinis & Medis</option>
              <option value="Elektronik & Utility" ${c?.category==="Elektronik & Utility"?"selected":""}>Elektronik & Utility</option>
              <option value="Pembersihan Khusus" ${c?.category==="Pembersihan Khusus"?"selected":""}>Pembersihan Khusus</option>
            </select>
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Takaran / Rasio Chemical</label>
            <input type="text" id="form-ik-ratio" class="form-control" value="${c?a(c.chemical_ratio||""):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-ik-target" class="form-control" required value="${c?a(c.target):""}" placeholder="Contoh: Kursi pasien, meja instrumen, mangkuk spittoon" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Bahan yang Digunakan (1 item per baris) *</label>
          <textarea id="form-ik-tools" class="form-control" rows="4" placeholder="Kain Microfiber Biru&#10;Sarung Tangan Karet&#10;Botol Spray Glass Cleaner" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${a(v)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah-Langkah Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-ik-steps" class="form-control" rows="6" placeholder="1. Pastikan posisi saklar OFF.&#10;2. Lap jok dental unit dengan kain microfiber biru.&#10;3. Sikat mangkuk spittoon..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${a(I)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Penting / Hal yang Perlu Diperhatikan (Safety)</label>
          <textarea id="form-ik-notes" class="form-control" rows="2" placeholder="Catatan larangan atau peringatan keselamatan..." style="width:100%; border-radius:6px; font-size:0.85rem;">${c?a(c.notes||""):""}</textarea>
        </div>
      </form>
    `;te({title:h?`\u270F\uFE0F Edit IK-${c.id}: ${c.title}`:"\u2795 Tambah Instruksi Kerja Baru",content:j,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(L,_)=>{let P=L.querySelector("#form-ik-title").value.trim(),K=L.querySelector("#form-ik-icon").value.trim()||"\u{1F4CB}",l=L.querySelector("#form-ik-cat").value,E=L.querySelector("#form-ik-ratio").value.trim(),D=L.querySelector("#form-ik-target").value.trim(),z=L.querySelector("#form-ik-tools").value.trim(),J=L.querySelector("#form-ik-steps").value.trim(),N=L.querySelector("#form-ik-notes").value.trim();if(!P||!D||!z||!J){q("Harap lengkapi seluruh field wajib (*)");return}let O=z.split(`
`).map(W=>W.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),U=J.split(`
`).map(W=>W.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),X={"Pembersihan Rutin":{color:"#2563EB",bg:"#EFF6FF"},"Sanitasi Toilet":{color:"#DC2626",bg:"#FEF2F2"},"Klinis & Medis":{color:"#E11D48",bg:"#FFF1F2"},"Elektronik & Utility":{color:"#0D9488",bg:"#F0FDFA"},"Pembersihan Khusus":{color:"#7C3AED",bg:"#F5F3FF"}}[l]||{color:"#2563EB",bg:"#EFF6FF"};if(h)c.title=P,c.icon=K,c.category=l,c.chemical_ratio=E,c.target=D,c.tools=O,c.steps=U,c.notes=N,c.color=X.color,c.bg=X.bg,Bt(d),G(`Instruksi Kerja "${P}" berhasil diperbarui`);else{let W=d.length>0?Math.max(...d.map(ee=>ee.id))+1:1,V={id:W,title:P,category:l,icon:K,color:X.color,bg:X.bg,target:D,chemical_ratio:E||"-",tools:O,steps:U,notes:N};d.push(V),Bt(d),G(`Instruksi Kerja baru IK-${W} berhasil ditambahkan`)}_(),w()}})}function p(c){c.innerHTML=`
      <!-- Sub-tab Switcher & Actions -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-tools-btn" class="btn btn-sm ${g==="tools"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9F9} Alat Kebersihan (${F.length} Item)
          </button>
          <button id="view-chem-btn" class="btn btn-sm ${g==="chemical"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9EA} Bahan Chemical (${S.length} Item)
          </button>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="tool-search" class="form-control" placeholder="Cari nama / peruntukan / rasio..." value="${s}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
            ${n?`
          <button id="btn-add-tool-chem" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u2795 ${g==="tools"?"Tambah Alat Baru":"Tambah Chemical Baru"}
          </button>
            `:""}
        </div>
      </div>

      <!-- Table Container -->
      <div id="tool-table-container"></div>
    `;let h=c.querySelector("#tool-table-container"),v=c.querySelector("#tool-search");function I(){let j=s.toLowerCase().trim();if(g==="tools"){let L=F.filter(_=>!j||_.name.toLowerCase().includes(j)||_.area.toLowerCase().includes(j));h.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Standar Peralatan Kebersihan (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-primary">${L.length} Item</span>
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
                  ${L.map((_,P)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${P%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${_.no||P+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${a(_.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#E0F2FE; color:#0369A1; font-weight:600;">${a(_.qty)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${a(_.area)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                    ${n?`
                        <button class="btn btn-sm btn-ghost btn-edit-tool" data-no="${_.no}" title="Edit Alat" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                    `:""}
                    ${i?`
                        <button class="btn btn-sm btn-ghost btn-delete-tool" data-no="${_.no}" title="Hapus Alat" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    `:""}
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,h.querySelectorAll(".btn-edit-tool").forEach(_=>{_.addEventListener("click",()=>{let P=parseInt(_.dataset.no,10),K=F.find(l=>l.no===P);K&&o(K)})}),h.querySelectorAll(".btn-delete-tool").forEach(_=>{_.addEventListener("click",()=>{let P=parseInt(_.dataset.no,10),K=F.find(l=>l.no===P);K&&ie(`Apakah Anda yakin ingin menghapus alat: "${K.name}"?`,()=>{F=F.filter(l=>l.no!==P),F.forEach((l,E)=>{l.no=E+1}),Tt(F),G("Alat kebersihan berhasil dihapus"),w()},"Hapus Alat Kebersihan")})})}else{let L=S.filter(_=>!j||_.name.toLowerCase().includes(j)||_.function.toLowerCase().includes(j));h.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Bahan Kimia & Standar Takaran (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-success">${L.length} Chemical</span>
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
                  ${L.map((_,P)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${P%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${_.no||P+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${a(_.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#FEF3C7; color:#92400E; font-weight:600;">${a(_.size)}</span></td>
                      <td style="padding:10px 14px;"><span style="background:#EFF6FF; color:#1E40AF; padding:3px 8px; border-radius:6px; font-weight:600; font-size:0.82rem;">${a(_.ratio)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${a(_.function)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                    ${n?`
                        <button class="btn btn-sm btn-ghost btn-edit-chem" data-no="${_.no}" title="Edit Chemical" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                    `:""}
                    ${i?`
                        <button class="btn btn-sm btn-ghost btn-delete-chem" data-no="${_.no}" title="Hapus Chemical" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    `:""}
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,h.querySelectorAll(".btn-edit-chem").forEach(_=>{_.addEventListener("click",()=>{let P=parseInt(_.dataset.no,10),K=S.find(l=>l.no===P);K&&f(K)})}),h.querySelectorAll(".btn-delete-chem").forEach(_=>{_.addEventListener("click",()=>{let P=parseInt(_.dataset.no,10),K=S.find(l=>l.no===P);K&&ie(`Apakah Anda yakin ingin menghapus chemical: "${K.name}"?`,()=>{S=S.filter(l=>l.no!==P),S.forEach((l,E)=>{l.no=E+1}),Dt(S),G("Bahan chemical berhasil dihapus"),w()},"Hapus Bahan Chemical")})})}}I(),c.querySelector("#view-tools-btn").addEventListener("click",()=>{g="tools",p(c)}),c.querySelector("#view-chem-btn").addEventListener("click",()=>{g="chemical",p(c)}),v&&v.addEventListener("input",j=>{s=j.target.value,I()}),c.querySelector("#btn-add-tool-chem").addEventListener("click",()=>{g==="tools"?o():f()})}function o(c=null){if(!n)return;let h=!!c,v=`
      <form id="tool-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Alat / Item *</label>
          <input type="text" id="form-tool-name" class="form-control" required value="${c?a(c.name):""}" placeholder="Contoh: Lobby Duster Set 60 cm" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kuantitas Standar *</label>
          <input type="text" id="form-tool-qty" class="form-control" required value="${c?a(c.qty):"1 unit"}" placeholder="Contoh: 1 set, 2 pcs, 1 paket" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Peruntukan Area & Kegunaan *</label>
          <textarea id="form-tool-area" class="form-control" rows="3" required placeholder="Contoh: Lobby utama dan koridor/lorong klinik" style="width:100%; border-radius:6px; font-size:0.85rem;">${c?a(c.area):""}</textarea>
        </div>
      </form>
    `;te({title:h?`\u270F\uFE0F Edit Alat: ${c.name}`:"\u2795 Tambah Alat Kebersihan Baru",content:v,confirmText:"Simpan",onConfirm:(I,j)=>{let L=I.querySelector("#form-tool-name").value.trim(),_=I.querySelector("#form-tool-qty").value.trim(),P=I.querySelector("#form-tool-area").value.trim();if(!L||!_||!P){q("Harap lengkapi semua kolom");return}if(h)c.name=L,c.qty=_,c.area=P,Tt(F),G(`Alat "${L}" berhasil diperbarui`);else{let K=F.length>0?Math.max(...F.map(l=>l.no))+1:1;F.push({no:K,name:L,qty:_,area:P}),Tt(F),G(`Alat baru "${L}" berhasil ditambahkan`)}j(),w()}})}function f(c=null){if(!n)return;let h=!!c,v=`
      <form id="chem-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Produk Chemical *</label>
          <input type="text" id="form-chem-name" class="form-control" required value="${c?a(c.name):""}" placeholder="Contoh: Floor Cleaner Prokleen Lavender" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kemasan Standar *</label>
            <input type="text" id="form-chem-size" class="form-control" required value="${c?a(c.size):"5 Liter"}" placeholder="Contoh: 5 Liter, 500 gram" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Rasio & Takaran Standar *</label>
            <input type="text" id="form-chem-ratio" class="form-control" required value="${c?a(c.ratio):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Fungsi & Cara Penggunaan *</label>
          <textarea id="form-chem-function" class="form-control" rows="3" required placeholder="Contoh: Pembersihan dan pengepelan rutin seluruh lantai ruangan" style="width:100%; border-radius:6px; font-size:0.85rem;">${c?a(c.function):""}</textarea>
        </div>
      </form>
    `;te({title:h?`\u270F\uFE0F Edit Chemical: ${c.name}`:"\u2795 Tambah Chemical Baru",content:v,confirmText:"Simpan",onConfirm:(I,j)=>{let L=I.querySelector("#form-chem-name").value.trim(),_=I.querySelector("#form-chem-size").value.trim(),P=I.querySelector("#form-chem-ratio").value.trim(),K=I.querySelector("#form-chem-function").value.trim();if(!L||!_||!P||!K){q("Harap lengkapi semua kolom");return}if(h)c.name=L,c.size=_,c.ratio=P,c.function=K,Dt(S),G(`Chemical "${L}" berhasil diperbarui`);else{let l=S.length>0?Math.max(...S.map(E=>E.no))+1:1;S.push({no:l,name:L,size:_,ratio:P,function:K}),Dt(S),G(`Chemical baru "${L}" berhasil ditambahkan`)}j(),w()}})}function y(c){c.innerHTML=`
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
    `}function x(c){c.innerHTML=`
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
    `}async function k(){try{let{downloadExcel:c}=await Promise.resolve().then(()=>(de(),xe)),h=d.map(I=>({"No IK":`IK-${I.id}`,"Judul Pekerjaan":I.title,Kategori:I.category,"Sasaran Target Area":I.target,"Rasio Chemical":I.chemical_ratio||"-","Alat & Bahan":I.tools.join(", "),"Langkah Kerja":I.steps.map((j,L)=>`${L+1}. ${j}`).join(" | "),"Catatan Safety":I.notes||"-"})),v=new Date().toISOString().slice(0,10);c(h,`SOP_Facility_Care_Instruksi_Kerja_${v}`),G("Data Excel Instruksi Kerja berhasil diunduh")}catch(c){console.error(c),q("Gagal mendownload data Excel: "+c.message)}}function a(c){return c?String(c).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}w()}me();var Xn=[{id:1,title:"Pengamanan Area & Aset Perusahaan",category:"Pengamanan Fisik & Patroli",icon:"\u{1F6E1}\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Seluruh area dalam & luar Head Quarter dan Klinik Cabang Audy Dental",tools:["Google Form Checklist Monitoring Klinik: https://forms.gle/1EAvkqDJVdctXEGn7","Senter Patroli","Komputer dengan Akses Monitor CCTV (Live View)","Alat Komunikasi / Handphone / HT"],steps:["Lakukan Safety Patrol secara rutin mengelilingi seluruh area sesuai jadwal yang telah ditentukan (Patrol 1, 2, dan 3).","Catat hasil temuan dan kondisi keamanan secara berkala ke dalam Google Form Safety Patrol resmi.","Awasi kendaraan yang terparkir di area luar klinik; pastikan khusus sepeda motor dalam keadaan terkunci ganda.","Pantau situasi lingkungan dan area sensitif secara berkala melalui layar monitor CCTV saat berada di pos security.","Segera lakukan tindakan persuasif ramah apabila ditemukan potensi gangguan keamanan, dan segera hubungi pihak Kepolisian terdekat bila situasi darurat tidak dapat diatasi.","Bila ditemukan permasalahan fisik, kerusakan fasilitas, atau kendala di area klinik, segera dokumentasikan (foto) dan laporkan langsung ke SPV Klinik dan Leader Facility Care.","Disaat malam hari atau saat klinik telah tutup operasional, pastikan seluruh AC dan lampu ruangan dalam keadaan mati, kecuali lampu keamanan tertentu yang memang harus dibiarkan menyala terang."],notes:"Prioritaskan keselamatan diri dan tindakan persuasif. Dilarang meninggalkan area pengawasan tanpa personil pengganti yang jelas."},{id:2,title:"Penertiban & Pengaturan Parkir Kendaraan",category:"Operasional Parkir & Luar",icon:"\u{1F17F}\uFE0F",color:"#D97706",bg:"#FFFBEB",target:"Area parkir mobil, motor, akses keluar masuk, dan teras depan klinik",tools:["Lampu Lalu Lintas (Lampu Lalin)","Cone / Kerucut Pembatas Parkir","Peluit Parkir"],steps:["Berdiri di posisi strategis dan bantu mengarahkan setiap kendaraan pasien atau tamu yang datang secara tepat dan rapi ke slot parkir yang tersedia.","Berikan informasi dan solusi dengan sikap ramah serta santun apabila seluruh slot parkir klinik dalam kondisi penuh.","Dalam kondisi tertentu (bila ada mobil yang terhalang mobil lain dan harus dipindahkan, sementara pasien sedang dalam tindakan perawatan): bantu memindahkan kendaraan HANYA JIKA petugas security memiliki kemampuan mengemudi dan memiliki SIM aktif.","Apabila petugas security tidak memiliki kemampuan mengemudikan mobil: sampaikan permohonan maaf dengan tulus kepada pasien/tamu, minta untuk menunggu sejenak, dan segera minta bantuan kepada staf Resepsionis untuk menghubungi pemilik mobil yang menghalangi.","Bantu pasien atau tamu yang akan keluar parkir dengan memberikan aba-aba dan arahan lalu lintas jalan raya yang aman dan tepat."],notes:"Dilarang memaksakan memindahkan mobil pasien jika tidak memiliki lisensi mengemudi demi menghindari risiko gesekan atau kerusakan kendaraan."},{id:3,title:"Penerimaan Tamu & Pasien Klinik",category:"Layanan & Hospitalitas",icon:"\u{1F91D}",color:"#059669",bg:"#ECFDF5",target:"Pintu masuk utama klinik, lobby, area resepsionis, dan pos penerimaan tamu",tools:["Alat Tulis Kantor (Pulpen)","Buku Tamu / Log Book Keamanan","Alat Komunikasi / Handphone"],steps:['Berdiri dengan tegap dan sigap, bukakan pintu untuk setiap pasien/tamu yang datang dengan menerapkan budaya 5S (Senyum, Salam, Sapa, Sopan, Santun) dan menyapa: "Selamat pagi / siang / sore, ada yang bisa kami bantu?".',"Setelah pasien atau tamu memberitahukan tujuannya, persilakan dan arahkan pasien langsung menuju meja Resepsionis.","Apabila yang datang adalah tamu instansi, vendor, atau pihak Manajemen/BOD: persilakan mengisi Buku Tamu (Log Book) terlebih dahulu dan segera hubungi staf/karyawan yang ingin ditemui.",'Apabila pasien langsung melangkah ke arah tujuannya, berikan gestur tubuh hormat menyilangkan kedua tangan di depan, dan jika pasien/tamu mengucapkan terima kasih, jawab dengan ramah: "Dengan Senang Hati".',"Dalam kondisi khusus (pasien lansia, ibu hamil, atau pengguna kursi roda): petugas security wajib mendampingi, mengawal, dan membantu mendorong kursi roda sampai ke depan meja resepsionis.",'Buka kembali pintu untuk setiap pasien yang telah selesai melakukan perawatan saat hendak keluar klinik sambil mengucapkan ucapan hangat: "Terima Kasih, semoga lekas sehat kembali".'],notes:"Jaga senyum tulus dan keramahan sebagai garda terdepan representasi citra Audy Dental."}],Zn=[{no:1,type:"Aset Tetap",item:"Rumah & Bangunan",description:"Gedung klinik, pos satpam, pagar, kanopi, teras beton",pic:"Building Maintenance & Security"},{no:2,type:"Aset Tetap",item:"Genset & Lift",description:"Mesin generator listrik cadangan dan elevator penumpang tetap",pic:"Teknisi & Security"},{no:3,type:"Aset Tetap",item:"AC Central & Split Tetap",description:"Outdoor dan indoor AC yang terpasang permanen di dinding/plafon",pic:"Teknisi AC & Security"},{no:4,type:"Aset Tetap",item:"Dental Unit Permanen",description:"Kursi dental unit terpasang instalasi air dan kompresor",pic:"Tim Klinik & Security"},{no:5,type:"Aset Tetap",item:"Instalasi Listrik, Pipa & Jaringan IT Permanen",description:"Panel MCB, kabel fiber optik, pipa drainase, server rack",pic:"IT & Teknisi"},{no:6,type:"Aset Tetap",item:"Furniture Melekat (Built-in)",description:"Meja resepsionis cor/melekat, partisi kaca mati, kitchen set pantry",pic:"General Support"},{no:7,type:"Aset Bergerak",item:"Komputer, Laptop & Printer",description:"PC kasir/resepsionis, laptop staf, printer rontgen/dokumen",pic:"Staf Terkait & IT"},{no:8,type:"Aset Bergerak",item:"Kendaraan Operasional",description:"Mobil operasional dan sepeda motor inventaris perusahaan",pic:"Security & Driver"},{no:9,type:"Aset Bergerak",item:"Furniture Lepasan",description:"Meja kerja, kursi resepsionis, kursi tunggu pasien, lemari file",pic:"Facility Care & Staf"},{no:10,type:"Aset Bergerak",item:"Hand Tools, Alat Kerja & Alat Dental Portabel",description:"Handpiece dental, alat rontgen portable, toolbox teknisi",pic:"Dokter/Perawat & Teknisi"},{no:11,type:"Low Value Asset",item:"Peralatan Kantor (Telepon, ATK)",description:"Pesawat telepon, kalkulator, stapler, perlengkapan meja",pic:"Resepsionis"},{no:12,type:"Low Value Asset",item:"Alat-alat Kebersihan",description:"Lobby duster, ember wringer pel, sapu, tangga, botol spray",pic:"Facility Care"},{no:13,type:"Low Value Asset",item:"Perlengkapan Security",description:"Lampu lalin, cone parkir, peluit, senter patroli, jas hujan",pic:"Security"}],Pa={cabang:[{zone:"Zona Hijau",level:"Akses Publik Bebas",color:"#059669",bg:"#ECFDF5",rooms:["Area Parkir Kendaraan","Lobby Lantai 1","Teras Luar Klinik"],desc:"Area dengan pengamanan umum yang bebas diakses pasien dan tamu setelah pengarahan security."},{zone:"Zona Kuning",level:"Akses Terbatas / Pendampingan",color:"#D97706",bg:"#FFFBEB",rooms:["Koridor Ruang Tindakan","Tangga & Lift Klinik","Ruang Tunggu Lantai 2 & 3","Kids Playground","Toilet Lantai 1 & 2","Mushola Staf & Pasien"],desc:"Area terbatas, tamu wajib didampingi petugas atau staf klinik terkait."},{zone:"Zona Merah",level:"Area Terlarang / Restricted",color:"#DC2626",bg:"#FEF2F2",rooms:["Area Meja Receptionist / Kasir","Seluruh Ruang Dental Unit (DU 101, 102, 201, 202, 301, 302)","Ruang Sterilisasi Alat","Ruang Rontgen Radiologi","Ruang Suction & Kompresor","Ruang Tim Staf & Gudang Material Medis"],desc:"Area sangat sensitif, hanya staf medis bertugas dan atasan pemilik area. Tamu wajib izin Manajemen/BOD dan didampingi penuh."}],hq:[{floor:"Lantai 1",hijau:["Area Parkir Mobil/Motor","Lobby Resepsionis Lantai 1"],kuning:["Seluruh Ruang Meeting Lantai 1","Ruang Training Lantai 1","Koridor Lift & Toilet"],merah:["Ruang Customer Services","Gudang Utama Lantai 1","Gudang Facility Care","Gudang Teknisi & Ruang Teknisi","Pos Satpam & Ruang Genset"]},{floor:"Lantai 2",hijau:["-"],kuning:["Ruang Meeting Lantai 2","Koridor Lift & Toilet","Pantry"],merah:["Ruang CEO","Ruang Sekretaris CEO","Ruang Human Capital","Ruang Dentist Relation","Ruang General Support","Ruang Eks. Patrel"]},{floor:"Lantai 3",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry","Mushola & Tempat Wudhu","Ruang Santai Karyawan"],merah:["Ruang Marketing","Studio Konten Marketing","Gudang Marketing","Ruang Patient Relation","Ruang OPS","Ruang Training Lantai 3","Gudang Lantai 3"]},{floor:"Lantai 4",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry"],merah:["Ruang Finance & Accounting","Ruang VP Finance & Accounting","Ruang Audit Internal","Ruang Training Lantai 4","Gudang Lantai 4"]},{floor:"Lantai 5",hijau:["-"],kuning:["Lobby Lantai 5","Pantry Lantai 5","Toilet Lantai 5","Ruang Training Besar Lantai 5"],merah:["Area Dak Toren Air & Rooftop"]}]},jt="fcms_sop_sec_ik_v1",Kt="fcms_sop_sec_assets_v1";function La(){try{let e=localStorage.getItem(jt);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK Sec",e)}return JSON.parse(JSON.stringify(Xn))}function zt(e){localStorage.setItem(jt,JSON.stringify(e))}function Ia(){try{let e=localStorage.getItem(Kt);if(e)return JSON.parse(e)}catch(e){console.error("Error load assets sec",e)}return JSON.parse(JSON.stringify(Zn))}function Mt(e){localStorage.setItem(Kt,JSON.stringify(e))}async function za(e){let t=ae(),n=Q(t,"sop","write"),i=Q(t,"sop","delete"),r="ik",m="",b="all",g="",s="cabang",d=La(),F=Ia();function S(){e.innerHTML=`
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
          ${i?`
          <button id="btn-reset-sec-default" class="btn btn-outline" title="Kembalikan data ke dokumen standar asli" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F504} Reset ke Standar
          </button>
          `:""}
          <button id="btn-print-sec" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F5A8}\uFE0F Cetak Ringkasan SOP
          </button>
          ${Q(t,"export")?`<button id="btn-export-sec" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F4E5} Download Data Excel
          </button>`:""}
        </div>
      </div>

      <!-- Navigation Tabs -->
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto; flex-wrap:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none;">
        <button class="sec-tab-btn ${r==="ik"?"active":""}" data-tab="ik" style="${w(r==="ik")}">
          \u{1F6E1}\uFE0F Instruksi Kerja (IK) <span class="badge" style="${C(r==="ik")}">${d.length}</span>
        </button>
        <button class="sec-tab-btn ${r==="zonasi"?"active":""}" data-tab="zonasi" style="${w(r==="zonasi")}">
          \u{1F5FA}\uFE0F Zonasi Akses & Denah
        </button>
        <button class="sec-tab-btn ${r==="assets"?"active":""}" data-tab="assets" style="${w(r==="assets")}">
          \u{1F4E6} Klasifikasi Aset <span class="badge" style="${C(r==="assets")}">${F.length}</span>
        </button>
        <button class="sec-tab-btn ${r==="policy"?"active":""}" data-tab="policy" style="${w(r==="policy")}">
          \u{1F4CB} Shift, Patrol & Insiden
        </button>
        <button class="sec-tab-btn ${r==="links"?"active":""}" data-tab="links" style="${w(r==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sec-content-area"></div>
    `,M(),A()}function w(a){return a?"border:none; background:#1E40AF; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function C(a){return a?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function M(){e.querySelectorAll(".sec-tab-btn").forEach(v=>{v.addEventListener("click",()=>{r=v.dataset.tab,S()})});let a=e.querySelector("#btn-reset-sec-default");a&&a.addEventListener("click",()=>{ie("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja dan Klasifikasi Aset Security ke standar dokumen asli?",()=>{localStorage.removeItem(jt),localStorage.removeItem(Kt),d=La(),F=Ia(),G("Data SOP Security berhasil di-reset ke standar dokumen"),S()},"Reset ke Standar Dokumen")});let c=e.querySelector("#btn-print-sec");c&&c.addEventListener("click",()=>{window.print()});let h=e.querySelector("#btn-export-sec");h&&h.addEventListener("click",()=>{x()})}function A(){let a=e.querySelector("#sec-content-area");a&&(r==="ik"?B(a):r==="zonasi"?$(a):r==="assets"?p(a):r==="policy"?f(a):r==="links"&&y(a))}function B(a){a.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="sec-ik-search" class="form-control" placeholder="Cari instruksi kerja security..." value="${m}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="sec-cat-pills-wrap">
              <button class="sec-cat-pill ${b==="all"?"active":""}" data-cat="all">Semua (${d.length})</button>
              <button class="sec-cat-pill ${b==="Pengamanan Fisik & Patroli"?"active":""}" data-cat="Pengamanan Fisik & Patroli">Patroli & Fisik (${d.filter(_=>_.category==="Pengamanan Fisik & Patroli").length})</button>
              <button class="sec-cat-pill ${b==="Operasional Parkir & Luar"?"active":""}" data-cat="Operasional Parkir & Luar">Parkir & Luar (${d.filter(_=>_.category==="Operasional Parkir & Luar").length})</button>
              <button class="sec-cat-pill ${b==="Layanan & Hospitalitas"?"active":""}" data-cat="Layanan & Hospitalitas">Tamu & Pasien (${d.filter(_=>_.category==="Layanan & Hospitalitas").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="sec-ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${d.length} Instruksi Kerja
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
    `;let c=document.createElement("style");c.innerHTML=`
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
    `,a.appendChild(c);let h=a.querySelector("#sec-ik-cards-grid-container"),v=a.querySelector("#sec-ik-count-label");function I(){let _=d.filter(P=>{let K=b==="all"||P.category===b,l=m.toLowerCase().trim(),E=!l||P.title.toLowerCase().includes(l)||P.target.toLowerCase().includes(l)||P.tools.some(D=>D.toLowerCase().includes(l))||P.steps.some(D=>D.toLowerCase().includes(l));return K&&E});if(v&&(v.textContent=`Menampilkan ${_.length} dari ${d.length} Instruksi Kerja`),_.length===0){h.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}h.innerHTML=`
        <div class="sec-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${_.map(P=>`
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
                    ${i?`
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
                    ${P.tools.slice(0,3).map(K=>`<li style="margin-bottom:2px;">${k(K)}</li>`).join("")}
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
      `,h.querySelectorAll(".btn-view-sec-ik").forEach(P=>{P.addEventListener("click",()=>{let K=parseInt(P.dataset.id,10),l=d.find(E=>E.id===K);l&&T(l)})}),h.querySelectorAll(".btn-edit-sec-ik").forEach(P=>{P.addEventListener("click",()=>{let K=parseInt(P.dataset.id,10),l=d.find(E=>E.id===K);l&&u(l)})}),h.querySelectorAll(".btn-delete-sec-ik").forEach(P=>{P.addEventListener("click",()=>{let K=parseInt(P.dataset.id,10),l=d.find(E=>E.id===K);l&&ie(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${l.title}"?`,()=>{d=d.filter(E=>E.id!==K),zt(d),G("Instruksi Kerja Security berhasil dihapus"),S()},"Hapus Instruksi Kerja")})})}I();let j=a.querySelector("#sec-ik-search");j&&j.addEventListener("input",_=>{m=_.target.value,I()}),a.querySelectorAll(".sec-cat-pill").forEach(_=>{_.addEventListener("click",()=>{a.querySelectorAll(".sec-cat-pill").forEach(P=>P.classList.remove("active")),_.classList.add("active"),b=_.dataset.cat,I()})});let L=a.querySelector("#btn-add-sec-ik");L&&L.addEventListener("click",()=>{u()})}function T(a){te({title:`${a.icon||"\u{1F6E1}\uFE0F"} IK-${String(a.id).padStart(2,"0")}: ${a.title}`,size:"lg",content:`
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
                ${a.tools.map(c=>{if(c.includes("http")){let h=c.split(": ");return`<li>${h[0]}: <a href="${h.slice(1).join(": ")}" target="_blank" rel="noopener" style="color:#2563EB; font-weight:600; text-decoration:underline;">Buka Tautan Form \u2794</a></li>`}return`<li>${k(c)}</li>`}).join("")}
              </ul>
            </div>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah & Prosedur Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${a.steps.map(c=>`<li style="margin-bottom:6px;">${k(c)}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${a.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Khusus & Ketentuan Penting:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${k(a.notes)}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(c,h)=>{h()}})}function u(a=null){if(!n)return;let c=!!a,h=a?a.tools.join(`
`):"",v=a?a.steps.join(`
`):"",I=`
      <form id="sec-ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-sec-ik-title" class="form-control" required value="${a?k(a.title):""}" placeholder="Contoh: Pengamanan Area & Aset Perusahaan" style="width:100%; border-radius:6px; height:36px;">
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
          <input type="text" id="form-sec-ik-target" class="form-control" required value="${a?k(a.target):""}" placeholder="Contoh: Seluruh area luar klinik dan akses parkir" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Perlengkapan (1 per baris) *</label>
          <textarea id="form-sec-ik-tools" class="form-control" rows="3" placeholder="Senter Patroli&#10;Lampu Lalin&#10;Peluit" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${k(h)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah & Prosedur Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-sec-ik-steps" class="form-control" rows="5" placeholder="1. Lakukan Safety Patrol sesuai jadwal.&#10;2. Arahkan kendaraan pasien dengan ramah..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${k(v)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Khusus / Ketentuan Penting</label>
          <textarea id="form-sec-ik-notes" class="form-control" rows="2" placeholder="Peringatan keselamatan atau pedoman khusus..." style="width:100%; border-radius:6px; font-size:0.85rem;">${a?k(a.notes||""):""}</textarea>
        </div>
      </form>
    `;te({title:c?`\u270F\uFE0F Edit IK-${a.id}: ${a.title}`:"\u2795 Tambah Instruksi Kerja Security Baru",content:I,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(j,L)=>{let _=j.querySelector("#form-sec-ik-title").value.trim(),P=j.querySelector("#form-sec-ik-icon").value.trim()||"\u{1F6E1}\uFE0F",K=j.querySelector("#form-sec-ik-cat").value,l=j.querySelector("#form-sec-ik-target").value.trim(),E=j.querySelector("#form-sec-ik-tools").value.trim(),D=j.querySelector("#form-sec-ik-steps").value.trim(),z=j.querySelector("#form-sec-ik-notes").value.trim();if(!_||!l||!E||!D){q("Harap lengkapi semua kolom bertanda bintang (*)");return}let J=E.split(`
`).map(H=>H.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),N=D.split(`
`).map(H=>H.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),U={"Pengamanan Fisik & Patroli":{color:"#2563EB",bg:"#EFF6FF"},"Operasional Parkir & Luar":{color:"#D97706",bg:"#FFFBEB"},"Layanan & Hospitalitas":{color:"#059669",bg:"#ECFDF5"}}[K]||{color:"#1E40AF",bg:"#EFF6FF"};if(c)a.title=_,a.icon=P,a.category=K,a.target=l,a.tools=J,a.steps=N,a.notes=z,a.color=U.color,a.bg=U.bg,zt(d),G(`Instruksi Kerja "${_}" berhasil diperbarui`);else{let H=d.length>0?Math.max(...d.map(W=>W.id))+1:1,X={id:H,title:_,category:K,icon:P,color:U.color,bg:U.bg,target:l,tools:J,steps:N,notes:z};d.push(X),zt(d),G(`Instruksi Kerja baru IK-${H} berhasil ditambahkan`)}L(),S()}})}function $(a){a.innerHTML=`
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:20px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-cabang-btn" class="btn btn-sm ${s==="cabang"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F3E5} Zonasi Klinik / Cabang
          </button>
          <button id="view-hq-btn" class="btn btn-sm ${s==="hq"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
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
    `;let c=a.querySelector("#zonasi-content-container");s==="cabang"?c.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
          ${Pa.cabang.map(h=>`
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; border-bottom:1px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.05rem; font-weight:700; color:${h.color}; display:flex; align-items:center; gap:8px;">
                  <span style="width:12px; height:12px; border-radius:50%; background:${h.color}; display:inline-block;"></span>
                  ${h.zone}
                </h3>
                <span class="badge" style="background:${h.bg}; color:${h.color}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:6px;">
                  ${h.level}
                </span>
              </div>
              <p style="font-size:0.82rem; color:var(--text-3); margin-bottom:14px; line-height:1.5;">${h.desc}</p>
              <div>
                <strong style="font-size:0.8rem; text-transform:uppercase; color:var(--text-2); display:block; margin-bottom:6px;">Daftar Ruangan Terkait:</strong>
                <ul style="margin:0 0 0 16px; padding:0; font-size:0.85rem; color:var(--text-1); line-height:1.7;">
                  ${h.rooms.map(v=>`<li>${v}</li>`).join("")}
                </ul>
              </div>
            </div>
          `).join("")}
        </div>
      `:c.innerHTML=`
        <div style="display:flex; flex-direction:column; gap:16px;">
          ${Pa.hq.map(h=>`
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; border-bottom:2px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1); display:flex; align-items:center; gap:8px;">
                  <span>\u{1F3E2}</span> Head Quarter \u2014 ${h.floor}
                </h3>
              </div>
              <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:14px;">
                <!-- Zona Hijau -->
                <div style="background:#ECFDF5; border:1px solid #A7F3D0; border-radius:8px; padding:12px;">
                  <strong style="color:#065F46; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E2} ZONA HIJAU (Publik):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#064E3B; line-height:1.6;">
                    ${h.hijau.map(v=>`<li>${v}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Kuning -->
                <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:8px; padding:12px;">
                  <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E1} ZONA KUNING (Terbatas):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#78350F; line-height:1.6;">
                    ${h.kuning.map(v=>`<li>${v}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Merah -->
                <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:8px; padding:12px;">
                  <strong style="color:#991B1B; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F534} ZONA MERAH (Restricted):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#7F1D1D; line-height:1.6;">
                    ${h.merah.map(v=>`<li>${v}</li>`).join("")}
                  </ul>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,a.querySelector("#view-cabang-btn").addEventListener("click",()=>{s="cabang",$(a)}),a.querySelector("#view-hq-btn").addEventListener("click",()=>{s="hq",$(a)})}function p(a){a.innerHTML=`
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
          <input type="text" id="sec-asset-search" class="form-control" placeholder="Cari nama aset / tipe / PIC..." value="${g}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
            ${n?`
          <button id="btn-add-sec-asset" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
            \u2795 Tambah Aset Baru
          </button>
            `:""}
        </div>
      </div>

      <div id="sec-asset-table-container"></div>
    `;let c=a.querySelector("#sec-asset-table-container"),h=a.querySelector("#sec-asset-search");function v(){let j=g.toLowerCase().trim(),L=F.filter(_=>!j||_.item.toLowerCase().includes(j)||_.type.toLowerCase().includes(j)||_.description.toLowerCase().includes(j)||_.pic&&_.pic.toLowerCase().includes(j));c.innerHTML=`
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
            <strong style="color:var(--text-1); font-size:0.95rem;">Rekap Klasifikasi Aset Perusahaan</strong>
            <span class="badge badge-primary">${L.length} Aset</span>
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
                ${L.map((_,P)=>`
                  <tr style="border-bottom:1px solid #F1F5F9; background:${P%2===0?"#fff":"#FAFAFA"};">
                    <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${_.no||P+1}</td>
                    <td style="padding:10px 14px;">
                      <span class="badge" style="${I(_.type)}">${k(_.type)}</span>
                    </td>
                    <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${k(_.item)}</td>
                    <td style="padding:10px 14px; color:var(--text-2);">${k(_.description)}</td>
                    <td style="padding:10px 14px; font-weight:600; color:#1E40AF;">${k(_.pic||"-")}</td>
                    <td style="padding:10px 14px; text-align:center;">
                    ${n?`
                      <button class="btn btn-sm btn-ghost btn-edit-asset" data-no="${_.no}" title="Edit Aset" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                    `:""}
                    ${i?`
                      <button class="btn btn-sm btn-ghost btn-delete-asset" data-no="${_.no}" title="Hapus Aset" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    `:""}
                    </td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </div>
      `,c.querySelectorAll(".btn-edit-asset").forEach(_=>{_.addEventListener("click",()=>{let P=parseInt(_.dataset.no,10),K=F.find(l=>l.no===P);K&&o(K)})}),c.querySelectorAll(".btn-delete-asset").forEach(_=>{_.addEventListener("click",()=>{let P=parseInt(_.dataset.no,10),K=F.find(l=>l.no===P);K&&ie(`Apakah Anda yakin ingin menghapus aset: "${K.item}"?`,()=>{F=F.filter(l=>l.no!==P),F.forEach((l,E)=>{l.no=E+1}),Mt(F),G("Aset berhasil dihapus"),S()},"Hapus Aset")})})}function I(j){return j==="Aset Tetap"?"background:#EFF6FF; color:#1E40AF; font-weight:700; padding:3px 8px; border-radius:6px;":j==="Aset Bergerak"?"background:#ECFDF5; color:#047857; font-weight:700; padding:3px 8px; border-radius:6px;":"background:#FEF3C7; color:#92400E; font-weight:700; padding:3px 8px; border-radius:6px;"}v(),h&&h.addEventListener("input",j=>{g=j.target.value,v()}),a.querySelector("#btn-add-sec-asset").addEventListener("click",()=>{o()})}function o(a=null){if(!n)return;let c=!!a,h=`
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
          <input type="text" id="form-asset-item" class="form-control" required value="${a?k(a.item):""}" placeholder="Contoh: Genset & Elevator" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Deskripsi / Contoh Item *</label>
          <textarea id="form-asset-desc" class="form-control" rows="3" required placeholder="Contoh: Generator daya darurat dan panel listrik utama" style="width:100%; border-radius:6px; font-size:0.85rem;">${a?k(a.description):""}</textarea>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">PIC Pengawas / Penanggung Jawab *</label>
          <input type="text" id="form-asset-pic" class="form-control" required value="${a?k(a.pic):""}" placeholder="Contoh: Teknisi & Security" style="width:100%; border-radius:6px; height:36px;">
        </div>
      </form>
    `;te({title:c?`\u270F\uFE0F Edit Aset: ${a.item}`:"\u2795 Tambah Klasifikasi Aset Baru",content:h,confirmText:"Simpan",onConfirm:(v,I)=>{let j=v.querySelector("#form-asset-type").value,L=v.querySelector("#form-asset-item").value.trim(),_=v.querySelector("#form-asset-desc").value.trim(),P=v.querySelector("#form-asset-pic").value.trim();if(!L||!_||!P){q("Harap lengkapi semua kolom");return}if(c)a.type=j,a.item=L,a.description=_,a.pic=P,Mt(F),G(`Aset "${L}" berhasil diperbarui`);else{let K=F.length>0?Math.max(...F.map(l=>l.no))+1:1;F.push({no:K,type:j,item:L,description:_,pic:P}),Mt(F),G(`Aset baru "${L}" berhasil ditambahkan`)}I(),S()}})}function f(a){a.innerHTML=`
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
    `}function y(a){a.innerHTML=`
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
    `}async function x(){try{let{downloadExcel:a}=await Promise.resolve().then(()=>(de(),xe)),c=d.map(v=>({"Kode IK":`IK-SEC-${String(v.id).padStart(2,"0")}`,"Judul Prosedur":v.title,Kategori:v.category,"Sasaran Target":v.target,"Alat & Perlengkapan":v.tools.join(", "),"Langkah Pelaksanaan":v.steps.map((I,j)=>`${j+1}. ${I}`).join(" | "),"Catatan Khusus":v.notes||"-"})),h=new Date().toISOString().slice(0,10);a(c,`SOP_Security_Instruksi_Kerja_${h}`),G("Data Excel SOP Security berhasil diunduh")}catch(a){console.error(a),q("Gagal mendownload data Excel: "+a.message)}}function k(a){return a?String(a).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}S()}me();var Ma=[{id:1,name:"Berita Acara Kerusakan Alat",category:"Berita Acara",document_link:"https://docs.google.com/document/d/1example_ba_kerusakan",description:"Format berita acara resmi pencatatan dan eskalasi kerusakan peralatan klinik dan aset fasilitas."},{id:2,name:"Checklist Apar",category:"Apar & K3",document_link:"https://docs.google.com/spreadsheets/d/1example_apar",description:"Lembar inspeksi berkala tabung pemadam api ringan (APAR), tekanan manometer, dan masa kadaluarsa."},{id:3,name:"Checklist Deep Cleaning",category:"Kegiatan Deep Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_deep_cleaning",description:"Panduan dan verifikasi checklist pembersihan mendalam (deep cleaning) area dental unit dan klinik."},{id:4,name:"Checklist Fulldraine",category:"Drain Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_fulldraine",description:"Checklist monitoring pembersihan saluran pembuangan air kotor, floor drain, dan sedotan suction."},{id:5,name:"Checklist General Cleaning",category:"Kegiatan General Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_general_cleaning",description:"Standar checklist kebersihan harian/mingguan seluruh ruangan, lobby, toilet, dan koridor."},{id:6,name:"Form Aktivity Reliefer",category:"Form Operasional",document_link:"https://forms.gle/1example_reliefer",description:"Formulir pencatatan absensi dan penugasan harian reliefer Facility Care di cabang pengganti."},{id:7,name:"Form Aktivity Security HQ",category:"Form Operasional",document_link:"https://forms.gle/1example_security_hq",description:"Formulir monitoring shift dan pengawasan keamanan harian di Kantor Pusat (Head Quarter)."}];function ja(e=""){let t=(e||"").toLowerCase();return t.includes("berita acara")?{label:e,bg:"#F3E8FF",color:"#7E22CE",border:"#D8B4FE",icon:"\u2696\uFE0F"}:t.includes("apar")||t.includes("k3")||t.includes("safety")?{label:e,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F9EF}"}:t.includes("deep cleaning")?{label:e,bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2728"}:t.includes("drain")?{label:e,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u{1F6BF}"}:t.includes("general cleaning")||t.includes("kebersihan")?{label:e,bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F9F9}"}:t.includes("form")?{label:e,bg:"#E0E7FF",color:"#4338CA",border:"#C7D2FE",icon:"\u{1F4DD}"}:{label:e||"Dokumen",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4C1}"}}function Ka(e=""){if(!e)return{type:"Tidak Ada Tautan",icon:"\u{1F4CE}",badgeClass:"bg-slate-100 text-slate-500"};let t=e.toLowerCase();return t.includes("forms.gle")||t.includes("google.com/forms")?{type:"Google Form",icon:"\u{1F7E3}",badgeClass:"badge-form",color:"#673AB7"}:t.includes("docs.google.com/spreadsheets")||t.includes("sheet")?{type:"Google Sheets",icon:"\u{1F7E2}",badgeClass:"badge-sheet",color:"#0F9D58"}:t.includes("docs.google.com/document")||t.includes("doc")?{type:"Google Docs",icon:"\u{1F535}",badgeClass:"badge-doc",color:"#4285F4"}:t.includes("drive.google.com")?{type:"Google Drive",icon:"\u{1F7E1}",badgeClass:"badge-drive",color:"#F4B400"}:t.endsWith(".pdf")?{type:"PDF Document",icon:"\u{1F534}",badgeClass:"badge-pdf",color:"#EA4335"}:{type:"Web Link",icon:"\u{1F310}",badgeClass:"badge-link",color:"#2563EB"}}async function Na(e){let t=ae(),n=Q(t,"checklist","write"),i=Q(t,"checklist","delete"),r=[],m="all",b="",g="fcms-master-doc-styles";if(!document.getElementById(g)){let T=document.createElement("style");T.id=g,T.textContent=`
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
    `,document.head.appendChild(T)}async function s(){try{let T=await R("/api/checklist?limit=1000");if(T.ok&&T.data?.data&&Array.isArray(T.data.data)&&T.data.data.length>0)r=T.data.data,localStorage.setItem("fcms_master_doc_cache",JSON.stringify(r));else{let u=localStorage.getItem("fcms_master_doc_cache");u?r=JSON.parse(u):(r=[...Ma],localStorage.setItem("fcms_master_doc_cache",JSON.stringify(r)))}}catch(T){console.warn("Fallback loading documents from cache/seed:",T);let u=localStorage.getItem("fcms_master_doc_cache");r=u?JSON.parse(u):[...Ma]}A()}function d(){return r.filter(T=>{let u=m==="all"||T.category===m,$=b.toLowerCase().trim(),p=!$||(T.name||"").toLowerCase().includes($)||(T.category||"").toLowerCase().includes($)||(T.description||"").toLowerCase().includes($);return u&&p})}function F(){let T={all:r.length};return r.forEach(u=>{let $=u.category||"Lainnya";T[$]=(T[$]||0)+1}),T}function S(T=null){if(!n)return;let u=!!T,$=Array.from(new Set(r.map(o=>o.category).filter(Boolean))),p=document.createElement("div");p.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Nama Dokumen / Checklist <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Checklist Pembersihan Wastafel & Toilet" value="${T?.name?String(T.name).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Kategori Dokumen <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-category" list="doc-categories-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Pilih atau ketik kategori baru..." value="${T?.category?String(T.category).replace(/"/g,"&quot;"):""}" required />
          <datalist id="doc-categories-list">
            ${$.map(o=>`<option value="${o}">${o}</option>`).join("")}
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
            placeholder="https://docs.google.com/..." value="${T?.document_link?String(T.document_link).replace(/"/g,"&quot;"):""}" />
          <span style="font-size:0.75rem; color:#64748B; margin-top:4px; display:block;">
            Bisa berupa tautan Google Forms untuk inspeksi harian, atau link Google Sheets / Docs sebagai acuan.
          </span>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Deskripsi / Petunjuk Penggunaan
          </label>
          <textarea id="doc-modal-desc" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan keterangan singkat mengenai tujuan dan instruksi pengisian dokumen ini...">${T?.description||""}</textarea>
        </div>
      </div>
    `,te({title:u?"\u270F\uFE0F Edit Master Dokumen":"\u2795 Tambah Master Dokumen Baru",content:p,size:"md",confirmText:u?"Simpan Perubahan":"Tambahkan Dokumen",confirmClass:"btn-primary",onConfirm:async(o,f)=>{let y=p.querySelector("#doc-modal-name"),x=p.querySelector("#doc-modal-category"),k=p.querySelector("#doc-modal-link"),a=p.querySelector("#doc-modal-desc"),c=(y.value||"").trim(),h=(x.value||"").trim(),v=(k.value||"").trim(),I=(a.value||"").trim();if(!c){q("Nama Dokumen wajib diisi"),y.focus();return}if(!h){q("Kategori Dokumen wajib diisi"),x.focus();return}let j={name:c,category:h,document_link:v,description:I};try{if(u)if((await R(`/api/checklist/${T.id}`,{method:"PUT",body:JSON.stringify(j)})).ok){let _=r.findIndex(P=>P.id===T.id);_!==-1&&(r[_]={...r[_],...j}),G("Dokumen berhasil diperbarui")}else{let _=r.findIndex(P=>P.id===T.id);_!==-1&&(r[_]={...r[_],...j}),G("Dokumen disimpan secara lokal")}else{let L=await R("/api/checklist",{method:"POST",body:JSON.stringify(j)});if(L.ok&&L.data?.id)r.unshift({id:L.data.id,...j}),G("Dokumen baru berhasil ditambahkan");else{let _=Date.now();r.unshift({id:_,...j}),G("Dokumen baru ditambahkan secara lokal")}}localStorage.setItem("fcms_master_doc_cache",JSON.stringify(r)),f(),A()}catch(L){console.error("Error saving document:",L),q("Gagal menyimpan dokumen: "+L.message)}}})}function w(T){ie(`Apakah Anda yakin ingin menghapus dokumen <strong>"${T.name}"</strong>?`,async()=>{try{let u=await R(`/api/checklist/${T.id}`,{method:"DELETE"});u.ok||console.warn("API delete warning:",u),r=r.filter($=>$.id!==T.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(r)),G(`Dokumen "${T.name}" berhasil dihapus`),A()}catch{r=r.filter($=>$.id!==T.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(r)),G(`Dokumen "${T.name}" berhasil dihapus`),A()}},"Hapus Dokumen")}function C(T,u){if(!T){wt("Dokumen belum memiliki tautan digital.");return}navigator.clipboard.writeText(T).then(()=>{G(`Tautan "${u}" disalin ke clipboard!`)}).catch(()=>{wt(`Tautan: ${T}`)})}async function M(){if(!Q(t,"export")){q("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:T}=await Promise.resolve().then(()=>(de(),xe)),u=r.map(($,p)=>({No:p+1,"Nama Dokumen":$.name||"",Kategori:$.category||"","Platform / Format":Ka($.document_link).type,"Link Dokumen":$.document_link||"","Deskripsi / Keterangan":$.description||""}));T(u,`Master_Dokumen_FCMS_${new Date().toISOString().slice(0,10)}`),G("Berhasil mengunduh data dokumen ke Excel.")}catch(T){q("Gagal mengekspor data: "+T.message)}}function A(){let T=d(),u=F(),$=Object.keys(u).filter(o=>o!=="all"),p=r.filter(o=>!!o.document_link).length;e.innerHTML=`
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
              ${Q(t,"export")?`<button id="doc-btn-excel" class="btn btn-ghost" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
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
                <div class="doc-stat-value">${r.length}</div>
                <div class="doc-stat-label">Total Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon">\u{1F5C2}\uFE0F</div>
              <div>
                <div class="doc-stat-value">${$.length}</div>
                <div class="doc-stat-label">Kategori Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon" style="background:rgba(16,185,129,0.18); color:#34D399;">\u{1F310}</div>
              <div>
                <div class="doc-stat-value">${p} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">/ ${r.length}</span></div>
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
              <input type="text" id="doc-search-input" class="doc-search-input" placeholder="Cari nama dokumen, kategori, atau deskripsi..." value="${b?b.replace(/"/g,"&quot;"):""}" />
            </div>
          </div>

          <!-- Category Pills -->
          <div class="doc-filter-pills">
            <button class="doc-pill ${m==="all"?"active":""}" data-cat="all">
              <span>Semua Dokumen</span>
              <span class="doc-pill-count">${u.all||0}</span>
            </button>
            ${$.map(o=>{let f=ja(o);return`
                <button class="doc-pill ${m===o?"active":""}" data-cat="${o}">
                  <span>${f.icon}</span>
                  <span>${o}</span>
                  <span class="doc-pill-count">${u[o]||0}</span>
                </button>
              `}).join("")}
          </div>
        </div>

        <!-- Documents View Container -->
        <div id="doc-list-view">
          ${T.length===0?`
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
              ${T.map(o=>{let f=ja(o.category),y=Ka(o.document_link);return`
                  <div class="doc-card" data-id="${o.id}">
                    <div>
                      <div class="doc-card-top">
                        <span class="doc-category-badge" style="background:${f.bg}; color:${f.color}; border:1px solid ${f.border};">
                          <span>${f.icon}</span>
                          <span>${f.label}</span>
                        </span>
                        <span class="doc-platform-badge">
                          <span>${y.icon}</span>
                          <span>${y.type}</span>
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
                        ${n?`
                        <button class="doc-icon-btn edit-doc-btn" title="Edit Dokumen" data-id="${o.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        `:""}
                        ${i?`
                        <button class="doc-icon-btn delete delete-doc-btn" title="Hapus Dokumen" data-id="${o.id}">
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
    `,B()}function B(){let T=e.querySelector("#doc-btn-add");T&&(T.onclick=()=>S());let u=e.querySelector("#doc-btn-excel");u&&(u.onclick=()=>M());let $=e.querySelector("#doc-btn-refresh");$&&($.onclick=()=>s());let p=e.querySelector("#doc-btn-reset-filter");p&&(p.onclick=()=>{b="",m="all",A()});let o=e.querySelector("#doc-search-input");o&&(o.oninput=f=>{b=f.target.value;let y=d();A();let x=e.querySelector("#doc-search-input");x&&(x.focus(),x.setSelectionRange(x.value.length,x.value.length))}),e.querySelectorAll(".doc-pill").forEach(f=>{f.onclick=()=>{m=f.getAttribute("data-cat")||"all",A()}}),e.querySelectorAll(".copy-link-btn").forEach(f=>{f.onclick=y=>{y.stopPropagation();let x=f.getAttribute("data-link"),k=f.getAttribute("data-title");C(x,k)}}),e.querySelectorAll(".edit-doc-btn").forEach(f=>{f.onclick=y=>{y.stopPropagation();let x=Number(f.getAttribute("data-id")),k=r.find(a=>a.id===x);k&&S(k)}}),e.querySelectorAll(".delete-doc-btn").forEach(f=>{f.onclick=y=>{y.stopPropagation();let x=Number(f.getAttribute("data-id")),k=r.find(a=>a.id===x);k&&w(k)}})}await s()}me();de();var Re={"Standar Umum":{icon:"\u{1F4CC}",bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE",rooms:["Lantai & Sudut Ruangan","Dinding, Partisi & Plint","Kaca, Cermin & Partisi Kaca","Sanitair & Area Basah Toilet","Plafon, Lampu & Kisi Ventilasi","Mebel, Meja & Kursi Kerja/Pasien","Peralatan Khusus & Medis Non-Steril","Wadah Sampah & Utilitas"]},"Area Depan & Pelayanan":{icon:"\u{1F6CB}\uFE0F",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",rooms:["Teras","Lobby 1","Receptionist","Snack Corner"]},"Area Klinis & Medis":{icon:"\u{1FA7A}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD",rooms:["Ruang Tindakan","Ronsen","Ruang Dokter","Ruang Steril"]},"Toilet & Sanitasi":{icon:"\u{1F6BF}",bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4",rooms:["Toilet 1","Janitor","Ruang Limbah"]},Sirkulasi:{icon:"\u{1F6B6}",bg:"#F1F5F9",color:"#334155",border:"#CBD5E1",rooms:["Area Tangga","Lift","Area Tangga Exit"]},"Area Anak & Fasilitas":{icon:"\u{1F9F8}",bg:"#FEF3C7",color:"#92400E",border:"#FDE68A",rooms:["Lobby 2","Playground","Musholla","Tempat Wudhu","Ruang Team"]},"Gudang & Teknis":{icon:"\u2699\uFE0F",bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA",rooms:["Gudang Farmasi & Bahan","RoofTop","Ruang Kompressor","Area Genset"]},Eksterior:{icon:"\u{1F333}",bg:"#DCFCE7",color:"#15803D",border:"#86EFAC",rooms:["Taman","Pos Security","Toilet Security","Halaman Parkir"]}};function Nt(e){for(let[t,n]of Object.entries(Re))if(n.rooms.includes(e))return{zoneName:t,...n};return{zoneName:"Lainnya",icon:"\u{1F4CD}",bg:"#F1F5F9",color:"#475569",border:"#E2E8F0"}}async function qa(e){let t=ae(),n=t&&(t.role==="superadmin"||t.role==="admin"),i=Q(t,"export");e.innerHTML=`
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
        ${i?`
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
        ${Object.entries(Re).map(([p,o])=>`
          <button class="zone-tab-btn" data-zone="${p}" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
            <span>${o.icon}</span> ${p}
            <span class="badge" style="background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-${p.replace(/[\s&]+/g,"-")}">-</span>
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
  `;let r="all",m="table",b=[];async function g(){try{let p=await R("/api/hygiene-standards?all=1");if(!p.ok)throw new Error("Gagal memuat standar hygiene");b=Array.isArray(p.data)?p.data:p.data?.data||[],d(),s(),S()}catch(p){console.error(p),q(p.message||"Terjadi kesalahan memuat data");let o=e.querySelector("#hygiene-view-container");o&&(o.innerHTML=`
          <div class="empty-state" style="text-align:center; padding:40px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:12px;">\u26A0\uFE0F</div>
            <h3 style="margin:0 0 8px 0; color:var(--text-1);">Gagal Memuat Data</h3>
            <p style="color:var(--text-3); margin-bottom:16px;">${p.message}</p>
            <button id="btn-retry-hygiene" class="btn btn-primary" style="padding:8px 16px; border-radius:8px;">Coba Lagi</button>
          </div>
        `,o.querySelector("#btn-retry-hygiene")?.addEventListener("click",g))}}function s(){let p=e.querySelector("#badge-count-all");p&&(p.textContent=b.length),Object.entries(Re).forEach(([y,x])=>{let k=y.replace(/[\s&]+/g,"-"),a=e.querySelector(`#badge-count-${k}`);if(a){let c=b.filter(h=>x.rooms.includes(h.room_name)).length;a.textContent=c}});let o=new Set(b.map(y=>y.room_name)),f=e.querySelector("#card-total-rooms");f&&(f.textContent=o.size)}function d(){let p=e.querySelector("#hygiene-room-filter");if(!p)return;let o=p.value,f=[];if(r==="all")f=Array.from(new Set(b.map(y=>y.room_name))).filter(Boolean).sort();else{let y=Re[r];y&&(f=y.rooms.filter(x=>b.some(k=>k.room_name===x)))}p.innerHTML=`<option value="">Semua Ruangan (${f.length})</option>`+f.map(y=>`<option value="${y}" ${y===o?"selected":""}>${y}</option>`).join("")}function F(){let p=(e.querySelector("#hygiene-search")?.value||"").toLowerCase().trim(),o=e.querySelector("#hygiene-room-filter")?.value||"";return b.filter(f=>{if(r!=="all"){let y=Re[r];if(!y||!y.rooms.includes(f.room_name))return!1}if(o&&f.room_name!==o)return!1;if(p){let y=(f.room_name||"").toLowerCase(),x=(f.item_name||"").toLowerCase(),k=(f.cleanliness_standard||"").toLowerCase();if(!y.includes(p)&&!x.includes(p)&&!k.includes(p))return!1}return!0})}function S(){let p=F(),o=e.querySelector("#hygiene-stats-text");o&&(o.innerHTML=`Menampilkan <strong>${p.length}</strong> dari <strong>${b.length}</strong> parameter`);let f=e.querySelector("#hygiene-view-container");if(f){if(p.length===0){f.innerHTML=`
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
      `,f.querySelector("#btn-clear-search-empty")?.addEventListener("click",()=>{let y=e.querySelector("#hygiene-search"),x=e.querySelector("#hygiene-room-filter");y&&(y.value=""),x&&(x.value=""),r="all",e.querySelectorAll(".zone-tab-btn").forEach(k=>k.classList.toggle("active",k.dataset.zone==="all")),e.querySelectorAll(".zone-tab-btn").forEach(k=>{k.dataset.zone==="all"?(k.style.background="#2563EB",k.style.color="#fff"):(k.style.background="#F1F5F9",k.style.color="var(--text-2)")}),d(),S()});return}m==="table"?w(f,p):C(f,p)}}function w(p,o){p.innerHTML=`
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
              ${o.map((f,y)=>{let x=Nt(f.room_name);return`
                  <tr style="border-bottom:1px solid #F1F5F9; transition:background-color 0.15s ease;" onmouseover="this.style.backgroundColor='#F8FAFC'" onmouseout="this.style.backgroundColor='transparent'">
                    <td style="padding:14px 16px; text-align:center; color:var(--text-3); font-weight:600; font-size:0.85rem;">
                      ${y+1}
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="margin-bottom:4px;">
                        <span style="display:inline-flex; align-items:center; gap:6px; background:${x.bg}; color:${x.color}; border:1px solid ${x.border}; padding:4px 10px; border-radius:6px; font-weight:700; font-size:0.83rem;">
                          <span>${x.icon}</span> ${f.room_name}
                        </span>
                      </div>
                      <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">
                        ${x.zoneName}
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
                    ${n?`
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
    `,M(p)}function C(p,o){let f={};o.forEach(y=>{let x=Nt(y.room_name);f[x.zoneName]||(f[x.zoneName]={info:x,rooms:{}}),f[x.zoneName].rooms[y.room_name]||(f[x.zoneName].rooms[y.room_name]=[]),f[x.zoneName].rooms[y.room_name].push(y)}),p.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:24px;">
        ${Object.entries(f).map(([y,x])=>{let k=Object.values(x.rooms).reduce((a,c)=>a+c.length,0);return`
            <div class="card" style="background:#fff; border:1px solid ${x.info.border}; border-radius:14px; overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.03);">
              <!-- Zone Card Header -->
              <div style="background:${x.info.bg}; border-bottom:1px solid ${x.info.border}; padding:14px 20px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="font-size:1.4rem;">${x.info.icon}</span>
                  <div>
                    <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:${x.info.color};">
                      ${y}
                    </h3>
                    <div style="font-size:0.78rem; color:${x.info.color}; opacity:0.85;">
                      ${Object.keys(x.rooms).length} Ruangan \xB7 ${k} Parameter Standar
                    </div>
                  </div>
                </div>
                <span class="badge" style="background:#fff; color:${x.info.color}; border:1px solid ${x.info.border}; font-weight:700; padding:4px 10px; border-radius:12px; font-size:0.8rem;">
                  ${k} Parameter
                </span>
              </div>

              <!-- Zone Rooms List -->
              <div style="padding:16px 20px;">
                ${Object.entries(x.rooms).map(([a,c])=>`
                  <div style="margin-bottom:20px; last-child:margin-bottom:0;">
                    <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                      <h4 style="margin:0; font-size:0.95rem; font-weight:700; color:var(--text-1);">
                        \u{1F6AA} ${a}
                      </h4>
                      <span class="badge badge-sm" style="background:#F1F5F9; color:var(--text-3); font-size:0.75rem;">
                        ${c.length} item
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
                          ${c.map((h,v)=>`
                            <tr style="border-bottom:1px solid #F1F5F9;">
                              <td style="padding:10px 12px; text-align:center; color:var(--text-3); font-size:0.8rem;">
                                ${v+1}
                              </td>
                              <td style="padding:10px 12px; font-weight:600; color:var(--text-1); vertical-align:top;">
                                ${h.item_name}
                              </td>
                              <td style="padding:10px 12px; vertical-align:top; color:#334155;">
                                <div style="display:flex; align-items:flex-start; gap:6px;">
                                  <span style="color:#10B981; font-weight:700;">\u2713</span>
                                  <span>${h.cleanliness_standard}</span>
                                </div>
                              </td>
                              ${n?`
                                <td style="padding:10px 12px; text-align:center; vertical-align:middle;">
                                  <div style="display:inline-flex; gap:4px;">
                                    <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${h.id}" title="Edit" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                                    <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${h.id}" data-name="${h.item_name}" title="Hapus" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
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
    `,M(p)}function M(p){p.querySelectorAll(".btn-edit-item").forEach(o=>{o.addEventListener("click",()=>{let f=o.dataset.id,y=b.find(x=>String(x.id)===String(f));y&&A(y)})}),p.querySelectorAll(".btn-delete-item").forEach(o=>{o.addEventListener("click",()=>{let f=o.dataset.id,y=o.dataset.name;ie(`Yakin ingin menghapus parameter standar <strong>"${y}"</strong>?`,async()=>{try{let x=await R(`/api/hygiene-standards/${f}`,{method:"DELETE"});if(!x.ok)throw new Error(x.data?.error||"Gagal menghapus parameter");G("Parameter kebersihan berhasil dihapus"),g()}catch(x){q(x.message)}},"Hapus Parameter Standar")})})}function A(p=null){if(!n){q("Anda tidak memiliki izin untuk menambah atau mengedit parameter kebersihan.");return}let o=!!p,f=Array.from(new Set(b.map(x=>x.room_name))).filter(Boolean).sort(),y=document.createElement("div");y.innerHTML=`
      <form id="form-hygiene-item" style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Ruangan / Lokasi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-room-name" list="room-suggestions" class="form-control" placeholder="Contoh: Ruang Tindakan, Lobby 1, Toilet 1..." value="${p?.room_name||""}" required style="width:100%; height:40px; border-radius:8px;">
          <datalist id="room-suggestions">
            ${f.map(x=>`<option value="${x}">`).join("")}
          </datalist>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Pilih dari daftar atau ketik nama ruangan baru</span>
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Item / Objek Inspeksi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-item-name" class="form-control" placeholder="Contoh: Eksterior Dental Unit & Kursi, Lantai & Sudut..." value="${p?.item_name||""}" required style="width:100%; height:40px; border-radius:8px;">
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Standar Kebersihan Fisik <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="modal-standard" class="form-control" rows="4" placeholder="Deskripsi kondisi fisik kebersihan yang dipersyaratkan..." required style="width:100%; border-radius:8px; line-height:1.5;">${p?.cleanliness_standard||""}</textarea>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Jelaskan kriteria fisik: bebas debu kasat mata, steril, tidak licin, tidak berbau, dll.</span>
        </div>
      </form>
    `,te({title:o?"\u270F\uFE0F Edit Parameter Kebersihan":"\u2795 Tambah Parameter Kebersihan",content:y,confirmText:o?"Simpan Perubahan":"Tambah Parameter",size:"md",onConfirm:async(x,k)=>{let a=y.querySelector("#modal-room-name").value.trim(),c=y.querySelector("#modal-item-name").value.trim(),h=y.querySelector("#modal-standard").value.trim();if(!a||!c||!h){q("Semua kolom wajib diisi");return}try{let v={room_name:a,item_name:c,cleanliness_standard:h},I=o?`/api/hygiene-standards/${p.id}`:"/api/hygiene-standards",L=await R(I,{method:o?"PUT":"POST",body:JSON.stringify(v)});if(!L.ok)throw new Error(L.data?.error||"Gagal menyimpan parameter");G(o?"Parameter berhasil diperbarui":"Parameter berhasil ditambahkan"),k(),g()}catch(v){q(v.message)}}})}function B(){let p=window.open("","_blank");if(!p){q("Pop-up printer diblokir oleh browser. Izinkan pop-up untuk mencetak SOP.");return}let f=`
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
            <div>Total: <strong>${b.length} Parameter Fisik</strong></div>
          </div>
        </div>

        <div class="title">
          <h1>STANDAR ACUAN KEBERSIHAN FISIK SELURUH CABANG</h1>
          <p>Buku Panduan Kriteria Kebersihan Fisik 8 Zona Area (Cleaning Service & Facility Care)</p>
        </div>

        ${Object.entries(Re).map(([y,x])=>{let k=b.filter(a=>x.rooms.includes(a.room_name));return k.length===0?"":`
            <div class="zone-title">${x.icon} ${y.toUpperCase()} (${k.length} PARAMETER)</div>
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
                ${k.map((a,c)=>`
                  <tr>
                    <td class="td-no">${c+1}</td>
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
    `;p.document.open(),p.document.write(f),p.document.close()}function T(){let p=b.map((f,y)=>{let x=Nt(f.room_name);return{No:y+1,"Kategori Zona":x.zoneName,"Ruangan / Lokasi":f.room_name,"Item / Objek Inspeksi":f.item_name,"Standar Kebersihan Fisik":f.cleanliness_standard}}),o=new Date().toISOString().slice(0,10);ne(p,`Master_Standar_Kebersihan_FCMS_${o}`),G("Data master kebersihan berhasil diekspor ke Excel")}e.querySelector("#hygiene-search")?.addEventListener("input",()=>S()),e.querySelector("#hygiene-room-filter")?.addEventListener("change",()=>S()),e.querySelector("#btn-reset-filters")?.addEventListener("click",()=>{let p=e.querySelector("#hygiene-search"),o=e.querySelector("#hygiene-room-filter");p&&(p.value=""),o&&(o.value=""),S()}),e.querySelectorAll(".zone-tab-btn").forEach(p=>{p.addEventListener("click",()=>{r=p.dataset.zone,e.querySelectorAll(".zone-tab-btn").forEach(o=>{let f=o.dataset.zone===r;if(o.classList.toggle("active",f),f){o.style.background="#2563EB",o.style.color="#fff";let y=o.querySelector(".badge");y&&(y.style.background="rgba(255,255,255,0.25)",y.style.color="#fff")}else{o.style.background="#F1F5F9",o.style.color="var(--text-2)";let y=o.querySelector(".badge");y&&(y.style.background="#CBD5E1",y.style.color="#334155")}}),d(),S()})});let u=e.querySelector("#view-table-btn"),$=e.querySelector("#view-zone-btn");u?.addEventListener("click",()=>{m="table",u.style.background="#fff",u.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",$.style.background="transparent",$.style.boxShadow="none",S()}),$?.addEventListener("click",()=>{m="zone",$.style.background="#fff",$.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",u.style.background="transparent",u.style.boxShadow="none",S()}),e.querySelector("#btn-print-hygiene")?.addEventListener("click",B),e.querySelector("#btn-export-hygiene")?.addEventListener("click",T),e.querySelector("#btn-add-hygiene")?.addEventListener("click",()=>A(null)),g()}me();de();async function Ra(e,t="supply"){return ei(e)}async function ei(e){let n=((await R("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}));be({container:e,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:i=>i?new Date(i).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(i,r)=>r.branch_name_ref||r.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:i=>{try{let r=JSON.parse(i);return Array.isArray(r)?r.join(", "):i}catch{return i||"-"}}},{key:"chemical_items",label:"Chemical",render:i=>{try{let r=JSON.parse(i);return Array.isArray(r)?r.join(", "):i}catch{return i||"-"}}},{key:"additional_notes",label:"Catatan",render:i=>i?.length>40?i.slice(0,40)+"\u2026":i||"-"},{key:"status",label:"Status",render:i=>fe(i)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:i=>{let r=i?.tools_items;try{r=Array.isArray(JSON.parse(r))?JSON.parse(r).join(", "):r}catch{}let m=i?.chemical_items;try{m=Array.isArray(JSON.parse(m))?JSON.parse(m).join(", "):m}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:i?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:i?.branch_id&&!n.find(b=>b.value==i.branch_id)?[...n,{value:i.branch_id,label:i.branch_name||i.branch_id}]:n,createApi:{path:"/api/branches",field:"full_name"},value:i?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:r},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:i?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:m},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:i?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:i?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:i?.status||""},{name:"processed_by",label:"Diproses Oleh",value:i?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async i=>{let r=new URLSearchParams(i||{}).toString(),m=await R(`/api/reports/supply?limit=10000&${r}`);if(m.ok){let b=m.data.data.map(g=>{let s=g.tools_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}let d=g.chemical_items;try{d=Array.isArray(JSON.parse(d))?JSON.parse(d).join(", "):d}catch{}return{Waktu:g.submitted_at||"",Pengirim:g.submitter_name||"",Cabang:g.branch_name_ref||g.branch_name||"","Alat/Barang":s||"",Chemical:d||"",Catatan:g.additional_notes||"",Status:g.status||"","Diproses Oleh":g.processed_by||""}});ne(b,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async i=>{let m=(await R("/api/branches?all=1")).data?.data||[],b=F=>{if(!F)return null;let S=String(F||"").toLowerCase(),w=m.find(C=>String(C.full_name||"").toLowerCase()===S||String(C.code||"").toLowerCase()===S||String(C.name||"").toLowerCase()===S);return w?w.id:null},g=F=>{if(F==null||F==="")return"";if(F instanceof Date&&!isNaN(F.getTime()))return F.toISOString().slice(0,10);let S=String(F).trim();if(S===""||S==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(S))return S.slice(0,10);if(/^\d{4,5}$/.test(S)){let C=Number(S);if(C>2e4&&C<99999){let M=new Date(Date.UTC(1899,11,30)+C*864e5);return isNaN(M.getTime())?"":M.toISOString().slice(0,10)}}let w=S.split(/[\/\-\.]/);if(w.length===3){let[C,M,A]=w.map(B=>B.trim());if(C.length===4&&M.length<=2&&A.length<=2)return`${C}-${M.padStart(2,"0")}-${A.padStart(2,"0")}`;if(A.length===4&&M.length<=2&&C.length<=2)return`${A}-${M.padStart(2,"0")}-${C.padStart(2,"0")}`}return S},s=i.map(F=>({submitted_at:g(F.Waktu||F.Tanggal),submitter_name:String(F.Pengirim||"").trim(),branch_id:b(String(F.Cabang||"").trim()),tools_items:String(F["Alat/Barang"]||F.Alat||"").trim(),chemical_items:String(F.Chemical||"").trim(),additional_notes:String(F.Catatan||F.Keterangan||"").trim(),status:String(F.Status||"").trim(),processed_by:String(F["Diproses Oleh"]||F.PIC||"").trim()})).filter(F=>F.submitted_at&&F.submitter_name&&F.branch_id),d=await R("/api/reports/supply/import",{method:"POST",body:JSON.stringify(s)});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(i,r)=>{let m=te({title:"Update Status Permintaan",content:`
              <div class="form-group">
                <label class="form-label">Status</label>
                <select class="form-control" id="supply-status">
                  <option value="Pending" ${i.status==="Pending"?"selected":""}>Pending</option>
                  <option value="Diproses" ${i.status==="Diproses"?"selected":""}>Diproses</option>
                  <option value="Selesai" ${i.status==="Selesai"?"selected":""}>Selesai</option>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label">Diproses Oleh</label>
                <input type="text" class="form-control" id="supply-processed-by" value="${i.processed_by||""}" placeholder="Nama">
              </div>
            `,onConfirm:async(b,g)=>{let s=b.querySelector("#supply-status").value,d=b.querySelector("#supply-processed-by").value;(await R(`/api/reports/supply/${i.id}`,{method:"PUT",body:JSON.stringify({status:s,processed_by:d})})).ok?(G("Status diperbarui."),g(),r()):q("Gagal update status.")}})}}]})}de();async function Oa(e){let t=ae();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}be({container:e,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:n=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",editor_khusus:"badge-warning",input_lembur:"badge-info",viewer:"badge-neutral"}[n]||"badge-neutral"}">${n}</span>`},{key:"is_active",label:"Status",render:n=>n?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:n=>n?new Date(n).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:n=>{let i=!!n;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:n?.full_name},{name:"username",label:"Username",required:!0,placeholder:"username",value:n?.username,autocomplete:"off"}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!0,placeholder:"email@contoh.com",value:n?.email,autocomplete:"off"},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"editor_khusus",label:"Editor (Hanya Lembur & Masalah)"},{value:"input_lembur",label:"Input Lembur (Hanya Isi Lembur)"},{value:"viewer",label:"Viewer"}],value:n?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:i?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!i,placeholder:"Min. 6 karakter",autocomplete:"new-password"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:i?n?.is_active:1}]}]},onBeforeSubmit:(n,i)=>(n.full_name&&(n.full_name=String(n.full_name).trim()),n.username&&(n.username=String(n.username).trim()),n.email&&(n.email=String(n.email).trim().toLowerCase()),i&&(!n.password||String(n.password).trim()==="")&&delete n.password,n),exportOptions:{moduleName:"users",onExport:async()=>{let n=await R(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let i=n.data.data.map(r=>({"Nama Lengkap":r.full_name||"",Username:r.username||"",Email:r.email||"",Role:r.role||"",Status:r.is_active?"Aktif":"Nonaktif"}));ne(i,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async n=>{let i=n.map(m=>({full_name:String(m["Nama Lengkap"]||"").trim(),username:String(m.Username||"").trim(),email:String(m.Email||"").trim().toLowerCase(),role:String(m.Role||"").trim()||"viewer",password:String(m.Password||"").trim()})).filter(m=>m.username&&m.password&&m.email&&m.full_name),r=await R("/api/users/import",{method:"POST",body:JSON.stringify(i)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}de();async function Ha(e){be({container:e,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:t=>t?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:t=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:t?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:t?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:t?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:t?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:t?.is_active!==void 0?t.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let t=await R(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(t.ok)ne(t.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async t=>{let n=t.map(r=>({code:String(r["Kode Cabang"]||"").trim(),name:String(r["Nama Pendek"]||"").trim(),full_name:String(r["Nama Lengkap"]||"").trim(),city:String(r.Kota||"").trim()})).filter(r=>r.code&&r.name),i=await R("/api/branches/import",{method:"POST",body:JSON.stringify(n)});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}async function Ua(e){let t=new Date,n=[];e.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{t.setMonth(t.getMonth()-1),r()}),document.getElementById("cal-next").addEventListener("click",()=>{t.setMonth(t.getMonth()+1),r()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(b=>b.addEventListener("change",r));async function i(){try{let b=`${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,"0")}`;n=(await R(`/api/dashboard/calendar?month=${b}`)).data?.data||[]}catch(b){console.warn("[Calendar] Failed to load events, rendering empty grid:",b),n=[]}}async function r(){let b=document.getElementById("calendar-grid");if(b){b.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await i();try{let g=t.getFullYear(),s=t.getMonth(),d=t.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),F=document.getElementById("cal-month-label");F&&(F.textContent=d);let S=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(o=>o.value)),w=n.filter(o=>S.has(o.type)),C={};w.forEach(o=>{let f=ti(o.event_date);f&&(C[f]||(C[f]=[]),C[f].push(o))});let M=new Date(g,s,1).getDay(),A=new Date(g,s+1,0).getDate(),B=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],T=new Date().toISOString().slice(0,10),u='<div class="calendar-grid">';B.forEach(o=>{u+=`<div class="cal-day-header">${o}</div>`});for(let o=0;o<M;o++)u+='<div class="cal-cell cal-cell-empty"></div>';for(let o=1;o<=A;o++){let f=`${g}-${String(s+1).padStart(2,"0")}-${String(o).padStart(2,"0")}`,y=C[f]||[],x=f===T;u+=`
          <div class="cal-cell ${x?"cal-today":""} ${y.length?"cal-has-events":""}"
               data-date="${f}" tabindex="0" role="button" aria-label="${f}">
            <div class="cal-day-num ${x?"today-num":""}">${o}</div>
            <div class="cal-events-preview">
              ${y.slice(0,3).map(k=>`
                <div class="cal-event-dot cal-color-${k.color||"gray"}" title="${Se(k.title||k.type)}">
                  <span class="cal-event-dot-label">${ai(k.title||k.branch_name||k.type,18)}</span>
                </div>
              `).join("")}
              ${y.length>3?`<div class="cal-more">+${y.length-3} lagi</div>`:""}
            </div>
          </div>`}let p=(M+A)%7;if(p!==0)for(let o=0;o<7-p;o++)u+='<div class="cal-cell cal-cell-empty"></div>';u+="</div>",b.innerHTML=u,b.querySelectorAll(".cal-cell[data-date]").forEach(o=>{o.addEventListener("click",()=>{let f=o.dataset.date,y=C[f]||[];if(!y.length)return;let x=document.getElementById("cal-event-list"),k=new Date(f+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=k,document.getElementById("cal-event-items").innerHTML=y.map(a=>`
            <div class="cal-event-item cal-color-border-${a.color||"gray"}">
              <div class="cal-event-type">
                ${ni(a.type)}
                ${a.date_type?` <span class="badge" style="font-size:0.75rem;padding:2px 6px;background:rgba(0,0,0,0.06);margin-left:4px;">${Se(a.date_type)}</span>`:""}
              </div>
              <div class="cal-event-title" style="font-weight:600;font-size:0.95rem;margin-top:2px;">${Se(a.title||"-")}</div>
              ${a.branch_name?`<div class="cal-event-branch" style="color:var(--text-2);font-size:0.85rem;margin-top:2px;">\u{1F4CD} ${Se(a.branch_name)}</div>`:""}
              ${a.pic||a.period||a.opening_date||a.target_date||a.completion_date||a.notes?`
                <div class="cal-event-meta" style="margin-top:6px;padding-top:6px;border-top:1px dashed var(--border);font-size:0.8rem;color:var(--text-2);display:flex;flex-direction:column;gap:3px;">
                  ${a.pic?`<div>\u{1F464} <b>PIC:</b> ${Se(a.pic)}</div>`:""}
                  ${a.period?`<div>\u{1F4C5} <b>Periode:</b> ${Se(a.period)}</div>`:""}
                  ${a.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${Se(qt(a.opening_date))}</div>`:""}
                  ${a.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${Se(qt(a.target_date))}</div>`:""}
                  ${a.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${Se(qt(a.completion_date))}</div>`:""}
                  ${a.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${Se(a.notes)}</div>`:""}
                </div>
              `:""}
              ${a.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${a.status==="Done"?"badge-success":"badge-warning"}">${Se(a.status)}</span></div>`:""}
              ${a.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${a.days_remaining} hari</div>`:""}
            </div>
          `).join(""),x.style.display="block"})})}catch(g){console.error("[Calendar] Render error:",g),b&&(b.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let m=()=>{i().then(()=>r())};dt.on("data:changed",m),r()}function ti(e){if(!e||e==="-"||String(e).trim()==="")return"";let t=String(e).trim();if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);let n=t.split(/[\/\-\.]/);if(n.length===3){let[i,r,m]=n.map(b=>b.trim());if(m.length===4)return`${m}-${r.padStart(2,"0")}-${i.padStart(2,"0")}`;if(i.length===4)return`${i}-${r.padStart(2,"0")}-${m.padStart(2,"0")}`}return t.slice(0,10)}function qt(e){if(!e||e==="-"||String(e).trim()==="")return"";let n=String(e).trim().split(/[\/\-\.]/);if(n.length===3){let[i,r,m]=n.map(b=>b.trim());if(i.length===4)return`${m.padStart(2,"0")}-${r.padStart(2,"0")}-${i}`;if(m.length===4)return`${i.padStart(2,"0")}-${r.padStart(2,"0")}-${m}`}return e}function ai(e,t){return e?e.length>t?e.slice(0,t)+"\u2026":e:""}function Se(e){return e?String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function ni(e){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging"}[e]||e}async function Ga(e){let t=ae(),n=(t?.full_name||t?.username||"U")[0].toUpperCase(),r={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[t?.role]||"#64748B";e.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${r},${r}99)">
            ${n}
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
  `;let m=localStorage.getItem("fm_token"),b=document.getElementById("session-info");if(m&&b)try{let g=JSON.parse(atob(m.split(".")[1])),s=new Date(g.exp*1e3);b.textContent=`Berakhir: ${s.toLocaleString("id-ID")}`}catch{b.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async g=>{g.preventDefault();let s=document.getElementById("pwd-error"),d=document.getElementById("pwd-success"),F=document.getElementById("btn-save-pwd");s.style.display="none",d.style.display="none";let S=g.target,w=S.current_password.value,C=S.new_password.value,M=S.confirm_password.value;if(C!==M){s.textContent="\u274C Konfirmasi password tidak cocok.",s.style.display="block";return}if(C.length<6){s.textContent="\u274C Password baru minimal 6 karakter.",s.style.display="block";return}F.disabled=!0,F.textContent="\u23F3 Menyimpan...";let A=await R("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:w,new_password:C})});F.disabled=!1,F.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',A.ok?(d.textContent="\u2705 Password berhasil diubah.",d.style.display="block",S.reset(),G("Password berhasil diubah.")):(s.textContent=A.data?.error||"Gagal mengubah password.",s.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}var bt={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function ue(e){if(e==null||e==="")return null;if(e instanceof Date)return isNaN(e.getTime())?null:e.toISOString().slice(0,10);let t=String(e).trim();if(t===""||t==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);if(/^\d{4,5}$/.test(t)){let r=Number(t);if(r>2e4&&r<99999){let m=new Date(Date.UTC(1899,11,30)+r*864e5);return isNaN(m.getTime())?null:m.toISOString().slice(0,10)}}let n=t.split(/[\/\-\.]/);if(n.length===3){let[r,m,b]=n.map(F=>F.trim()),g=Number(r),s=Number(m),d=Number(b);if(r.length===4&&g>1900)return`${r}-${m.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&d>1900)return g>12?`${b}-${m.padStart(2,"0")}-${r.padStart(2,"0")}`:s>12?`${b}-${r.padStart(2,"0")}-${m.padStart(2,"0")}`:`${b}-${m.padStart(2,"0")}-${r.padStart(2,"0")}`;if(b.length===2&&!isNaN(d)){let F=d>=50?`19${b}`:`20${b}`;return g>12?`${F}-${m.padStart(2,"0")}-${r.padStart(2,"0")}`:`${F}-${m.padStart(2,"0")}-${r.padStart(2,"0")}`}}let i=new Date(t);return isNaN(i.getTime())?null:i.toISOString().slice(0,10)}function Ja(e){return Object.values(e).every(t=>t==null||String(t).trim()==="")}var ii={validation:{required:[],map:e=>({cabang:e.CABANG,pic:e.PIC,kegiatan:e.KEGIATAN,quartal:e.QUARTAL,masa_pkwt:e["MASA PKWT"],pic_pelapor:e["PIC PELAPOR"],kontrak:e.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:e=>({name:e["Nama SOP"],category:e.Kategori||"Umum",document_link:e["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({full_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||e.Divisi||"FACILITY CARE",phone:e["No. Hp"]||e["No. HP"],join_date:ue(e["Tanggal Masuk"]||e["Tgl Masuk"]),status:e.Status||"",target_pindah_os:ue(e["Target Pindah OS"]),target_selesai_os:ue(e["Target Selesai OS"]),notes:e.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({employee_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||"FACILITY CARE",start_date:ue(e["Tanggal Mulai"]),end_date:ue(e["Tanggal Selesai"]),contract_type:e["Tipe Kontrak"]||"",pkwt_number:e.PKWT||"",status:e.Status||"",notes:e.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:e=>({report_date:ue(e["Tanggal Info"]),branch_name:e.Cabang,category:e.Kategori,source:e["Sumber Laporan"],complaint:e.Keluhan,employee_name:e["Nama FC"],fc_specialist:e["FC Spesialis"],solution:e.Solusi,status:e.Status||"",completion_date:ue(e["Tanggal Selesai"])})},one_on_one:{required:[],map:e=>({meeting_date:ue(e.Tanggal),branch_name:e.Cabang,employee_name:e["Nama Karyawan"],pic:e.Pic,problem:e.Masalah,solution:e.Solusi,status:e.Status||"",completion_date:ue(e["Tanggal Selesai"]),document_link:e["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:e=>({branch_name:e.Cabang,activity_type:e.Kegiatan,period:e.Periode,pic:e.Pic||e.PIC,opening_date:ue(e["Tanggal Opening"]||e["Tgl Opening"]),target_date:ue(e["Tanggal Target"]||e["Tgl Target"]),completion_date:ue(e["Tanggal Selesai"]||e["Tgl Selesai"]),status:e.Status||"",notes:e.Keterangan||e.Catatan})},inspection:{required:[],map:e=>({inspection_date:ue(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",fc_score:e["Point FC SP"]!==void 0&&e["Point FC SP"]!==null?parseFloat(String(e["Point FC SP"]).replace(",",".")):null,spv_score:e["Point SPV"]!==void 0&&e["Point SPV"]!==null?parseFloat(String(e["Point SPV"]).replace(",",".")):null,document_link:e.Link,notes:""})},cleaning:{required:[],map:e=>({activity_date:ue(e.Tanggal),branch_name:e.Cabang,activity_type:e["Jenis Kegiatan"]||"General Cleaning",period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},fogging:{required:[],map:e=>({activity_date:ue(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:e=>({info_date:ue(e["Tgl Info"]),branch_name:e.Cabang,problem:e.Permasalahan,pic:e.PIC,done_date:ue(e["Tgl Done"]),status:e.Status||"",notes:e.Ket})},relievers:{required:[],map:e=>({branch_name:e.Cabang,original_fc_name:e["Nama Facility care"],period:e.Periode,reliever_name:e.Relifer,backup_date:ue(e["Tanggal Back Up"]),completion_date:ue(e["Tanggal Selesai"]),reason:e.Keterangan,shift:e.Shift,status:e.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:e=>({training_date:ue(e.Tanggal),batch:e.Batch,subject:e.Materi,participants:e.Peserta,branch_name:e.Cabang,trainer:e.Trainer,score:e.Nilai!==void 0&&e.Nilai!==null?parseFloat(String(e.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:e=>({name:e["Master Checklist"],category:"Umum",document_link:e["Link Document"],description:""})},supply:{required:[],map:e=>({submitted_at:ue(e.Timestamp),submitter_name:e["Nama Lengkap"],branch_name:e["Kebutuhan Untuk Cabang"],tools_items:e["Alat - Alat / Barang"],tools_quantity:e["Jumlah Permintaan Alat / Barang"],chemical_items:e.Chemical,chemical_quantity:e["Jumlah Permintaan Chemical"],additional_notes:e["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:e.Status||""})}};function ri(e,t){let n=bt[e];if(!n)return{valid:[],errors:[],mapped:[],skipped:!0};let i=ii[n.module];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let r=[],m=[],b=[];return t.filter(s=>!Ja(s)).forEach((s,d)=>{let F=t.indexOf(s)+2,S=[];i.required.forEach(({key:C,label:M})=>{let A=s[C];if(A==null||String(A).trim()===""){let B=Object.keys(s).filter(T=>T.trim()).join(", ");S.push({column:M,originalValue:A||"",reason:`Kolom "${M}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${B.slice(0,120)}`})}});let w=i.map(s);S.length>0?m.push({row:F,data:w,raw:s,errors:S}):(r.push(s),b.push(w))}),{valid:r,errors:m,mapped:b}}function Ya(e){let t=[];return e.SheetNames.forEach(n=>{let i=bt[n];if(!i)return;let r=e.Sheets[n],m=window.XLSX.utils.sheet_to_json(r,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),b=ri(n,m),g=m.filter(s=>!Ja(s));t.push({sheetName:n,module:i.module,label:i.label,total:g.length,valid:b.mapped.length,errorCount:b.errors.length,errors:b.errors,mapped:b.mapped,skipped:!1})}),t}function Va(){let e=window.XLSX,t=e.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([i,r])=>{e.utils.book_append_sheet(t,e.utils.json_to_sheet(r),i)}),e.writeFile(t,"Template_Import_Data_Awal_FCMS.xlsx")}function Wa(e){let t=window.XLSX,n=t.utils.book_new(),i=!1;return e.forEach(r=>{if(!r.errors||r.errors.length===0)return;i=!0;let m=r.errors.map(g=>({"No. Baris":g.row,"Kolom Gagal":(g.errors||[]).map(s=>s.column||s).join("; "),"Alasan Error":(g.errors||[]).map(s=>s.reason||s).join("; "),...Object.fromEntries(Object.entries(g.data||{}).map(([s,d])=>[s,d??""]))})),b=t.utils.json_to_sheet(m);t.utils.book_append_sheet(n,b,r.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),i?(t.writeFile(n,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var oi=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","sop","inspection","cleaning","fogging","basecamp","supply"];function Qa(e){e.innerHTML=`
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
              ${Object.entries(bt).map(([A,{label:B}])=>`<span class="import-sheet-tag">\u{1F4C4} ${A} \u2192 ${B}</span>`).join("")}
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
  `;let t=null,n=null,i=0,r={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function m(A){Object.entries(r).forEach(([B,T])=>{T.style.display=B===A?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let A=document.getElementById("btn-backup-db");A.disabled=!0,A.textContent="\u23F3 Memproses Backup...";try{let B=await R("/api/import/backup");if(B.ok){if(!window.XLSX){q("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let T=window.XLSX,u=T.utils.book_new();Object.entries(B.data.database).forEach(([$,p])=>{let o=p.length>0?p:[{}],f=T.utils.json_to_sheet(o);T.utils.book_append_sheet(u,f,$.substring(0,31))}),T.writeFile(u,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),G("Backup berhasil diunduh!")}else q("Gagal memproses backup: "+(B.data?.error||"Unknown error"))}catch(B){q("Gagal memproses backup: "+B.message)}finally{A.disabled=!1,A.textContent="\u{1F4E6} Backup Database"}});let b=document.getElementById("btn-sync-google");b&&b.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let A=b.innerHTML;b.innerHTML='<span class="spinner"></span> Menyinkronkan...',b.disabled=!0;try{let B=await R("/api/sync/google-sheets",{method:"POST"});B.ok?alert("Sinkronisasi Berhasil: "+(B.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(B.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{b.innerHTML=A,b.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Va(),G("Template Excel berhasil didownload!")});let g=document.getElementById("file-input"),s=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",A=>{A.stopPropagation(),g.click()}),g.addEventListener("change",A=>{A.target.files[0]&&d(A.target.files[0])}),s.addEventListener("dragover",A=>{A.preventDefault(),s.classList.add("drag-over")}),s.addEventListener("dragleave",()=>s.classList.remove("drag-over")),s.addEventListener("drop",A=>{A.preventDefault(),s.classList.remove("drag-over");let B=A.dataTransfer.files[0];B&&B.name.match(/\.xlsx?$/i)?d(B):q("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{t=null,g.value="",document.getElementById("file-info").style.display="none",s.style.display="",m("upload")});async function d(A){t=A,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${A.name} (${(A.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",s.style.display="none",await F(A)}async function F(A){m("validating");let B=document.getElementById("validation-status"),T=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");B.textContent="Membaca file Excel...",T.style.width="20%",await Xe(200);let u=await A.arrayBuffer(),$=window.XLSX.read(u,{type:"array",cellDates:!0});B.textContent=`Memvalidasi ${$.SheetNames.length} sheet...`,T.style.width="50%",await Xe(100),n=Ya($),T.style.width="100%",B.textContent="Validasi selesai!",await Xe(300),S()}catch(u){m("upload"),q("Gagal memproses file: "+u.message),document.getElementById("file-info").style.display="flex",s.style.display="none"}}function S(){m("preview");let A=n.filter(x=>!x.skipped).length,B=n.reduce((x,k)=>x+k.total,0),T=n.reduce((x,k)=>x+k.valid,0),u=n.reduce((x,k)=>x+k.errorCount,0),$=B>0?Math.round(T/B*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${A} sheet</span>
      <span class="badge badge-secondary">${B} baris</span>
      <span class="badge badge-success">${T} valid (${$}%)</span>
      ${u>0?`<span class="badge badge-danger">${u} error</span>`:""}
    `;let p=document.getElementById("preview-table-container");p.innerHTML=`
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
          ${n.map((x,k)=>`
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
                ${x.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${k}">\u{1F50D} ${x.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,p.querySelectorAll(".btn-detail-error").forEach(x=>{x.addEventListener("click",()=>{let k=n[Number(x.dataset.idx)];w(k)})});let o=document.getElementById("error-detail-section"),f=document.getElementById("error-detail-container");f.innerHTML="",o.style.display="none";let y=document.getElementById("btn-start-import");T===0?(y.disabled=!0,y.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(y.disabled=!1,u>0?(y.innerHTML=`\u{1F680} Import ${T} Data Valid (${u} dilewati)`,y.title="Baris error akan dilewati, baris valid tetap diimport"):y.innerHTML=`\u{1F680} Mulai Import ${T} Data`)}function w(A){let B=document.getElementById("error-detail-section"),T=document.getElementById("error-detail-container");B.style.display="";let u=A.errors.slice(0,100).map($=>(Array.isArray($.errors)?$.errors:[]).map(o=>{let f=typeof o=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${$.row}</span></td>
            <td><strong>${f?o.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${f&&o.originalValue!==void 0?o.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${f?o.reason:o}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${f&&o.aliases?`Gunakan salah satu nama kolom:<br><em>${o.aliases}</em>`:f&&o.hint?o.hint:""}
            </td>
          </tr>
        `}).join("")).join("");T.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${A.sheetName} \u2014 ${A.errorCount} baris error dari ${A.total} total
          ${A.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
            <tbody>${u||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${A.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,B.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{m("upload"),document.getElementById("file-info").style.display="none",s.style.display="",t=null,g.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!n)return;Wa(n)?G("Log error berhasil didownload."):G("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let A=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";C(A)});async function C(A){m("importing"),i=Date.now();let B=[];oi.forEach(o=>{let f=n?.find(y=>y.module===o&&y.mapped?.length>0);f&&B.push(f)});let T=document.getElementById("import-steps-list");T.innerHTML=B.map(o=>`
      <div class="import-step-item" id="step-item-${o.module}">
        <span class="step-item-icon" id="step-icon-${o.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${o.label} <span class="step-item-count">(${o.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${o.module}"></span>
      </div>
    `).join("");let u=document.getElementById("import-bar"),$=document.getElementById("import-current-status"),p={totalSheets:B.length,totalRows:B.reduce((o,f)=>o+f.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let o=0;o<B.length;o++){let f=B[o],y=document.getElementById(`step-icon-${f.module}`),x=document.getElementById(`step-status-${f.module}`);y.textContent="\u{1F504}",x.textContent="Mengimport...",$.textContent=`Mengimport ${f.label}...`,u.style.width=`${Math.round(o/B.length*100)}%`;try{let k=await R(`/api/import/${f.module}`,{method:"POST",body:JSON.stringify({rows:f.mapped,onDuplicate:A})});if(k.ok){let a=k.data;p.inserted+=a.inserted||0,p.skipped+=a.skipped||0,p.moduleResults.push({label:f.label,inserted:a.inserted||0,skipped:a.skipped||0,status:"ok"}),y.textContent="\u2705",x.innerHTML=`<span class="badge badge-success">${a.inserted||0} berhasil</span>${a.skipped>0?` <span class="badge badge-neutral">${a.skipped} skip</span>`:""}`}else p.failed++,p.moduleResults.push({label:f.label,inserted:0,skipped:0,status:"error",error:k.data?.error}),y.textContent="\u274C",x.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(k){p.failed++,p.moduleResults.push({label:f.label,inserted:0,skipped:0,status:"error",error:k.message}),y.textContent="\u274C",x.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Xe(150)}u.style.width="100%",$.textContent="Selesai!",await Xe(400),M(p)}function M(A){m("summary");let B=((Date.now()-i)/1e3).toFixed(1),T=A.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${T?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${T?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${A.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${A.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${A.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${A.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${A.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${A.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
      <div class="summary-stat-card info">
        <div class="stat-value">${B}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${A.moduleResults.map(u=>`
            <tr>
              <td>${u.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${u.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${u.skipped}</span></td>
              <td style="text-align:center">
                ${u.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${u.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{t=null,n=null,g.value="",document.getElementById("file-info").style.display="none",s.style.display="",m("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Xe(e){return new Promise(t=>setTimeout(t,e))}me();var Xa=[{id:1,employee_name:"Teguh Ariyandi",division:"FACILITY CARE",branch_id:18,branch_name:"018. Harapan Indah",tanggal:"2025-08-19",akhir_sp:"2026-02-18",sp_type:"SP 1",status:"Aktif",document_link:"",notes:"Keterlambatan berulang dan ketidaklengkapan checklist harian lantai dasar."},{id:2,employee_name:"Andri",division:"FACILITY CARE",branch_id:1,branch_name:"001. Pondok Bambu",tanggal:"2025-04-10",akhir_sp:"2025-10-09",sp_type:"SP 1",status:"Selesai",document_link:"",notes:"Pelanggaran grooming dan standar pembersihan area tindakan tanpa APD lengkap."},{id:3,employee_name:"M. FIRMANSYAH",division:"FACILITY CARE",branch_id:25,branch_name:"025. Bogor Baru",tanggal:"2024-03-21",akhir_sp:"2024-09-20",sp_type:"SP 2",status:"Selesai",document_link:"",notes:"Meninggalkan area tugas tanpa konfirmasi supervisor selama jam operasional klinik."}];function Pe(e){if(!e)return{status:"Tidak Ditentukan",label:"Tidak Ada Tanggal",bg:"#F1F5F9",color:"#64748B",border:"#CBD5E1",icon:"\u26AA",daysLeft:0,isExpired:!1,isExpiringSoon:!1};let t=new Date;t.setHours(0,0,0,0);let n=new Date(e);n.setHours(0,0,0,0);let i=n.getTime()-t.getTime(),r=Math.ceil(i/(1e3*60*60*24));return r<0?{status:"Selesai",label:`Masa Sanksi Selesai (${Math.abs(r)} hari lalu)`,shortLabel:"Selesai",bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2705",daysLeft:r,isExpired:!0,isExpiringSoon:!1}:r<=30?{status:"Mendekati Berakhir",label:`Segera Berakhir (Sisa ${r} hari)`,shortLabel:`Sisa ${r} hari`,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F",daysLeft:r,isExpired:!1,isExpiringSoon:!0}:{status:"Aktif",label:`Aktif Berlaku (Sisa ${r} hari)`,shortLabel:`Aktif (${r} hari)`,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F534}",daysLeft:r,isExpired:!1,isExpiringSoon:!1}}function Rt(e=""){let t=(e||"").toUpperCase().trim();return t==="SP 3"?{label:"SP 3 (Terakhir)",bg:"#7F1D1D",color:"#FFFFFF",border:"#991B1B",icon:"\u{1F6A8}"}:t==="SP 2"?{label:"SP 2",bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F6D1}"}:t==="SP 1"?{label:"SP 1",bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F"}:t.includes("TEGURAN")?{label:"Teguran Lisan",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F4AC}"}:{label:e||"SP",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4DC}"}}function Oe(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Za(e){let t=ae(),n=Q(t,"sp","write"),i=Q(t,"sp","write"),r=Q(t,"sp","delete"),m=[],b=[],g=[],s="all",d="all",F="all",S="",w=localStorage.getItem("fcms_sp_view_mode")||"card",C="fcms-sp-page-styles",M=document.getElementById(C);M||(M=document.createElement("style"),M.id=C,document.head.appendChild(M)),M.textContent=`
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
    `;async function A(){try{let[x,k]=await Promise.all([se(),he()]);b=x||[],g=k||[];let a=await R("/api/sp?limit=1000");if(a.ok&&a.data?.data&&Array.isArray(a.data.data)&&a.data.data.length>0)m=a.data.data,localStorage.setItem("fcms_sp_data_cache",JSON.stringify(m));else{let c=localStorage.getItem("fcms_sp_data_cache");c?m=JSON.parse(c):(m=[...Xa],localStorage.setItem("fcms_sp_data_cache",JSON.stringify(m)))}}catch(x){console.warn("Fallback loading SP data from cache/seed:",x);let k=localStorage.getItem("fcms_sp_data_cache");m=k?JSON.parse(k):[...Xa]}f()}function B(){return m.filter(x=>{let k=S.toLowerCase().trim(),a=!k||(x.employee_name||"").toLowerCase().includes(k)||(x.branch_name||"").toLowerCase().includes(k)||(x.division||"").toLowerCase().includes(k)||(x.sp_type||"").toLowerCase().includes(k),c=d==="all"||String(x.branch_id)===String(d),h=F==="all"||x.division===F,v=Pe(x.akhir_sp),I=!0;return s==="active"?I=!v.isExpired:s==="expiring"?I=v.isExpiringSoon:s==="expired"?I=v.isExpired:s==="sp1"?I=(x.sp_type||"").toUpperCase()==="SP 1":s==="sp2"?I=(x.sp_type||"").toUpperCase()==="SP 2":s==="sp3"?I=(x.sp_type||"").toUpperCase()==="SP 3":s==="teguran"&&(I=(x.sp_type||"").toUpperCase().includes("TEGURAN")),a&&c&&h&&I})}function T(){let x=0,k=0,a=0,c=0,h=0,v=0,I=0,j=new Set;return m.forEach(L=>{L.employee_name&&j.add(L.employee_name.toLowerCase().trim());let _=Pe(L.akhir_sp);_.isExpired?a++:x++,_.isExpiringSoon&&k++;let P=(L.sp_type||"").toUpperCase();P==="SP 1"?c++:P==="SP 2"?h++:P==="SP 3"?v++:P.includes("TEGURAN")&&I++}),{total:m.length,active:x,expiring:k,expired:a,sp1:c,sp2:h,sp3:v,teguran:I,uniqueEmployees:j.size}}function u(x=null){let k=!!x;if(k&&!i){q("Anda tidak memiliki izin untuk mengedit data SP.");return}if(!k&&!n){q("Anda tidak memiliki izin untuk menambah data SP.");return}let a=!!x,c=x?.tanggal?x.tanggal.slice(0,10):new Date().toISOString().slice(0,10),h=x?.akhir_sp?x.akhir_sp.slice(0,10):"";if(!h){let L=new Date(c);L.setMonth(L.getMonth()+6),h=L.toISOString().slice(0,10)}let v=document.createElement("div");v.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <input type="text" id="sp-modal-emp" list="sp-emp-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih atau ketik nama..." value="${x?.employee_name?String(x.employee_name).replace(/"/g,"&quot;"):""}" required />
            <datalist id="sp-emp-list">
              ${g.map(L=>`<option value="${L.label}">${L.label}</option>`).join("")}
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
              ${b.map(L=>`<option value="${L.value}" ${String(x?.branch_id)===String(L.value)?"selected":""}>${L.label}</option>`).join("")}
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
              value="${c}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Akhir Masa Berlaku SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-akhir" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${h}" required />
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
    `;let I=v.querySelector("#sp-modal-tanggal"),j=v.querySelector("#sp-modal-akhir");I.addEventListener("change",()=>{if(I.value){let L=new Date(I.value);L.setMonth(L.getMonth()+6),j.value=L.toISOString().slice(0,10)}}),te({title:a?"\u270F\uFE0F Edit Data Surat Peringatan (SP)":"\u2795 Terbitkan Surat Peringatan (SP) Baru",content:v,size:"md",confirmText:a?"Simpan Perubahan":"Terbitkan SP",confirmClass:"btn-primary",onConfirm:async(L,_)=>{let P=(v.querySelector("#sp-modal-emp").value||"").trim(),K=v.querySelector("#sp-modal-div").value,l=v.querySelector("#sp-modal-branch").value,E=v.querySelector("#sp-modal-type").value,D=v.querySelector("#sp-modal-tanggal").value,z=v.querySelector("#sp-modal-akhir").value,J=(v.querySelector("#sp-modal-link").value||"").trim(),N=(v.querySelector("#sp-modal-notes").value||"").trim();if(!P){q("Nama Karyawan wajib diisi");return}if(!l){q("Cabang penempatan wajib dipilih");return}if(!D||!z){q("Tanggal SP dan Akhir SP wajib diisi");return}let O=b.find(X=>String(X.value)===String(l)),U=O?O.label:"Cabang",H={employee_name:P,division:K,branch_id:Number(l),branch_name:U,sp_type:E,tanggal:D,akhir_sp:z,document_link:J,notes:N,status:Pe(z).status};try{if(a){let X=await R(`/api/sp/${x.id}`,{method:"PUT",body:JSON.stringify(H)}),W=m.findIndex(V=>V.id===x.id);W!==-1&&(m[W]={...m[W],...H}),G("Data SP berhasil diperbarui")}else{let X=await R("/api/sp",{method:"POST",body:JSON.stringify(H)}),W=X.ok&&X.data?.id?X.data.id:Date.now();m.unshift({id:W,...H}),G("Surat Peringatan baru berhasil diterbitkan")}localStorage.setItem("fcms_sp_data_cache",JSON.stringify(m)),_(),f()}catch(X){console.error("Error saving SP:",X),q("Gagal menyimpan SP: "+X.message)}}})}function $(x){ie(`Hapus data Surat Peringatan untuk karyawan <strong>${x.employee_name}</strong> (${x.sp_type})?`,async()=>{try{await R(`/api/sp/${x.id}`,{method:"DELETE"})}catch(k){console.warn("API delete error fallback to local:",k)}m=m.filter(k=>k.id!==x.id),localStorage.setItem("fcms_sp_data_cache",JSON.stringify(m)),G(`Data SP ${x.employee_name} berhasil dihapus`),f()},"Hapus Data SP")}function p(x){let k=Pe(x.akhir_sp),a=Rt(x.sp_type),c=document.createElement("div");c.innerHTML=`
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
            <span class="sp-status-badge" style="background:${k.bg}; color:${k.color}; border:1px solid ${k.border}; font-size:0.85rem;">
              <span>${k.icon}</span>
              <span>${k.label}</span>
            </span>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Tanggal Terbit SP</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${Oe(x.tanggal)}</div>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Akhir Masa Berlaku</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${Oe(x.akhir_sp)}</div>
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
    `,te({title:"\u{1F4CB} Lembar Riwayat Surat Peringatan (SP)",content:c,size:"md",cancelText:"Tutup"})}async function o(){if(!Q(t,"export")){q("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:x}=await Promise.resolve().then(()=>(de(),xe)),k=m.map((a,c)=>{let h=Pe(a.akhir_sp);return{No:c+1,"Nama Karyawan":a.employee_name||"",Divisi:a.division||"",Cabang:a.branch_name||"","Jenis SP":a.sp_type||"","Tanggal SP":a.tanggal||"","Akhir SP":a.akhir_sp||"","Status Masa Berlaku":h.status,"Sisa Hari":h.daysLeft>0?h.daysLeft:0,"Link BAP / Foto":a.document_link||"","Catatan / Kronologi":a.notes||a.description||""}});x(k,`Data_Surat_Peringatan_FCMS_${new Date().toISOString().slice(0,10)}`),G("Berhasil mengunduh data SP ke format Excel.")}catch(x){q("Gagal ekspor: "+x.message)}}function f(){let x=B(),k=T();e.innerHTML=`
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
              
              ${Q(t,"export")?`<button id="sp-btn-excel" class="btn btn-ghost sp-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
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
                <div class="sp-stat-value">${k.active}</div>
                <div class="sp-stat-label">SP Aktif Berlaku</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">\u26A0\uFE0F</div>
              <div>
                <div class="sp-stat-value">${k.expiring}</div>
                <div class="sp-stat-label">Segera Berakhir (&lt;30 Hari)</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">\u2705</div>
              <div>
                <div class="sp-stat-value">${k.expired}</div>
                <div class="sp-stat-label">Masa Sanksi Selesai</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">\u{1F465}</div>
              <div>
                <div class="sp-stat-value">${k.uniqueEmployees} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">karyawan</span></div>
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
              <input type="text" id="sp-search-input" class="sp-search-input" placeholder="Cari nama karyawan, cabang, atau divisi..." value="${S?S.replace(/"/g,"&quot;"):""}" />
            </div>

            <!-- Branch Dropdown -->
            <select id="sp-filter-branch" class="sp-filter-select">
              <option value="all">Semua Cabang (${b.length})</option>
              ${b.map(a=>`<option value="${a.value}" ${String(d)===String(a.value)?"selected":""}>${a.label}</option>`).join("")}
            </select>

            <!-- Division Dropdown -->
            <select id="sp-filter-division" class="sp-filter-select">
              <option value="all" ${F==="all"?"selected":""}>Semua Divisi</option>
              <option value="FACILITY CARE" ${F==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${F==="SECURITY"?"selected":""}>SECURITY</option>
            </select>

            <div class="sp-view-toggle">
              <button type="button" class="sp-view-btn ${w==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="sp-view-btn ${w==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Filter Pills -->
          <div class="sp-filter-pills">
            <button class="sp-pill ${s==="all"?"active":""}" data-filter="all">
              <span>Semua SP</span>
              <span class="sp-pill-count">${k.total}</span>
            </button>
            <button class="sp-pill ${s==="active"?"active":""}" data-filter="active">
              <span>\u{1F534} SP Aktif</span>
              <span class="sp-pill-count">${k.active}</span>
            </button>
            <button class="sp-pill ${s==="expiring"?"active":""}" data-filter="expiring">
              <span>\u26A0\uFE0F &lt;30 Hari</span>
              <span class="sp-pill-count">${k.expiring}</span>
            </button>
            <button class="sp-pill ${s==="expired"?"active":""}" data-filter="expired">
              <span>\u2705 Selesai</span>
              <span class="sp-pill-count">${k.expired}</span>
            </button>
            <button class="sp-pill ${s==="sp1"?"active":""}" data-filter="sp1">
              <span>SP 1</span>
              <span class="sp-pill-count">${k.sp1}</span>
            </button>
            <button class="sp-pill ${s==="sp2"?"active":""}" data-filter="sp2">
              <span>SP 2</span>
              <span class="sp-pill-count">${k.sp2}</span>
            </button>
            <button class="sp-pill ${s==="sp3"?"active":""}" data-filter="sp3">
              <span>SP 3</span>
              <span class="sp-pill-count">${k.sp3}</span>
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
          `:w==="table"?`
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
                    ${x.map((a,c)=>{let h=Pe(a.akhir_sp),v=Rt(a.sp_type);return`
                        <tr data-id="${a.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${c+1}</td>
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
                            <span class="sp-type-badge" style="background:${v.bg}; color:${v.color}; border:1px solid ${v.border};">
                              <span>${v.icon}</span>
                              <span>${v.label}</span>
                            </span>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            ${Oe(a.tanggal)}
                          </td>
                          <td style="white-space:nowrap; font-weight:700; color:#0F172A;">
                            ${Oe(a.akhir_sp)}
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="sp-status-badge" style="background:${h.bg}; color:${h.color}; border:1px solid ${h.border};">
                              <span>${h.icon}</span>
                              <span>${h.shortLabel}</span>
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
                              ${i?`
                                <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${a.id}">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              `:""}
                              ${r?`
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
              ${x.map(a=>{let c=Pe(a.akhir_sp),h=Rt(a.sp_type);return`
                  <div class="sp-card" data-id="${a.id}">
                    <div>
                      <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px;">
                        <span class="sp-type-badge" style="background:${h.bg}; color:${h.color}; border:1px solid ${h.border};">
                          <span>${h.icon}</span>
                          <span>${h.label}</span>
                        </span>
                        <span class="sp-status-badge" style="background:${c.bg}; color:${c.color}; border:1px solid ${c.border};">
                          <span>${c.icon}</span>
                          <span>${c.shortLabel}</span>
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
                          <div style="font-weight:700; color:#1E293B;">${Oe(a.tanggal)}</div>
                        </div>
                        <div style="text-align:right;">
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Akhir SP</div>
                          <div style="font-weight:700; color:#0F172A;">${Oe(a.akhir_sp)}</div>
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
                        ${i?`
                          <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${a.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                        `:""}
                        ${r?`
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
    `,y()}function y(){let x=e.querySelector("#sp-btn-add");x&&(x.onclick=()=>u());let k=e.querySelector("#sp-btn-excel");k&&(k.onclick=()=>o());let a=e.querySelector("#sp-btn-refresh");a&&(a.onclick=()=>A());let c=e.querySelector("#sp-btn-reset-filter");c&&(c.onclick=()=>{S="",d="all",F="all",s="all",f()});let h=e.querySelector("#sp-search-input");h&&(h.oninput=j=>{S=j.target.value,f();let L=e.querySelector("#sp-search-input");L&&(L.focus(),L.setSelectionRange(L.value.length,L.value.length))});let v=e.querySelector("#sp-filter-branch");v&&(v.onchange=j=>{d=j.target.value,f()});let I=e.querySelector("#sp-filter-division");I&&(I.onchange=j=>{F=j.target.value,f()}),e.querySelectorAll(".sp-pill").forEach(j=>{j.onclick=()=>{s=j.getAttribute("data-filter")||"all",f()}}),e.querySelectorAll(".sp-view-btn").forEach(j=>{j.onclick=()=>{let L=j.getAttribute("data-view");L&&L!==w&&(w=L,localStorage.setItem("fcms_sp_view_mode",L),f())}}),e.querySelectorAll(".sp-detail-btn, .sp-view-detail-link").forEach(j=>{j.onclick=L=>{L.stopPropagation();let _=Number(j.getAttribute("data-id")),P=m.find(K=>K.id===_);P&&p(P)}}),e.querySelectorAll(".sp-edit-btn").forEach(j=>{j.onclick=L=>{L.stopPropagation();let _=Number(j.getAttribute("data-id")),P=m.find(K=>K.id===_);P&&u(P)}}),e.querySelectorAll(".sp-delete-btn").forEach(j=>{j.onclick=L=>{L.stopPropagation();let _=Number(j.getAttribute("data-id")),P=m.find(K=>K.id===_);P&&$(P)}})}await A()}me();var Ot=[{id:1,tanggal:"2026-10-02",employee_name:"TAUFIK HIDAYAT",from_branch_id:12,from_branch_name:"012. BSD",to_branch_id:0,to_branch_name:"000. Management",status:"Selesai",document_link:""},{id:2,tanggal:"2026-09-30",employee_name:"ELVAN JANGGA PRAMANA",from_branch_id:55,from_branch_name:"055. Gajah Mada",to_branch_id:0,to_branch_name:"000. Management",status:"Proses",document_link:""},{id:3,tanggal:"2026-06-23",employee_name:"Widya Astuti",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:7,to_branch_name:"007. Bekasi",status:"Selesai",document_link:""},{id:4,tanggal:"2026-06-22",employee_name:"Gema Azan Saputra",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:21,to_branch_name:"021. Kelapa Gading Boulevard",status:"Selesai",document_link:""},{id:5,tanggal:"2026-06-22",employee_name:"Ahmad Farhan",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:62,to_branch_name:"062. Summarecon Bekasi",status:"Selesai",document_link:""}];function en(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function tn(e){let t=[],n=[],i=[],r="all",m="all",b="all",g="",s=!1,d=localStorage.getItem("fcms_mutasi_view_mode")||"card",F=ae(),S=Q(F,"mutasi","write"),w=Q(F,"mutasi","write"),C=Q(F,"mutasi","delete");try{[n,i]=await Promise.all([se(),he()])}catch(o){console.error("Error fetching cache:",o)}async function M(){s=!0,$();try{let o=await R("/api/mutasi?limit=10000");if(o.ok){let f=o.data?.data||o.data||[];t=Array.isArray(f)&&f.length>0?f:Ot}else t=Ot}catch(o){console.warn("Gagal load mutasi, menggunakan data lokal:",o),t=Ot}finally{s=!1,$()}}async function A(){if(!Q(F,"export")){q("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:o}=await Promise.resolve().then(()=>(de(),xe)),f=t.map((y,x)=>({No:x+1,Tanggal:y.tanggal||"","Nama Karyawan":y.employee_name||"","Cabang Asal":y.from_branch_name||"","Cabang Tujuan":y.to_branch_name||"",Status:y.status||"","Link Dokumen":y.document_link||""}));o(f,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`),G("Data mutasi berhasil diekspor ke Excel!")}catch(o){q("Gagal ekspor Excel: "+o.message)}}async function B(o){if(!w)return;let f=t.find(y=>y.id===o);ie({title:"Hapus Data Mutasi",message:`Apakah Anda yakin ingin menghapus data mutasi untuk "${f?.employee_name||"karyawan ini"}"?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let y=await R(`/api/mutasi/${o}`,{method:"DELETE"});y.ok?(G("Data mutasi berhasil dihapus"),t=t.filter(x=>x.id!==o),$()):q(y.data?.error||"Gagal menghapus mutasi")}catch(y){q("Gagal menghapus data mutasi: "+y.message)}}})}function T(o=null){let f=!!o;if(f&&!w){q("Anda tidak memiliki izin untuk mengedit data mutasi.");return}if(!f&&!S){q("Anda tidak memiliki izin untuk menambah data mutasi.");return}if(!w)return;let y=!!o,x=o?.tanggal?o.tanggal.slice(0,10):new Date().toISOString().slice(0,10),k=document.createElement("div");k.innerHTML=`
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
            ${i.map(a=>`<option value="${typeof a=="string"?a:a.label||a.name||""}">`).join("")}
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
                <option value="${a.value}" ${String(o?.from_branch_id)===String(a.value)?"selected":""}>
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
                <option value="${a.value}" ${String(o?.to_branch_id)===String(a.value)?"selected":""}>
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
            placeholder="https://drive.google.com/file/d/..." value="${o?.document_link?String(o.document_link).replace(/"/g,"&quot;"):""}" />
        </div>
      </div>
    `,te({title:y?"\u270F\uFE0F Edit Data Mutasi Karyawan":"\u2795 Tambah Data Mutasi Karyawan",content:k,size:"md",confirmText:y?"Simpan Perubahan":"Simpan Mutasi",confirmClass:"btn-primary",onConfirm:async(a,c)=>{let h=k.querySelector("#mutasi-modal-tanggal").value,v=(k.querySelector("#mutasi-modal-emp").value||"").trim(),I=k.querySelector("#mutasi-modal-from").value,j=k.querySelector("#mutasi-modal-to").value,L=k.querySelector("#mutasi-modal-status").value,_=(k.querySelector("#mutasi-modal-link").value||"").trim();if(!h){q("Tanggal mutasi wajib diisi");return}if(!v){q("Nama karyawan wajib diisi");return}if(!I){q("Cabang asal wajib dipilih");return}if(!j){q("Cabang tujuan wajib dipilih");return}let P=n.find(E=>String(E.value)===String(I)),K=n.find(E=>String(E.value)===String(j)),l={tanggal:h,employee_name:v,from_branch_id:Number(I),from_branch_name:P?P.label:"Cabang Asal",to_branch_id:Number(j),to_branch_name:K?K.label:"Cabang Tujuan",status:L,document_link:_};try{if(y){let E=await R(`/api/mutasi/${o.id}`,{method:"PUT",body:JSON.stringify(l)});if(E.ok){G("Data mutasi berhasil diperbarui");let D=t.findIndex(z=>z.id===o.id);D!==-1&&(t[D]={...t[D],...l}),c(),$()}else q(E.data?.error||"Gagal memperbarui data mutasi")}else{let E=await R("/api/mutasi",{method:"POST",body:JSON.stringify(l)});if(E.ok){G("Data mutasi baru berhasil ditambahkan");let D=E.data?.id||Date.now();t.unshift({id:D,...l}),c(),$()}else q(E.data?.error||"Gagal menambahkan mutasi")}}catch(E){q("Terjadi kesalahan: "+E.message)}}})}function u(){return t.filter(o=>{if(g){let f=g.toLowerCase(),y=(o.employee_name||"").toLowerCase().includes(f),x=(o.from_branch_name||"").toLowerCase().includes(f),k=(o.to_branch_name||"").toLowerCase().includes(f);if(!y&&!x&&!k)return!1}return!(m!=="all"&&String(o.from_branch_id)!==String(m)||b!=="all"&&String(o.to_branch_id)!==String(b)||r!=="all"&&o.status!==r)})}function $(){let o=u(),f=t.length,y=t.filter(v=>v.status==="Proses").length,x=t.filter(v=>v.status==="Selesai").length,k=new Date,a=k.getFullYear(),c=k.getMonth(),h=t.filter(v=>{if(!v.tanggal)return!1;let I=new Date(v.tanggal);return!isNaN(I.getTime())&&I.getFullYear()===a&&I.getMonth()===c}).length;e.innerHTML=`
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
              ${S?`
                <button id="mutasi-btn-add" class="btn btn-primary mutasi-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Mutasi</span>
                </button>
              `:""}
              
              ${Q(F,"export")?`<button id="mutasi-btn-excel" class="btn btn-ghost mutasi-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
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
                <div class="mutasi-stat-val">${f}</div>
                <div class="mutasi-stat-lbl">Total Mutasi</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u23F3
              </div>
              <div>
                <div class="mutasi-stat-val">${y}</div>
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
                <div class="mutasi-stat-val">${h}</div>
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
              <input type="text" id="mutasi-search-input" class="mutasi-search-input" placeholder="Cari nama karyawan, cabang asal, atau cabang tujuan..." value="${g}" />
            </div>

            <select id="mutasi-filter-from" class="mutasi-select">
              <option value="all">Semua Cabang Asal (${n.length})</option>
              ${n.map(v=>`<option value="${v.value}" ${String(m)===String(v.value)?"selected":""}>${v.label}</option>`).join("")}
            </select>

            <select id="mutasi-filter-to" class="mutasi-select">
              <option value="all">Semua Cabang Tujuan (${n.length})</option>
              ${n.map(v=>`<option value="${v.value}" ${String(b)===String(v.value)?"selected":""}>${v.label}</option>`).join("")}
            </select>

            <div class="mutasi-view-toggle">
              <button type="button" class="mutasi-view-btn ${d==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="mutasi-view-btn ${d==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Filter Pills (Swipeable on Mobile) -->
          <div class="mutasi-filter-pills">
            <div class="mutasi-pill ${r==="all"?"active":""}" data-status="all">
              <span>Semua Mutasi</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${f}</span>
            </div>
            <div class="mutasi-pill ${r==="Proses"?"active":""}" data-status="Proses">
              <span style="color:#D97706;">\u23F3</span>
              <span>Proses</span>
              <span class="badge" style="background:rgba(217,119,6,0.15); color:#B45309; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${y}</span>
            </div>
            <div class="mutasi-pill ${r==="Selesai"?"active":""}" data-status="Selesai">
              <span style="color:#16A34A;">\u2705</span>
              <span>Selesai</span>
              <span class="badge" style="background:rgba(22,163,74,0.15); color:#15803D; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${x}</span>
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
          `:d==="table"?`
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
                      ${w||C?'<th style="width:90px; text-align:center;">Aksi</th>':""}
                    </tr>
                  </thead>
                  <tbody>
                    ${o.map((v,I)=>{let j=v.status==="Selesai",L=(v.employee_name||"Karyawan").split(" ").map(_=>_[0]).filter(Boolean).slice(0,2).join("").toUpperCase();return`
                        <tr data-id="${v.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${I+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="mutasi-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${L}
                              </div>
                              <div>
                                <div style="font-weight:700; color:#0F172A; white-space:nowrap;">${v.employee_name}</div>
                                <div style="font-size:0.72rem; color:#64748B;">ID Mutasi: #${v.id}</div>
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            \u{1F5D3}\uFE0F ${en(v.tanggal)}
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#334155; font-size:0.85rem; display:inline-flex; align-items:center; gap:4px;">
                              <span>\u{1F4CD}</span>
                              <span>${v.from_branch_name||"-"}</span>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:700; color:#2563EB; font-size:0.85rem; display:inline-flex; align-items:center; gap:5px;">
                              <span>\u2794</span>
                              <span>${v.to_branch_name||"-"}</span>
                            </div>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="mutasi-status-badge ${j?"selesai":"proses"}">
                              <span>${j?"\u2705":"\u23F3"}</span>
                              <span>${v.status||"Proses"}</span>
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            ${v.document_link?`
                              <a href="${v.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.76rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                                <span>\u{1F4C4} Buka SK</span>
                              </a>
                            `:`
                              <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                            `}
                          </td>
                          ${w||C?`
                            <td style="text-align:center; white-space:nowrap;">
                              <div style="display:inline-flex; align-items:center; gap:6px;">
                                ${w?`
                                  <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${v.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                  </button>
                                `:""}
                                ${C?`
                                  <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${v.id}">
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
              ${o.map(v=>{let I=v.status==="Selesai",j=(v.employee_name||"Karyawan").split(" ").map(L=>L[0]).filter(Boolean).slice(0,2).join("").toUpperCase();return`
                  <div class="mutasi-card" data-id="${v.id}">
                    <div>
                      <!-- Card Header: Status & Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="mutasi-status-badge ${I?"selesai":"proses"}">
                          <span>${I?"\u2705":"\u23F3"}</span>
                          <span>${v.status||"Proses"}</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B; display:inline-flex; align-items:center; gap:4px;">
                          <span>\u{1F5D3}\uFE0F</span>
                          <span>${en(v.tanggal)}</span>
                        </span>
                      </div>

                      <!-- Employee Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="mutasi-avatar-box">
                          ${j}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${v.employee_name}
                          </h4>
                          <div style="font-size:0.75rem; color:#64748B;">
                            ID Mutasi: #${v.id}
                          </div>
                        </div>
                      </div>

                      <!-- Route Visual Box -->
                      <div class="mutasi-route-box">
                        <div class="mutasi-route-node">
                          <div class="mutasi-route-label">Cabang Asal</div>
                          <div class="mutasi-route-branch" title="${v.from_branch_name||"-"}">
                            ${v.from_branch_name||"-"}
                          </div>
                        </div>

                        <div class="mutasi-route-arrow">
                          \u2794
                        </div>

                        <div class="mutasi-route-node" style="text-align:right;">
                          <div class="mutasi-route-label">Cabang Tujuan</div>
                          <div class="mutasi-route-branch" style="color:#2563EB;" title="${v.to_branch_name||"-"}">
                            ${v.to_branch_name||"-"}
                          </div>
                        </div>
                      </div>
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:8px;">
                      <div>
                        ${v.document_link?`
                          <a href="${v.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>\u{1F4C4} Buka SK</span>
                          </a>
                        `:`
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${w||C?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${w?`
                            <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${v.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                          `:""}
                          ${C?`
                            <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${v.id}">
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
    `,p()}function p(){let o=e.querySelector("#mutasi-btn-add");o&&(o.onclick=()=>T());let f=e.querySelector("#mutasi-btn-excel");f&&(f.onclick=()=>A());let y=e.querySelector("#mutasi-btn-refresh");y&&(y.onclick=()=>M());let x=e.querySelector("#mutasi-btn-reset-filter");x&&(x.onclick=()=>{g="",m="all",b="all",r="all",$()});let k=e.querySelector("#mutasi-search-input");k&&(k.oninput=h=>{g=h.target.value,$();let v=e.querySelector("#mutasi-search-input");v&&(v.focus(),v.setSelectionRange(v.value.length,v.value.length))});let a=e.querySelector("#mutasi-filter-from");a&&(a.onchange=h=>{m=h.target.value,$()});let c=e.querySelector("#mutasi-filter-to");c&&(c.onchange=h=>{b=h.target.value,$()}),e.querySelectorAll(".mutasi-pill").forEach(h=>{h.onclick=()=>{r=h.getAttribute("data-status")||"all",$()}}),e.querySelectorAll(".mutasi-view-btn").forEach(h=>{h.onclick=()=>{let v=h.getAttribute("data-view");v&&v!==d&&(d=v,localStorage.setItem("fcms_mutasi_view_mode",v),$())}}),e.querySelectorAll(".mutasi-edit-btn").forEach(h=>{h.onclick=v=>{v.stopPropagation();let I=Number(h.getAttribute("data-id")),j=t.find(L=>L.id===I);j&&T(j)}}),e.querySelectorAll(".mutasi-delete-btn").forEach(h=>{h.onclick=v=>{v.stopPropagation();let I=Number(h.getAttribute("data-id"));B(I)}})}await M()}me();var Ht=[{id:1,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:1,employee_name:"ADE SURAHMAN",start_time:"23:00",end_time:"08:00",break_hours:1,total_hours:8,reason:"General Cleaning"},{id:2,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:2,employee_name:"TAUFIK HIDAYAT",start_time:"00:00",end_time:"08:00",break_hours:1,total_hours:7,reason:"General Cleaning"},{id:3,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:3,employee_name:"YUDA WILDAN WIGUNA",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"},{id:4,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:4,employee_name:"REZA RIZKI",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"},{id:5,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:5,employee_name:"ADITYA PRATAMA",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"}],an=["General Cleaning","Deep Cleaning","Cover manpower","Longshift","Pekerjaan urgent"];function nn(e,t,n){if(!e||!t)return 0;try{let[i,r]=e.split(":").map(Number),[m,b]=t.split(":").map(Number),g=m*60+b-(i*60+r);g<0&&(g+=24*60);let s=g/60-(parseFloat(n)||0);return s<0&&(s=0),Number.isInteger(s)?s:parseFloat(s.toFixed(2))}catch{return 0}}function Ut(e=""){let t=(e||"").toLowerCase();return t.includes("general")?{bg:"#EFF6FF",color:"#1D4ED8",border:"#BFDBFE",icon:"\u{1F9F9}"}:t.includes("deep")?{bg:"#F5F3FF",color:"#6D28D9",border:"#DDD6FE",icon:"\u2728"}:t.includes("cover")?{bg:"#ECFDF5",color:"#047857",border:"#A7F3D0",icon:"\u{1F465}"}:t.includes("longshift")?{bg:"#FFFBEB",color:"#B45309",border:"#FDE68A",icon:"\u23F3"}:t.includes("urgent")?{bg:"#FFF1F2",color:"#BE123C",border:"#FECDD3",icon:"\u{1F6A8}"}:{bg:"#F8FAFC",color:"#475569",border:"#CBD5E1",icon:"\u23F1\uFE0F"}}function Gt(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function rn(e){let t=[],n=[],i=[],r="all",m="all",b="",g=!1,s=localStorage.getItem("fcms_overtime_view_mode")||"card",d=ae(),F=Q(d,"overtime","write"),S=Q(d,"overtime","write"),w=Q(d,"overtime","delete");try{[n,i]=await Promise.all([se(),Te()])}catch(p){console.error("Error fetching cache:",p)}async function C(){g=!0,u();try{let p=await R("/api/overtime?limit=10000");if(p.ok){let o=p.data?.data||p.data||[];t=Array.isArray(o)&&o.length>0?o:Ht}else t=Ht}catch(p){console.warn("Gagal load lembur, menggunakan data lokal:",p),t=Ht}finally{g=!1,u()}}async function M(){if(!Q(d,"export")){q("Anda tidak memiliki izin untuk mengekspor data.");return}try{let{downloadExcel:p}=await Promise.resolve().then(()=>(de(),xe)),o=t.map((f,y)=>({No:y+1,Tanggal:f.date||"","Cabang Lembur":f.branch_name||"","Nama Karyawan":f.employee_name||"","Jam Mulai":f.start_time||"","Jam Selesai":f.end_time||"","Istirahat (Jam)":f.break_hours!=null?f.break_hours:0,"Total Jam":f.total_hours!=null?f.total_hours:0,"Alasan / Keterangan":f.reason||""}));p(o,`Data_Lembur_${new Date().toISOString().slice(0,10)}`),G("Data lembur berhasil diekspor ke Excel!")}catch(p){q("Gagal ekspor Excel: "+p.message)}}async function A(p){if(!w)return;let o=t.find(f=>f.id===p);ie({title:"Hapus Catatan Lembur",message:`Apakah Anda yakin ingin menghapus catatan lembur untuk "${o?.employee_name||"karyawan ini"}" pada tanggal ${Gt(o?.date)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let f=await R(`/api/overtime/${p}`,{method:"DELETE"});f.ok?(G("Catatan lembur berhasil dihapus"),t=t.filter(y=>y.id!==p),u()):q(f.data?.error||"Gagal menghapus lembur")}catch(f){q("Gagal menghapus catatan lembur: "+f.message)}}})}function B(p=null){if(!S)return;let o=!!p,f=p?.date?p.date.slice(0,10):new Date().toISOString().slice(0,10),y=p?.start_time||"22:00",x=p?.end_time||"08:00",k=p?.break_hours!=null?p.break_hours:1,a=p?.total_hours!=null?p.total_hours:nn(y,x,k),c=document.createElement("div");c.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="ot-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${f}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penugasan <span style="color:#DC2626;">*</span>
            </label>
            <select id="ot-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${n.map(_=>`
                <option value="${_.value}" ${String(p?.branch_id)===String(_.value)?"selected":""}>
                  ${_.label}
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
            ${i.map(_=>{let P=typeof _=="string"?_:_.label||_.full_name||_.name||"",K=typeof _=="object"&&_!==null?_.value!=null?_.value:_.id!=null?_.id:P:_,l=String(p?.employee_id)===String(K)||p?.employee_name===P;return`<option value="${K}" ${l?"selected":""}>${P}</option>`}).join("")}
          </select>
        </div>

        <!-- Shift & Waktu Lembur -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Mulai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${y}" required />
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
              value="${k}" required />
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
            ${an.map(_=>`
              <option value="${_}" ${p?.reason===_?"selected":""}>${_}</option>
            `).join("")}
          </select>
        </div>
      </div>
    `;let h=c.querySelector("#ot-modal-start"),v=c.querySelector("#ot-modal-end"),I=c.querySelector("#ot-modal-break"),j=c.querySelector("#ot-modal-total");function L(){let _=h.value,P=v.value,K=I.value;j.value=nn(_,P,K)}h.addEventListener("input",L),v.addEventListener("input",L),I.addEventListener("input",L),te({title:o?"\u270F\uFE0F Edit Data Lembur Karyawan":"\u2795 Catat Penugasan Lembur Baru",content:c,size:"md",confirmText:o?"Simpan Perubahan":"Catat Lembur",confirmClass:"btn-primary",onConfirm:async(_,P)=>{let K=c.querySelector("#ot-modal-date").value,l=c.querySelector("#ot-modal-branch").value,E=c.querySelector("#ot-modal-emp").value,D=c.querySelector("#ot-modal-start").value,z=c.querySelector("#ot-modal-end").value,J=parseFloat(c.querySelector("#ot-modal-break").value)||0,N=parseFloat(c.querySelector("#ot-modal-total").value)||0,O=c.querySelector("#ot-modal-reason").value;if(!K){q("Tanggal lembur wajib diisi");return}if(!l){q("Cabang penugasan wajib dipilih");return}if(!E){q("Nama karyawan wajib dipilih");return}if(!D||!z){q("Jam mulai dan jam selesai wajib diisi");return}let U=n.find(V=>String(V.value)===String(l)),H=i.find(V=>String(V.value)===String(E)||String(V.id)===String(E)),X=H?H.label||H.full_name||H.name||"Karyawan":typeof E=="string"&&isNaN(Number(E))?E:"Karyawan",W={date:K,branch_id:Number(l),branch_name:U?U.label:"Cabang Lembur",employee_id:Number(E),employee_name:H?H.label:"Karyawan",start_time:D,end_time:z,break_hours:J,total_hours:N,reason:O};try{if(o){let V=await R(`/api/overtime/${p.id}`,{method:"PUT",body:JSON.stringify(W)});if(V.ok){G("Data lembur berhasil diperbarui");let ee=t.findIndex(Z=>Z.id===p.id);ee!==-1&&(t[ee]={...t[ee],...W}),P(),u()}else q(V.data?.error||"Gagal memperbarui data lembur")}else{let V=await R("/api/overtime",{method:"POST",body:JSON.stringify(W)});if(V.ok){G("Data lembur berhasil ditambahkan");let ee=V.data?.id||Date.now();t.unshift({id:ee,...W}),P(),u()}else q(V.data?.error||"Gagal menambahkan lembur")}}catch(V){q("Terjadi kesalahan: "+V.message)}}})}function T(){return t.filter(p=>{if(b){let o=b.toLowerCase(),f=(p.employee_name||"").toLowerCase().includes(o),y=(p.branch_name||"").toLowerCase().includes(o),x=(p.reason||"").toLowerCase().includes(o);if(!f&&!y&&!x)return!1}return!(m!=="all"&&String(p.branch_id)!==String(m)||r!=="all"&&p.reason!==r)})}function u(){let p=T(),o=t.length,f=t.reduce((h,v)=>h+(parseFloat(v.total_hours)||0),0),y=o>0?(f/o).toFixed(1):0,x=new Date,k=x.getFullYear(),a=x.getMonth(),c=t.filter(h=>{if(!h.date)return!1;let v=new Date(h.date);return!isNaN(v.getTime())&&v.getFullYear()===k&&v.getMonth()===a}).length;e.innerHTML=`
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
              ${F?`
                <button id="lembur-btn-add" class="btn btn-primary lembur-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Lembur</span>
                </button>
              `:""}
              
              ${Q(d,"export")?`<button id="lembur-btn-excel" class="btn btn-ghost lembur-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
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
                <div class="lembur-stat-val">${f} Jam</div>
                <div class="lembur-stat-lbl">Total Jam Lembur</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F4DD}
              </div>
              <div>
                <div class="lembur-stat-val">${o}</div>
                <div class="lembur-stat-lbl">Catatan Penugasan</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4CA}
              </div>
              <div>
                <div class="lembur-stat-val">${y} Jam</div>
                <div class="lembur-stat-lbl">Rata-rata Durasi</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="lembur-stat-val">${c}</div>
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
              <input type="text" id="lembur-search-input" class="lembur-search-input" placeholder="Cari nama karyawan, cabang, atau alasan lembur..." value="${b}" />
            </div>

            <select id="lembur-filter-branch" class="lembur-select">
              <option value="all">Semua Cabang (${n.length})</option>
              ${n.map(h=>`<option value="${h.value}" ${String(m)===String(h.value)?"selected":""}>${h.label}</option>`).join("")}
            </select>

            <div class="lembur-view-toggle">
              <button type="button" class="lembur-view-btn ${s==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="lembur-view-btn ${s==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Reason Filter Pills (Swipeable on Mobile) -->
          <div class="lembur-filter-pills">
            <div class="lembur-pill ${r==="all"?"active":""}" data-reason="all">
              <span>Semua Lembur</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${o}</span>
            </div>
            ${an.map(h=>{let v=t.filter(j=>j.reason===h).length,I=Ut(h);return`
                <div class="lembur-pill ${r===h?"active":""}" data-reason="${h}">
                  <span>${I.icon}</span>
                  <span>${h}</span>
                  ${v>0?`<span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${v}</span>`:""}
                </div>
              `}).join("")}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="lembur-list-container">
          ${p.length===0?`
            <div class="lembur-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data Lembur Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="lembur-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:s==="table"?`
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
                      ${S||w?'<th style="width:90px; text-align:center;">Aksi</th>':""}
                    </tr>
                  </thead>
                  <tbody>
                    ${p.map((h,v)=>{let I=(h.employee_name||"Karyawan").split(" ").map(L=>L[0]).filter(Boolean).slice(0,2).join("").toUpperCase(),j=Ut(h.reason);return`
                        <tr data-id="${h.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${v+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="lembur-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${I}
                              </div>
                              <div style="font-weight:700; color:#0F172A; white-space:nowrap;">
                                ${h.employee_name}
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#334155; font-size:0.85rem; display:inline-flex; align-items:center; gap:4px;">
                              <span>\u{1F4CD}</span>
                              <span>${h.branch_name||"Cabang"}</span>
                            </div>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            \u{1F4C5} ${Gt(h.date)}
                          </td>
                          <td style="white-space:nowrap;">
                            <span style="font-weight:700; color:#2563EB;">\u23F0 ${h.start_time||"00:00"} - ${h.end_time||"00:00"}</span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="lembur-hours-badge" style="padding:3px 8px; font-size:0.8rem;">
                              <span>\u23F1\uFE0F</span>
                              <span>${h.total_hours!=null?h.total_hours:0} Jam</span>
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap; color:#64748B; font-weight:600; font-size:0.82rem;">
                            \u2615 ${h.break_hours!=null?h.break_hours:0} Jam
                          </td>
                          <td style="white-space:nowrap;">
                            <span class="lembur-reason-badge" style="background:${j.bg}; color:${j.color}; border:1px solid ${j.border};">
                              <span>${j.icon}</span>
                              <span>${h.reason||"Lembur Operasional"}</span>
                            </span>
                          </td>
                          ${S||w?`
                            <td style="text-align:center; white-space:nowrap;">
                              <div style="display:inline-flex; align-items:center; gap:6px;">
                                ${S?`
                                  <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${h.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                  </button>
                                `:""}
                                ${w?`
                                  <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${h.id}">
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
              ${p.map(h=>{let v=(h.employee_name||"Karyawan").split(" ").map(j=>j[0]).filter(Boolean).slice(0,2).join("").toUpperCase(),I=Ut(h.reason);return`
                  <div class="lembur-card" data-id="${h.id}">
                    <div>
                      <!-- Card Header: Jam & Istirahat -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="lembur-hours-badge">
                          <span>\u23F1\uFE0F</span>
                          <span>${h.total_hours!=null?h.total_hours:0} Jam Lembur</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                          \u2615 Istirahat: ${h.break_hours!=null?h.break_hours:0} Jam
                        </span>
                      </div>

                      <!-- Employee & Branch -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="lembur-avatar-box">
                          ${v}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${h.employee_name}
                          </h4>
                          <div style="font-size:0.8rem; color:#64748B; display:flex; align-items:center; gap:4px;">
                            <span>\u{1F4CD}</span>
                            <span style="font-weight:600; color:#334155;">${h.branch_name||"Cabang"}</span>
                          </div>
                        </div>
                      </div>

                      <!-- Shift Times & Date Box -->
                      <div class="lembur-time-box">
                        <div>
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Tanggal</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#1E293B;">
                            \u{1F5D3}\uFE0F ${Gt(h.date)}
                          </div>
                        </div>

                        <div style="text-align:right;">
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Waktu Shift</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#2563EB;">
                            \u23F0 ${h.start_time||"00:00"} - ${h.end_time||"00:00"}
                          </div>
                        </div>
                      </div>

                      <!-- Reason Badge -->
                      <div>
                        <span class="lembur-reason-badge" style="background:${I.bg}; color:${I.color}; border:1px solid ${I.border};">
                          <span>${I.icon}</span>
                          <span>${h.reason||"Lembur Operasional"}</span>
                        </span>
                      </div>
                    </div>

                    <!-- Card Footer: Aksi -->
                    ${S||w?`
                      <div style="display:flex; justify-content:flex-end; align-items:center; gap:6px; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:12px;">
                        ${S?`
                          <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${h.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                        `:""}
                        ${w?`
                          <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${h.id}">
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
    `,$()}function $(){let p=e.querySelector("#lembur-btn-add");p&&(p.onclick=()=>B());let o=e.querySelector("#lembur-btn-excel");o&&(o.onclick=()=>M());let f=e.querySelector("#lembur-btn-refresh");f&&(f.onclick=()=>C());let y=e.querySelector("#lembur-btn-reset-filter");y&&(y.onclick=()=>{b="",m="all",r="all",u()});let x=e.querySelector("#lembur-search-input");x&&(x.oninput=a=>{b=a.target.value,u();let c=e.querySelector("#lembur-search-input");c&&(c.focus(),c.setSelectionRange(c.value.length,c.value.length))});let k=e.querySelector("#lembur-filter-branch");k&&(k.onchange=a=>{m=a.target.value,u()}),e.querySelectorAll(".lembur-pill").forEach(a=>{a.onclick=()=>{r=a.getAttribute("data-reason")||"all",u()}}),e.querySelectorAll(".lembur-view-btn").forEach(a=>{a.onclick=()=>{let c=a.getAttribute("data-view");c&&c!==s&&(s=c,localStorage.setItem("fcms_overtime_view_mode",c),u())}}),e.querySelectorAll(".lembur-edit-btn").forEach(a=>{a.onclick=c=>{c.stopPropagation();let h=Number(a.getAttribute("data-id")),v=t.find(I=>I.id===h);v&&B(v)}}),e.querySelectorAll(".lembur-delete-btn").forEach(a=>{a.onclick=c=>{c.stopPropagation();let h=Number(a.getAttribute("data-id"));A(h)}})}await C()}async function on(e){let t=ae();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}be({container:e,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:n=>new Date(n).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(n,i)=>`<strong>${n||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${i.user_role||""}</small>`},{key:"action",label:"Aksi",render:n=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[n]||"badge-neutral"}">${n}</span>`},{key:"module",label:"Modul",render:n=>`<span style="text-transform:capitalize">${(n||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(n,i)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${n}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async n=>{try{let m=((await(await fetch(`/api/audit-logs?search=${n}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(d=>String(d.id)===String(n));if(!m)return alert("Data tidak ditemukan");let b=d=>{if(!d)return"Tidak ada data";try{return JSON.stringify(JSON.parse(d),null,2)}catch{return d}},g=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${b(m.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${b(m.new_data)}</pre>
           </div>
         </div>
       `,{createModal:s}=await Promise.resolve().then(()=>(me(),Zt));s({title:`Detail Audit Log #${n}`,content:g,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}me();de();var Jt={"2026-09":"September 2026","2026-10":"Oktober 2026","2026-11":"November 2026","2026-12":"Desember 2026"},Ze={scheduled:{label:"Terjadwal",badgeClass:"badge-warning",icon:"\u{1F7E1}",bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},in_progress:{label:"Sedang Proses",badgeClass:"badge-info",icon:"\u{1F535}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},completed:{label:"Selesai",badgeClass:"badge-success",icon:"\u{1F7E2}",bg:"#DCFCE7",color:"#166534",border:"#86EFAC"}};async function sn(e){let t=ae(),n=t&&(t.role==="superadmin"||t.role==="admin"),i=Q(t,"export");e.innerHTML=`
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
        ${i?`
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
  `;let r="board",m=null,b=[];R("/api/branches?all=1").then(u=>{u.ok&&u.data?.data&&(b=u.data.data)}).catch(()=>{});async function g(){try{let u=await R("/api/os-migrations/summary");if(!u.ok)throw new Error("Gagal memuat data migrasi");m=u.data?.data||u.data,s(m),F()}catch(u){document.getElementById("migration-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${u.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function s(u){let $=document.getElementById("migration-summary-cards");if(!$||!u)return;let p=u.total_branches||0,o=u.vendors?.AEON?.total||0,f=u.vendors?.["Triple S"]?.total||0,y=u.by_status?.completed||0,x=u.by_status?.in_progress||0,k=u.by_status?.scheduled||0,a=p>0?Math.round(y/p*100):0;$.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${p} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.8rem; font-weight:800; color:#1E40AF; line-height:1.2;">${o} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
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
          <span style="font-size:0.85rem; font-weight:700; color:#15803D;">${a}%</span>
        </div>
        <div style="width:100%; height:8px; background:#E2E8F0; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${a}%; height:100%; background:#16A34A; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.75rem; color:var(--text-2); font-weight:500;">
          <span>\u{1F7E2} Selesai: <strong>${y}</strong></span>
          <span>\u{1F535} Proses: <strong>${x}</strong></span>
          <span>\u{1F7E1} Terjadwal: <strong>${k}</strong></span>
        </div>
      </div>
    `}function d(){if(!m||!m.matrix)return[];let u=(document.getElementById("migration-search")?.value||"").toLowerCase().trim(),$=document.getElementById("filter-vendor")?.value||"",p=document.getElementById("filter-month")?.value||"",o=document.getElementById("filter-status")?.value||"";return m.matrix.filter(f=>!($&&f.vendor!==$||p&&f.target_month!==p||o&&f.status!==o||u&&!`${f.branch_code} ${f.branch_name} ${f.branch_full_name} ${f.vendor} ${f.notes||""} ${f.pic_vendor||""}`.toLowerCase().includes(u)))}function F(){let u=d(),$=document.getElementById("filter-stats-text");$&&($.innerHTML=`Menampilkan <strong>${u.length}</strong> dari <strong>${m?.total_branches||0}</strong> cabang`);let p=document.getElementById("migration-view-container");r==="board"?S(p,u):w(p,u)}function S(u,$){if(!$.length){u.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;return}let p=["AEON","Triple S"],o=ta(Array.from(new Set($.map(y=>y.target_month).filter(Boolean)))),f="";p.forEach(y=>{let x=$.filter(I=>I.vendor===y);if(document.getElementById("filter-vendor")?.value&&document.getElementById("filter-vendor").value!==y||x.length===0)return;let k=x.length,a=y==="AEON"?"#1E40AF":"#0F766E",c=y==="AEON"?"#EFF6FF":"#F0FDFA",h=y==="AEON"?"#BFDBFE":"#99F6E4",v=o.filter(I=>x.some(j=>j.target_month===I));f+=`
        <div class="vendor-board-section" style="margin-bottom:32px; background:#fff; border:1px solid var(--border); border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <!-- Vendor Section Header -->
          <div style="padding:14px 20px; background:${c}; border-bottom:1px solid ${h}; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.1rem; font-weight:800; color:${a};">${y==="AEON"?"\u{1F3E2} AEON":"\u{1F6E1}\uFE0F Triple S"}</span>
              <span class="badge" style="background:#fff; color:${a}; border:1px solid ${h}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px;">
                ${k} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
              Format Timeline Berdasarkan Dokumen Migrasi
            </div>
          </div>

          <!-- Months Grid (Hanya menampilkan bulan yang memiliki jadwal) -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px; padding:18px;">
            ${v.map(I=>{let j=x.filter(P=>P.target_month===I);return`
                <div class="month-column" style="background:#F8FAFC; border:1px solid var(--border); border-radius:10px; display:flex; flex-direction:column; min-height:220px;">
                  <!-- Month Header -->
                  <div style="padding:12px 14px; background:#F1F5F9; border-bottom:1px solid var(--border); border-radius:9px 9px 0 0; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">
                      \u{1F4C5} ${(Jt[I]||I).split(" ")[0]}
                    </div>
                    <span class="badge" style="background:#E2E8F0; color:var(--text-2); font-weight:700; font-size:0.75rem; padding:2px 7px; border-radius:4px;">
                      ${j.length}
                    </span>
                  </div>

                  <!-- Branch Cards Container -->
                  <div style="padding:10px; display:flex; flex-direction:column; gap:8px; flex:1; max-height:480px; overflow-y:auto;">
                    ${j.map(P=>{let K=Ze[P.status]||Ze.scheduled;return`
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
                              <button class="status-toggle-btn" data-id="${P.id}" data-status="${P.status}" style="background:${K.bg}; color:${K.color}; border:1px solid ${K.border}; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; display:inline-flex; align-items:center; gap:4px; cursor:pointer;" title="Klik untuk ubah status">
                                ${K.icon} ${K.label} \u25BE
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
      `}),u.innerHTML=f,C(u)}function w(u,$){if(!$.length){u.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
        </div>
      `;return}u.innerHTML=`
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
              ${$.map((p,o)=>{let f=Ze[p.status]||Ze.scheduled;return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3);">${o+1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">${p.branch_full_name}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${p.vendor==="AEON"?"#DBEAFE":"#CCFBF1"}; color:${p.vendor==="AEON"?"#1E40AF":"#0F766E"}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${p.vendor}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:var(--text-2);">
                      ${Jt[p.target_month]||p.target_month}
                    </td>
                    <td style="padding:12px 16px;">
                      <button class="status-toggle-btn" data-id="${p.id}" data-status="${p.status}" style="background:${f.bg}; color:${f.color}; border:1px solid ${f.border}; padding:4px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; cursor:pointer;" title="Klik untuk ubah status">
                        ${f.icon} ${f.label} \u25BE
                      </button>
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-2);">
                      ${p.actual_date||"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; font-weight:600; color:var(--text-2);">
                      ${p.employee_count>0?`\u{1F464} ${p.employee_count}`:"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3); max-width:200px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${p.notes||"-"}
                    </td>
                    ${n?`
                      <td style="padding:12px 16px; text-align:right;">
                        <button class="btn btn-outline btn-sm edit-card-btn" data-id="${p.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-card-btn" data-id="${p.id}" data-name="${p.branch_name}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,C(u)}function C(u){u.querySelectorAll(".status-toggle-btn").forEach($=>{$.addEventListener("click",p=>{p.stopPropagation();let o=$.dataset.id,f=$.dataset.status;M($,o,f)})}),u.querySelectorAll(".edit-card-btn").forEach($=>{$.addEventListener("click",()=>{let p=$.dataset.id,o=m?.matrix?.find(f=>String(f.id)===String(p));o&&A(o)})}),u.querySelectorAll(".delete-card-btn").forEach($=>{$.addEventListener("click",()=>{let p=$.dataset.id,o=$.dataset.name||"Cabang";ie(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${o}</strong>?`,async()=>{try{if(!(await R(`/api/os-migrations/${p}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus");G(`Jadwal migrasi ${o} berhasil dihapus`),g()}catch(f){q(f.message)}})})})}function M(u,$,p){document.querySelectorAll(".status-popup-menu").forEach(k=>k.remove());let o=u.getBoundingClientRect(),f=document.createElement("div");f.className="status-popup-menu",f.style.cssText=`
      position: fixed;
      top: ${o.bottom+window.scrollY+4}px;
      left: ${o.left+window.scrollX}px;
      background: #fff;
      border: 1px solid var(--border);
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 9999;
      padding: 6px 0;
      min-width: 160px;
    `;let y=[{key:"scheduled",label:"\u{1F7E1} Terjadwal"},{key:"in_progress",label:"\u{1F535} Sedang Proses"},{key:"completed",label:"\u{1F7E2} Selesai"}];f.innerHTML=y.map(k=>`
      <button class="status-option-item" data-key="${k.key}" style="
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
        color: ${k.key===p?"#2563EB":"var(--text-1)"};
        font-weight: ${k.key===p?"700":"500"};
      ">
        ${k.label} ${k.key===p?"\u2713":""}
      </button>
    `).join("");let x=k=>{!f.contains(k.target)&&k.target!==u&&(f.remove(),document.removeEventListener("click",x))};setTimeout(()=>document.addEventListener("click",x),50),f.querySelectorAll(".status-option-item").forEach(k=>{k.addEventListener("click",async()=>{let a=k.dataset.key;if(f.remove(),document.removeEventListener("click",x),a!==p)try{if(!(await R(`/api/os-migrations/${$}/status`,{method:"PATCH",body:JSON.stringify({status:a})})).ok)throw new Error("Gagal memperbarui status");G("Status berhasil diubah"),g()}catch(c){q(c.message)}})}),document.body.appendChild(f)}function A(u=null){if(!n){q("Anda tidak memiliki izin untuk menambah atau mengedit migrasi OS.");return}let $=!!u,p=$?"Edit Jadwal Migrasi OS":"Tambah Jadwal Migrasi OS",f=`
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${b.map(y=>`
      <option value="${y.id}" ${u&&u.branch_id===y.id?"selected":""}>
        ${y.full_name}
      </option>
    `).join("")}
          </select>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Vendor Outsourcing *</label>
            <select name="vendor" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="AEON" ${u?.vendor==="AEON"?"selected":""}>AEON</option>
              <option value="Triple S" ${u?.vendor==="Triple S"?"selected":""}>Triple S</option>
              <option value="Lainnya" ${u&&!["AEON","Triple S"].includes(u.vendor)?"selected":""}>Lainnya</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Target Bulan *</label>
            <select name="target_month" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="2026-09" ${u?.target_month==="2026-09"?"selected":""}>September 2026</option>
              <option value="2026-10" ${u?.target_month==="2026-10"?"selected":""}>Oktober 2026</option>
              <option value="2026-11" ${u?.target_month==="2026-11"?"selected":""}>November 2026</option>
              <option value="2026-12" ${u?.target_month==="2026-12"?"selected":""}>Desember 2026</option>
            </select>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Status Pelaksanaan</label>
            <select name="status" class="form-control" style="width:100%; height:38px; border-radius:8px;">
              <option value="scheduled" ${u?.status==="scheduled"?"selected":""}>\u{1F7E1} Terjadwal</option>
              <option value="in_progress" ${u?.status==="in_progress"?"selected":""}>\u{1F535} Sedang Proses</option>
              <option value="completed" ${u?.status==="completed"?"selected":""}>\u{1F7E2} Selesai</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Tanggal Realisasi (Opsional)</label>
            <input type="date" name="actual_date" class="form-control" value="${u?.actual_date||""}" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">PIC / Koordinator Vendor (Opsional)</label>
          <input type="text" name="pic_vendor" class="form-control" placeholder="Nama PIC vendor / no HP" value="${u?.pic_vendor||""}" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Catatan / Keterangan</label>
          <textarea name="notes" class="form-control" rows="2" placeholder="Catatan serah terima / info tambahan..." style="width:100%; border-radius:8px;">${u?.notes||""}</textarea>
        </div>
      </form>
    `;te({title:p,content:f,confirmText:$?"Simpan Perubahan":"Tambah Cabang",onConfirm:async(y,x)=>{let k=y.querySelector("#migration-form");if(!k.checkValidity()){k.reportValidity();return}let a=new FormData(k),c={branch_id:parseInt(a.get("branch_id"),10),vendor:a.get("vendor"),target_month:a.get("target_month"),status:a.get("status"),actual_date:a.get("actual_date")||null,pic_vendor:a.get("pic_vendor")||null,notes:a.get("notes")||null};try{let h=$?`/api/os-migrations/${u.id}`:"/api/os-migrations",I=await R(h,{method:$?"PUT":"POST",body:JSON.stringify(c)});if(!I.ok){let j=I.data?.error||I.data?.message||"Gagal menyimpan data";throw new Error(j)}G($?"Jadwal migrasi berhasil diupdate":"Jadwal migrasi berhasil ditambahkan"),x(),g()}catch(h){q(h.message)}}})}let B=document.getElementById("view-board-btn"),T=document.getElementById("view-table-btn");B.addEventListener("click",()=>{r="board",B.style.background="#fff",B.classList.remove("btn-ghost"),T.style.background="transparent",T.classList.add("btn-ghost"),F()}),T.addEventListener("click",()=>{r="table",T.style.background="#fff",T.classList.remove("btn-ghost"),B.style.background="transparent",B.classList.add("btn-ghost"),F()}),["migration-search","filter-vendor","filter-month","filter-status"].forEach(u=>{let $=document.getElementById(u);$&&($.addEventListener("input",F),$.addEventListener("change",F))}),document.getElementById("btn-add-migration")?.addEventListener("click",()=>{A()}),document.getElementById("btn-export-migrations")?.addEventListener("click",()=>{let u=d();if(!u.length){q("Tidak ada data untuk diexport");return}let $=u.map((p,o)=>({No:o+1,"Kode Cabang":p.branch_code,"Nama Cabang":p.branch_name,"Cabang Lengkap":p.branch_full_name,"Vendor OS":p.vendor,"Target Bulan":Jt[p.target_month]||p.target_month,Status:Ze[p.status]?.label||p.status,"Tanggal Realisasi":p.actual_date||"","Jml Personil FC":p.employee_count||0,"PIC Vendor":p.pic_vendor||"",Catatan:p.notes||""}));ne($,`Timeline_Migrasi_OS_${new Date().toISOString().slice(0,10)}`),G("File Excel berhasil diunduh")}),g()}me();de();var ln={Persiapan:{bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE"},"Area Umum":{bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD"},"Ruang Tindakan":{bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},"Ruang Penunjang":{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},Toilet:{bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4"},"Area Luar":{bg:"#FEF3C7",color:"#92400E",border:"#FDE68A"},"Kontrol & Check":{bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA"},Istirahat:{bg:"#E2E8F0",color:"#475569",border:"#CBD5E1"},"Layanan Pasien":{bg:"#DCFCE7",color:"#15803D",border:"#86EFAC"},"Sampah Umum":{bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},"Sampah Medis":{bg:"#FEE2E2",color:"#991B1B",border:"#FCA5A5"},Penutup:{bg:"#E0E7FF",color:"#3730A3",border:"#C7D2FE"}};async function dn(e){let t=ae(),n=t&&(t.role==="superadmin"||t.role==="admin"),i=Q(t,"export");e.innerHTML=`
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
        ${i?`
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
  `;let r="pagi",m="timeline",b=[];async function g(){try{let u=await R("/api/dac?all=1");if(!u.ok)throw new Error("Gagal memuat standar DAC");b=u.data?.data||u.data||[];let $=b.filter(y=>y.shift==="pagi").length,p=b.filter(y=>y.shift==="siang").length;document.getElementById("badge-count-pagi").textContent=$,document.getElementById("badge-count-siang").textContent=p,document.getElementById("badge-count-all").textContent=b.length;let o=document.getElementById("dac-category-filter"),f=Array.from(new Set(b.map(y=>y.category))).sort();o.innerHTML='<option value="">Semua Kategori</option>'+f.map(y=>`<option value="${y}">${y}</option>`).join(""),d(),F()}catch(u){document.getElementById("dac-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${u.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function s(){let u=(document.getElementById("dac-search")?.value||"").toLowerCase().trim(),$=document.getElementById("dac-category-filter")?.value||"",p=document.getElementById("dac-period-filter")?.value||"";return b.filter(o=>{if(r!=="all"&&o.shift!==r||$&&o.category!==$)return!1;if(p==="morning"){let f=parseInt(o.start_time.split(":")[0],10);if(o.shift==="pagi"&&f>=12||o.shift==="siang"&&f>=17)return!1}else if(p==="afternoon"){let f=parseInt(o.start_time.split(":")[0],10);if(o.shift==="pagi"&&f<12||o.shift==="siang"&&f<17)return!1}return!(u&&!`${o.start_time} ${o.end_time} ${o.category} ${o.location} ${o.activity}`.toLowerCase().includes(u))})}function d(){let u=document.getElementById("dac-summary-cards");if(!u)return;let $=b.filter(x=>x.shift==="pagi"),p=b.filter(x=>x.shift==="siang"),o=r==="pagi"?$:r==="siang"?p:b,f=r==="pagi"?"08:00 \u2013 17:00":r==="siang"?"13:00 \u2013 22:00":"08:00 \u2013 22:00",y=r==="pagi"?"Shift Pagi":r==="siang"?"Shift Siang":"Semua Shift";u.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Aktivitas (${y})</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${o.length} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Aktivitas</span></div>
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
    `}function F(){let u=s(),$=document.getElementById("dac-stats-text");$&&($.innerHTML=`Menampilkan <strong>${u.length}</strong> aktivitas`);let p=document.getElementById("dac-view-container");m==="timeline"?S(p,u):w(p,u)}function S(u,$){if(!$.length){u.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;return}u.innerHTML=`
      <div class="timeline-container" style="display:flex; flex-direction:column; gap:12px; max-width:960px; margin:0 auto;">
        ${$.map(p=>{let o=ln[p.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},f=p.shift==="pagi"?"\u{1F305} Pagi":"\u{1F307} Siang",y=p.shift==="pagi"?"#FEF3C7":"#E0E7FF",x=p.shift==="pagi"?"#92400E":"#3730A3";return`
            <div class="dac-timeline-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:16px 20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; gap:16px; align-items:flex-start; position:relative;">
              <!-- Time Column -->
              <div style="min-width:130px; text-align:center; background:#F8FAFC; border:1px solid var(--border); border-radius:8px; padding:10px 8px;">
                <div style="font-size:1.1rem; font-weight:800; color:var(--text-1); font-family:monospace;">
                  ${p.start_time} - ${p.end_time}
                </div>
                <div style="font-size:0.75rem; color:var(--text-3); font-weight:600; margin-top:4px;">
                  Item #${p.item_no}
                </div>
                ${r==="all"?`
                  <div style="margin-top:6px;">
                    <span class="badge" style="background:${y}; color:${x}; font-size:0.7rem; font-weight:700; padding:2px 6px; border-radius:4px;">
                      ${f}
                    </span>
                  </div>
                `:""}
              </div>

              <!-- Content Column -->
              <div style="flex:1;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                  <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span class="badge" style="background:${o.bg}; color:${o.color}; border:1px solid ${o.border}; font-weight:700; font-size:0.75rem; padding:3px 10px; border-radius:6px;">
                      ${p.category}
                    </span>
                    <span style="font-weight:700; font-size:0.95rem; color:var(--text-1);">
                      \u{1F4CD} ${p.location}
                    </span>
                  </div>
                  ${n?`
                    <div style="display:flex; gap:4px;">
                      <button class="btn btn-ghost btn-sm edit-dac-btn" data-id="${p.id}" title="Edit" style="padding:4px 8px; font-size:0.8rem; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-ghost btn-sm delete-dac-btn" data-id="${p.id}" data-name="${p.location}" title="Hapus" style="padding:4px 8px; font-size:0.8rem; color:var(--danger); cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </div>
                  `:""}
                </div>

                <div style="font-size:0.9rem; color:var(--text-2); line-height:1.5; background:#FAFBFD; padding:10px 14px; border-radius:8px; border-left:3px solid ${o.color}; margin-top:8px;">
                  ${p.activity}
                </div>
              </div>
            </div>
          `}).join("")}
      </div>
    `,C(u)}function w(u,$){if(!$.length){u.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
        </div>
      `;return}u.innerHTML=`
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div class="table-responsive">
          <table class="table" style="width:100%; border-collapse:collapse;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:1px solid var(--border); text-align:left;">
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:50px;">NO</th>
                ${r==="all"?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">SHIFT</th>':""}
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:110px;">JAM</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:140px;">KATEGORI</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:200px;">LOKASI / AREA</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">AKTIVITAS</th>
                ${n?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right; width:100px;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${$.map(p=>{let o=ln[p.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"};return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 14px; font-size:0.85rem; color:var(--text-3); font-weight:600;">${p.item_no}</td>
                    ${r==="all"?`
                      <td style="padding:12px 14px;">
                        <span class="badge" style="background:${p.shift==="pagi"?"#FEF3C7":"#E0E7FF"}; color:${p.shift==="pagi"?"#92400E":"#3730A3"}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:4px;">
                          ${p.shift==="pagi"?"Pagi":"Siang"}
                        </span>
                      </td>
                    `:""}
                    <td style="padding:12px 14px; font-family:monospace; font-weight:700; font-size:0.85rem; color:var(--text-1); white-space:nowrap;">
                      ${p.start_time} - ${p.end_time}
                    </td>
                    <td style="padding:12px 14px;">
                      <span class="badge" style="background:${o.bg}; color:${o.color}; border:1px solid ${o.border}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px; white-space:nowrap;">
                        ${p.category}
                      </span>
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; font-weight:600; color:var(--text-1);">
                      ${p.location}
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; color:var(--text-2); line-height:1.4;">
                      ${p.activity}
                    </td>
                    ${n?`
                      <td style="padding:12px 14px; text-align:right; white-space:nowrap;">
                        <button class="btn btn-outline btn-sm edit-dac-btn" data-id="${p.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-dac-btn" data-id="${p.id}" data-name="${p.location}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,C(u)}function C(u){u.querySelectorAll(".edit-dac-btn").forEach($=>{$.addEventListener("click",()=>{let p=$.dataset.id,o=b.find(f=>String(f.id)===String(p));o&&M(o)})}),u.querySelectorAll(".delete-dac-btn").forEach($=>{$.addEventListener("click",()=>{let p=$.dataset.id,o=$.dataset.name||"item";ie(`Apakah Anda yakin ingin menghapus aktivitas pada lokasi <strong>${o}</strong>?`,async()=>{try{if(!(await R(`/api/dac/${p}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus item");G("Aktivitas DAC berhasil dihapus"),g()}catch(f){q(f.message)}})})})}function M(u=null){if(!n){q("Anda tidak memiliki izin untuk menambah atau mengedit aktivitas.");return}let $=!!u,p=$?"Edit Aktivitas DAC":"Tambah Aktivitas DAC",o=["Persiapan","Area Umum","Ruang Tindakan","Ruang Penunjang","Toilet","Area Luar","Kontrol & Check","Istirahat","Layanan Pasien","Sampah Umum","Sampah Medis","Penutup"],f=`
      <form id="dac-form" style="display:flex; flex-direction:column; gap:14px;">
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Shift *</label>
            <select name="shift" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="pagi" ${u?.shift==="pagi"||!u&&r==="pagi"?"selected":""}>\u{1F305} Shift Pagi (08:00 - 17:00)</option>
              <option value="siang" ${u?.shift==="siang"||!u&&r==="siang"?"selected":""}>\u{1F307} Shift Siang (13:00 - 22:00)</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Nomor Urut *</label>
            <input type="number" name="item_no" class="form-control" value="${u?.item_no||""}" placeholder="Otomatis jika kosong" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Mulai (HH:MM) *</label>
            <input type="time" name="start_time" class="form-control" value="${u?.start_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Selesai (HH:MM) *</label>
            <input type="time" name="end_time" class="form-control" value="${u?.end_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Kategori *</label>
          <input list="cat-suggestions" name="category" class="form-control" value="${u?.category||""}" required placeholder="Pilih atau ketik kategori..." style="width:100%; height:38px; border-radius:8px;">
          <datalist id="cat-suggestions">
            ${o.map(y=>`<option value="${y}">`).join("")}
          </datalist>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Lokasi / Area *</label>
          <input type="text" name="location" class="form-control" value="${u?.location||""}" required placeholder="Contoh: Ruang tindakan Du 101, Lobby, Toilet GF" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Aktivitas Detail *</label>
          <textarea name="activity" class="form-control" rows="3" required placeholder="Deskripsi aktivitas pembersihan..." style="width:100%; border-radius:8px;">${u?.activity||""}</textarea>
        </div>
      </form>
    `;te({title:p,content:f,confirmText:$?"Simpan Perubahan":"Tambah Aktivitas",onConfirm:async(y,x)=>{let k=y.querySelector("#dac-form");if(!k.checkValidity()){k.reportValidity();return}let a=new FormData(k),c={shift:a.get("shift"),item_no:a.get("item_no")?parseInt(a.get("item_no"),10):void 0,start_time:a.get("start_time"),end_time:a.get("end_time"),category:a.get("category"),location:a.get("location"),activity:a.get("activity")};try{let h=$?`/api/dac/${u.id}`:"/api/dac",I=await R(h,{method:$?"PUT":"POST",body:JSON.stringify(c)});if(!I.ok){let j=I.data?.error||I.data?.message||"Gagal menyimpan aktivitas";throw new Error(j)}G($?"Aktivitas berhasil diperbarui":"Aktivitas berhasil ditambahkan"),x(),g()}catch(h){q(h.message)}}})}function A(){let u=s();if(!u.length){q("Tidak ada data untuk dicetak");return}let $=r==="pagi"?"SHIFT PAGI (08:00 - 17:00)":r==="siang"?"SHIFT SIANG (13:00 - 22:00)":"SELURUH SHIFT (PAGI & SIANG)",p=window.open("","_blank");if(!p){q("Popup diblokir browser, mohon izinkan popup untuk mencetak");return}p.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>STANDAR ACUAN DAC - ${$}</title>
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
          <div class="subtitle">STANDAR RUTINITAS DAILY ACTIVITY CLEANING (DAC) \u2014 ${$}</div>
        </div>
        <table>
          <thead>
            <tr>
              <th style="width:35px; text-align:center;">NO</th>
              ${r==="all"?'<th style="width:55px;">SHIFT</th>':""}
              <th style="width:105px;">JAM</th>
              <th style="width:130px;">KATEGORI</th>
              <th style="width:180px;">LOKASI / AREA</th>
              <th>RINCIAN AKTIVITAS</th>
            </tr>
          </thead>
          <tbody>
            ${u.map(o=>`
              <tr>
                <td style="text-align:center; font-weight:bold;">${o.item_no}</td>
                ${r==="all"?`<td>${o.shift==="pagi"?"Pagi":"Siang"}</td>`:""}
                <td class="jam">${o.start_time} - ${o.end_time}</td>
                <td><strong>${o.category}</strong></td>
                <td><strong>${o.location}</strong></td>
                <td>${o.activity}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
        <div style="margin-top:20px; text-align:right; font-size:9pt; color:#64748b;">
          Dicetak pada: ${new Date().toLocaleString("id-ID")} | Standar Resmi Operasional Facility Care Audy Dental
        </div>
      </body>
      </html>
    `),p.document.close(),p.focus(),setTimeout(()=>{p.print()},400)}document.querySelectorAll(".shift-tab-btn").forEach(u=>{u.addEventListener("click",()=>{document.querySelectorAll(".shift-tab-btn").forEach($=>{$.style.background="#F1F5F9",$.style.color="var(--text-2)"}),u.style.background="#2563EB",u.style.color="#fff",r=u.dataset.shift,d(),F()})});let B=document.getElementById("view-timeline-btn"),T=document.getElementById("view-table-btn");B.addEventListener("click",()=>{m="timeline",B.style.background="#fff",B.classList.remove("btn-ghost"),T.style.background="transparent",T.classList.add("btn-ghost"),F()}),T.addEventListener("click",()=>{m="table",T.style.background="#fff",T.classList.remove("btn-ghost"),B.style.background="transparent",B.classList.add("btn-ghost"),F()}),["dac-search","dac-category-filter","dac-period-filter"].forEach(u=>{let $=document.getElementById(u);$&&($.addEventListener("input",F),$.addEventListener("change",F))}),document.getElementById("btn-add-dac")?.addEventListener("click",()=>{M()}),document.getElementById("btn-print-dac").addEventListener("click",()=>{A()}),document.getElementById("btn-export-dac")?.addEventListener("click",()=>{let u=s();if(!u.length){q("Tidak ada data untuk diexport");return}let $=u.map(p=>({No:p.item_no,Shift:p.shift==="pagi"?"Shift Pagi":"Shift Siang","Jam Mulai":p.start_time,"Jam Selesai":p.end_time,Kategori:p.category,"Lokasi / Area":p.location,"Rincian Aktivitas":p.activity}));ne($,`DAC_Daily_Activity_Cleaning_${r}_${new Date().toISOString().slice(0,10)}`),G("File Excel berhasil diunduh")}),g()}window.parseFlexibleDate=e=>{if(!e||e==="-")return"";if(e=String(e).trim(),/^\d{5}$/.test(e)){let n=Math.floor(Number(e)-25569);return new Date(n*86400*1e3).toISOString().split("T")[0]}if(e.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let n=e.split(/[\/\-]/);return`${n[2]}-${n[1]}-${n[0]}`}let t=e.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(t){let n=t[1],i=parseInt(t[2],10),r=parseInt(t[3],10);if(i>12&&r<=12)return`${n}-${t[3]}-${t[2]}`}return e.split("T")[0]};window.formatDate=e=>{let t=window.parseFlexibleDate(e);if(!t)return"";let n=t.split("-");if(n.length===3&&n[0].length===4){let i=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],r=parseInt(n[2],10),m=i[parseInt(n[1],10)-1];return`${r} ${m} ${n[0]}`}return t};function le(e){return async t=>{if(!tt()){_e("/login");return}return e(t)}}var et=null;function si(){et&&clearInterval(et);let e=()=>{let t=new Date,n=t.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),i=t.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),r=document.getElementById("header-clock-time"),m=document.getElementById("header-clock-date");r&&(r.textContent=n),m&&(m.textContent=i)};e(),et=setInterval(e,1e3)}async function li(){try{let e=await R("/api/dashboard/kpi");if(!e.ok)return;let t=e.data?.data||e.data||{},n=(i,r)=>{let m=document.getElementById(i);m&&(m.textContent=r>0?r:"",m.style.display=r>0?"inline-flex":"none")};n("badge-issues",t.issues?.current||0),n("badge-contracts",t.expiring30?.current||0),n("badge-oo1",t.one_on_one?.current||0),n("badge-schedule",t.schedule?.current||0),n("badge-supply",t.supply?.current||0)}catch{}}var He=[];async function di(){try{let e=await R("/api/dashboard/notifications");if(!e.ok)return;He=e.data?.data||e.data||[];let t=document.getElementById("notif-dot");t&&(t.style.display=He.length>0?"block":"none",t.textContent=He.length)}catch{}}function pi(){if(!He.length){te({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(t,n)=>n()});return}let e=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${He.map(t=>`
        <div class="notif-item notif-severity-${t.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${t.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${t.date}</span>
            <span class="badge badge-${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"info"}">${t.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;te({title:`Notifikasi (${He.length})`,content:e,confirmText:"Tutup",onConfirm:(t,n)=>n()})}function pn(){let e=ae(),t=(e?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
  `;let n=document.getElementById("sidebar"),i=document.getElementById("sidebar-overlay"),r=document.getElementById("topbar-menu-btn"),m=document.getElementById("sidebar-close"),b=()=>{n.classList.add("open"),i.classList.add("show")},g=()=>{n.classList.remove("open"),i.classList.remove("show")};r?.addEventListener("click",b),m?.addEventListener("click",g),i?.addEventListener("click",g),document.querySelectorAll(".nav-item").forEach(d=>d.addEventListener("click",g)),e&&e.role==="input_lembur"&&(document.querySelectorAll(".nav-item").forEach(d=>{d.dataset.route!=="/overtime"&&(d.style.display="none")}),document.querySelectorAll(".nav-section").forEach(d=>{Array.from(d.querySelectorAll(".nav-item")).some(S=>S.style.display!=="none")||(d.style.display="none")}));function s(){let d=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(w=>{let C=w.dataset.route;w.classList.toggle("active",d===C||C!=="/dashboard"&&d.startsWith(C))});let F=document.getElementById("topbar-title"),S=document.querySelector(".nav-item.active .nav-label");F&&S&&(F.textContent=S.textContent)}window.addEventListener("hashchange",s),s(),si(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await R("/api/auth/logout",{method:"POST"}),at(),et&&clearInterval(et),_e("/login")}),li(),di(),document.getElementById("btn-notif")?.addEventListener("click",d=>{d.preventDefault(),pi()})}async function ci(){oe("/login",({main:t})=>ya(t)),oe("/dashboard",le(({main:t})=>ma(t))),oe("/calendar",le(({main:t})=>Ua(t))),oe("/employees",le(({main:t,params:n})=>ra(t,n))),oe("/contracts",le(({main:t,params:n})=>sa(t,n))),oe("/sp",le(({main:t})=>Za(t))),oe("/mutasi",le(({main:t})=>tn(t))),oe("/overtime",le(({main:t})=>rn(t))),oe("/timeline",le(({main:t,params:n})=>aa(t,n))),oe("/os-migrations",le(({main:t})=>sn(t))),oe("/issues",le(({main:t,params:n})=>pa(t,n))),oe("/one-on-one",le(({main:t,params:n})=>ca(t,n))),oe("/training",le(({main:t})=>wa(t))),oe("/relievers",le(({main:t,params:n})=>Sa(t,n))),oe("/reports/inspection",le(({main:t})=>Fa(t))),oe("/reports/cleaning",le(({main:t})=>Ca(t))),oe("/reports/fogging",le(({main:t})=>$a(t))),oe("/reports/basecamp",le(({main:t})=>_a(t))),oe("/reports/supply",le(({main:t})=>Ra(t,"supply"))),oe("/sop",()=>{window.location.hash="#/sop-fc"}),oe("/sop-fc",le(({main:t})=>Da(t))),oe("/sop-security",le(({main:t})=>za(t))),oe("/checklist",le(({main:t})=>Na(t))),oe("/hygiene-standards",le(({main:t})=>qa(t))),oe("/dac",le(({main:t})=>dn(t))),oe("/users",le(({main:t})=>Oa(t))),oe("/branches",le(({main:t})=>Ha(t))),oe("/profile",le(({main:t})=>Ga(t))),oe("/settings/import",le(({main:t})=>{let n=ae();if(!n||n.role!=="superadmin"&&n.role!=="admin"){_e("/dashboard");return}Qa(t)})),oe("/audit-logs",le(({main:t})=>on(t)));let e=tt();if(!e&&window.location.hash!=="#/login"&&_e("/login"),e){let t=await R("/api/auth/me");if(t.ok){nt(t.data.data),pn();let n=ae(),i=window.location.hash;n&&n.role==="input_lembur"&&(i==="#/dashboard"||i===""||i==="#/")&&_e("/overtime")}else at(),_e("/login")}window.addEventListener("fm:login",()=>{pn();let t=ae();t&&t.role==="input_lembur"?_e("/overtime"):_e("/dashboard")}),Xt()}ci();
