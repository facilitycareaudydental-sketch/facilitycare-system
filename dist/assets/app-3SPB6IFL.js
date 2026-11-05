var gn=Object.defineProperty;var Ut=(e,t)=>()=>(e&&(t=e(e=0)),t);var Gt=(e,t)=>{for(var a in t)gn(e,a,{get:t[a],enumerable:!0})};var Wt={};Gt(Wt,{confirmDialog:()=>ie,createModal:()=>X});function X({title:e,content:t,onConfirm:a,onCancel:i,confirmText:s="Simpan",cancelText:b="Batal",size:x="md",confirmClass:g="btn-primary"}){let o={sm:"400px",md:"560px",lg:"720px",xl:"900px"},l=document.createElement("div");l.className="modal-overlay",l.innerHTML=`
    <div class="modal" style="max-width:${o[x]||o.md}">
      <div class="modal-header">
        <h3 class="modal-title">${e}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof t=="string"?t:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${b}</button>
        ${a?`<button class="btn ${g} modal-confirm">${s}</button>`:""}
      </div>
    </div>
  `,t instanceof HTMLElement&&l.querySelector(".modal-body").appendChild(t);let A=()=>{l.classList.remove("show"),setTimeout(()=>l.remove(),250)};return l.querySelector(".modal-close").addEventListener("click",()=>{i&&i(),A()}),l.querySelector(".modal-cancel").addEventListener("click",()=>{i&&i(),A()}),a&&l.querySelector(".modal-confirm").addEventListener("click",()=>a(l,A)),l.addEventListener("click",T=>{T.target===l&&(i&&i(),A())}),document.body.appendChild(l),requestAnimationFrame(()=>l.classList.add("show")),{overlay:l,close:A}}function ie(e,t,a="Konfirmasi"){if(typeof e=="object"&&e!==null){let i=e;return X({title:i.title||"Konfirmasi",content:typeof i.message=="string"?`<p>${i.message}</p>`:i.message,onConfirm:async(s,b)=>{i.onConfirm&&await i.onConfirm(),b()},onCancel:i.onCancel,confirmText:i.confirmText||"Ya, Lanjutkan",cancelText:i.cancelText||"Batal",confirmClass:i.confirmClass||"btn-danger"})}return X({title:a,content:typeof e=="string"?`<p>${e}</p>`:e,onConfirm:async(i,s)=>{t&&await t(),s()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var ce=Ut(()=>{});var fe={};Gt(fe,{downloadExcel:()=>ne,parseExcel:()=>Ue,renderExcelButtons:()=>fn});function Ue(e){return new Promise((t,a)=>{let i=new FileReader;i.onload=s=>{try{let b=new Uint8Array(s.target.result),x=XLSX.read(b,{type:"array"}),g=x.SheetNames[0],o=x.Sheets[g];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${e.name}`),console.log(`File Size: ${(e.size/1024).toFixed(2)} KB`),console.log(`File Type: ${e.type||"unknown"}`),console.log(`Sheets Found: ${x.SheetNames.join(", ")}`),console.log(`Sheet Used: ${g}`);let l=XLSX.utils.decode_range(o["!ref"]||"A1:A1"),A=l.e.r-l.s.r+1,T=l.e.c-l.s.c+1;console.log(`Total Rows (including empty): ${A}`),console.log(`Total Columns: ${T}`);let C=[];for(let z=l.s.c;z<=l.e.c;++z){let B=o[XLSX.utils.encode_cell({c:z,r:l.s.r})];B&&B.v&&C.push(B.v)}console.log(`Headers Found: ${C.join(", ")}`),console.log("---------------------------");let _=XLSX.utils.sheet_to_json(o,{defval:""});Object.defineProperty(_,"__worksheet",{value:o,enumerable:!1}),Object.defineProperty(_,"__headers",{value:C,enumerable:!1}),t(_)}catch(b){a(b)}},i.onerror=s=>a(s),i.readAsArrayBuffer(e)})}function ne(e,t){try{let a=XLSX.utils.json_to_sheet(e),i=XLSX.utils.book_new();XLSX.utils.book_append_sheet(i,a,"Data"),XLSX.writeFile(i,`${t}.xlsx`)}catch(a){throw console.error("Error generating Excel file:",a),a}}function fn(e){return`
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
  `}var de=Ut(()=>{});var mn="https://fm-operations-api.facilitycare-audydental.workers.dev",bn=mn,Pe=1e4;function Xe(){return localStorage.getItem("fm_token")}function Jt(e){localStorage.setItem("fm_token",e)}function Ze(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function Z(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function et(e){localStorage.setItem("fm_user",JSON.stringify(e))}function ee(e,t="read",a=null){if(!e||!e.role)return!1;let i=e.role;if(i==="superadmin"||i==="admin")return!0;let s=null,b="read";if(a!==null){if(s=t,b=a,i==="input_lembur"){if(s==="overtime"){if(b==="read"||b==="write")return!0;if(b==="delete")return!1}else if(b==="read")return!0;return!1}if(i==="editor_khusus"){if(s==="issues"||s==="overtime"){if(b==="read"||b==="write"||b==="delete")return!0}else if(b==="read")return!0;return!1}}else b=t||"read";let g={superadmin:5,admin:4,manager:3,spv:2,editor_khusus:1,input_lembur:1,viewer:1}[i]||0;return b==="read"?g>=1:b==="write"?g>=2:b==="delete"?g>=3:b==="admin"?g>=4:!1}async function N(e,t={}){let a=Xe(),i={"Content-Type":"application/json",...a?{Authorization:`Bearer ${a}`}:{},...t.headers||{}};try{let s=`cb=${Date.now()}`,b=e.includes("?")?"&":"?",x=`${bn}${e}${b}${s}`,g=await fetch(x,{...t,headers:i}),o;try{let l=await g.text();try{o=JSON.parse(l)}catch{o={error:`Server Error (${g.status}): ${l.substring(0,80)}...`}}}catch{o={error:"Gagal membaca respon dari server"}}return g.status===401&&(Ze(),window.location.hash="#/login"),{ok:g.ok,status:g.status,data:o}}catch(s){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${s.message})`}}}}var bt={},tt=null;function oe(e,t){bt[e]=t}function Ae(e){window.location.hash=e}function Yt(){async function e(){let t=window.location.hash.replace("#","")||"/dashboard",[a,...i]=t.split("?"),s=bt[a];if(!s){for(let[x,g]of Object.entries(bt))if(x.endsWith("/*")&&a.startsWith(x.slice(0,-2))){s=g;break}}tt&&(tt(),tt=null);let b=document.getElementById("main-content");if(b&&(b.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),s){let x=new URLSearchParams(i.join("?")),g=a.split("/").filter(Boolean),o=await s({path:a,params:x,segments:g,main:b});o&&(tt=o)}else{let x=b||document.getElementById("app");x&&(x.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",e),e()}var Oe;function hn(){return Oe||(Oe=document.createElement("div"),Oe.id="toast-container",document.body.appendChild(Oe)),Oe}function ht(e,t="info",a=3500){let i=hn(),s=document.createElement("div");s.className=`toast toast-${t}`;let b={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};s.innerHTML=`<span class="toast-icon">${b[t]||"\u2139"}</span><span class="toast-msg">${e}</span>`,i.appendChild(s),requestAnimationFrame(()=>s.classList.add("show")),setTimeout(()=>{s.classList.remove("show"),setTimeout(()=>s.remove(),350)},a)}var H=e=>ht(e,"success"),q=e=>ht(e,"error");var ft=e=>ht(e,"info");ce();ce();function at({columns:e,data:t,onEdit:a,onDelete:i,onView:s,actions:b=[],emptyText:x="Tidak ada data",bulkSelect:g=null}){let o=document.createElement("div");if(o.className="table-wrapper",!t||t.length===0)return o.innerHTML=`<div class="empty-state"><p>${x}</p></div>`,o;let l=document.createElement("table");l.className="data-table";let A=document.createElement("thead"),T=document.createElement("tr");if(g){let _=document.createElement("th");_.style.width="40px",_.style.textAlign="center";let z=document.createElement("input");z.type="checkbox",z.id="select-all-checkbox",z.title="Pilih semua",z.addEventListener("change",()=>{t.forEach(B=>{z.checked?g.selectedIds.add(B.id):g.selectedIds.delete(B.id)}),o.querySelectorAll(".row-checkbox").forEach(B=>B.checked=z.checked),g.onToggle()}),_.appendChild(z),T.appendChild(_)}if(e.forEach(_=>{let z=document.createElement("th");z.textContent=_.label,_.width&&(z.style.width=_.width),T.appendChild(z)}),a||i||s||b.length>0){let _=document.createElement("th");_.textContent="Aksi",_.style.width="120px",T.appendChild(_)}A.appendChild(T),l.appendChild(A);let C=document.createElement("tbody");return t.forEach(_=>{let z=document.createElement("tr");if(g){let B=document.createElement("td");B.style.textAlign="center",B.style.width="40px";let P=document.createElement("input");P.type="checkbox",P.className="row-checkbox",P.checked=g.selectedIds.has(_.id),P.addEventListener("change",()=>{if(P.checked)g.selectedIds.add(_.id);else{g.selectedIds.delete(_.id);let c=document.getElementById("select-all-checkbox");c&&(c.checked=!1)}g.onToggle()}),B.appendChild(P),z.appendChild(B)}if(e.forEach(B=>{let P=document.createElement("td");if(B.render){let c=B.render(_[B.key],_);c instanceof HTMLElement?P.appendChild(c):P.innerHTML=c||""}else P.textContent=_[B.key]!==null&&_[B.key]!==void 0&&_[B.key]!==""?_[B.key]:"";B.nowrap&&(P.style.whiteSpace="nowrap"),z.appendChild(P)}),a||i||s||b.length>0){let B=document.createElement("td");B.className="actions-cell";let P=document.createElement("div");if(P.className="btn-group",s){let c=document.createElement("button");c.className="btn btn-xs btn-ghost",c.innerHTML="\u{1F441}",c.title="Lihat",c.addEventListener("click",()=>s(_)),P.appendChild(c)}if(a){let c=document.createElement("button");c.className="btn btn-xs btn-secondary",c.innerHTML="\u270F\uFE0F",c.title="Edit",c.addEventListener("click",()=>a(_)),P.appendChild(c)}b.forEach(c=>{let y=document.createElement("button");y.className=`btn btn-xs ${c.class||"btn-ghost"}`,y.innerHTML=c.icon||c.label,y.title=c.label,y.addEventListener("click",()=>c.handler(_)),P.appendChild(y)}),B.appendChild(P),z.appendChild(B)}C.appendChild(z)}),l.appendChild(C),o.appendChild(l),o}function nt({page:e,pages:t,total:a,limit:i,onPage:s}){if(t<=1)return null;let b=document.createElement("div");b.className="pagination";let x=document.createElement("span");x.className="pagination-info",x.textContent=`Total: ${a} data`,b.appendChild(x);let g=document.createElement("div");g.className="pagination-btns";let o=(T,C,_=!1,z=!1)=>{let B=document.createElement("button");B.className=`btn btn-sm ${z?"btn-primary":"btn-ghost"} pagination-btn`,B.textContent=T,B.disabled=_,B.addEventListener("click",()=>s(C)),g.appendChild(B)};o("\xAB",1,e===1),o("\u2039",e-1,e===1);let l=Math.max(1,e-2),A=Math.min(t,e+2);for(let T=l;T<=A;T++)o(T,T,!1,T===e);return o("\u203A",e+1,e===t),o("\xBB",t,e===t),b.appendChild(g),b}ce();function He(e){return e.map(t=>{if(t.type==="hidden")return`<input type="hidden" name="${t.name}" value="${t.value||""}">`;if(t.type==="html")return t.html||"";if(t.type==="row")return`<div class="form-row">${He(t.fields)}</div>`;let a=t.required?"required":"",i=t.label?`<label class="form-label">${t.label}${t.required?' <span class="required">*</span>':""}</label>`:"",s="";switch(t.type){case"textarea":s=`<textarea name="${t.name}" class="form-control" placeholder="${t.placeholder||""}" ${a} rows="${t.rows||3}">${t.value||""}</textarea>`;break;case"select":let x=(t.options||[]).map(_=>{let z=typeof _=="object"?_.value:_,B=typeof _=="object"?_.label:_,P=t.value==z?"selected":"";return`<option value="${z}" ${P}>${B}</option>`}).join("");s=`<select name="${t.name}" class="form-control" ${a}><option value="">-- Pilih ${t.label||""} --</option>${x}</select>`;break;case"combobox":let g=`dl-${t.name}-${Math.random().toString(36).substring(7)}`,o=(t.options||[]).map(_=>{let z=typeof _=="object"?_.value:_,B=typeof _=="object"?_.label||_.value||"":_||"";return(B==="undefined"||B==="[object Object]"||B==="null")&&(B=""),B?`<option value="${B}"></option>`:""}).join(""),l=t.value||"";if(t.value){let _=(t.options||[]).find(z=>(typeof z=="object"?z.value:z)==t.value);if(_){let z=typeof _=="object"?_.label||_.value||"":_||"";z&&z!=="undefined"&&z!=="[object Object]"&&z!=="null"&&(l=z)}}s=`
          <input type="text" name="${t.name}" list="${g}" class="form-control" value="${l}" placeholder="Pilih atau ketik baru..." ${a} autocomplete="off">
          <datalist id="${g}">${o}</datalist>
        `;break;case"checkbox":s=`<label class="checkbox-label"><input type="checkbox" name="${t.name}" value="1" ${t.value?"checked":""}> ${t.checkLabel||t.label}</label>`;break;case"date":let A=window.parseFlexibleDate&&t.value?window.parseFlexibleDate(t.value):t.value||"";s=`<input type="date" name="${t.name}" class="form-control" value="${A}" ${a}>`;break;case"month":let T=t.value?String(t.value).slice(0,7):"";s=`<input type="month" name="${t.name}" class="form-control" value="${T}" ${a}>`;break;case"number":let C=t.readonly?"readonly":"";s=`<input type="number" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" min="${t.min||""}" max="${t.max||""}" step="${t.step||"1"}" ${a} ${C}>`;break;case"email":s=`<input type="email" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${a} autocomplete="${t.autocomplete||"off"}">`;break;case"password":s=`<input type="password" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${a} autocomplete="${t.autocomplete||"new-password"}">`;break;case"url":s=`<input type="url" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||"https://..."}" ${a}>`;break;default:s=`<input type="${t.type||"text"}" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${a} autocomplete="${t.autocomplete||"off"}">`}let b=t.hint?`<div class="form-hint">${t.hint}</div>`:"";return`<div class="form-group ${t.class||""}">${i}${s}${b}</div>`}).join("")}function it(e){let t={},a=new FormData(e);for(let[i,s]of a.entries())t[i]=s===""?null:s;return e.querySelectorAll("input[type=checkbox]").forEach(i=>{i.checked||(t[i.name]=null)}),t}function rt(e,t){t&&Object.entries(t).forEach(([a,i])=>{let s=e.querySelector(`[name="${a}"]`);s&&(s.hasAttribute("list")||(s.type==="checkbox"?s.checked=!!i:s.type==="date"&&i&&window.parseFlexibleDate?s.value=window.parseFlexibleDate(i):s.type==="month"&&i?s.value=String(i).slice(0,7):s.value=i??""))})}de();var $e={},ot={on(e,t){$e[e]||($e[e]=new Set),$e[e].add(t)},off(e,t){$e[e]&&$e[e].delete(t)},emit(e,t){$e[e]&&$e[e].forEach(a=>{try{a(t)}catch(i){console.warn("[calendarBus] Handler error:",i)}})},clear(){Object.keys($e).forEach(e=>delete $e[e])}},xn=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function he(e){if(!e){ot.emit("data:changed",{module:"unknown"});return}let t=String(e).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");ot.emit("data:changed",{module:t,relevant:xn.has(t)})}function ue({container:e,title:t,icon:a,apiPath:i,columns:s,formFields:b,filterFields:x,defaultFilters:g={},itemLabel:o="Data",canCreate:l=!0,canEdit:A=!0,canDelete:T=!0,onBeforeSubmit:C,onAfterLoad:_,onDataLoaded:z,extraActions:B=[],initialSearch:P="",exportOptions:c=null,bulkDelete:y=!1,paginationMode:v="server"}){let m=Z();ee(m,"delete")||(T=!1,y=!1),m&&m.role==="viewer"&&(l=!1,A=!1,T=!1,y=!1,c=null),m&&m.role==="editor_khusus"&&i!=="/api/issues"&&i!=="/api/overtime"&&(l=!1,A=!1,T=!1,y=!1,c=null),m&&m.role==="input_lembur"&&(i==="/api/overtime"?(T=!1,y=!1,c=null):(l=!1,A=!1,T=!1,y=!1,c=null));let r=1,h={...g};P&&(h.search=P);let F=new Set;e.innerHTML=`
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
      <h1 class="page-title">${a} ${t}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${l?`<button class="btn btn-primary" id="btn-create">+ Tambah ${o}</button>`:""}
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
    

    ${x&&x.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${x.filter(p=>p.type==="search").map(p=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${p.placeholder||"Cari..."}" id="filter-search" value="${h.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${x.filter(p=>p.type!=="search").map(p=>{if(p.type==="select"||p.type==="combobox"){let S=(p.label||"").startsWith("Pilih")?p.label:`Pilih ${p.label||""}`;return`<select class="filter-select" name="${p.name}" id="filter-${p.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${S}</option>${(p.options||[]).map(u=>`<option value="${typeof u=="object"?u.value:u}" ${h[p.name]===(typeof u=="object"?u.value:u)?"selected":""}>${typeof u=="object"?u.label:u}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${x.filter(p=>p.type!=="search").map(p=>{if(p.type==="select"||p.type==="combobox"){let S=(p.label||"").startsWith("Pilih")?p.label:`Pilih ${p.label||""}`;return`<select class="filter-select filter-select-sheet" name="${p.name}-sheet" id="filter-sheet-${p.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${S}</option>${(p.options||[]).map(u=>`<option value="${typeof u=="object"?u.value:u}" ${h[p.name]===(typeof u=="object"?u.value:u)?"selected":""}>${typeof u=="object"?u.label:u}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function k(){let p=document.getElementById("bulk-toolbar");if(!p)return;let S=document.getElementById("bulk-count"),u=document.getElementById("btn-bulk-delete"),M=document.getElementById("btn-bulk-cancel");S.textContent=`${F.size} item dipilih`,F.size>0?(p.style.display="flex",u.disabled=!1,M.disabled=!1):(p.style.display="none",u.disabled=!0,M.disabled=!0);let R=document.getElementById("select-all-checkbox");if(R){let K=document.querySelectorAll(".row-checkbox");if(K.length>0){let O=[...K].every(Y=>Y.checked),D=[...K].some(Y=>Y.checked);R.checked=O,R.indeterminate=D&&!O}else R.checked=!1,R.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{F.clear(),document.querySelectorAll(".row-checkbox").forEach(S=>S.checked=!1);let p=document.getElementById("select-all-checkbox");p&&(p.checked=!1),k()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(F.size===0)return;let p=[...F],S=document.createElement("div");S.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",S.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${p.length} ${o}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${p.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(S),S.querySelector("#bulk-cancel-btn").addEventListener("click",()=>S.remove()),S.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let u=S.querySelector("#bulk-confirm-btn");u.disabled=!0,u.textContent="Menghapus...";let M=await N(`${i}/bulk`,{method:"DELETE",body:JSON.stringify({ids:p})});S.remove(),M.ok?(H(`${p.length} ${o} berhasil dihapus.`),F.clear(),k(),he(i),L()):q(M.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),n;E?.addEventListener("input",p=>{clearTimeout(n),n=setTimeout(()=>{h.search=p.target.value,r=1,F.clear(),k(),L()},400)}),x?.forEach(p=>{(p.type==="select"||p.type==="combobox")&&(document.getElementById(`filter-${p.name}`)?.addEventListener("change",S=>{h[p.name]=S.target.value;let u=document.getElementById(`filter-sheet-${p.name}`);u&&(u.value=S.target.value),r=1,F.clear(),k(),L()}),document.getElementById(`filter-sheet-${p.name}`)?.addEventListener("change",S=>{h[p.name]=S.target.value;let u=document.getElementById(`filter-${p.name}`);u&&(u.value=S.target.value),r=1,F.clear(),k(),L(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{h={},E&&(E.value=""),x?.forEach(p=>{let S=document.getElementById(`filter-${p.name}`);S&&(S.value="");let u=document.getElementById(`filter-sheet-${p.name}`);u&&(u.value="")}),r=1,F.clear(),k(),L()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{h={},E&&(E.value=""),x?.forEach(p=>{let S=document.getElementById(`filter-${p.name}`);S&&(S.value="");let u=document.getElementById(`filter-sheet-${p.name}`);u&&(u.value="")}),r=1,F.clear(),k(),L(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>j(null)),c&&document.addEventListener("click",function(p){let S=document.getElementById("aksi-menu-main"),u=document.getElementById("btn-aksi-main");S&&u&&!u.contains(p.target)&&!S.contains(p.target)&&S.classList.remove("show-aksi-menu")});let d=document.getElementById("btn-mobile-filter"),f=document.getElementById("filter-options-wrapper"),w=document.getElementById("btn-close-filter-sheet");if(d&&f&&(d.addEventListener("click",p=>{p.preventDefault(),f.classList.add("sheet-open")}),w&&w.addEventListener("click",p=>{p.preventDefault(),f.classList.remove("sheet-open")})),c){document.getElementById(`btn-export-${c.moduleName}`)?.addEventListener("click",async S=>{let u=S.target,M=u.innerHTML;u.innerHTML="\u23F3 Loading...",u.disabled=!0;try{await c.onExport()}catch{q("Gagal export data")}finally{u.innerHTML=M,u.disabled=!1}}),document.getElementById(`btn-template-${c.moduleName}`)?.addEventListener("click",()=>{c.onTemplate()});let p=document.getElementById(`input-import-${c.moduleName}`);p?.addEventListener("change",async S=>{let u=S.target.files[0];if(!u)return;let M=document.getElementById(`label-import-${c.moduleName}`),R=M?M.querySelector(".import-text"):null,K=R?R.innerText:"";R&&(R.innerText="\u231B Memproses..."),M&&(M.style.pointerEvents="none"),p.disabled=!0;try{let O=await Ue(u);if(O.length===0)throw new Error("File kosong atau format salah");await c.onImport(O),H("Import berhasil!"),he(i),L()}catch(O){q(O.message||"Gagal import data")}finally{R&&(R.innerText=K),M&&(M.style.pointerEvents="auto"),p.disabled=!1,p.value=""}})}async function L(){k();let p=document.getElementById("table-container");if(!p)return;p.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let S=v==="client",u=S?1:r,M=S?Pe:20,R=new URLSearchParams({page:u,limit:M,...Object.fromEntries(Object.entries(h).filter(([,U])=>U))}),K=await N(`${i}?${R}`);if(!K.ok){p.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${K.data?.error||"Error"}</p></div>`;return}let O=K.data?.data||K.data||[],D=K.data?.pagination,Y=O.length;if(S){O=z(O);let U=O.length,W=20,Q=Math.ceil(U/W);r>Q&&Q>0&&(r=Q);let V=(r-1)*W,te=r*W;O=O.slice(V,te),D={page:r,limit:W,total:U,pages:Q}}!1,_&&_(O);let J=at({columns:s,data:O,onEdit:A?U=>j(U):null,actions:B.map(U=>({...U,handler:W=>U.handler(W,L)})),emptyText:`Tidak ada ${String(o||"").toLowerCase()}`,bulkSelect:y?{selectedIds:F,onToggle:k}:null});p.innerHTML="",p.appendChild(J);let G=document.getElementById("pagination-container");if(G&&(G.innerHTML="",D&&D.pages>1)){let U=nt({page:D.page,pages:D.pages,total:D.total,limit:D.limit,onPage:W=>{r=W,L()}});U&&G.appendChild(U)}}function I(p){let S=typeof b=="function"?b(p):b;return He(S)}function j(p){let S=!!p,u=document.createElement("form");if(u.noValidate=!0,u.innerHTML=I(p),S){let R=typeof b=="function"?b(p):b;rt(u,p)}let{close:M}=X({title:S?`Edit ${o}`:`Tambah ${o}`,content:u,size:"lg",confirmText:S?"Simpan Perubahan":`Tambah ${o}`,onConfirm:async(R,K)=>{if(!u.reportValidity())return;let O=R.querySelector(".modal-confirm");O.disabled=!0,O.textContent="Menyimpan...";let D=it(u),Y=typeof b=="function"?b(p):b,J=async Q=>{for(let V of Q)if(V.type==="row")await J(V.fields);else if(V.type==="combobox"&&D[V.name]){let te=D[V.name],se=(V.options||[]).find(ae=>{let le=String(typeof ae=="object"?ae.value:ae),Ce=String(typeof ae=="object"?ae.label:ae);return le===te||Ce===te});if(se)D[V.name]=typeof se=="object"?se.value:se;else if(V.createApi){let ae={};ae[V.createApi.field]=te,V.createApi.extra&&Object.assign(ae,V.createApi.extra);let le=await N(V.createApi.path,{method:"POST",body:JSON.stringify(ae)});if(le.ok&&le.data?.id)D[V.name]=le.data.id;else if(le.ok&&!le.data?.id)D[V.name]=te;else throw new Error(`Gagal membuat master data: ${le.data?.error||"Unknown error"}`)}}};try{await J(Y)}catch(Q){q(Q.message),O.disabled=!1,O.textContent=S?"Simpan Perubahan":`Tambah ${o}`;return}C&&(D=await C(D,p));let G=S?"PUT":"POST",U=S?`${i}/${p.id}`:i,W=await N(U,{method:G,body:JSON.stringify(D)});W.ok?(H(S?`${o} berhasil diperbarui.`:`${o} berhasil ditambahkan.`),K(),he(i),L()):(q(W.data?.error||"Gagal menyimpan data."),O.disabled=!1,O.textContent=S?"Simpan Perubahan":`Tambah ${o}`)}})}function $(p){ie(`Hapus ${o} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let S=await N(`${i}/${p.id}`,{method:"DELETE"});S.ok?(H(`${o} berhasil dihapus.`),he(i),L()):q(S.data?.error||"Gagal menghapus.")},`Hapus ${o}`)}return L(),L}var Le=null,st=null;async function Te(e=!1){if(Le&&!e)return console.log("Employees Raw (Cache Hit)",Le.slice(0,5)),Le;let t=await N(`/api/employees?limit=${Pe}&status=Aktif`);return Le=(t.data?.data||[]).map(a=>({value:a.id,label:a.full_name})),console.log("Employees Raw",t.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",Le.slice(0,5)),Le}async function me(e=!1){let a=(await Te(e)).map(i=>({value:i.label,label:i.label}));return console.log("Employee Options",a.slice(0,5)),a}async function re(e=!1){return st&&!e||(st=((await N("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}))),st}function be(e){let t={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!e||e==="-"||String(e).trim()===""?"":`<span class="badge ${t[e]||"badge-neutral"}">${e}</span>`}function yt(e){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[e]||"badge-neutral"}">${e||"-"}</span>`}function ye(e){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[e]||"badge-neutral"}">${e||"-"}</span>`}de();ce();ce();de();ce();var _e={};function Me(e){if(_e[e]){try{_e[e].destroy()}catch{}delete _e[e]}}function yn(){Object.keys(_e).forEach(Me)}var xe=(e,t=0)=>{let a=Number(e);return isNaN(a)||e===null||e===void 0?t:a},Ie=(e,t="\u2014")=>{if(e==null||e==="")return t;let a=String(e).trim();return a===""||a==="[object Object]"?t:a};function lt(e,t,a=900){if(!e)return;e._animFrame&&cancelAnimationFrame(e._animFrame);let i=Math.max(0,Math.round(xe(t))),s=e.textContent?e.textContent.replace(/[^\d]/g,""):"",b=s!==""&&parseInt(s)||0;if(i===b){e.textContent=i.toLocaleString("id-ID"),e._animFrame=null;return}let x=Date.now(),g=()=>{let o=Math.min((Date.now()-x)/a,1),l=1-Math.pow(1-o,3);e.textContent=Math.round(b+l*(i-b)).toLocaleString("id-ID"),o<1?e._animFrame=requestAnimationFrame(g):(e.textContent=i.toLocaleString("id-ID"),e._animFrame=null)};e._animFrame=requestAnimationFrame(g)}var vn={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},kn=e=>{let t=Ie(e,"\u2014");return`<span class="status-pill ${vn[t]||"pill-neutral"}">${t}</span>`};var ke={family:"Inter",size:11},Se="#94A3B8",ze="#F1F5F9",vt=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],wn=()=>window.innerWidth<768;function dt(e={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:wn()?"bottom":"top",labels:{font:ke,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ke,titleFont:{...ke,weight:"700"}}},scales:{x:{grid:{color:ze},ticks:{font:ke,color:Se,maxRotation:0}},y:{grid:{color:ze},ticks:{font:ke,color:Se},beginAtZero:!0}},...e}}var En=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),Sn=()=>Array(5).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Vt(e=3){return Array(e).fill(0).map((t,a)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${a<e-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function ve(e,t,a=8e3){try{let i=new AbortController,s=setTimeout(()=>i.abort(),a),b=await N(e,{signal:i.signal}).catch(()=>null);if(clearTimeout(s),!b||!b.ok)return t;let x=b.data;return x?x.data!==void 0?x.data??t:x:t}catch{return t}}function Fn(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(i=>{let s=document.getElementById(i);s&&(s.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(i=>{let s=document.getElementById(i);if(s&&s.style.display==="none"){s.style.display="block";let b=s.parentElement;if(b&&!b.querySelector(".chart-empty")){let x=document.createElement("div");x.className="chart-empty",x.textContent="Belum ada data",s.style.display="none",b.appendChild(x)}}});let e=document.getElementById("kpi-row");e&&e.querySelector(".skeleton")&&Zt({});let t=document.getElementById("mini-stats-row");t&&t.querySelector(".skeleton")&&ea({}),["table-contracts","table-issues"].forEach(i=>{let s=document.getElementById(i);s&&s.querySelector(".skeleton")&&(s.innerHTML='<div class="chart-empty">Belum ada data</div>')});let a=document.getElementById("activity-log");a&&a.querySelector(".skeleton")&&(a.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Xt(e){yn(),e._dashRefresh&&clearInterval(e._dashRefresh),e._skelTimeout&&clearTimeout(e._skelTimeout);let t=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});e.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">
      
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h2 style="margin:0; font-size:1.2rem; color:var(--text-1);">Dashboard Overview</h2>
        <button id="btn-dash-refresh" class="btn-primary" style="padding:6px 12px; font-size:0.8rem; display:flex; align-items:center; gap:6px; cursor:pointer; background:var(--primary); color:#fff; border:none; border-radius:6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
          Refresh Data
        </button>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${En()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${Sn()}</div>

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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${Vt(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${Vt(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>Qt(e)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async a=>{let i=a.target.value,s=document.getElementById("jadwal-year-label");s&&(s.textContent=i);let b=document.getElementById("skel-jadwal"),x=document.getElementById("chart-jadwal");b&&(b.style.display="block",b.style.position="absolute"),x&&(x.style.display="none");let g=await ve(`/api/dashboard/schedule-chart?year=${i}`,{},8e3);try{ta(g)}catch(o){console.warn("ScheduleChart render:",o),we("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async a=>{let i=a.target.value,s=i?`/api/dashboard/inspection-bar?month=${i}`:"/api/dashboard/inspection-bar",b=document.getElementById("skel-insp"),x=document.getElementById("chart-insp");b&&(b.style.display="block",b.style.position="absolute"),x&&(x.style.display="none");let g=await ve(s,{},8e3);try{aa(g)}catch(o){console.warn("InspBar render:",o),we("skel-insp","chart-insp")}}),e._skelTimeout=setTimeout(()=>Fn(),5e3),await Qt(e)}async function Qt(e){e._skelTimeout&&(clearTimeout(e._skelTimeout),e._skelTimeout=null);let[t,a,i,s,b,x,g,o]=await Promise.all([ve("/api/dashboard/kpi",{},8e3),ve("/api/dashboard/issues-trend",{},8e3),ve("/api/dashboard/issues-summary",{},8e3),ve("/api/dashboard/stats",{},8e3),ve("/api/dashboard/calendar",[],8e3),ve("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),ve(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),ve("/api/dashboard/dropdown-aggregates",{},8e3)]),l=document.getElementById("filter-insp-month"),A=l?l.value:"",T=A?`/api/dashboard/inspection-bar?month=${A}`:"/api/dashboard/inspection-bar",C=await ve(T,{},8e3);window.dashboardAggregates=o||{};try{Zt(t)}catch(_){console.warn("KPI render:",_)}try{ea(t)}catch(_){console.warn("MiniStats render:",_)}try{ta(g)}catch(_){console.warn("ScheduleChart render:",_),we("skel-jadwal","chart-jadwal")}try{Cn(Array.isArray(i?.by_category)?i.by_category:[])}catch(_){console.warn("Donut render:",_),we("skel-donut","chart-donut")}try{$n(a)}catch(_){console.warn("Trend render:",_),we("skel-trend","chart-trend")}try{aa(C)}catch(_){console.warn("InspBar render:",_),we("skel-insp","chart-insp")}try{let _=Array.isArray(s)?s:Array.isArray(s?.recent_issues)?s.recent_issues:[];An(_)}catch(_){console.warn("IssuesTable render:",_)}try{let _=Array.isArray(s?.expiring_contracts)?s.expiring_contracts:[];_n(x)}catch(_){console.warn("ContractsTable render:",_)}try{Tn(Array.isArray(b)?b:[])}catch(_){console.warn("Agenda render:",_)}try{Bn()}catch(_){console.warn("Quick Actions render:",_)}}function Zt(e){let t=document.getElementById("kpi-row");if(!t)return;e=e||{};let a=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,i=window.dashboardAggregates||{},s=o=>{let l=window.dashboardAggregates?.schedule_quarter_details?.[o];return l?`Insp: ${l.insp} \u2022 GC: ${l.gc} \u2022 DC: ${l.dc} \u2022 Fog: ${l.fog}`:""},b=o=>{let l=window.dashboardAggregates?.schedule_quarter_details?.[o];return l?`Rincian ${o}: Inspeksi: ${l.insp} | General Cleaning: ${l.gc} | Deep Cleaning: ${l.dc} | Fogging: ${l.fog}`:`Total Kegiatan ${o}`},x=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
          <option value="Q1" ${a==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${a==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${a==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${a==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:i.schedule_by_quarter?.[a]??e.schedule?.current??0,href:`#/timeline?dash_filter=period_${a.toLowerCase()}`,color:"kpi-blue",sub:s(a),subId:"mini-jadwal-detail",tooltip:b(a)}];t.innerHTML=x.map(o=>{let l=o.val!==void 0?xe(o.val,0):xe(e[o.key]?.current,0);return o.dropdown?`
        <div class="kpi-card ${o.color}" ${o.id?`id="${o.id}"`:""} ${o.tooltip?`title="${o.tooltip}"`:""} data-href="${o.href}" tabindex="0" role="link" style="cursor:pointer;padding:10px 12px">
          <div style="display:flex; gap:10px; align-items:center;">
            <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${o.icon}</span></div>
            <div style="flex:1;min-width:0;">
              <div style="display:flex; align-items:baseline; gap:3px;">
                <div class="kpi-value" data-target="${l}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${l}</div>
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
              <div class="kpi-value" data-target="${l}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${l}</div>
              ${o.dropdown?o.dropdown:""}
            </div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${o.label}</div>
            <div class="kpi-subtitle" ${o.subId?`id="${o.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${o.tooltip||""}">${o.sub}</div>
          </div>
        </div>
      </a>`}).join(""),t.querySelectorAll(".kpi-card[data-href]").forEach(o=>{o.addEventListener("click",l=>{l.target.closest("select")||o.dataset.href&&(window.location.hash=o.dataset.href)}),o.addEventListener("keydown",l=>{l.key==="Enter"&&!l.target.closest("select")&&o.dataset.href&&(window.location.hash=o.dataset.href)})}),t.querySelectorAll(".kpi-value").forEach(o=>{lt(o,parseInt(o.dataset.target)||0)});let g=document.getElementById("dash-jadwal-period");if(g){let o=l=>{let T=(window.dashboardAggregates?.schedule_by_quarter||{})[l]||0,C=document.querySelector("#mini-jadwal .kpi-value");C&&(C.dataset.target=T,lt(C,T,400));let _=document.getElementById("mini-jadwal-detail");_&&(_.textContent=s(l),_.title=b(l));let z=document.getElementById("mini-jadwal");if(z){let B=`#/timeline?dash_filter=period_${l.toLowerCase()}`;z.tagName==="A"?z.href=B:z.dataset.href=B,z.title=b(l)}};g.addEventListener("change",l=>o(l.target.value))}}function ea(e){let t=document.getElementById("mini-stats-row");if(!t)return;e=e||{};let a=new Date().getFullYear(),i=String(new Date().getMonth()+1).padStart(2,"0"),s=`${a}-${i}`,b=l=>`
    <select id="${l}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
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
  `,x=window.dashboardAggregates||{},g=[{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:b("dash-inspeksi-month"),val:x.inspeksi_by_month?.[s]??e.inspection_month?.current??0,href:`#/timeline?dash_filter=inspeksi&month=${s}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:b("dash-gcdc-month"),val:x.gcdc_by_month?.[s]??e.cleaning_month?.current??0,href:`#/timeline?dash_filter=gcdc&month=${s}`,color:"mini-green"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:b("dash-fogging-month"),val:x.fogging_by_month?.[s]??e.fogging_month?.current??0,href:`#/reports/fogging?dash_filter=fogging&month=${s}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:e.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:e.branches?.current,href:"#/branches",color:"mini-teal"}];t.innerHTML=g.map(l=>l.dropdown?`
        <div class="mini-stat ${l.color}" style="cursor:pointer;" id="${l.id||""}" data-href="${l.href}" tabindex="0" role="link">
          <div class="mini-stat-icon">${l.icon}</div>
          <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="mini-stat-value" data-target="${xe(l.val)}">0</div>
              ${l.dropdown}
            </div>
            <div class="mini-stat-text">${l.label}</div>
            ${l.extra||""}
          </div>
        </div>`:`
      <a href="${l.href}" class="mini-stat ${l.color}" style="text-decoration:none" id="${l.id||""}">
        <div class="mini-stat-icon">${l.icon}</div>
        <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
          <div style="display:flex; align-items:baseline; gap:3px;">
            <div class="mini-stat-value" data-target="${xe(l.val)}">0</div>
            ${l.dropdown?l.dropdown:""}
          </div>
          <div class="mini-stat-text">${l.label}</div>
          ${l.extra||""}
        </div>
      </a>`).join(""),t.querySelectorAll(".mini-stat[data-href]").forEach(l=>{l.addEventListener("click",A=>{A.target.closest("select")||l.dataset.href&&(window.location.hash=l.dataset.href)}),l.addEventListener("keydown",A=>{A.key==="Enter"&&!A.target.closest("select")&&l.dataset.href&&(window.location.hash=l.dataset.href)})}),t.querySelectorAll(".mini-stat-value").forEach(l=>lt(l,parseInt(l.dataset.target)||0,700));let o=(l,A,T,C)=>{let _=document.getElementById(l);if(_){let z=B=>{let c=(window.dashboardAggregates?.[T]||{})[B]||0,y=document.querySelector(`#${A} .mini-stat-value`);y&&(y.dataset.target=c,lt(y,c,400));let v=document.getElementById(A);if(v){let m=`${C}&month=${B}`;v.tagName==="A"?v.href=m:v.dataset.href=m}};_.addEventListener("change",B=>z(B.target.value))}};o("dash-inspeksi-month","mini-inspeksi","inspeksi_by_month","#/timeline?dash_filter=inspeksi"),o("dash-gcdc-month","mini-gcdc","gcdc_by_month","#/timeline?dash_filter=gcdc"),o("dash-fogging-month","mini-fogging","fogging_by_month","#/reports/fogging?dash_filter=fogging")}function Cn(e){we("skel-donut","chart-donut");let t=document.getElementById("chart-donut"),a=document.getElementById("donut-legend");if(!t||!a)return;Me("donut");let i=(e||[]).filter(o=>xe(o.count)>0);if(!i.length){Ge(t,"Belum ada data permasalahan");return}let s=i.map(o=>`${Ie(o.category,"Lainnya")}`),b=i.map(o=>xe(o.count)),x=b.reduce((o,l)=>o+l,0);a.innerHTML=i.map((o,l)=>{let A=vt[l%vt.length],T=x>0?Math.round(o.count/x*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${A}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${o.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${T}%)</span></div>
          <div class="donut-legend-label">${s[l]}</div>
        </div>
      </div>
    `}).join("");let g={id:"centerText",beforeDraw:function(o){let l=o.width,A=o.height,T=o.ctx;T.restore();let C=(A/80).toFixed(2);T.font="bold "+C+"em Inter",T.textBaseline="middle",T.fillStyle="#1E293B";let _=x.toString(),z=Math.round((l-T.measureText(_).width)/2),B=A/2;T.fillText(_,z,B-4),T.font="600 "+(C*.35).toFixed(2)+"em Inter",T.fillStyle="#64748B";let P="Total",c=Math.round((l-T.measureText(P).width)/2);T.fillText(P,c,B+10),T.save()}};_e.donut=new Chart(t,{type:"doughnut",data:{labels:s,datasets:[{data:b,backgroundColor:vt,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ke,titleFont:{...ke,weight:"700"},callbacks:{label:o=>` ${o.label}: ${o.parsed} kasus`}}},cutout:"75%"},plugins:[g]})}function $n(e){we("skel-trend","chart-trend");let t=document.getElementById("chart-trend");if(!t)return;Me("trend"),e=e||{};let a=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],i=(e.labels||[]).map(x=>{if(!x||typeof x!="string")return"";try{let[g,o]=x.split("-");return(a[Number(o)-1]||o)+" "+String(g).slice(-2)}catch{return x}}),s=(e.open||[]).map(x=>xe(x)),b=(e.closed||[]).map(x=>xe(x));if(!i.length){Ge(t,"Belum ada data trend");return}_e.trend=new Chart(t,{type:"line",data:{labels:i,datasets:[{label:"Open",data:s,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:b,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:dt({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Se,maxRotation:45,autoSkip:!0}},y:{grid:{color:ze},ticks:{font:{family:"Inter",size:9},color:Se},beginAtZero:!0}}})})}function ta(e){we("skel-jadwal","chart-jadwal");let t=document.getElementById("chart-jadwal");if(!t)return;Me("jadwal"),e=e||{};let a=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(e).some(o=>Array.isArray(o)&&o.some(l=>l>0))){Ge(t,"Belum ada data jadwal");return}let s=e["Inspeksi Hygiene"]||Array(12).fill(0),b=e["General Cleaning"]||Array(12).fill(0),x=e["Deep Cleaning"]||Array(12).fill(0),g=e.Fogging||Array(12).fill(0);_e.jadwal=new Chart(t,{type:"bar",data:{labels:a,datasets:[{label:"Inspeksi",data:s,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:b,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:x,backgroundColor:"#F59E0B"},{label:"Fogging",data:g,backgroundColor:"#EF4444"}]},options:dt({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Se,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:ze},ticks:{font:{family:"Inter",size:9},color:Se},min:0}}})})}function aa(e){we("skel-insp","chart-insp");let t=document.getElementById("chart-insp");if(!t)return;Me("inspBar"),e=e||{};let a=e.labels||[],i=(e.fc||[]).map(b=>xe(b)),s=(e.spv||[]).map(b=>xe(b));if(!a.length){Ge(t,"Belum ada data inspeksi");return}_e.inspBar=new Chart(t,{type:"bar",data:{labels:a,datasets:[{label:"Skor FC",data:i,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:s,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:dt({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ke,color:Se,maxRotation:45,minRotation:45}},y:{grid:{color:ze},ticks:{font:ke,color:Se},min:0,max:100}}})})}function _n(e){we("skel-contract-mini","chart-contract-mini");let t=document.getElementById("chart-contract-mini");if(!t)return;Me("contractMiniBar"),e=e||{};let a={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},i=(e.labels||[]).map(x=>{let g=x.split("-")[1];return a[g]||x}),s=(e.data||[]).map(x=>xe(x));if(!i.length){Ge(t,"Belum ada data");return}let b=t.getContext("2d");_e.contractMiniBar=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Kontrak Habis",data:s,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:dt({onClick:(x,g)=>{if(g&&g.length>0){let o=g[0].index,l=(e.labels||[])[o];l&&(window.location.hash="#/contracts?month_expiry="+l)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ke,color:Se,maxRotation:0,autoSkip:!1}},y:{grid:{color:ze,borderDash:[4,4],drawBorder:!1},ticks:{font:ke,color:Se,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function An(e){let t=document.getElementById("table-issues");if(!t)return;let a=(e||[]).slice(0,8);if(!a.length){t.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}t.innerHTML=`
    <div class="dash-list">
      ${a.map(i=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${kn(i.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Ie(i.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Ie(i.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function Tn(e){let t=document.getElementById("widget-agenda");if(!t)return;let a=new Date,i=`${a.getFullYear()}-${String(a.getMonth()+1).padStart(2,"0")}-${String(a.getDate()).padStart(2,"0")}`,b=(e||[]).filter(x=>(x.event_date||"").startsWith(i)).slice(0,10);if(!b.length){t.innerHTML="";return}t.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${b.map(x=>{let g="#3B82F6",o="#EFF6FF",l="Agenda",A=(x.title||"").toLowerCase();return A.includes("inspeksi")?(g="#10B981",o="#ECFDF5",l="Inspeksi"):A.includes("cleaning")||A.includes("gcdc")?(g="#3B82F6",o="#EFF6FF",l="Cleaning"):A.includes("reliefer")?(g="#F59E0B",o="#FFFBEB",l="Reliefer"):A.includes("fogging")&&(g="#8B5CF6",o="#F5F3FF",l="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(x.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${g};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Ie(x.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Ie(x.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${o};color:${g}">${l}</div>
        </div>
      `}).join("")}
    </div>
  `}function Bn(){let e=document.getElementById("quick-actions");if(!e)return;let t=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];e.innerHTML=t.map(a=>`
    <a href="${a.href}" class="action-btn">
      <div class="action-icon" style="background:${a.bg}">${a.icon}</div>
      ${a.label}
    </a>
  `).join("")}function we(e,t){let a=document.getElementById(e),i=document.getElementById(t);if(a&&(a.style.display="none",a.style.position=""),i){i.style.display="block";let s=i.parentElement;if(s){let b=s.querySelector(".chart-empty");b&&b.remove()}}}function Ge(e,t="Belum ada data"){if(!e)return;e.style.display="none";let a=e.parentElement;if(!a)return;if(!a.querySelector(".chart-empty")){let s=document.createElement("div");s.className="chart-empty",s.textContent=t,a.appendChild(s)}}async function na(e){document.getElementById("app").innerHTML=`
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
  `;let t=document.getElementById("login-form"),a=document.getElementById("login-error"),i=document.getElementById("login-btn"),s=document.getElementById("toggle-password"),b=document.getElementById("login-password");s?.addEventListener("click",()=>{let x=b.type==="text";b.type=x?"password":"text",s.style.color=x?"":"var(--primary)"}),t?.addEventListener("submit",async x=>{x.preventDefault(),a.style.display="none";let g=t.username.value.trim(),o=t.password.value;if(!g||!o){a.textContent="Username dan password wajib diisi.",a.style.display="block";return}i.querySelector(".btn-text").style.display="none",i.querySelector(".btn-spinner").style.display="",i.disabled=!0;try{let l=await N("/api/auth/login",{method:"POST",body:JSON.stringify({username:g,password:o})});l.ok&&l.data.success?(Jt(l.data.data.token),et(l.data.data.user),H("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(a.textContent=l.data.error||"Username atau password salah.",a.style.display="block",i.classList.add("shake"),setTimeout(()=>i.classList.remove("shake"),600))}catch{a.textContent="Gagal terhubung ke server. Periksa koneksi internet.",a.style.display="block"}finally{i.querySelector(".btn-text").style.display="",i.querySelector(".btn-spinner").style.display="none",i.disabled=!1}})}ce();var kt=[{id:1,full_name:"ACHMAD NURYAINI",branch_id:7,branch_name:"007. Bekasi",division:"FACILITY CARE",phone:"081298765431",join_date:"2023-10-09",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:2,full_name:"ADAM JHORDY",branch_id:32,branch_name:"032. Cawang",division:"FACILITY CARE",phone:"085712345678",join_date:"2024-03-18",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:3,full_name:"ADE MUSTIKA REZDAFITRA",branch_id:16,branch_name:"016. Cibubur",division:"FACILITY CARE",phone:"081345678901",join_date:"2022-06-21",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:4,full_name:"ADE SURAHMAN",branch_id:0,branch_name:"000. Management",division:"FACILITY CARE",phone:"081223344556",join_date:"2021-01-15",status:"Aktif",notes:"Trainer & Supervisor Operasional",target_pindah_os:"",target_selesai_os:""},{id:5,full_name:"ADITYA",branch_id:26,branch_name:"026. Tangerang City",division:"FACILITY CARE",phone:"087811223344",join_date:"2024-10-19",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:6,full_name:"ADITYA PRATAMA",branch_id:54,branch_name:"054. Bogor Gunung Batu",division:"FACILITY CARE",phone:"089699887766",join_date:"2025-09-04",status:"Aktif",notes:"Lulusan Training Batch 5 nilai tertinggi",target_pindah_os:"",target_selesai_os:""},{id:7,full_name:"AGUNG SEPTIADI",branch_id:0,branch_name:"000. Management",division:"FC - RELIEFER",phone:"082155667788",join_date:"2026-02-12",status:"Aktif",notes:"Personil cover cabang mobile",target_pindah_os:"",target_selesai_os:""}],ia=["FACILITY CARE","SECURITY","FC - RELIEFER"],Dn=["Aktif","Tidak Aktif","Resign","Cut"];function Ke(e){if(!e)return"-";try{let t=String(e).trim();if(/^\d{5}(\.\d+)?$/.test(t)){let i=new Date(Math.round((Number(t)-25569)*86400*1e3));if(!isNaN(i.getTime()))return i.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}let a=new Date(t);return isNaN(a.getTime())?t:a.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function ra(e){if(!e||e==="-")return"\u{1F464}";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function oa(e,t){let a=[],i=[],s="all",b="all",x="all",g="",o=!1,l=localStorage.getItem("fcms_employees_view_mode")||"card",A=t?t.get("dash_filter"):null;A==="active"?s="Aktif":A==="reliefer"&&(b="FC - RELIEFER");let T=Z(),C=ee(T,"employees","write"),_=ee(T,"employees","write"),z=ee(T,"employees","delete");try{i=await re()}catch(r){console.error("Error fetching branches:",r)}async function B(){o=!0,m();try{let r=await N("/api/employees?limit=10000");if(r.ok){let h=r.data?.data||r.data||[];a=Array.isArray(h)&&h.length>0?h:kt}else a=kt}catch(r){console.warn("Gagal load employees, menggunakan data lokal:",r),a=kt}finally{o=!1,m()}}async function P(){try{let{downloadExcel:r}=await Promise.resolve().then(()=>(de(),fe)),h=a.map((F,k)=>({No:k+1,"Nama Lengkap":F.full_name||"",Cabang:F.branch_name||"",Divisi:F.division||"","No. HP":F.phone||"","Tgl Masuk":F.join_date||"",Status:F.status||"","Target Pindah OS":F.target_pindah_os||"","Target Selesai OS":F.target_selesai_os||"",Catatan:F.notes||""}));r(h,`Data_Karyawan_FCMS_${new Date().toISOString().split("T")[0]}`),H("Berhasil mengekspor data karyawan ke Excel")}catch(r){q("Gagal ekspor Excel: "+r.message)}}function c(r=null){let h=!!r,F=r?.division||"FACILITY CARE",k=r?.status||"Aktif",E=document.createElement("div");E.className="emp-modal-body",E.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Lengkap Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <input type="text" id="emp-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: Achmad Nuryaini" value="${r?.full_name?String(r.full_name).replace(/"/g,"&quot;"):""}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              No. Handphone / WhatsApp
            </label>
            <input type="tel" id="emp-modal-phone" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="08xx-xxxx-xxxx" value="${r?.phone?String(r.phone).replace(/"/g,"&quot;"):""}" />
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${i.map(n=>`
                <option value="${n.value}" ${String(r?.branch_id)===String(n.value)?"selected":""}>
                  ${n.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi / Bagian <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-division" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${ia.map(n=>`
                <option value="${n}" ${F===n?"selected":""}>${n}</option>
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
              value="${r?.join_date||""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${Dn.map(n=>`
                <option value="${n}" ${k===n?"selected":""}>${n}</option>
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
              value="${r?.target_pindah_os?String(r.target_pindah_os).slice(0,7):""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Target Selesai OS (Opsional)
            </label>
            <input type="month" id="emp-modal-target-selesai" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${r?.target_selesai_os?String(r.target_selesai_os).slice(0,7):""}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan / Keterangan Karyawan
          </label>
          <textarea id="emp-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Catatan keahlian, tugas khusus, riwayat cover...">${r?.notes||""}</textarea>
        </div>
      </div>
    `,X({title:h?"\u270F\uFE0F Edit Data Karyawan":"\u2795 Tambah Data Karyawan",content:E,size:"md",confirmText:h?"Simpan Perubahan":"Simpan Karyawan",confirmClass:"btn-primary",onConfirm:async(n,d)=>{let f=(E.querySelector("#emp-modal-name").value||"").trim(),w=(E.querySelector("#emp-modal-phone").value||"").trim(),L=E.querySelector("#emp-modal-branch").value,I=E.querySelector("#emp-modal-division").value,j=E.querySelector("#emp-modal-joindate").value||null,$=E.querySelector("#emp-modal-status").value,p=E.querySelector("#emp-modal-target-pindah").value||null,S=E.querySelector("#emp-modal-target-selesai").value||null,u=(E.querySelector("#emp-modal-notes").value||"").trim();if(!f||!L||!I||!$){q("Harap lengkapi Nama Lengkap, Cabang, Divisi, dan Status!");return}let M={full_name:f,phone:w,branch_id:Number(L),division:I,join_date:j,status:$,target_pindah_os:p,target_selesai_os:S,notes:u};try{let R;h?R=await N(`/api/employees/${r.id}`,{method:"PUT",body:JSON.stringify(M)}):R=await N("/api/employees",{method:"POST",body:JSON.stringify(M)}),R.ok?(H(h?"Data karyawan berhasil diperbarui!":"Karyawan baru berhasil ditambahkan!"),d(),await B()):q(R.data?.error||"Gagal menyimpan data karyawan")}catch(R){q("Terjadi kesalahan: "+R.message)}}})}async function y(r){ie({title:"Nonaktifkan Karyawan?",message:`Apakah Anda yakin ingin menonaktifkan personil <strong>${r.full_name}</strong> dari cabang <em>${r.branch_name||"cabang"}</em>?`,confirmText:"Ya, Nonaktifkan",confirmClass:"btn-danger",onConfirm:async()=>{try{let h=await N(`/api/employees/${r.id}`,{method:"DELETE"});h.ok?(H("Status karyawan berhasil diubah menjadi Tidak Aktif"),await B()):q(h.data?.error||"Gagal menonaktifkan data")}catch(h){q("Terjadi kesalahan: "+h.message)}}})}function v(){return a.filter(r=>{if(g){let h=g.toLowerCase(),F=(r.full_name||"").toLowerCase().includes(h),k=(r.branch_name||"").toLowerCase().includes(h),E=(r.phone||"").toLowerCase().includes(h),n=(r.division||"").toLowerCase().includes(h),d=(r.notes||"").toLowerCase().includes(h);if(!F&&!k&&!E&&!n&&!d)return!1}return!(s!=="all"&&(s==="Aktif"&&r.status!=="Aktif"||s==="Nonaktif"&&r.status==="Aktif")||b!=="all"&&r.division!==b||x!=="all"&&String(r.branch_id)!==String(x))})}function m(){let r=v(),h=a.length,F=a.filter(u=>u.status==="Aktif").length,k=a.filter(u=>u.division==="FACILITY CARE"&&u.status==="Aktif").length,E=a.filter(u=>(u.division==="SECURITY"||u.division==="FC - RELIEFER")&&u.status==="Aktif").length,n=a.length,d=a.filter(u=>u.status==="Aktif").length,f=a.filter(u=>u.status!=="Aktif").length;e.innerHTML=`
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
              ${C?`
                <button id="emp-btn-add" class="btn btn-primary emp-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Karyawan</span>
                </button>
              `:""}
              
              <button id="emp-btn-excel" class="btn btn-ghost emp-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

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
                <div class="emp-stat-val">${h}</div>
                <div class="emp-stat-lbl">Total Karyawan</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">\u{1F7E2}</div>
              <div>
                <div class="emp-stat-val">${F}</div>
                <div class="emp-stat-lbl">Personil Aktif</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(14, 165, 233, 0.2); color:#7DD3FC;">\u{1F3E2}</div>
              <div>
                <div class="emp-stat-val">${k}</div>
                <div class="emp-stat-lbl">Facility Care</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(124, 58, 237, 0.2); color:#C4B5FD;">\u{1F6E1}\uFE0F</div>
              <div>
                <div class="emp-stat-val">${E}</div>
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
              ${i.map(u=>`
                <option value="${u.value}" ${String(x)===String(u.value)?"selected":""}>
                  ${u.label}
                </option>
              `).join("")}
            </select>

            <select id="emp-division-filter" class="emp-select">
              <option value="all">Semua Divisi</option>
              ${ia.map(u=>`
                <option value="${u}" ${b===u?"selected":""}>
                  ${u}
                </option>
              `).join("")}
            </select>

            <div class="emp-view-toggle" title="Pilih Tampilan">
              <button type="button" class="emp-view-btn ${l==="card"?"active":""}" data-view="card">
                <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="emp-view-btn ${l==="table"?"active":""}" data-view="table">
                <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="emp-filter-pills">
            <div class="emp-pill ${s==="all"&&b==="all"?"active":""}" data-status="all" data-div="all">
              <span>Semua Karyawan</span>
              <span class="emp-pill-count">${n}</span>
            </div>
            <div class="emp-pill ${s==="Aktif"?"active":""}" data-status="Aktif" data-div="all">
              <span class="emp-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="emp-pill-count">${d}</span>
            </div>
            <div class="emp-pill ${b==="FC - RELIEFER"?"active":""}" data-status="all" data-div="FC - RELIEFER">
              <span class="emp-pill-dot" style="background:#8B5CF6;"></span>
              <span>FC - Reliefer</span>
              <span class="emp-pill-count">${a.filter(u=>u.division==="FC - RELIEFER").length}</span>
            </div>
            <div class="emp-pill ${s==="Nonaktif"?"active":""}" data-status="Nonaktif" data-div="all">
              <span class="emp-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="emp-pill-count">${f}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive List Container (Card / Table) -->
        <div id="emp-list-container">
          ${r.length===0?`
            <div class="emp-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada karyawan yang cocok</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian, filter cabang, atau filter divisi.</p>
              <button id="emp-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
          `:l==="table"?`
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
                    ${r.map((u,M)=>{let R=u.division||"FACILITY CARE",K="emp-div-fc",O="";R==="SECURITY"?(K="emp-div-sec",O="sec"):R==="FC - RELIEFER"&&(K="emp-div-rel",O="rel");let D=u.status==="Aktif",Y=D?"emp-status-aktif":"emp-status-nonaktif",J=D?"\u{1F7E2} Aktif":`\u26AA ${u.status||"Nonaktif"}`;return`
                        <tr data-id="${u.id}">
                          <td style="text-align:center; font-weight:700; color:#64748B;">${M+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="emp-avatar ${O}" style="width:36px; height:36px; border-radius:9px; font-size:0.85rem;">
                                ${ra(u.full_name)}
                              </div>
                              <div>
                                <div style="font-weight:700; color:#0F172A; white-space:nowrap;">${u.full_name||"-"}</div>
                                <div style="font-size:0.75rem; color:#64748B;">ID #${u.id}</div>
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#334155;">
                            \u{1F4CD} ${u.branch_name||"Cabang Tidak Ditentukan"}
                          </td>
                          <td style="white-space:nowrap;">
                            <span class="emp-div-badge ${K}">
                              ${R==="SECURITY"?"\u{1F6E1}\uFE0F":R==="FC - RELIEFER"?"\u{1F504}":"\u{1F9F9}"} ${R}
                            </span>
                          </td>
                          <td style="white-space:nowrap;">
                            ${u.phone?`
                              <a href="tel:${u.phone}" class="emp-phone-link">
                                \u{1F4DE} ${u.phone}
                              </a>
                            `:'<span style="color:#94A3B8;">-</span>'}
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            ${Ke(u.join_date)}
                          </td>
                          <td style="max-width:260px;">
                            ${u.target_pindah_os||u.target_selesai_os?`
                              <div style="font-size:0.78rem; font-weight:600; color:#4F46E5; margin-bottom:2px;">
                                \u{1F504} ${u.target_pindah_os?"Pindah: "+Ke(u.target_pindah_os):""}
                                ${u.target_selesai_os?(u.target_pindah_os?" \u2192 ":"")+"Selesai: "+Ke(u.target_selesai_os):""}
                              </div>
                            `:""}
                            ${u.notes?`
                              <div style="font-size:0.78rem; color:#475569; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${String(u.notes).replace(/"/g,"&quot;")}">
                                \u{1F4DD} ${u.notes}
                              </div>
                            `:!u.target_pindah_os&&!u.target_selesai_os?'<span style="color:#94A3B8;">-</span>':""}
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="emp-status-badge ${Y}">
                              ${J}
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <div style="display:inline-flex; align-items:center; gap:6px;">
                              ${u.phone?`
                                <a href="https://wa.me/${u.phone.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="emp-icon-btn" title="Chat WhatsApp" style="color:#16A34A;">
                                  <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                                </a>
                              `:""}
                              ${_?`
                                <button class="emp-icon-btn edit" data-action="edit" data-id="${u.id}" title="Edit Data Karyawan">
                                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              `:""}
                              ${z?`
                                <button class="emp-icon-btn delete" data-action="delete" data-id="${u.id}" title="Nonaktifkan Karyawan">
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
              ${r.map(u=>{let M=u.division||"FACILITY CARE",R="emp-div-fc",K="";M==="SECURITY"?(R="emp-div-sec",K="sec"):M==="FC - RELIEFER"&&(R="emp-div-rel",K="rel");let O=u.status==="Aktif",D=O?"emp-status-aktif":"emp-status-nonaktif",Y=O?"\u{1F7E2} Aktif":`\u26AA ${u.status||"Nonaktif"}`;return`
              <div class="emp-card" data-id="${u.id}">
                <!-- Top Badges Row -->
                <div class="emp-card-top">
                  <span class="emp-div-badge ${R}">
                    ${M==="SECURITY"?"\u{1F6E1}\uFE0F":M==="FC - RELIEFER"?"\u{1F504}":"\u{1F9F9}"} ${M}
                  </span>
                  <div style="display:flex; align-items:center; gap:8px;">
                    <span class="emp-status-badge ${D}">
                      ${Y}
                    </span>
                  </div>
                </div>

                <!-- Profile Box -->
                <div class="emp-profile-box">
                  <div class="emp-avatar ${K}">
                    ${ra(u.full_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="emp-name" style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${u.full_name||"-"}
                    </div>
                    <div class="emp-subtitle">
                      Personil ${M}
                    </div>
                  </div>
                </div>

                <!-- Information Grid -->
                <div class="emp-info-card">
                  <div class="emp-info-row">
                    <span class="emp-info-label">\u{1F4CD} Cabang:</span>
                    <span class="emp-info-val">${u.branch_name||"Cabang Tidak Ditentukan"}</span>
                  </div>
                  <div class="emp-info-row">
                    <span class="emp-info-label">\u{1F4DE} No. HP:</span>
                    <span class="emp-info-val">
                      ${u.phone?`
                        <a href="tel:${u.phone}" class="emp-phone-link">
                          <span>${u.phone}</span>
                        </a>
                      `:'<span style="color:#94A3B8;">-</span>'}
                    </span>
                  </div>
                  <div class="emp-info-row">
                    <span class="emp-info-label">\u{1F5D3}\uFE0F Tgl Masuk:</span>
                    <span class="emp-info-val">${Ke(u.join_date)}</span>
                  </div>
                  ${u.target_pindah_os||u.target_selesai_os?`
                    <div class="emp-info-row target-os" style="border-top:1px dashed #E2E8F0; padding-top:6px; margin-top:2px; flex-direction:column; align-items:flex-start; gap:3px; width:100%;">
                      <span class="emp-info-label" style="font-weight:600; color:#475569;">\u{1F504} Target Migrasi OS:</span>
                      <span class="emp-info-val" style="font-size:0.78rem; color:#4F46E5; text-align:left; width:100%; word-break:break-word; overflow-wrap:anywhere;">
                        ${u.target_pindah_os?"Pindah: "+Ke(u.target_pindah_os):""}
                        ${u.target_selesai_os?(u.target_pindah_os?" \u2192 ":"")+"Selesai: "+Ke(u.target_selesai_os):""}
                      </span>
                    </div>
                  `:""}
                </div>

                <!-- Notes box (if any) -->
                ${u.notes?`
                  <div class="emp-notes-box">
                    \u{1F4DD} <strong>Catatan:</strong> ${u.notes}
                  </div>
                `:""}

                <!-- Footer Actions -->
                <div class="emp-card-footer">
                  <div style="font-size:0.75rem; color:#94A3B8;">
                    ID #${u.id}
                  </div>

                  <div class="emp-action-btns">
                    ${u.phone?`
                      <a href="https://wa.me/${u.phone.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="emp-icon-btn" title="Chat WhatsApp" style="color:#16A34A;">
                        <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                      </a>
                    `:""}

                    ${_?`
                      <button class="emp-icon-btn edit" data-action="edit" data-id="${u.id}" title="Edit Data Karyawan">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                    `:""}
                    ${z?`
                      <button class="emp-icon-btn delete" data-action="delete" data-id="${u.id}" title="Nonaktifkan Karyawan">
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
    `,e.querySelectorAll(".emp-view-btn").forEach(u=>{u.addEventListener("click",()=>{let M=u.getAttribute("data-view");M&&M!==l&&(l=M,localStorage.setItem("fcms_employees_view_mode",M),m())})});let w=e.querySelector("#emp-search-input");w&&w.addEventListener("input",u=>{g=u.target.value,m();let M=e.querySelector("#emp-search-input");M&&(M.focus(),M.setSelectionRange(M.value.length,M.value.length))});let L=e.querySelector("#emp-branch-filter");L&&L.addEventListener("change",u=>{x=u.target.value,m()});let I=e.querySelector("#emp-division-filter");I&&I.addEventListener("change",u=>{b=u.target.value,m()}),e.querySelectorAll(".emp-pill").forEach(u=>{u.addEventListener("click",()=>{s=u.dataset.status;let M=u.dataset.div;M&&M!=="all"?b=M:u.dataset.status==="all"&&(b="all"),m()})});let j=e.querySelector("#emp-reset-filter");j&&j.addEventListener("click",()=>{g="",x="all",b="all",s="all",m()});let $=e.querySelector("#emp-btn-add");$&&$.addEventListener("click",()=>c());let p=e.querySelector("#emp-btn-excel");p&&p.addEventListener("click",()=>P());let S=e.querySelector("#emp-btn-refresh");S&&S.addEventListener("click",()=>B()),e.querySelectorAll(".emp-icon-btn.edit").forEach(u=>{u.addEventListener("click",M=>{let R=Number(M.currentTarget.getAttribute("data-id")),K=a.find(O=>O.id===R);K&&c(K)})}),e.querySelectorAll(".emp-icon-btn.delete").forEach(u=>{u.addEventListener("click",M=>{let R=Number(M.currentTarget.getAttribute("data-id")),K=a.find(O=>O.id===R);K&&y(K)})})}await B()}ce();de();function Je(e){if(!e||e==="-")return"-";if(String(e).startsWith("2099"))return"Tetap / PKWTT";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Fe(e){return e?String(e).startsWith("2099"):!1}function sa(e){if(!e||e==="-")return"\u{1F464}";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function la(e,t){let a=[],i=[],s=[],b={},x="",g="all",o="all",l="all",A="all",T=!1,C=localStorage.getItem("fcms_contracts_view_mode")||"card",_=t?t.get("dash_filter"):null;_==="expiring30"?(A="30",l="Aktif"):_==="active"&&(l="Aktif");let z=Z(),B=ee(z,"contracts","write"),P=ee(z,"contracts","write"),c=ee(z,"contracts","delete");try{let[n,d]=await Promise.all([re(),Te()]);i=Array.isArray(n)?n:[],s=Array.isArray(d)?d:[],s.forEach(f=>{let w=f.value!==void 0?f.value:f.id,L=f.label||f.full_name;w&&f.phone&&(b[w]=f.phone),L&&f.phone&&(b[L.toLowerCase()]=f.phone)})}catch(n){console.error("Error fetching cache options for contracts:",n)}async function y(){T=!0,E();try{let n=await N("/api/contracts?limit=10000");if(n.ok){let d=n.data?.data||n.data||[];a=Array.isArray(d)?d:[]}else a=[],q("Gagal memuat data kontrak dari server.")}catch(n){console.error("Error load contracts:",n),a=[],q("Terjadi kesalahan koneksi saat memuat kontrak.")}finally{T=!1,E()}}function v(){return a.filter(n=>{if(x.trim()){let d=x.toLowerCase().trim(),f=String(n.employee_name||"").toLowerCase(),w=String(n.branch_name||"").toLowerCase(),L=String(n.pkwt_number||"").toLowerCase(),I=String(n.contract_type||"").toLowerCase(),j=String(n.notes||"").toLowerCase();if(!(f.includes(d)||w.includes(d)||L.includes(d)||I.includes(d)||j.includes(d)))return!1}if(g!=="all"&&String(n.branch_id)!==String(g)||o!=="all"&&n.division!==o||l==="Aktif"&&n.status!=="Aktif")return!1;if(l==="Expiring30"){if(!!!(n.end_date&&String(n.end_date).trim()&&n.end_date!=="-"))return!1;let f=Fe(n.end_date),w=n.days_remaining!==null&&n.days_remaining!==void 0&&n.days_remaining!==""?Number(n.days_remaining):null;if(n.status!=="Aktif"||f||w===null||isNaN(w)||w<0||w>30)return!1}if(l==="Tetap"&&(!Fe(n.end_date)||n.status!=="Aktif")||l==="Nonaktif"&&n.status==="Aktif")return!1;if(A!=="all"){if(!!!(n.end_date&&String(n.end_date).trim()&&n.end_date!=="-"))return!1;let f=parseInt(A,10),w=Fe(n.end_date),L=n.days_remaining!==null&&n.days_remaining!==void 0&&n.days_remaining!==""?Number(n.days_remaining):null;if(w||n.status!=="Aktif"||L===null||isNaN(L)||L<0||L>f)return!1}return!0})}async function m(){X({title:"\u{1F50D} Analisis Selisih Kontrak Karyawan",content:`
        <div style="text-align:center; padding:24px 10px;">
          <div style="font-size:2rem; margin-bottom:10px;">\u23F3</div>
          <p style="color:#64748B; font-size:0.9rem;">Sedang menyinkronkan data karyawan aktif dengan kontrak aktif di database...</p>
        </div>
      `,cancelText:"Tutup"});try{let[n,d]=await Promise.all([N("/api/employees?limit=10000&status=Aktif"),N("/api/contracts?limit=10000")]),f=n.ok?n.data?.data||n.data||[]:[],w=d.ok?d.data?.data||d.data||[]:[],L=new Set(w.filter($=>$.status==="Aktif").map($=>$.employee_id)),I=f.filter($=>!L.has($.id)),j="";I.length===0?j=`
          <div style="text-align:center; padding:20px 10px;">
            <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F389}</div>
            <h4 style="font-size:1.1rem; font-weight:800; color:#10B981; margin:0 0 8px 0;">Semua Karyawan Aktif Memiliki Kontrak!</h4>
            <p style="font-size:0.86rem; color:#64748B; margin:0;">
              Total <strong>${f.length}</strong> karyawan aktif tercatat memiliki perikatan kontrak aktif yang sah di sistem.
            </p>
          </div>
        `:j=`
          <div style="display:flex; flex-direction:column; gap:12px;">
            <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:10px; padding:12px; font-size:0.84rem; color:#991B1B;">
              Ditemukan <strong>${I.length}</strong> dari <strong>${f.length}</strong> karyawan aktif yang <strong>belum memiliki Kontrak Aktif</strong>.
            </div>
            <div style="max-height:360px; overflow-y:auto; display:flex; flex-direction:column; gap:8px; padding-right:4px;">
              ${I.map($=>{let p=w.filter(u=>u.employee_id===$.id),S='<span style="color:#D97706; font-weight:600;">Belum pernah ada data kontrak</span>';if(p.length>0){let u=p[0];S=`Riwayat: Status <strong>${u.status}</strong> (Selesai: ${Je(u.end_date)})`}return`
                  <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:8px; padding:10px 12px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">
                    <div>
                      <div style="font-weight:700; font-size:0.88rem; color:#0F172A;">${$.full_name}</div>
                      <div style="font-size:0.76rem; color:#64748B;">
                        Cabang: ${$.branch_name||"-"} \u2022 Divisi: ${$.division||"FACILITY CARE"}
                      </div>
                      <div style="font-size:0.75rem; color:#475569; margin-top:2px;">
                        ${S}
                      </div>
                    </div>
                    ${P?`
                      <button class="btn btn-primary btn-add-contract-for-emp" data-empid="${$.id}" data-empname="${String($.full_name).replace(/"/g,"&quot;")}" data-branchid="${$.branch_id||""}" data-division="${$.division||"FACILITY CARE"}" style="font-size:0.76rem; padding:5px 10px; border-radius:6px; background:#2563EB;">
                        + Buat Kontrak
                      </button>
                    `:""}
                  </div>
                `}).join("")}
            </div>
          </div>
        `,X({title:"Analisis Selisih Kontrak Karyawan",content:j,cancelText:"Tutup"}),setTimeout(()=>{document.querySelectorAll(".btn-add-contract-for-emp").forEach($=>{$.addEventListener("click",p=>{let S=p.currentTarget,u=S.getAttribute("data-empid"),M=S.getAttribute("data-empname"),R=S.getAttribute("data-branchid"),K=S.getAttribute("data-division"),O=document.querySelector(".modal-close");O&&O.click(),h({employee_id:u,employee_name:M,branch_id:R,division:K,status:"Aktif"})})})},100)}catch(n){console.error(n),q("Gagal melakukan analisis selisih data kontrak.")}}function r(n,d){if(!n)return"";let f=String(n).split("-").map(Number);if(f.length!==3||!f[0]||!f[1]||!f[2])return"";let w=new Date(f[0],f[1]-1,f[2]);if(isNaN(w.getTime()))return"";if(d==="KONTRAK 6 BULAN")w.setMonth(w.getMonth()+6),w.setDate(w.getDate()-1);else if(d==="KONTRAK 1 TAHUN")w.setFullYear(w.getFullYear()+1),w.setDate(w.getDate()-1);else if(d==="KONTRAK 2 TAHUN")w.setFullYear(w.getFullYear()+2),w.setDate(w.getDate()-1);else return"";let L=w.getFullYear(),I=String(w.getMonth()+1).padStart(2,"0"),j=String(w.getDate()).padStart(2,"0");return`${L}-${I}-${j}`}function h(n=null){let d=!!(n&&n.id),f=d?`\u270F\uFE0F Edit Kontrak: ${n.employee_name||"Karyawan"}`:"\u2795 Tambah Kontrak Karyawan Baru",w=n?.employee_id||"",L=n?.branch_id||"",I=n?.division||"FACILITY CARE",j=n?.status||"Aktif",$=n?.contract_type||"KONTRAK 1 TAHUN",p=n?.pkwt_number||"PKWT 1",S=n?.notes||"",u=!1;n&&(n.contract_type==="PKWTT (TETAP)"||n.pkwt_number==="TETAP"?u=!0:n.contract_type&&(n.contract_type.startsWith("KONTRAK")||n.contract_type.includes("BULAN")||n.contract_type.includes("TAHUN"))||n.pkwt_number&&n.pkwt_number.startsWith("PKWT")?u=!1:u=Fe(n.end_date));let M=n?.start_date?String(n.start_date).slice(0,10):"",R=n?.end_date&&!Fe(n.end_date)?String(n.end_date).slice(0,10):"";!u&&M&&!R&&(R=r(M,$));let K=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-employee" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Karyawan --</option>
              ${s.map(D=>{let Y=D.value!==void 0?D.value:D.id,J=D.label||D.full_name||"Karyawan",G=D.division||D.branch_name?` (${D.division||"FC"} - ${D.branch_name||"Cabang -"})`:"";return`
                <option value="${Y}" ${String(w)===String(Y)?"selected":""}>
                  ${J}${G}
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
              ${i.map(D=>`
                <option value="${D.value}" ${String(L)===String(D.value)?"selected":""}>
                  ${D.label}
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
              <option value="FACILITY CARE" ${I==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${I==="SECURITY"?"selected":""}>SECURITY</option>
              <option value="FC - RELIEFER" ${I==="FC - RELIEFER"?"selected":""}>FC - RELIEFER</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Kontrak <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Aktif" ${j==="Aktif"?"selected":""}>\u{1F7E2} Aktif</option>
              <option value="Tidak Aktif" ${j==="Tidak Aktif"?"selected":""}>\u26AA Tidak Aktif</option>
              <option value="Resign" ${j==="Resign"?"selected":""}>\u{1F534} Resign</option>
              <option value="Cut" ${j==="Cut"?"selected":""}>\u2702\uFE0F Cut / Pemutusan</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tipe / Durasi Kontrak
            </label>
            <select id="con-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="KONTRAK 6 BULAN" ${$==="KONTRAK 6 BULAN"?"selected":""}>KONTRAK 6 BULAN</option>
              <option value="KONTRAK 1 TAHUN" ${$==="KONTRAK 1 TAHUN"?"selected":""}>KONTRAK 1 TAHUN</option>
              <option value="KONTRAK 2 TAHUN" ${$==="KONTRAK 2 TAHUN"?"selected":""}>KONTRAK 2 TAHUN</option>
              <option value="PKWTT (TETAP)" ${$==="PKWTT (TETAP)"?"selected":""}>PKWTT (TETAP)</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tingkat / No. PKWT
            </label>
            <select id="con-modal-pkwt" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="PKWT 1" ${p==="PKWT 1"?"selected":""}>PKWT 1</option>
              <option value="PKWT 2" ${p==="PKWT 2"?"selected":""}>PKWT 2</option>
              <option value="PKWT 3" ${p==="PKWT 3"?"selected":""}>PKWT 3</option>
              <option value="PKWT 4" ${p==="PKWT 4"?"selected":""}>PKWT 4</option>
              <option value="PKWT 5" ${p==="PKWT 5"?"selected":""}>PKWT 5</option>
              <option value="PKWT 6" ${p==="PKWT 6"?"selected":""}>PKWT 6</option>
              <option value="TETAP" ${p==="TETAP"?"selected":""}>PKWTT / TETAP</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Mulai Kontrak
            </label>
            <input type="date" id="con-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${M}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Selesai Kontrak
            </label>
            <input type="date" id="con-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; ${u?"opacity:0.5; background-color:#F8FAFC; cursor:not-allowed;":"background-color:#FFFFFF;"}" 
              value="${R}" ${u?"disabled":""} />
            <div style="margin-top:6px; display:flex; align-items:center; gap:6px;">
              <input type="checkbox" id="con-modal-is-tetap" ${u?"checked":""} style="cursor:pointer;" />
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
            placeholder="Keterangan perpanjangan, klausul khusus, atau nomor SK...">${S}</textarea>
        </div>
      </div>
    `,O=X({title:f,content:K,confirmText:d?"\u{1F4BE} Simpan Perubahan":"\u2795 Tambah Kontrak",cancelText:"Batal",onConfirm:async(D,Y)=>{let J=D.querySelector("#con-modal-employee"),G=D.querySelector("#con-modal-branch"),U=D.querySelector("#con-modal-division"),W=D.querySelector("#con-modal-status"),Q=D.querySelector("#con-modal-start"),V=D.querySelector("#con-modal-end"),te=D.querySelector("#con-modal-is-tetap"),se=D.querySelector("#con-modal-type"),ae=D.querySelector("#con-modal-pkwt"),le=D.querySelector("#con-modal-notes"),Ce=J?.value,Ot=G?.value,rn=U?.value||"FACILITY CARE",on=W?.value||"Aktif",ct=Q?.value?.trim()||null,Ht=te?.checked,sn=V?.value?.trim(),ut=Ht?"2099-12-31":sn||null,ln=se?.value,dn=ae?.value,pn=le?.value?.trim()||null;if(!Ce){q("Harap pilih nama karyawan!"),J?.focus();return}if(!Ot){q("Harap pilih cabang penempatan!"),G?.focus();return}if(!Ht&&ct&&ut&&ut<ct){q("Tanggal selesai kontrak tidak boleh lebih awal dari tanggal mulai kontrak!"),V?.focus();return}let cn={employee_id:parseInt(Ce,10),branch_id:parseInt(Ot,10),division:rn,status:on,start_date:ct,end_date:ut,contract_type:ln,pkwt_number:dn,notes:pn};try{let gt=d?`/api/contracts/${n.id}`:"/api/contracts",mt=await N(gt,{method:d?"PUT":"POST",body:JSON.stringify(cn)});if(mt.ok)H(d?"Data kontrak berhasil diperbarui!":"Kontrak baru berhasil ditambahkan!"),Y(),await y();else{let un=mt.data?.error||mt.data?.message||"Gagal menyimpan data kontrak.";q(un)}}catch(gt){console.error(gt),q("Terjadi kesalahan koneksi ke server.")}}});setTimeout(()=>{let D=O?.overlay||document.querySelector(".modal-overlay:last-child");if(!D)return;let Y=D.querySelector("#con-modal-is-tetap"),J=D.querySelector("#con-modal-end"),G=D.querySelector("#con-modal-start"),U=D.querySelector("#con-modal-type"),W=D.querySelector("#con-modal-pkwt");if(!Y||!J)return;function Q(V,te=!0){V?(Y.checked=!0,J.disabled=!0,J.style.opacity="0.5",J.style.backgroundColor="#F8FAFC",J.style.cursor="not-allowed",J.value="",te&&(U&&(U.value="PKWTT (TETAP)"),W&&(W.value="TETAP"))):(Y.checked=!1,J.disabled=!1,J.style.opacity="1",J.style.backgroundColor="#FFFFFF",J.style.cursor="default",te&&(U&&U.value==="PKWTT (TETAP)"&&(U.value="KONTRAK 1 TAHUN"),W&&W.value==="TETAP"&&(W.value="PKWT 1")))}Y.addEventListener("change",V=>{Q(V.target.checked,!0),V.target.checked||J.focus()}),U&&U.addEventListener("change",V=>{let te=V.target.value;te==="PKWTT (TETAP)"?(Q(!0,!1),W&&(W.value="TETAP")):(Q(!1,!1),W&&W.value==="TETAP"&&(W.value="PKWT 1"),G&&G.value&&(J.value=r(G.value,te)))}),W&&W.addEventListener("change",V=>{V.target.value==="TETAP"?(U&&(U.value="PKWTT (TETAP)"),Q(!0,!1)):(U&&U.value==="PKWTT (TETAP)"&&(U.value="KONTRAK 1 TAHUN"),Q(!1,!1),G&&G.value&&U&&(J.value=r(G.value,U.value)))}),G&&G.addEventListener("change",V=>{!Y.checked&&V.target.value&&U&&(J.value=r(V.target.value,U.value))})},20)}function F(n){X({title:"\u{1F5D1}\uFE0F Hapus Data Kontrak",content:`
        <div style="font-size:0.92rem; color:#334155; line-height:1.5;">
          Apakah Anda yakin ingin menghapus catatan kontrak untuk <strong>${n.employee_name||"karyawan ini"}</strong> (ID #${n.id})?
          <div style="margin-top:10px; padding:10px; background:#FEF2F2; border-radius:8px; border:1px solid #FECACA; font-size:0.8rem; color:#991B1B;">
            \u26A0\uFE0F Tindakan ini akan menghapus riwayat masa berlaku kontrak terkait dari sistem.
          </div>
        </div>
      `,confirmText:"Ya, Hapus",cancelText:"Batal",onConfirm:async(d,f)=>{try{let w=await N(`/api/contracts/${n.id}`,{method:"DELETE"});w.ok?(H("Data kontrak berhasil dihapus."),f(),await y()):q("Gagal menghapus kontrak: "+(w.data?.message||"Terjadi kesalahan"))}catch(w){console.error(w),q("Terjadi kendala koneksi.")}}})}function k(){if(!a.length){q("Belum ada data kontrak untuk diekspor.");return}let n=a.map(d=>{let f=Fe(d.end_date),w=!!(d.end_date&&String(d.end_date).trim()&&d.end_date!=="-"),L=d.days_remaining!==null&&d.days_remaining!==void 0&&d.days_remaining!==""?Number(d.days_remaining):null;return{"Nama Lengkap":d.employee_name||"",Cabang:d.branch_name||"",Divisi:d.division||"","No. PKWT":d.pkwt_number||"","Tipe Kontrak":d.contract_type||"","Tanggal Mulai":d.start_date||"","Tanggal Selesai":f?"Tetap / PKWTT":w?d.end_date:"-","Sisa Hari":f?"Tetap":w&&L!==null&&!isNaN(L)?L+" Hari":"-",Status:d.status||"",Catatan:d.notes||""}});ne(n,`Data_Kontrak_Karyawan_${new Date().toISOString().slice(0,10)}`),H("File Excel Data Kontrak berhasil diunduh!")}function E(){let n=v(),d=a.length,f=a.filter(D=>D.status==="Aktif").length,w=a.filter(D=>{if(!!!(D.end_date&&String(D.end_date).trim()&&D.end_date!=="-"))return!1;let J=Fe(D.end_date),G=D.days_remaining!==null&&D.days_remaining!==void 0&&D.days_remaining!==""?Number(D.days_remaining):null;return D.status==="Aktif"&&!J&&G!==null&&!isNaN(G)&&G>=0&&G<=30}).length,L=a.filter(D=>Fe(D.end_date)&&D.status==="Aktif").length,I=a.filter(D=>D.status!=="Aktif").length;e.innerHTML=`
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
              ${P?`
                <button id="con-btn-add" class="btn btn-primary con-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Kontrak</span>
                </button>
              `:""}

              <button id="con-btn-excel" class="btn btn-ghost con-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="con-btn-missing" class="btn btn-ghost con-btn-missing" style="background:rgba(239,68,68,0.15); border:1px solid rgba(239,68,68,0.4); color:#FCA5A5; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;" title="Cek Karyawan Aktif yang belum memiliki Kontrak Aktif">
                <span>\u{1F50D}</span>
                <span>Cek Selisih</span>
              </button>

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
                <div class="con-stat-val" style="color:#34D399;">${f}</div>
                <div class="con-stat-lbl">Kontrak Aktif</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FBBF24;">
                \u26A0\uFE0F
              </div>
              <div>
                <div class="con-stat-val" style="color:#FBBF24;">${w}</div>
                <div class="con-stat-lbl">Habis &lt; 30 Hari</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(99, 102, 241, 0.2); color:#A5B4FC;">
                \u{1F3DB}\uFE0F
              </div>
              <div>
                <div class="con-stat-val" style="color:#A5B4FC;">${L}</div>
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
                placeholder="Cari nama karyawan, cabang, no PKWT, catatan..." value="${x?String(x).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="con-branch-filter" class="con-select">
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(D=>`
                <option value="${D.value}" ${String(g)===String(D.value)?"selected":""}>
                  ${D.label}
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
              <option value="all" ${A==="all"?"selected":""}>Semua Durasi Masa Habis</option>
              <option value="7" ${A==="7"?"selected":""}>Akan Habis \u2264 7 Hari</option>
              <option value="14" ${A==="14"?"selected":""}>Akan Habis \u2264 14 Hari</option>
              <option value="30" ${A==="30"?"selected":""}>Akan Habis \u2264 30 Hari</option>
              <option value="60" ${A==="60"?"selected":""}>Akan Habis \u2264 60 Hari</option>
            </select>

            <div class="con-view-toggle">
              <button type="button" class="con-view-btn ${C==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="con-view-btn ${C==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="con-filter-pills">
            <div class="con-pill ${l==="all"&&A==="all"?"active":""}" data-status="all">
              <span>Semua Kontrak</span>
              <span class="con-pill-count">${d}</span>
            </div>
            <div class="con-pill ${l==="Aktif"&&A==="all"?"active":""}" data-status="Aktif">
              <span class="con-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="con-pill-count">${f}</span>
            </div>
            <div class="con-pill ${l==="Expiring30"?"active":""}" data-status="Expiring30">
              <span class="con-pill-dot" style="background:#F59E0B;"></span>
              <span>Akan Habis (&lt;30 Hari)</span>
              <span class="con-pill-count">${w}</span>
            </div>
            <div class="con-pill ${l==="Tetap"?"active":""}" data-status="Tetap">
              <span class="con-pill-dot" style="background:#6366F1;"></span>
              <span>PKWTT (Tetap)</span>
              <span class="con-pill-count">${L}</span>
            </div>
            <div class="con-pill ${l==="Nonaktif"?"active":""}" data-status="Nonaktif">
              <span class="con-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="con-pill-count">${I}</span>
            </div>
          </div>
        </div>

        <!-- Content Area: Pilihan Kartu atau Tabel -->
        <div id="con-list-container">
          ${n.length===0?`
            <div class="con-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada kontrak yang sesuai kriteria</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba atur ulang kata kunci pencarian, filter cabang, atau filter masa berlaku kontrak.</p>
              <button id="con-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
          `:C==="table"?`
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
                    ${n.map((D,Y)=>{let J=D.division||"FACILITY CARE",G="con-div-fc",U="";J==="SECURITY"?(G="con-div-sec",U="sec"):J==="FC - RELIEFER"&&(G="con-div-rel",U="rel");let W=!!(D.end_date&&String(D.end_date).trim()&&D.end_date!=="-"),Q=Fe(D.end_date),V=D.status==="Aktif",te=D.days_remaining!==null&&D.days_remaining!==void 0&&D.days_remaining!==""?Number(D.days_remaining):null,se=W&&!Q&&te!==null&&!isNaN(te)?te:null,ae="con-status-aktif",le="\u{1F7E2} Aktif";Q?(ae="con-status-tetap",le="\u{1F3DB}\uFE0F PKWTT (Tetap)"):V?se!==null?se<=0?(ae="con-status-danger",le="\u26A0\uFE0F Habis Hari Ini"):se<=30?(ae="con-status-warning",le=`\u26A0\uFE0F Sisa ${se} Hari`):(ae="con-status-aktif",le=`\u{1F7E2} Aktif (${se} Hari)`):(ae="con-status-aktif",le="\u{1F7E2} Aktif"):(ae="con-status-danger",le=`\u26AA ${D.status||"Nonaktif"}`);let Ce=b[D.employee_id]||b[String(D.employee_name||"").toLowerCase()]||"";return`
                        <tr data-id="${D.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${Y+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="con-avatar ${U}" style="width:36px; height:36px; font-size:0.8rem; border-radius:9px;">
                                ${sa(D.employee_name)}
                              </div>
                              <div>
                                <div style="font-weight:700; color:#0F172A; white-space:nowrap;">${D.employee_name||"Karyawan Tanpa Nama"}</div>
                                <div style="font-size:0.74rem; color:#64748B;">${D.contract_type||"Kontrak Kerja"} \u2022 ID #${D.id}</div>
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#0F172A; font-size:0.84rem; margin-bottom:4px;">\u{1F4CD} ${D.branch_name||"Cabang Tidak Ditentukan"}</div>
                            <span class="con-div-badge ${G}" style="padding:2px 8px; font-size:0.7rem;">
                              ${J==="SECURITY"?"\u{1F6E1}\uFE0F":J==="FC - RELIEFER"?"\u{1F504}":"\u{1F9F9}"} ${J}
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            ${D.pkwt_number?`
                              <span class="con-pkwt-badge">
                                \u{1F4DC} ${D.pkwt_number}
                              </span>
                            `:'<span style="color:#94A3B8;">-</span>'}
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            ${Je(D.start_date)}
                          </td>
                          <td style="white-space:nowrap; font-weight:600; ${Q?"color:#4338CA; font-weight:700;":"color:#1E293B;"}">
                            ${Q?"Tetap / PKWTT":W?Je(D.end_date):"-"}
                          </td>
                          <td style="white-space:nowrap; font-size:0.82rem; font-weight:600; ${Q?"color:#15803D;":se!==null&&se<=30&&V?"color:#D97706; font-weight:700;":"color:#475569;"}">
                            ${Q?"Karyawan Tetap (Permanen)":se!==null?se<=0?"Masa Kontrak Telah Habis":`${se} Hari Tersisa`:"-"}
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="con-status-badge ${ae}">
                              ${le}
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <div style="display:inline-flex; align-items:center; gap:6px;">
                              ${Ce?`
                                <a href="https://wa.me/${Ce.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="con-icon-btn" title="Chat WhatsApp Karyawan" style="color:#16A34A;">
                                  <svg width="15" height="15" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                                </a>
                              `:""}
                              ${B?`
                                <button class="con-icon-btn edit" data-action="edit" data-id="${D.id}" title="Edit Data Kontrak">
                                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              `:""}
                              ${c?`
                                <button class="con-icon-btn delete" data-action="delete" data-id="${D.id}" title="Hapus Kontrak">
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
              ${n.map(D=>{let Y=D.division||"FACILITY CARE",J="con-div-fc",G="";Y==="SECURITY"?(J="con-div-sec",G="sec"):Y==="FC - RELIEFER"&&(J="con-div-rel",G="rel");let U=!!(D.end_date&&String(D.end_date).trim()&&D.end_date!=="-"),W=Fe(D.end_date),Q=D.status==="Aktif",V=D.days_remaining!==null&&D.days_remaining!==void 0&&D.days_remaining!==""?Number(D.days_remaining):null,te=U&&!W&&V!==null&&!isNaN(V)?V:null,se="con-status-aktif",ae="\u{1F7E2} Aktif";W?(se="con-status-tetap",ae="\u{1F3DB}\uFE0F PKWTT (Tetap)"):Q?te!==null?te<=0?(se="con-status-danger",ae="\u26A0\uFE0F Habis Hari Ini"):te<=30?(se="con-status-warning",ae=`\u26A0\uFE0F Sisa ${te} Hari`):(se="con-status-aktif",ae=`\u{1F7E2} Aktif (${te} Hari)`):(se="con-status-aktif",ae="\u{1F7E2} Aktif"):(se="con-status-danger",ae=`\u26AA ${D.status||"Nonaktif"}`);let le=b[D.employee_id]||b[String(D.employee_name||"").toLowerCase()]||"";return`
              <div class="con-card" data-id="${D.id}">
                <!-- Top Badges Row -->
                <div class="con-card-top">
                  <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
                    <span class="con-div-badge ${J}">
                      ${Y==="SECURITY"?"\u{1F6E1}\uFE0F":Y==="FC - RELIEFER"?"\u{1F504}":"\u{1F9F9}"} ${Y}
                    </span>
                    ${D.pkwt_number?`
                      <span class="con-pkwt-badge">
                        \u{1F4DC} ${D.pkwt_number}
                      </span>
                    `:""}
                  </div>
                  <span class="con-status-badge ${se}">
                    ${ae}
                  </span>
                </div>

                <!-- Profile Box -->
                <div class="con-profile-box">
                  <div class="con-avatar ${G}">
                    ${sa(D.employee_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="con-name">
                      ${D.employee_name||"Karyawan Tanpa Nama"}
                    </div>
                    <div class="con-subtitle">
                      Personil ${Y} \u2022 ${D.contract_type||"Kontrak Kerja"}
                    </div>
                  </div>
                </div>

                <!-- Information Grid -->
                <div class="con-info-card">
                  <div class="con-info-row">
                    <span class="con-info-label">\u{1F4CD} Cabang:</span>
                    <span class="con-info-val">${D.branch_name||"Cabang Tidak Ditentukan"}</span>
                  </div>
                  <div class="con-info-row">
                    <span class="con-info-label">\u{1F5D3}\uFE0F Mulai Kontrak:</span>
                    <span class="con-info-val">${Je(D.start_date)}</span>
                  </div>
                  <div class="con-info-row">
                    <span class="con-info-label">\u{1F3C1} Selesai Kontrak:</span>
                    <span class="con-info-val" style="${W?"color:#4338CA; font-weight:700;":""}">
                      ${W?"Tetap / PKWTT":U?Je(D.end_date):"-"}
                    </span>
                  </div>
                  <div class="con-info-row">
                    <span class="con-info-label">\u23F3 Masa Berlaku:</span>
                    <span class="con-info-val" style="${W?"color:#15803D;":te!==null&&te<=30&&Q?"color:#D97706; font-weight:700;":""}">
                      ${W?"Karyawan Tetap (Permanen)":te!==null?te<=0?"Masa Kontrak Telah Habis":`${te} Hari Tersisa`:"-"}
                    </span>
                  </div>
                </div>

                <!-- Notes box (if any) -->
                ${D.notes?`
                  <div class="con-notes-box">
                    \u{1F4DD} <strong>Catatan:</strong> ${D.notes}
                  </div>
                `:""}

                <!-- Footer Actions -->
                <div class="con-card-footer">
                  <div style="font-size:0.75rem; color:#94A3B8;">
                    ID #${D.id} \u2022 ${D.pkwt_number||"PKWT"}
                  </div>

                  <div class="con-action-btns">
                    ${le?`
                      <a href="https://wa.me/${le.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="con-icon-btn" title="Chat WhatsApp Karyawan" style="color:#16A34A;">
                        <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                      </a>
                    `:""}

                    ${B?`
                      <button class="con-icon-btn edit" data-action="edit" data-id="${D.id}" title="Edit Data Kontrak">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                    `:""}
                    ${c?`
                      <button class="con-icon-btn delete" data-action="delete" data-id="${D.id}" title="Hapus Kontrak">
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
    `,e.querySelectorAll(".con-view-btn").forEach(D=>{D.addEventListener("click",()=>{let Y=D.getAttribute("data-view");Y&&Y!==C&&(C=Y,localStorage.setItem("fcms_contracts_view_mode",Y),E())})});let j=e.querySelector("#con-search-input");j&&j.addEventListener("input",D=>{x=D.target.value,E()});let $=e.querySelector("#con-branch-filter");$&&$.addEventListener("change",D=>{g=D.target.value,E()});let p=e.querySelector("#con-division-filter");p&&p.addEventListener("change",D=>{o=D.target.value,E()});let S=e.querySelector("#con-expiring-filter");S&&S.addEventListener("change",D=>{A=D.target.value,E()}),e.querySelectorAll(".con-pill").forEach(D=>{D.addEventListener("click",Y=>{let J=Y.currentTarget.getAttribute("data-status");l=J,J==="Expiring30"?A="30":A="all",E()})});let u=e.querySelector("#con-reset-filter");u&&u.addEventListener("click",()=>{x="",g="all",o="all",l="all",A="all",E()});let M=e.querySelector("#con-btn-add");M&&M.addEventListener("click",()=>h());let R=e.querySelector("#con-btn-excel");R&&R.addEventListener("click",k);let K=e.querySelector("#con-btn-missing");K&&K.addEventListener("click",m);let O=e.querySelector("#con-btn-refresh");O&&O.addEventListener("click",()=>y()),e.querySelectorAll(".con-icon-btn.edit").forEach(D=>{D.addEventListener("click",Y=>{let J=parseInt(Y.currentTarget.getAttribute("data-id"),10),G=a.find(U=>U.id===J);G&&h(G)})}),e.querySelectorAll(".con-icon-btn.delete").forEach(D=>{D.addEventListener("click",Y=>{let J=parseInt(Y.currentTarget.getAttribute("data-id"),10),G=a.find(U=>U.id===J);G&&F(G)})})}await y()}de();var wt=[],Ye=[];function Pn(e){if(!Array.isArray(e))return"Q3";let t=["Q4","Q3","Q2","Q1"];for(let a of t)if(e.some(i=>i.period===a))return a;return"Q3"}async function da(e,t){wt=await re();let a=await me();Ye=["Berlin Ariansyah","Ade Surahman"];let i=P=>P&&!Ye.find(c=>String(typeof c=="object"?c.value:c).toLowerCase()===String(P).toLowerCase())?[...Ye,P]:Ye,s=await N(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),b=P=>{if(!P||P==="-"||String(P).trim()==="")return"";let c=String(P).split("-");return c.length===3&&c[0].length===4?`${c[2]}-${c[1]}-${c[0]}`:P},x=s.data?.data||[],g=Pn(x),o=t?t.get("dash_filter"):null,l=new Date,A=`${l.getFullYear()}-${String(l.getMonth()+1).padStart(2,"0")}`,T={},C=t&&t.get("month")?t.get("month"):null;o==="inspeksi"?T={status:"Done",activity_type:"Inspeksi Hygiene",month:C}:o==="gcdc"?T={status:"Done",activity_type:"GCDC",month:C}:o&&o.startsWith("period_")&&(T={period:o.replace("period_","").toUpperCase()});let _=new Date().getFullYear(),B=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((P,c)=>{let y=String(c+1).padStart(2,"0");return{value:`${_}-${y}`,label:`${P} ${_}`}});ue({container:e,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!1,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:T,onDataLoaded:P=>P.sort((c,y)=>{let v=c.opening_date?new Date(c.opening_date).getTime():0;return(y.opening_date?new Date(y.opening_date).getTime():0)-v}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:P=>yt(P)},{key:"period",label:"Periode",render:P=>ye(P)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:P=>b(P)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:P=>b(P)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:P=>b(P)},{key:"status",label:"Status",render:P=>be(P)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:wt},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:B},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Ye}],formFields:P=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:wt,value:P?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:P?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:P?.period},{name:"pic",label:"PIC",type:"combobox",options:i(P?.pic),value:P?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:P?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:P?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:P?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:P?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:P?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let P=await N(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(P.ok){let c=P.data.data.map(y=>({Cabang:y.branch_name||"",Kegiatan:y.activity_type||"",Periode:y.period||"",PIC:y.pic||"","Tgl Opening":y.opening_date||"","Tgl Target":y.target_date||"","Tgl Selesai":y.completion_date||"",Status:y.status||""}));ne(c,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async P=>{let y=(await N("/api/branches?all=1")).data?.data||[],v=F=>{if(!F)return null;let k=String(F||"").toLowerCase().trim(),E=k.match(/^(\d{3})/);if(E){let d=y.find(f=>String(f.code||"")===E[1]);if(d)return d.id}let n=y.find(d=>String(d.full_name||"").toLowerCase()===k||String(d.code||"").toLowerCase()===k||String(d.name||"").toLowerCase()===k);return n?n.id:null},m=F=>{if(F==null||F==="")return"";if(F instanceof Date&&!isNaN(F.getTime()))return F.toISOString().slice(0,10);let k=String(F).trim();if(k===""||k==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(k))return k.slice(0,10);if(/^\d{4,5}$/.test(k)){let n=Number(k);if(n>2e4&&n<99999){let d=new Date(Date.UTC(1899,11,30)+n*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let E=k.split(/[\/\-\.]/);if(E.length===3){let[n,d,f]=E.map(w=>w.trim());if(n.length===4&&d.length<=2&&f.length<=2)return`${n}-${d.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&d.length<=2&&n.length<=2)return`${f}-${d.padStart(2,"0")}-${n.padStart(2,"0")}`}return k},r=P.map(F=>({branch_id:v(String(F.Cabang||"").trim()),branch_name:String(F.Cabang||"").trim(),activity_type:String(F.Kegiatan||"").trim(),period:String(F.Periode||"").trim(),pic:String(F.PIC||F.Pic||"").trim(),opening_date:m(F["Tgl Opening"]||F["Tanggal Opening"]||F["Tgl Openir"]),target_date:m(F["Tgl Target"]||F["Tanggal Target"]),completion_date:m(F["Tgl Selesai"]||F["Tanggal Selesai"]),status:String(F.Status||"").trim(),notes:String(F.Catatan||F.Keterangan||"").trim()})).filter(F=>F.activity_type&&F.period),h=await N("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!h.ok)throw new Error(h.data?.error||"Import gagal");return he("schedule"),h.data}}})}ce();var Et=[{id:1,report_date:"2026-08-27",branch_id:19,branch_name:"019. Cinere",category:"SDM",source:"AM",complaint:"Keluhan AM fc atas nama Malik bau badan, dan sudah dibilangin berulang kali terkait kebersihan seragam dan deodoran.",employee_name:"MALIK",fc_specialist:"Mizwar",solution:"Diberikan teguran lisan, edukasi kebersihan diri, dan pendampingan seragam cadangan di klinik.",status:"Done",completion_date:"2026-08-28",day_count:1},{id:2,report_date:"2026-08-26",branch_id:0,branch_name:"000. Management",category:"SDM",source:"Berlin",complaint:"Security atas nama Reza suhelmi meninggalkan area tugas jaga tanpa personil pengganti saat jam operasional klinik.",employee_name:"REZA SUHELMI",fc_specialist:"Ade Surahman",solution:"Pemanggilan investigasi dan penerbitan Surat Peringatan (SP) sesuai SOP security.",status:"Done",completion_date:"2026-08-27",day_count:1},{id:3,report_date:"2026-07-27",branch_id:30,branch_name:"030. Ciputat",category:"SDM",source:"AM",complaint:"Terinfo kembali permasalahan FC Ciputat, terkait standar pembersihan area dental unit yang belum memenuhi checklist hygiene.",employee_name:"HENDRA",fc_specialist:"Berlin Ariansyah",solution:"Re-training modul pembersihan dental unit dan monitoring harian oleh supervisor.",status:"In Progress",completion_date:null,day_count:null},{id:4,report_date:"2026-06-24",branch_id:30,branch_name:"030. Ciputat",category:"SDM",source:"AM",complaint:"FC atas nama Hendra mengintimidasi rcp pada saat briefing pagi mengenai pembagian area kerja.",employee_name:"HENDRA",fc_specialist:"Berlin Ariansyah",solution:"Mediasi bersama AM dan HR, diberikan pembinaan konseling perilaku kerja.",status:"Done",completion_date:"2026-06-25",day_count:1},{id:5,report_date:"2026-06-05",branch_id:0,branch_name:"000. Management",category:"SDM",source:"Berlin",complaint:"Kinerja kurang dan ada komplain dari cabang mengenai ketidakhadiran berkala personil pengganti.",employee_name:"Indra",fc_specialist:"Ade",solution:"Evaluasi penempatan dan pembinaan kedisiplinan kerja.",status:"Done",completion_date:"2026-06-08",day_count:3},{id:6,report_date:"2026-06-05",branch_id:0,branch_name:"000. Management",category:"SDM",source:"Ade",complaint:"Sering tidak masuk kerja dan memalsukan surat dokter saat izin sakit.",employee_name:"Kristian",fc_specialist:"Ade Surahman",solution:"Investigasi ke klinik dokter terkait dan eskalasi ke HRD untuk tindakan tegas.",status:"Done",completion_date:"2026-06-07",day_count:2},{id:7,report_date:"2026-05-25",branch_id:19,branch_name:"019. Cinere",category:"SDM",source:"SPV",complaint:"Tidak masuk kerja tanpa kabar (alpa), dan sudah di-one on one oleh SPV terkait absensi.",employee_name:"Malik",fc_specialist:"Mizwar",solution:"Pemberian SP 1 dan penandatanganan komitmen kehadiran kerja.",status:"Done",completion_date:"2026-05-26",day_count:1}],pa=["SDM","Cleaning","Aset","K3","Lainnya"],Ln=["Open","In Progress","Done"],In=["SPV","AM","RCP","Perawat","FC","Security","Berlin","Ade","Pattrel","Dentrel"];function ca(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function zn(e){if(!e||e==="-")return"\u26A0\uFE0F";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function ua(e,t){let a=[],i=[],s=[],b="all",x="all",g="all",o="",l=!1;(t?t.get("dash_filter"):null)==="open"&&(b="Open");let T=Z(),C=ee(T,"issues","write");try{[i,s]=await Promise.all([re(),me()])}catch(m){console.error("Error fetching cache:",m)}async function _(){l=!0,v();try{let m=await N("/api/issues?limit=10000");if(m.ok){let r=m.data?.data||m.data||[];a=Array.isArray(r)&&r.length>0?r:Et}else a=Et}catch(m){console.warn("Gagal load issues, menggunakan data lokal:",m),a=Et}finally{l=!1,v()}}async function z(){try{let{downloadExcel:m}=await Promise.resolve().then(()=>(de(),fe)),r=a.map((h,F)=>({No:F+1,"Tanggal Info":h.report_date||"",Cabang:h.branch_name||"",Kategori:h.category||"","Sumber Laporan":h.source||"",Keluhan:h.complaint||"","Nama FC / Security":h.employee_name||"","FC Spesialis":h.fc_specialist||"","Solusi / Tindak Lanjut":h.solution||"",Status:h.status||"","Tanggal Selesai":h.completion_date||"","Durasi Hari":h.day_count!=null?h.day_count:""}));m(r,`Data_Permasalahan_FCMS_${new Date().toISOString().split("T")[0]}`),H("Berhasil mengekspor data permasalahan ke Excel")}catch(m){q("Gagal ekspor Excel: "+m.message)}}async function B(m,r){try{let h=r==="Done"?m.completion_date||new Date().toISOString().split("T")[0]:r==="Open"?null:m.completion_date,F=m.day_count;if(m.report_date&&h){let n=new Date(m.report_date),d=new Date(h);F=Math.max(0,Math.floor((d-n)/(1e3*60*60*24)))}else r==="Open"&&(F=null);let k={...m,status:r,completion_date:h,day_count:F},E=await N(`/api/issues/${m.id}`,{method:"PUT",body:JSON.stringify(k)});E.ok?(H(`Status masalah diubah ke "${r}"`),await _()):q(E.data?.error||"Gagal memperbarui status")}catch(h){q("Terjadi kesalahan: "+h.message)}}function P(m=null){let r=!!m,h=new Date().toISOString().split("T")[0],F=m?.report_date||h,k=m?.category||"SDM",E=m?.source||"AM",n=m?.status||"Open",d=document.createElement("div");d.className="issue-modal-body",d.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Info <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="iss-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${F}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${i.map($=>`
                <option value="${$.value}" ${String(m?.branch_id)===String($.value)?"selected":""}>
                  ${$.label}
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
              ${pa.map($=>`
                <option value="${$}" ${k===$?"selected":""}>${$}</option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Sumber Laporan
            </label>
            <input list="iss-source-options" id="iss-modal-source" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: AM, SPV, Berlin..." value="${E}" />
            <datalist id="iss-source-options">
              ${In.map($=>`<option value="${$}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Detail Keluhan / Permasalahan <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="iss-modal-complaint" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Jelaskan kendala, kronologi, atau komplain..." required>${m?.complaint||""}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Personil FC / Security
            </label>
            <input list="iss-emp-options" id="iss-modal-employee" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Ketik nama karyawan terkait..." value="${m?.employee_name?String(m.employee_name).replace(/"/g,"&quot;"):""}" />
            <datalist id="iss-emp-options">
              ${s.map($=>`<option value="${$.label}">`).join("")}
            </datalist>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              FC Spesialis / Penanggung Jawab
            </label>
            <input list="iss-spec-options" id="iss-modal-spec" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: Mizwar, Ade, Berlin..." value="${m?.fc_specialist?String(m.fc_specialist).replace(/"/g,"&quot;"):""}" />
            <datalist id="iss-spec-options">
              ${s.map($=>`<option value="${$.label}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Solusi / Tindak Lanjut Penanganan
          </label>
          <textarea id="iss-modal-solution" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tindakan yang telah atau akan diambil untuk menyelesaikan masalah...">${m?.solution||""}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Masalah <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${Ln.map($=>`
                <option value="${$}" ${n===$?"selected":""}>${$}</option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Selesai
            </label>
            <input type="date" id="iss-modal-completion" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${m?.completion_date||""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Durasi (Hari)
            </label>
            <input type="number" id="iss-modal-daycount" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; background:#F8FAFC;" 
              value="${m?.day_count!=null?m.day_count:""}" readonly placeholder="Otomatis" />
          </div>
        </div>
      </div>
    `;let f=d.querySelector("#iss-modal-date"),w=d.querySelector("#iss-modal-completion"),L=d.querySelector("#iss-modal-daycount"),I=d.querySelector("#iss-modal-status");function j(){if(f.value&&w.value){let $=new Date(f.value),p=new Date(w.value),S=Math.floor((p-$)/(1e3*60*60*24));L.value=isNaN(S)?"":Math.max(0,S)}else L.value=""}f.addEventListener("change",j),w.addEventListener("change",j),I.addEventListener("change",()=>{I.value==="Done"&&!w.value?(w.value=new Date().toISOString().split("T")[0],j()):I.value==="Open"&&(w.value="",L.value="")}),X({title:r?"\u270F\uFE0F Edit Data Permasalahan":"\u2795 Tambah Data Permasalahan",content:d,size:"md",confirmText:r?"Simpan Perubahan":"Simpan Permasalahan",confirmClass:"btn-primary",onConfirm:async($,p)=>{let S=d.querySelector("#iss-modal-date").value,u=d.querySelector("#iss-modal-branch").value,M=d.querySelector("#iss-modal-category").value,R=(d.querySelector("#iss-modal-source").value||"").trim(),K=(d.querySelector("#iss-modal-complaint").value||"").trim(),O=(d.querySelector("#iss-modal-employee").value||"").trim(),D=(d.querySelector("#iss-modal-spec").value||"").trim(),Y=(d.querySelector("#iss-modal-solution").value||"").trim(),J=d.querySelector("#iss-modal-status").value,G=d.querySelector("#iss-modal-completion").value||null;if(!S||!u||!K||!M){q("Harap lengkapi Tanggal, Cabang, Kategori, dan Keluhan!");return}let U=null;if(S&&G){let Q=new Date(S),V=new Date(G);U=Math.max(0,Math.floor((V-Q)/(1e3*60*60*24)))}let W={report_date:S,branch_id:Number(u),category:M,source:R,complaint:K,employee_name:O,fc_specialist:D,solution:Y,status:J,completion_date:G,day_count:U};try{let Q;r?Q=await N(`/api/issues/${m.id}`,{method:"PUT",body:JSON.stringify(W)}):Q=await N("/api/issues",{method:"POST",body:JSON.stringify(W)}),Q.ok?(H(r?"Permasalahan berhasil diperbarui!":"Permasalahan baru berhasil ditambahkan!"),p(),await _()):q(Q.data?.error||"Gagal menyimpan permasalahan")}catch(Q){q("Terjadi kesalahan: "+Q.message)}}})}async function c(m){ie({title:"Hapus Permasalahan?",message:`Apakah Anda yakin ingin menghapus tiket permasalahan dari <strong>${m.branch_name||"cabang"}</strong>:<br><br><em>"${m.complaint}"</em>?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let r=await N(`/api/issues/${m.id}`,{method:"DELETE"});r.ok?(H("Permasalahan berhasil dihapus"),await _()):q(r.data?.error||"Gagal menghapus data")}catch(r){q("Terjadi kesalahan: "+r.message)}}})}function y(){return a.filter(m=>{if(o){let r=o.toLowerCase(),h=(m.complaint||"").toLowerCase().includes(r),F=(m.employee_name||"").toLowerCase().includes(r),k=(m.branch_name||"").toLowerCase().includes(r),E=(m.source||"").toLowerCase().includes(r),n=(m.category||"").toLowerCase().includes(r),d=(m.solution||"").toLowerCase().includes(r);if(!h&&!F&&!k&&!E&&!n&&!d)return!1}return!(b!=="all"&&m.status!==b||x!=="all"&&m.category!==x||g!=="all"&&String(m.branch_id)!==String(g))})}function v(){let m=y(),r=a.length,h=a.filter(K=>K.status==="Open"||K.status==="In Progress").length,F=a.filter(K=>K.status==="Done").length,k=new Date,E=k.getFullYear(),n=k.getMonth(),d=a.filter(K=>{if(!K.report_date)return!1;let O=new Date(K.report_date);return!isNaN(O.getTime())&&O.getFullYear()===E&&O.getMonth()===n}).length,f=a.length,w=a.filter(K=>K.status==="Open").length,L=a.filter(K=>K.status==="In Progress").length,I=a.filter(K=>K.status==="Done").length;e.innerHTML=`
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
              ${C?`
                <button id="issues-btn-add" class="btn btn-primary issues-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Permasalahan</span>
                </button>
              `:""}
              
              <button id="issues-btn-excel" class="btn btn-ghost issues-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="issues-btn-refresh" class="btn btn-ghost issues-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${l?"spin-icon":""}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="issues-stats-row">
            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(239, 68, 68, 0.2); color:#FCA5A5;">\u26A0\uFE0F</div>
              <div>
                <div class="issues-stat-val">${r}</div>
                <div class="issues-stat-lbl">Total Permasalahan</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FCD34D;">\u23F3</div>
              <div>
                <div class="issues-stat-val">${h}</div>
                <div class="issues-stat-lbl">Open & In Progress</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">\u2705</div>
              <div>
                <div class="issues-stat-val">${F}</div>
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
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(K=>`
                <option value="${K.value}" ${String(g)===String(K.value)?"selected":""}>
                  ${K.label}
                </option>
              `).join("")}
            </select>

            <select id="issues-category-filter" class="issues-select">
              <option value="all">Semua Kategori</option>
              ${pa.map(K=>`
                <option value="${K}" ${x===K?"selected":""}>
                  Kategori: ${K}
                </option>
              `).join("")}
            </select>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="issues-filter-pills">
            <div class="issues-pill ${b==="all"?"active":""}" data-status="all">
              <span>Semua Status</span>
              <span class="issues-pill-count">${f}</span>
            </div>
            <div class="issues-pill ${b==="Open"?"active":""}" data-status="Open">
              <span class="issues-pill-dot" style="background:#F59E0B;"></span>
              <span>Open</span>
              <span class="issues-pill-count">${w}</span>
            </div>
            <div class="issues-pill ${b==="In Progress"?"active":""}" data-status="In Progress">
              <span class="issues-pill-dot" style="background:#3B82F6;"></span>
              <span>In Progress</span>
              <span class="issues-pill-count">${L}</span>
            </div>
            <div class="issues-pill ${b==="Done"?"active":""}" data-status="Done">
              <span class="issues-pill-dot" style="background:#10B981;"></span>
              <span>Done</span>
              <span class="issues-pill-count">${I}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive Cards Grid -->
        <div class="issues-cards-grid">
          ${m.length===0?`
            <div class="issues-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada permasalahan yang cocok</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian atau filter status yang dipilih.</p>
              <button id="issues-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
          `:m.map(K=>{let O=K.category||"Lainnya",D=`issues-category-${O.toLowerCase().replace(/[^a-z0-9]/g,"")}`||"issues-category-lainnya",Y=K.status||"Open",J="issues-status-open",G="\u26A0\uFE0F Open";Y==="In Progress"?(J="issues-status-inprogress",G="\u{1F504} In Progress"):Y==="Done"&&(J="issues-status-done",G="\u2705 Done");let U=Y==="Done"?"status-done":Y==="In Progress"?"status-inprogress":"";return`
              <div class="issues-card" data-id="${K.id}">
                <!-- Top Badges Row -->
                <div class="issues-card-top">
                  <span class="issues-category-badge ${D}">
                    \u{1F3F7}\uFE0F ${O}
                  </span>
                  <div style="display:flex; align-items:center; gap:8px;">
                    <span class="issues-status-badge ${J}">
                      ${G}
                    </span>
                    <span style="font-size:0.8rem; color:#64748B; display:flex; align-items:center; gap:4px;">
                      \u{1F5D3}\uFE0F ${ca(K.report_date)}
                    </span>
                  </div>
                </div>

                <!-- Complaint Box -->
                <div class="issues-complaint-box ${U}">
                  "${K.complaint||"-"}"
                </div>

                <!-- Person & Location Box -->
                <div class="issues-person-box">
                  <div class="issues-avatar-box">
                    ${zn(K.employee_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="issues-person-sub">PERSONIL FC / SECURITY</div>
                    <div class="issues-person-name" style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${K.employee_name||"Tidak ada personil spesifik"}
                    </div>
                  </div>
                </div>

                <!-- Details Grid -->
                <div class="issues-details-grid">
                  <div class="issues-detail-item" style="grid-column: 1 / -1;">
                    <span style="color:#64748B;">\u{1F4CD} Cabang:</span>
                    <strong style="color:#0F172A;">${K.branch_name||"Cabang Tidak Diketahui"}</strong>
                  </div>
                  <div class="issues-detail-item">
                    <span style="color:#64748B;">\u{1F5E3}\uFE0F Sumber:</span>
                    <strong>${K.source||"-"}</strong>
                  </div>
                  <div class="issues-detail-item">
                    <span style="color:#64748B;">\u{1F6E1}\uFE0F Spesialis:</span>
                    <strong>${K.fc_specialist||"-"}</strong>
                  </div>
                </div>

                <!-- Solution & Completion Box (if available) -->
                ${K.solution||K.completion_date?`
                  <div class="issues-solution-box">
                    ${K.solution?`
                      <div class="issues-solution-title">
                        <span>\u{1F4A1} Tindak Lanjut / Solusi:</span>
                      </div>
                      <div style="margin-bottom:${K.completion_date?"6px":"0"};">
                        ${K.solution}
                      </div>
                    `:""}
                    ${K.completion_date?`
                      <div style="font-size:0.76rem; color:#15803D; font-weight:600; display:flex; align-items:center; gap:4px;">
                        <span>\u23F1\uFE0F Selesai: ${ca(K.completion_date)}</span>
                        ${K.day_count!=null?`<span>(Durasi: ${K.day_count} Hari)</span>`:""}
                      </div>
                    `:""}
                  </div>
                `:""}

                <!-- Footer Actions -->
                <div class="issues-card-footer">
                  <div>
                    ${C?Y==="Open"?`
                        <button class="issues-quick-btn start" data-action="quick-progress" data-id="${K.id}" title="Ubah status ke In Progress">
                          <span>\u{1F504} Mulai Tangani</span>
                        </button>
                      `:Y==="In Progress"?`
                        <button class="issues-quick-btn complete" data-action="quick-done" data-id="${K.id}" title="Tandai selesai">
                          <span>\u2705 Selesaikan</span>
                        </button>
                      `:`
                        <button class="issues-quick-btn reopen" data-action="quick-reopen" data-id="${K.id}" title="Buka kembali tiket masalah">
                          <span>\u21A9\uFE0F Buka Kembali</span>
                        </button>
                      `:""}
                  </div>

                  <div class="issues-action-btns">
                    ${C?`
                      <button class="issues-icon-btn edit" data-action="edit" data-id="${K.id}" title="Edit Masalah">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                      <button class="issues-icon-btn delete" data-action="delete" data-id="${K.id}" title="Hapus Masalah">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                      </button>
                    `:""}
                  </div>
                </div>
              </div>
            `}).join("")}
        </div>

      </div>
    `;let j=e.querySelector("#issues-search-input");j&&j.addEventListener("input",K=>{o=K.target.value,v();let O=e.querySelector("#issues-search-input");O&&(O.focus(),O.setSelectionRange(O.value.length,O.value.length))});let $=e.querySelector("#issues-branch-filter");$&&$.addEventListener("change",K=>{g=K.target.value,v()});let p=e.querySelector("#issues-category-filter");p&&p.addEventListener("change",K=>{x=K.target.value,v()}),e.querySelectorAll(".issues-pill").forEach(K=>{K.addEventListener("click",()=>{b=K.dataset.status,v()})});let S=e.querySelector("#issues-reset-filter");S&&S.addEventListener("click",()=>{o="",g="all",x="all",b="all",v()});let u=e.querySelector("#issues-btn-add");u&&u.addEventListener("click",()=>P());let M=e.querySelector("#issues-btn-excel");M&&M.addEventListener("click",()=>z());let R=e.querySelector("#issues-btn-refresh");R&&R.addEventListener("click",()=>_()),e.querySelectorAll(".issues-card").forEach(K=>{let O=Number(K.dataset.id),D=a.find(Q=>Q.id===O);if(!D)return;let Y=K.querySelector('[data-action="edit"]');Y&&Y.addEventListener("click",()=>P(D));let J=K.querySelector('[data-action="delete"]');J&&J.addEventListener("click",()=>c(D));let G=K.querySelector('[data-action="quick-progress"]');G&&G.addEventListener("click",()=>B(D,"In Progress"));let U=K.querySelector('[data-action="quick-done"]');U&&U.addEventListener("click",()=>B(D,"Done"));let W=K.querySelector('[data-action="quick-reopen"]');W&&W.addEventListener("click",()=>B(D,"Open"))})}await _()}var je=[];function Mn(e,t){let a=String(e.status||"").toLowerCase();return t==="pending"?a==="pending":!1}async function ga(e,t){let a=t?t.get("dash_filter"):null;je=await re();let i=await me(),s=["Ade","Berlin"],b=g=>g&&!i.find(o=>o.value===g)?[...i,{value:g,label:g}]:i,x=g=>g&&!s.find(o=>(typeof o=="object"?o.value:o)===g)?[...s,g]:s;ue({container:e,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!1,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:g=>a?g.filter(o=>Mn(o,a)):g,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:g=>window.formatDate(g)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:g=>`<span title="${g||""}">${g?.length>50?g.slice(0,50)+"\u2026":g||"-"}</span>`},{key:"solution",label:"Solusi",render:g=>`<span title="${g||""}">${g?.length>40?g.slice(0,40)+"\u2026":g||"-"}</span>`},{key:"status",label:"Status",render:g=>be(g)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:g=>window.formatDate(g)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:g=>g?`<a href="${g}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:je},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async g=>{let o=new URLSearchParams(g||{}).toString(),l=await N(`/api/one-on-one?limit=10000&${o}`);if(l.ok){let A=l.data.data.map(C=>({Tanggal:C.meeting_date||"",Cabang:C.branch_name||"","Nama Karyawan":C.employee_name||"",PIC:C.pic||"",Masalah:C.problem||"",Solusi:C.solution||"",Status:C.status||"","Tgl Selesai":C.completion_date||"",Dokumen:C.document_link||""})),{downloadExcel:T}=await Promise.resolve().then(()=>(de(),fe));T(A,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let g=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:o}=await Promise.resolve().then(()=>(de(),fe));o(g,"Template_Import_OneOnOne")},onImport:async g=>{let o=C=>{if(!C)return null;let _=String(C||"").toLowerCase(),z=je.find(B=>String(B.label||"").toLowerCase()===_);return z?z.value:null},l=C=>{if(!C)return"";if(C instanceof Date&&!isNaN(C.getTime()))return C.toISOString().slice(0,10);let _=String(C).trim();if(/^\d{4,5}$/.test(_)){let B=Number(_);if(B>2e4&&B<99999){let P=new Date(Date.UTC(1899,11,30)+B*864e5);return isNaN(P.getTime())?"":P.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(_))return _.slice(0,10);let z=_.split(/[\/\-\.]/);if(z.length===3){let[B,P,c]=z.map(y=>y.trim());if(B.length===4&&P.length<=2&&c.length<=2)return`${B}-${P.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&P.length<=2&&B.length<=2)return`${c}-${P.padStart(2,"0")}-${B.padStart(2,"0")}`}return _},A=g.map(C=>({meeting_date:l(C.Tanggal),employee_name:String(C["Nama Karyawan"]||"").trim(),branch_id:o(String(C.Cabang||"").trim()),pic:String(C.PIC||"").trim(),problem:String(C.Masalah||"").trim(),solution:String(C.Solusi||"").trim(),status:String(C.Status||"").trim(),completion_date:l(C["Tgl Selesai"]),document_link:String(C.Dokumen||"").trim()})).filter(C=>C.meeting_date&&C.employee_name&&C.branch_id),T=await N("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:A,onDuplicate:"update"})});if(!T.ok)throw new Error(T.data?.error||"Import gagal");return T.data}},formFields:g=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:g?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:g?.branch_id&&!je.find(o=>o.value==g.branch_id)?[...je,{value:g.branch_id,label:g.branch_name||g.branch_id}]:je,createApi:{path:"/api/branches",field:"full_name"},value:g?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:b(g?.employee_name),value:g?.employee_name},{name:"pic",label:"PIC",type:"select",options:x(g?.pic),createApi:{path:"/api/pic",field:"name"},value:g?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:g?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:g?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:g?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:g?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:g?.document_link}]})}ce();var St=[{id:1,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:6,branch_name:"006. Kelapa Gading MOI",trainer:"Berlin Ariansyah",participants:"MUHAMMAD PRAKAS AGAZI",score:90,document_link:""},{id:2,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:11,branch_name:"011. Kuningan",trainer:"Berlin Ariansyah",participants:"KAMALLUDIN",score:88,document_link:""},{id:3,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:42,branch_name:"042. Muara Karang",trainer:"Berlin Ariansyah",participants:"SAHRUL ICHSANI",score:85,document_link:""},{id:4,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:51,branch_name:"051. Bintaro Sektor 9",trainer:"Berlin Ariansyah",participants:"MUHAMAD VICKRI FATHURROHMAN",score:92,document_link:""},{id:5,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:45,branch_name:"045. FX Mall",trainer:"Berlin Ariansyah",participants:"DEDE AHDORI",score:86,document_link:""},{id:6,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:54,branch_name:"054. Bogor Gunung Batu",trainer:"Berlin Ariansyah",participants:"ADITYA PRATAMA",score:95,document_link:""},{id:7,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:55,branch_name:"055. Gajah Mada",trainer:"Berlin Ariansyah",participants:"ELVAN JANGGA PRAMANA",score:89,document_link:""}],ma=["Batch 5","Batch 4","Batch 3","Batch 2","Batch 1"],ba=["Berlin Ariansyah","Ade Surahman","Mizwar","Fajar","Ade","Berlin"];function Kn(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Be(e){if(!e)return"-";try{let t=JSON.parse(e);if(Array.isArray(t))return t.join(", ")}catch{}return String(e)}async function ha(e){let t=[],a=[],i=[],s="all",b="all",x="all",g="",o=!1,l=Z(),A=ee(l,"training","write");try{[a,i]=await Promise.all([re(),me()])}catch(y){console.error("Error fetching cache:",y)}async function T(){o=!0,P();try{let y=await N("/api/training?limit=10000");if(y.ok){let v=y.data?.data||y.data||[];t=Array.isArray(v)&&v.length>0?v:St}else t=St}catch(y){console.warn("Gagal load training, menggunakan data lokal:",y),t=St}finally{o=!1,P()}}async function C(){try{let{downloadExcel:y}=await Promise.resolve().then(()=>(de(),fe)),v=t.map((m,r)=>({No:r+1,Tanggal:m.training_date||"",Batch:m.batch||"",Materi:m.subject||"",Cabang:m.branch_name||"",Trainer:m.trainer||"",Peserta:Be(m.participants),Nilai:m.score!=null?m.score:"",Dokumen:m.document_link||"",Catatan:m.notes||""}));y(v,`Data_Training_${new Date().toISOString().slice(0,10)}`),H("Data training berhasil diekspor ke Excel!")}catch(y){q("Gagal ekspor Excel: "+y.message)}}async function _(y){if(!A)return;let v=t.find(m=>m.id===y);ie({title:"Hapus Data Training",message:`Apakah Anda yakin ingin menghapus data sesi training "${v?.subject||"ini"}" untuk ${Be(v?.participants)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let m=await N(`/api/training/${y}`,{method:"DELETE"});m.ok?(H("Data training berhasil dihapus"),t=t.filter(r=>r.id!==y),P()):q(m.data?.error||"Gagal menghapus data training")}catch(m){q("Gagal menghapus training: "+m.message)}}})}function z(y=null){if(!A)return;let v=!!y,m=y?.training_date?y.training_date.slice(0,10):new Date().toISOString().slice(0,10),r=y?.batch||"Batch 5",h=y?.trainer||"Berlin Ariansyah",F=Be(y?.participants)!=="-"?Be(y?.participants):"",k=document.createElement("div");k.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Training <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="tr-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${m}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Batch Pelatihan <span style="color:#DC2626;">*</span>
            </label>
            <input list="tr-batch-options" id="tr-modal-batch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih atau ketik batch..." value="${r}" required />
            <datalist id="tr-batch-options">
              ${ma.map(E=>`<option value="${E}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Materi / Topik Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="tr-modal-subject" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Basic Knowledge & Practical Skills" value="${y?.subject?String(y.subject).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="tr-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${a.map(E=>`
                <option value="${E.value}" ${String(y?.branch_id)===String(E.value)?"selected":""}>
                  ${E.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Trainer / Instruktur <span style="color:#DC2626;">*</span>
            </label>
            <input list="tr-trainer-options" id="tr-modal-trainer" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih trainer..." value="${h}" required />
            <datalist id="tr-trainer-options">
              ${ba.map(E=>`<option value="${E}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Peserta Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="tr-modal-participants" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Ketik nama karyawan peserta training..." required>${F}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nilai / Score Evaluasi (0 - 100)
            </label>
            <input type="number" min="0" max="100" step="1" id="tr-modal-score" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: 90" value="${y?.score!=null?y.score:""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Link Dokumen / Sertifikat (Opsional)
            </label>
            <input type="url" id="tr-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="https://drive.google.com/..." value="${y?.document_link?String(y.document_link).replace(/"/g,"&quot;"):""}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan Pelatihan
          </label>
          <textarea id="tr-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Catatan hasil pelatihan, rekomendasi materi lanjutan...">${y?.notes||""}</textarea>
        </div>
      </div>
    `,X({title:v?"\u270F\uFE0F Edit Data Sesi Training":"\u2795 Tambah Data Sesi Training",content:k,size:"md",confirmText:v?"Simpan Perubahan":"Simpan Training",confirmClass:"btn-primary",onConfirm:async(E,n)=>{let d=k.querySelector("#tr-modal-date").value,f=(k.querySelector("#tr-modal-batch").value||"").trim(),w=(k.querySelector("#tr-modal-subject").value||"").trim(),L=k.querySelector("#tr-modal-branch").value,I=(k.querySelector("#tr-modal-trainer").value||"").trim(),j=(k.querySelector("#tr-modal-participants").value||"").trim(),$=k.querySelector("#tr-modal-score").value,p=$!==""?Number($):null,S=(k.querySelector("#tr-modal-link").value||"").trim(),u=(k.querySelector("#tr-modal-notes").value||"").trim();if(!d){q("Tanggal training wajib diisi");return}if(!w){q("Materi / topik pelatihan wajib diisi");return}if(!L){q("Cabang klinik wajib dipilih");return}if(!I){q("Trainer wajib diisi");return}if(!j){q("Nama peserta training wajib diisi");return}let M=a.find(K=>String(K.value)===String(L)),R={training_date:d,batch:f,subject:w,branch_id:Number(L),branch_name:M?M.label:"Cabang",trainer:I,participants:j,score:p,document_link:S,notes:u};try{if(v){let K=await N(`/api/training/${y.id}`,{method:"PUT",body:JSON.stringify(R)});if(K.ok){H("Data training berhasil diperbarui");let O=t.findIndex(D=>D.id===y.id);O!==-1&&(t[O]={...t[O],...R}),n(),P()}else q(K.data?.error||"Gagal memperbarui data training")}else{let K=await N("/api/training",{method:"POST",body:JSON.stringify(R)});if(K.ok){H("Sesi training baru berhasil ditambahkan");let O=K.data?.id||Date.now();t.unshift({id:O,...R}),n(),P()}else q(K.data?.error||"Gagal menambahkan training")}}catch(K){q("Terjadi kesalahan: "+K.message)}}})}function B(){return t.filter(y=>{if(g){let v=g.toLowerCase(),m=(y.subject||"").toLowerCase().includes(v),r=(y.trainer||"").toLowerCase().includes(v),h=Be(y.participants).toLowerCase().includes(v),F=(y.branch_name||"").toLowerCase().includes(v),k=(y.batch||"").toLowerCase().includes(v);if(!m&&!r&&!h&&!F&&!k)return!1}return!(s!=="all"&&y.batch!==s||b!=="all"&&String(y.branch_id)!==String(b)||x!=="all"&&y.trainer!==x)})}function P(){let y=B(),v=t.length,m=new Set;t.forEach(f=>{Be(f.participants).split(",").forEach(L=>{let I=L.trim();I&&I!=="-"&&m.add(I)})});let r=m.size,F=new Set(t.map(f=>f.batch).filter(Boolean)).size,k=new Date,E=k.getFullYear(),n=k.getMonth(),d=t.filter(f=>{if(!f.training_date)return!1;let w=new Date(f.training_date);return!isNaN(w.getTime())&&w.getFullYear()===E&&w.getMonth()===n}).length;e.innerHTML=`
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
              ${A?`
                <button id="tr-btn-add" class="btn btn-primary training-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Training</span>
                </button>
              `:""}
              
              <button id="tr-btn-excel" class="btn btn-ghost training-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

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
                <div class="training-stat-val">${v}</div>
                <div class="training-stat-lbl">Sesi Training</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F465}
              </div>
              <div>
                <div class="training-stat-val">${r}</div>
                <div class="training-stat-lbl">Peserta Terlatih</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4DA}
              </div>
              <div>
                <div class="training-stat-val">${F} Batch</div>
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
              <input type="text" id="tr-search-input" class="training-search-input" placeholder="Cari materi, trainer, peserta, atau cabang..." value="${g}" />
            </div>

            <select id="tr-filter-branch" class="training-select">
              <option value="all">Semua Cabang (${a.length})</option>
              ${a.map(f=>`<option value="${f.value}" ${String(b)===String(f.value)?"selected":""}>${f.label}</option>`).join("")}
            </select>

            <select id="tr-filter-trainer" class="training-select">
              <option value="all">Semua Trainer</option>
              ${ba.map(f=>`<option value="${f}" ${x===f?"selected":""}>${f}</option>`).join("")}
            </select>
          </div>

          <!-- Batch Filter Pills (Swipeable on Mobile) -->
          <div class="training-filter-pills">
            <div class="training-pill ${s==="all"?"active":""}" data-batch="all">
              <span>Semua Batch</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${v}</span>
            </div>
            ${ma.map(f=>{let w=t.filter(L=>L.batch===f).length;return`
                <div class="training-pill ${s===f?"active":""}" data-batch="${f}">
                  <span>\u{1F393}</span>
                  <span>${f}</span>
                  ${w>0?`<span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${w}</span>`:""}
                </div>
              `}).join("")}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="training-list-container">
          ${y.length===0?`
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
              ${y.map(f=>{let w=Be(f.participants),I=w.split(",")[0].trim().split(" ").map(j=>j[0]).filter(Boolean).slice(0,2).join("").toUpperCase()||"TR";return`
                  <div class="training-card" data-id="${f.id}">
                    <div>
                      <!-- Card Header: Batch & Nilai/Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="training-batch-badge">
                          <span>\u{1F393}</span>
                          <span>${f.batch||"Reguler"}</span>
                        </span>
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${f.score!=null?`
                            <span class="training-score-badge">
                              <span>\u2B50</span>
                              <span>${f.score}</span>
                            </span>
                          `:""}
                          <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                            \u{1F5D3}\uFE0F ${Kn(f.training_date)}
                          </span>
                        </div>
                      </div>

                      <!-- Subject / Materi Title -->
                      <h4 style="margin:0 0 10px 0; font-size:1.05rem; font-weight:700; color:#0F172A; line-height:1.4;">
                        ${f.subject||"Materi Pelatihan"}
                      </h4>

                      <!-- Participant & Branch Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:10px;">
                        <div class="training-avatar-box">
                          ${I}
                        </div>
                        <div style="min-width:0;">
                          <div style="font-size:0.72rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:1px;">Peserta Training</div>
                          <div style="font-size:0.95rem; font-weight:700; color:#1E293B; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${w}">
                            ${w}
                          </div>
                        </div>
                      </div>

                      <!-- Detail Box: Cabang & Trainer -->
                      <div class="training-info-box">
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#334155;">
                          <span>\u{1F4CD}</span>
                          <span style="font-weight:600;">${f.branch_name||"Cabang Belum Ditentukan"}</span>
                        </div>
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#64748B;">
                          <span>\u{1F468}\u200D\u{1F3EB}</span>
                          <span>Trainer: <strong style="color:#0F172A;">${f.trainer||"-"}</strong></span>
                        </div>
                      </div>

                      ${f.notes?`
                        <div style="font-size:0.78rem; color:#64748B; font-style:italic; line-height:1.4; margin-bottom:8px;">
                          \u{1F4DD} ${f.notes}
                        </div>
                      `:""}
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:8px;">
                      <div>
                        ${f.document_link?`
                          <a href="${f.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>\u{1F4C4} Materi / Sertifikat</span>
                          </a>
                        `:`
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${A?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          <button class="training-icon-btn training-edit-btn" title="Edit Data Training" data-id="${f.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                          <button class="training-icon-btn delete training-delete-btn" title="Hapus Training" data-id="${f.id}">
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
    `,c()}function c(){let y=e.querySelector("#tr-btn-add");y&&(y.onclick=()=>z());let v=e.querySelector("#tr-btn-excel");v&&(v.onclick=()=>C());let m=e.querySelector("#tr-btn-refresh");m&&(m.onclick=()=>T());let r=e.querySelector("#tr-btn-reset-filter");r&&(r.onclick=()=>{g="",b="all",s="all",x="all",P()});let h=e.querySelector("#tr-search-input");h&&(h.oninput=E=>{g=E.target.value,P();let n=e.querySelector("#tr-search-input");n&&(n.focus(),n.setSelectionRange(n.value.length,n.value.length))});let F=e.querySelector("#tr-filter-branch");F&&(F.onchange=E=>{b=E.target.value,P()});let k=e.querySelector("#tr-filter-trainer");k&&(k.onchange=E=>{x=E.target.value,P()}),e.querySelectorAll(".training-pill").forEach(E=>{E.onclick=()=>{s=E.getAttribute("data-batch")||"all",P()}}),e.querySelectorAll(".training-edit-btn").forEach(E=>{E.onclick=n=>{n.stopPropagation();let d=Number(E.getAttribute("data-id")),f=t.find(w=>w.id===d);f&&z(f)}}),e.querySelectorAll(".training-delete-btn").forEach(E=>{E.onclick=n=>{n.stopPropagation();let d=Number(E.getAttribute("data-id"));_(d)}})}await T()}ce();de();function fa({container:e,title:t,icon:a,apiPath:i,columns:s,formFields:b,filterFields:x,defaultFilters:g={},itemLabel:o="Data",canCreate:l=!0,canEdit:A=!0,canDelete:T=!0,onBeforeSubmit:C,onAfterLoad:_,onDataLoaded:z,extraActions:B=[],initialSearch:P="",exportOptions:c=null,bulkDelete:y=!1,paginationMode:v="server"}){let m=Z();ee(m,"delete")||(T=!1,y=!1),m&&m.role==="viewer"&&(l=!1,A=!1,T=!1,y=!1,c=null),m&&m.role==="editor_khusus"&&i!=="/api/issues"&&i!=="/api/overtime"&&(l=!1,A=!1,T=!1,y=!1,c=null),m&&m.role==="input_lembur"&&(i==="/api/overtime"?(T=!1,y=!1,c=null):(l=!1,A=!1,T=!1,y=!1,c=null));let r=1,h={...g};P&&(h.search=P);let F=new Set;e.innerHTML=`
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
      <h1 class="page-title">${a} ${t}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${l?`<button class="btn btn-primary" id="btn-create">+ Tambah ${o}</button>`:""}
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
    

    ${x&&x.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${x.filter(p=>p.type==="search").map(p=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${p.placeholder||"Cari..."}" id="filter-search" value="${h.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${x.filter(p=>p.type!=="search").map(p=>{if(p.type==="select"||p.type==="combobox"){let S=(p.label||"").startsWith("Pilih")?p.label:`Pilih ${p.label||""}`;return`<select class="filter-select" name="${p.name}" id="filter-${p.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${S}</option>${(p.options||[]).map(u=>`<option value="${typeof u=="object"?u.value:u}" ${h[p.name]===(typeof u=="object"?u.value:u)?"selected":""}>${typeof u=="object"?u.label:u}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${x.filter(p=>p.type!=="search").map(p=>{if(p.type==="select"||p.type==="combobox"){let S=(p.label||"").startsWith("Pilih")?p.label:`Pilih ${p.label||""}`;return`<select class="filter-select filter-select-sheet" name="${p.name}-sheet" id="filter-sheet-${p.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${S}</option>${(p.options||[]).map(u=>`<option value="${typeof u=="object"?u.value:u}" ${h[p.name]===(typeof u=="object"?u.value:u)?"selected":""}>${typeof u=="object"?u.label:u}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function k(){let p=document.getElementById("bulk-toolbar");if(!p)return;let S=document.getElementById("bulk-count"),u=document.getElementById("btn-bulk-delete"),M=document.getElementById("btn-bulk-cancel");S.textContent=`${F.size} item dipilih`,F.size>0?(p.style.display="flex",u.disabled=!1,M.disabled=!1):(p.style.display="none",u.disabled=!0,M.disabled=!0);let R=document.getElementById("select-all-checkbox");if(R){let K=document.querySelectorAll(".row-checkbox");if(K.length>0){let O=[...K].every(Y=>Y.checked),D=[...K].some(Y=>Y.checked);R.checked=O,R.indeterminate=D&&!O}else R.checked=!1,R.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{F.clear(),document.querySelectorAll(".row-checkbox").forEach(S=>S.checked=!1);let p=document.getElementById("select-all-checkbox");p&&(p.checked=!1),k()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(F.size===0)return;let p=[...F],S=document.createElement("div");S.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",S.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${p.length} ${o}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${p.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(S),S.querySelector("#bulk-cancel-btn").addEventListener("click",()=>S.remove()),S.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let u=S.querySelector("#bulk-confirm-btn");u.disabled=!0,u.textContent="Menghapus...";let M=await N(`${i}/bulk`,{method:"DELETE",body:JSON.stringify({ids:p})});S.remove(),M.ok?(H(`${p.length} ${o} berhasil dihapus.`),F.clear(),k(),he(i),L()):q(M.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),n;E?.addEventListener("input",p=>{clearTimeout(n),n=setTimeout(()=>{h.search=p.target.value,r=1,F.clear(),k(),L()},400)}),x?.forEach(p=>{(p.type==="select"||p.type==="combobox")&&(document.getElementById(`filter-${p.name}`)?.addEventListener("change",S=>{h[p.name]=S.target.value;let u=document.getElementById(`filter-sheet-${p.name}`);u&&(u.value=S.target.value),r=1,F.clear(),k(),L()}),document.getElementById(`filter-sheet-${p.name}`)?.addEventListener("change",S=>{h[p.name]=S.target.value;let u=document.getElementById(`filter-${p.name}`);u&&(u.value=S.target.value),r=1,F.clear(),k(),L(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{h={},E&&(E.value=""),x?.forEach(p=>{let S=document.getElementById(`filter-${p.name}`);S&&(S.value="");let u=document.getElementById(`filter-sheet-${p.name}`);u&&(u.value="")}),r=1,F.clear(),k(),L()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{h={},E&&(E.value=""),x?.forEach(p=>{let S=document.getElementById(`filter-${p.name}`);S&&(S.value="");let u=document.getElementById(`filter-sheet-${p.name}`);u&&(u.value="")}),r=1,F.clear(),k(),L(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>j(null)),c&&document.addEventListener("click",function(p){let S=document.getElementById("aksi-menu-main"),u=document.getElementById("btn-aksi-main");S&&u&&!u.contains(p.target)&&!S.contains(p.target)&&S.classList.remove("show-aksi-menu")});let d=document.getElementById("btn-mobile-filter"),f=document.getElementById("filter-options-wrapper"),w=document.getElementById("btn-close-filter-sheet");if(d&&f&&(d.addEventListener("click",p=>{p.preventDefault(),f.classList.add("sheet-open")}),w&&w.addEventListener("click",p=>{p.preventDefault(),f.classList.remove("sheet-open")})),c){document.getElementById(`btn-export-${c.moduleName}`)?.addEventListener("click",async S=>{let u=S.target,M=u.innerHTML;u.innerHTML="\u23F3 Loading...",u.disabled=!0;try{await c.onExport()}catch{q("Gagal export data")}finally{u.innerHTML=M,u.disabled=!1}}),document.getElementById(`btn-template-${c.moduleName}`)?.addEventListener("click",()=>{c.onTemplate()});let p=document.getElementById(`input-import-${c.moduleName}`);p?.addEventListener("change",async S=>{let u=S.target.files[0];if(!u)return;let M=document.getElementById(`label-import-${c.moduleName}`),R=M?M.querySelector(".import-text"):null,K=R?R.innerText:"";R&&(R.innerText="\u231B Memproses..."),M&&(M.style.pointerEvents="none"),p.disabled=!0;try{let O=await Ue(u);if(O.length===0)throw new Error("File kosong atau format salah");await c.onImport(O),H("Import berhasil!"),he(i),L()}catch(O){q(O.message||"Gagal import data")}finally{R&&(R.innerText=K),M&&(M.style.pointerEvents="auto"),p.disabled=!1,p.value=""}})}async function L(){k();let p=document.getElementById("table-container");if(!p)return;p.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let S=v==="client",u=S?1:r,M=S?Pe:20,R=new URLSearchParams({page:u,limit:M,...Object.fromEntries(Object.entries(h).filter(([,U])=>U))}),K=await N(`${i}?${R}`);if(!K.ok){p.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${K.data?.error||"Error"}</p></div>`;return}let O=K.data?.data||K.data||[],D=K.data?.pagination,Y=O.length;if(S){O=z(O);let U=O.length,W=20,Q=Math.ceil(U/W);r>Q&&Q>0&&(r=Q);let V=(r-1)*W,te=r*W;O=O.slice(V,te),D={page:r,limit:W,total:U,pages:Q}}!1,_&&_(O);let J=at({columns:s,data:O,onEdit:A?U=>j(U):null,actions:B.map(U=>({...U,handler:W=>U.handler(W,L)})),emptyText:`Tidak ada ${String(o||"").toLowerCase()}`,bulkSelect:y?{selectedIds:F,onToggle:k}:null});p.innerHTML="",p.appendChild(J);let G=document.getElementById("pagination-container");if(G&&(G.innerHTML="",D&&D.pages>1)){let U=nt({page:D.page,pages:D.pages,total:D.total,limit:D.limit,onPage:W=>{r=W,L()}});U&&G.appendChild(U)}}function I(p){let S=typeof b=="function"?b(p):b;return He(S)}function j(p){let S=!!p,u=document.createElement("form");if(u.noValidate=!0,u.innerHTML=I(p),S){let R=typeof b=="function"?b(p):b;rt(u,p)}let{close:M}=X({title:S?`Edit ${o}`:`Tambah ${o}`,content:u,size:"lg",confirmText:S?"Simpan Perubahan":`Tambah ${o}`,onConfirm:async(R,K)=>{if(!u.reportValidity())return;let O=R.querySelector(".modal-confirm");O.disabled=!0,O.textContent="Menyimpan...";let D=it(u),Y=typeof b=="function"?b(p):b,J=async Q=>{for(let V of Q)if(V.type==="row")await J(V.fields);else if(V.type==="combobox"&&D[V.name]){let te=D[V.name],se=(V.options||[]).find(ae=>{let le=String(typeof ae=="object"?ae.value:ae),Ce=String(typeof ae=="object"?ae.label:ae);return le===te||Ce===te});if(se)D[V.name]=typeof se=="object"?se.value:se;else if(V.createApi){let ae={};ae[V.createApi.field]=te,V.createApi.extra&&Object.assign(ae,V.createApi.extra);let le=await N(V.createApi.path,{method:"POST",body:JSON.stringify(ae)});if(le.ok&&le.data?.id)D[V.name]=le.data.id;else if(le.ok&&!le.data?.id)D[V.name]=te;else throw new Error(`Gagal membuat master data: ${le.data?.error||"Unknown error"}`)}}};try{await J(Y)}catch(Q){q(Q.message),O.disabled=!1,O.textContent=S?"Simpan Perubahan":`Tambah ${o}`;return}C&&(D=await C(D,p));let G=S?"PUT":"POST",U=S?`${i}/${p.id}`:i,W=await N(U,{method:G,body:JSON.stringify(D)});W.ok?(H(S?`${o} berhasil diperbarui.`:`${o} berhasil ditambahkan.`),K(),he(i),L()):(q(W.data?.error||"Gagal menyimpan data."),O.disabled=!1,O.textContent=S?"Simpan Perubahan":`Tambah ${o}`)}})}function $(p){ie(`Hapus ${o} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let S=await N(`${i}/${p.id}`,{method:"DELETE"});S.ok?(H(`${o} berhasil dihapus.`),he(i),L()):q(S.data?.error||"Gagal menghapus.")},`Hapus ${o}`)}return L(),L}de();async function xa(e,t){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let a=await re(),i=await me(),s=t?t.get("dash_filter"):null,b={};if(s==="reliever"){let C=new Date,_=`${C.getFullYear()}-${String(C.getMonth()+1).padStart(2,"0")}`;b={status:"Done",month:t&&t.get("month")?t.get("month"):_}}console.log("RAW",await Te()),console.log("OPTIONS",i);let x=C=>C&&!i.find(_=>_.value===C)?[...i,{value:C,label:C}]:i,g=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],o=C=>C&&!g.includes(C)?[...g,C]:g,l=new Date().getFullYear(),T=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((C,_)=>{let z=String(_+1).padStart(2,"0");return{value:`${l}-${z}`,label:`${C} ${l}`}});fa({container:e,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!1,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:b,onDataLoaded:C=>C.sort((_,z)=>{let B=_.backup_date?new Date(_.backup_date).getTime():0;return(z.backup_date?new Date(z.backup_date).getTime():0)-B}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:C=>ye(C)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:C=>window.formatDate(C)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:C=>window.formatDate(C)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:C=>C?`<span class="badge badge-info">${C}</span>`:"-"},{key:"status",label:"Status",render:C=>be(C)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:g},{type:"select",name:"branch_id",label:"Cabang",options:a},{type:"select",name:"month",label:"Bulan",options:T},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:C=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:a,value:C?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:C?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:x(C?.original_fc_name),value:C?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:o(C?.reliever_name),value:C?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:C?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:C?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:C?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:C?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:C?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let C=await N(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(C.ok){let _=C.data.data.map(z=>({Cabang:z.branch_name||"","Nama Facility care":z.original_fc_name||"",Periode:z.period||"",Relifer:z.reliever_name||"","Tanggal Back Up":z.backup_date||"","Tanggal Selesai":z.completion_date||"",Keterangan:z.reason||"",Shift:z.shift||"",Status:z.status||""}));_.length===0&&_.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),ne(_,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async C=>{let z=(await N("/api/branches?all=1")).data?.data||[],B=y=>{if(!y)return null;let v=String(y||"").toLowerCase(),m=z.find(r=>String(r.full_name||"").toLowerCase()===v||String(r.code||"").toLowerCase()===v||String(r.name||"").toLowerCase()===v);return m?m.id:null},P=C.map(y=>({branch_name:String(y.Cabang||"").trim(),backup_date:String(y["Tanggal Back Up"]||y["Tanggal Backup"]||"").trim(),original_fc_name:String(y["Nama Facility care"]||y["FC Digantikan"]||"").trim(),reliever_name:String(y.Relifer||y.Reliefer||"").trim(),period:String(y.Periode||"").trim(),reason:String(y.Keterangan||"").trim(),shift:String(y.Shift||"").trim(),completion_date:String(y["Tanggal Selesai"]||"").trim(),status:String(y.Status||"").trim()})).filter(y=>y.reliever_name&&y.backup_date),c=await N("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:P,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}}})}de();async function ya(e){let t=await re(),a=Array.from({length:4},(i,s)=>String(new Date().getFullYear()-s));ue({container:e,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:i=>ye(i)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"fc_score",label:"Point FC",render:i=>i!=null?`<strong class="${i>=80?"text-success":i>=60?"text-warning":"text-danger"}">${i}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:i=>i!=null?`<strong>${i}</strong>`:"-"},{key:"status",label:"Status",render:i=>be(i)},{key:"document_link",label:"Dokumen",render:i=>i?`<a href="${i}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:i=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:i?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:i?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:i?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:i?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:i?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:i?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:i?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:i?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async i=>{let s=new URLSearchParams(i||{}).toString(),b=await N(`/api/reports/inspection?limit=10000&${s}`);if(b.ok){let x=b.data.data.map(g=>({Cabang:g.branch_name||"",Periode:g.period||"",Tanggal:g.inspection_date||"","Point FC":g.fc_score!==null&&g.fc_score!==void 0?g.fc_score:"","Point SPV":g.spv_score!==null&&g.spv_score!==void 0?g.spv_score:"",Status:g.status||"","Link Dokumen":g.document_link||""}));ne(x,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async i=>{let s=o=>{if(!o)return null;let l=String(o||"").toLowerCase(),A=t.find(T=>String(T.label||"").toLowerCase()===l);return A?A.value:null},b=o=>{if(o==null||o==="")return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let l=String(o).trim();if(l===""||l==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);if(/^\d{4,5}$/.test(l)){let T=Number(l);if(T>2e4&&T<99999){let C=new Date(Date.UTC(1899,11,30)+T*864e5);return isNaN(C.getTime())?"":C.toISOString().slice(0,10)}}let A=l.split(/[\/\-\.]/);if(A.length===3){let[T,C,_]=A.map(z=>z.trim());if(T.length===4&&C.length<=2&&_.length<=2)return`${T}-${C.padStart(2,"0")}-${_.padStart(2,"0")}`;if(_.length===4&&C.length<=2&&T.length<=2)return`${_}-${C.padStart(2,"0")}-${T.padStart(2,"0")}`}return l},x=i.map(o=>({branch_id:s(String(o.Cabang||"").trim()),period:String(o.Periode||"").trim(),inspection_date:b(o.Tanggal),fc_score:o["Point FC"]!==void 0&&o["Point FC"]!==""?Number(o["Point FC"]):null,spv_score:o["Point SPV"]!==void 0&&o["Point SPV"]!==""?Number(o["Point SPV"]):null,status:String(o.Status||"").trim(),document_link:String(o["Link Dokumen"]||"").trim(),notes:String(o.Catatan||o.Keterangan||"").trim()})).filter(o=>o.branch_id&&o.period&&o.inspection_date),g=await N("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:x,onDuplicate:"update"})});if(!g.ok)throw new Error(g.data?.error||"Import gagal");return g.data}}})}de();async function va(e){let t=await re(),a=Array.from({length:4},(i,s)=>String(new Date().getFullYear()-s));ue({container:e,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:i=>`<span class="badge ${i==="Deep Cleaning"?"badge-purple":"badge-success"}">${i}</span>`},{key:"period",label:"Periode",render:i=>ye(i)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"status",label:"Status",render:i=>be(i)},{key:"document_link",label:"Dokumen",render:i=>i?`<a href="${i}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:i=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t,value:i?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:i?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:i?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:i?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:i?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:i?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:i?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async i=>{let s=new URLSearchParams(i||{}).toString(),b=await N(`/api/reports/cleaning?limit=10000&${s}`);if(b.ok){let x=b.data.data.map(g=>({Cabang:g.branch_name||"",Jenis:g.activity_type||"",Periode:g.period||"",Tanggal:g.activity_date||"",Status:g.status||"","Link Dokumen":g.document_link||""}));ne(x,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async i=>{let s=o=>{if(!o)return null;let l=String(o||"").toLowerCase(),A=t.find(T=>String(T.label||"").toLowerCase()===l);return A?A.value:null},b=o=>{if(o==null||o==="")return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let l=String(o).trim();if(l===""||l==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);if(/^\d{4,5}$/.test(l)){let T=Number(l);if(T>2e4&&T<99999){let C=new Date(Date.UTC(1899,11,30)+T*864e5);return isNaN(C.getTime())?"":C.toISOString().slice(0,10)}}let A=l.split(/[\/\-\.]/);if(A.length===3){let[T,C,_]=A.map(z=>z.trim());if(T.length===4&&C.length<=2&&_.length<=2)return`${T}-${C.padStart(2,"0")}-${_.padStart(2,"0")}`;if(_.length===4&&C.length<=2&&T.length<=2)return`${_}-${C.padStart(2,"0")}-${T.padStart(2,"0")}`}return l},x=i.map(o=>({branch_id:s(String(o.Cabang||"").trim()),activity_type:String(o.Jenis||o.Kegiatan||"").trim(),period:String(o.Periode||"").trim(),activity_date:b(o.Tanggal),status:String(o.Status||"").trim(),document_link:String(o["Link Dokumen"]||"").trim(),notes:String(o.Catatan||o.Keterangan||"").trim()})).filter(o=>o.branch_id&&o.activity_type&&o.period&&o.activity_date),g=await N("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:x,onDuplicate:"update"})});if(!g.ok)throw new Error(g.data?.error||"Import gagal");return g.data}}})}de();async function ka(e,t){let a=await re(),i=Array.from({length:4},(x,g)=>String(new Date().getFullYear()-g)),s=t?t.get("dash_filter"):null,b={};if(s==="fogging"){let x=new Date,g=String(x.getMonth()+1).padStart(2,"0"),o=String(x.getFullYear()),l=t?t.get("month"):null;l&&l.length===7&&(o=l.split("-")[0],g=l.split("-")[1]),b={status:"Done",month:g,year:o}}ue({container:e,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:b,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:x=>`<span class="badge badge-warning">${x}</span>`},{key:"period",label:"Periode",render:x=>ye(x)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:x=>window.formatDate(x)},{key:"status",label:"Status",render:x=>be(x)},{key:"document_link",label:"Dokumen",render:x=>x?`<a href="${x}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:x=>x||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:a},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:x=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:a,value:x?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:x?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:x?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:x?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:x?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:x?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async x=>{let g=new URLSearchParams(x||{}).toString(),o=await N(`/api/reports/fogging?limit=10000&${g}`);if(o.ok){let l=o.data.data.map(A=>({Cabang:A.branch_name||"",Jenis:A.activity_type||"Fogging",Periode:A.period||"",Tanggal:A.activity_date||"",Status:A.status||"","Link Dokumen":A.document_link||""}));ne(l,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async x=>{let g=T=>{if(!T)return null;let C=String(T||"").toLowerCase(),_=a.find(z=>String(z.label||"").toLowerCase()===C);return _?_.value:null},o=T=>{if(T==null||T==="")return"";if(T instanceof Date&&!isNaN(T.getTime()))return T.toISOString().slice(0,10);let C=String(T).trim();if(C===""||C==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(C))return C.slice(0,10);if(/^\d{4,5}$/.test(C)){let z=Number(C);if(z>2e4&&z<99999){let B=new Date(Date.UTC(1899,11,30)+z*864e5);return isNaN(B.getTime())?"":B.toISOString().slice(0,10)}}let _=C.split(/[\/\-\.]/);if(_.length===3){let[z,B,P]=_.map(c=>c.trim());if(z.length===4&&B.length<=2&&P.length<=2)return`${z}-${B.padStart(2,"0")}-${P.padStart(2,"0")}`;if(P.length===4&&B.length<=2&&z.length<=2)return`${P}-${B.padStart(2,"0")}-${z.padStart(2,"0")}`}return C},l=x.map(T=>({branch_id:g(String(T.Cabang||"").trim()),activity_type:String(T.Jenis||T.Kegiatan||"Fogging").trim(),period:String(T.Periode||"").trim(),activity_date:o(T.Tanggal),status:String(T.Status||"").trim(),document_link:String(T["Link Dokumen"]||"").trim(),notes:String(T.Catatan||T.Keterangan||"").trim()})).filter(T=>T.branch_id&&T.period&&T.activity_date),A=await N("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(l)});if(!A.ok)throw new Error(A.data?.error||"Import gagal");return A.data}}})}de();async function wa(e){let t=await re(),a=await me(),i=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],s=Array.from({length:4},(g,o)=>String(new Date().getFullYear()-o)),b=g=>g&&!a.find(o=>o.value===g)?[...a,{value:g,label:g}]:a,x=g=>g&&!i.find(o=>o.value===g)?[...i,{value:g,label:g}]:i;ue({container:e,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:g=>window.formatDate(g)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:g=>`<span title="${g||""}">${g?.length>60?g.slice(0,60)+"\u2026":g||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:g=>window.formatDate(g)},{key:"status",label:"Status",render:g=>be(g)},{key:"notes",label:"Keterangan",render:g=>g?.length>40?g.slice(0,40)+"\u2026":g||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:s}],formFields:g=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:g?.branch_id},{name:"pic",label:"PIC",type:"select",options:x(g?.pic),value:g?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:g?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:g?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:g?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:g?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:g?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async g=>{let o=new URLSearchParams(g||{}).toString(),l=await N(`/api/reports/basecamp?limit=10000&${o}`);if(l.ok){let A=l.data.data.map(T=>({"Tgl Info":T.info_date||"",Cabang:T.branch_name||"",Permasalahan:T.problem||"",PIC:T.pic||"","Tgl Done":T.done_date||"",Status:T.status||"",Keterangan:T.notes||""}));ne(A,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async g=>{let o=C=>{if(!C)return null;let _=String(C||"").toLowerCase(),z=t.find(B=>String(B.label||"").toLowerCase()===_);return z?z.value:null},l=C=>{if(C==null||C==="")return"";if(C instanceof Date&&!isNaN(C.getTime()))return C.toISOString().slice(0,10);let _=String(C).trim();if(_===""||_==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(_))return _.slice(0,10);if(/^\d{4,5}$/.test(_)){let B=Number(_);if(B>2e4&&B<99999){let P=new Date(Date.UTC(1899,11,30)+B*864e5);return isNaN(P.getTime())?"":P.toISOString().slice(0,10)}}let z=_.split(/[\/\-\.]/);if(z.length===3){let[B,P,c]=z.map(y=>y.trim());if(B.length===4&&P.length<=2&&c.length<=2)return`${B}-${P.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&P.length<=2&&B.length<=2)return`${c}-${P.padStart(2,"0")}-${B.padStart(2,"0")}`}return _},A=g.map(C=>({info_date:l(C["Tgl Info"]||C["Tanggal Info"]),branch_id:o(String(C.Cabang||"").trim()),problem:String(C.Permasalahan||"").trim(),pic:String(C.PIC||"").trim(),done_date:l(C["Tgl Done"]||C["Tanggal Done"]),status:String(C.Status||"").trim(),notes:String(C.Keterangan||C.Catatan||"").trim()})).filter(C=>C.info_date&&C.branch_id&&C.problem),T=await N("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(A)});if(!T.ok)throw new Error(T.data?.error||"Import gagal");return T.data}}})}ce();var jn=[{id:1,title:"Dusting (Pembersihan Debu)",category:"Pembersihan Rutin",icon:"\u{1F9F9}",color:"#2563EB",bg:"#EFF6FF",target:"Meja resepsionis, meja pasien, meja dokter, furniture, kaca, monitor & keyboard",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Kain Lap Microfiber Biru (Area Umum)","Kanebo (Kaca & Layar Monitor)","Kemoceng","Botol Spray Glass Cleaner","Kuas 4 Inch (Keyboard & Sela Sempit)"],steps:["Siapkan larutan glass cleaner di botol spray khusus kaca dan perabot.","Lipat kain microfiber menjadi 8 bidang kerja agar efisien dan bersih merata.","Semprotkan cairan pembersih ke permukaan kain lap, bukan langsung ke monitor atau bodi elektronik.","Usap permukaan perabot dari bagian paling atas ke bagian bawah secara searah atau pola S.","Gunakan kuas cat 4 inch untuk membersihkan sela-sela keyboard dan lubang ventilasi.","Keringkan kaca atau monitor menggunakan lap microfiber kering atau kanebo bersih."],notes:"Khusus layar monitor medis & dental unit, dilarang menyemprotkan cairan langsung ke layar."},{id:2,title:"Sweeping (Lobby Duster & Sapu Nylon)",category:"Pembersihan Rutin",icon:"\u{1F33E}",color:"#D97706",bg:"#FFFBEB",target:"Lobby utama, lorong, ruang tindakan, balkon, pantry, teras",chemical_ratio:"-",tools:["Lobby Duster Set 60cm (Area Lobby & Lorong)","Lobby Duster Set 40cm (Area Ruang Tindakan)","Sapu & Pengki Nagata 742s","Kapi Scrapper (Kerak/Permen Karet)"],steps:["Pastikan kain lobby duster dalam keadaan kering, bersih, dan bebas pasir kasar.","Dorong lobby duster menyusuri tepi dinding ruangan terlebih dahulu membentuk perimeter.","Lanjutkan mendorong dari sudut terjauh menuju arah pintu keluar tanpa mengangkat kain pel lobi dari lantai.","Kumpulkan debu dan kotoran di dekat pintu keluar, lalu angkat menggunakan sapu dan pengki Nagata.","Jika ada noda kerak atau bekas stiker/permen karet menempel di lantai, kikis perlahan menggunakan kapi scrapper.","Kibaskan kain duster di luar ruangan atau sedot sisa debu pada kain duster."],notes:"Jangan pernah menggunakan lobby duster di area lantai yang masih basah atau tergenang air."},{id:3,title:"Mopping (Pengepelan Lantai)",category:"Pembersihan Rutin",icon:"\u{1FAA3}",color:"#059669",bg:"#ECFDF5",target:"Seluruh lantai ruangan, koridor, dan area tunggu klinik",chemical_ratio:"25 ml Floor Cleaner per 1 Liter air bersih",tools:["Single Bucket Mop Wringer 20L / Double Bucket 46L","Mopping Stick / Tongkat & Kain Pel","Chemical Floor Cleaner Prokleen Lavender","Warning Sign Lantai Licin"],steps:['Pasang papan peringatan "Caution Wet Floor" di titik strategis yang mudah terlihat.',"Campurkan cairan Floor Cleaner ke dalam ember air bersih dengan takaran 25 ml per 1 Liter air.","Celupkan kain pel, masukkan ke dalam alat pemeras (wringer), dan peras hingga kondisi kain pel lembap (tidak menetes).","Mulai mengepel dari sudut terjauh ruangan mundur ke arah pintu keluar dengan gerakan membentuk pola angka delapan (8).","Bilas kain pel secara berkala apabila air pada ember bilas sudah terlihat keruh.","Biarkan lantai mengering sempurna secara alami sebelum papan peringatan wet floor dilepas."],notes:'Wajib pasang papan peringatan "Caution Wet Floor" sebelum mulai mengepel untuk mencegah risiko terpeleset.'},{id:4,title:"Membersihkan Toilet & Wastafel",category:"Sanitasi Toilet",icon:"\u{1F6BD}",color:"#DC2626",bg:"#FEF2F2",target:"Toilet pasien, toilet tim, closet, wastafel, dinding keramik, floor drain",chemical_ratio:"Karbol 25 ml:1 L Air / Bowl Cleaner 100 ml spray (kerak) / Bubuk KIFA",tools:["Sarung Tangan Karet (Wajib)","Sikat Kloset Scotch-Brite 3M","Kain Lap Microfiber Merah (Khusus Toilet)","Sponge Scotch Brite & Sikat Lantai ID-51","Floor Squeegee 75cm","Chemical Karbol Pine, Bowl Cleaner, Bubuk KIFA"],steps:["Kenakan sarung tangan karet dan siapkan perlengkapan sanitasi toilet.","Flush kloset, tuangkan Toilet Bowl Cleaner ke bagian dalam dinding kloset dan leher angsa, diamkan 5\u201310 menit.","Sikat bagian dalam kloset secara menyeluruh dengan sikat kloset, lalu bilas (flush) kembali hingga bersih.","Seka dudukan, penutup, dan tabung luar kloset menggunakan kain lap microfiber merah yang dibasahi larutan disinfektan karbol.","Bersihkan wastafel dan kran stainless menggunakan sponge tapas dengan bubuk pembersih KIFA (jangan gunakan Bowl Cleaner pada kran stainless).","Sikat lantai kamar mandi dengan sikat lantai ID-51 dan cairan karbol disinfektan.","Tarik genangan air ke arah floor drain menggunakan floor squeegee hingga lantai tidak licin dan cepat kering.","Pasang kamper/urinal block, isi ulang sabun cuci tangan, dan pasang plastik sampah baru."],notes:"Kain lap warna MERAH mutlak hanya untuk toilet dan mangkuk spittoon. Dilarang mencampur cairan kimia sembarangan."},{id:5,title:"Pembersihan Area Parkir & Luar",category:"Pembersihan Rutin",icon:"\u{1F17F}\uFE0F",color:"#4F46E5",bg:"#EEF2FF",target:"Halaman parkir mobil & motor, saluran air luar, teras depan",chemical_ratio:"-",tools:["Sapu Lidi Gagang Panjang","Pengki Nagata","Plastik Sampah Hitam Besar (90x120)","Selang Air"],steps:["Sapu daun kering, puntung rokok, plastik, dan pasir dari batas trotoar/pagar menuju satu titik.","Kumpulkan dan angkat sampah ke dalam pengki, masukkan langsung ke plastik sampah hitam besar.","Periksa dan bersihkan saluran drainase air di area parkir dari sumbatan sampah padat.","Bila ada tumpahan oli atau kotoran tanah membandel, siram dengan air dan sikat.","Rapikan pembatas parkir dan pastikan area depan klinik terlihat bersih dan rapi menyambut pasien."],notes:"Lakukan sebelum jam operasional klinik dibuka (pukul 07.30 - 08.00)."},{id:6,title:"Pembersihan Kaca & Jendela",category:"Pembersihan Rutin",icon:"\u{1FA9F}",color:"#0891B2",bg:"#ECFEFF",target:"Pintu kaca utama, jendela kaca klinik, partisi kaca, cermin toilet",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Botol Spray Glass Cleaner Prokleen","Window Wiper / Window Squeegee","Kain Microfiber Kering / Kanebo","Tongkat Teleskopik (Kaca Tinggi)"],steps:["Semprotkan larutan glass cleaner merata pada permukaan kaca yang hendak dibersihkan.","Tarik wiper karet dari arah atas ke bawah secara lurus vertikal atau zigzag horizontal tanpa terputus.","Lap bilah karet wiper setiap selesai satu tarikan menggunakan kain lap bersih agar tidak meninggalkan garis bekas air.","Gunakan kain microfiber kering untuk membersihkan sisa tetesan air di sudut-sudut lis bingkai kaca.","Untuk kaca bagian luar yang tinggi, pasang wiper pada tongkat sambungan teleskopik."],notes:"Hindari membersihkan kaca saat terpapar terik sinar matahari langsung karena cairan akan cepat mengering dan membekas noda."},{id:7,title:"Mencuci Bola Mandi Bola (Playground)",category:"Pembersihan Khusus",icon:"\u{1F939}",color:"#7C3AED",bg:"#F5F3FF",target:"Area bermain anak (Kids Playground) dan bola-bola mandi bola",chemical_ratio:"30 ml Dishwashing Liquid per 1 Liter Air + Karbol Disinfektan",tools:["Baskom Plastik Besar 55cm","Dishwashing Liquid Prokleen","Karbol Disinfektan Pine","Lap Microfiber Bersih","Karung / Wadah Pengering Jaring"],steps:["Kumpulkan seluruh bola mandi bola ke dalam wadah atau karung bersih.","Siapkan baskom plastik besar berisi air hangat kuku dicampur Dishwashing Liquid dan sedikit karbol disinfektan.","Rendam bola dan cuci secara bertahap sambil digosok perlahan menggunakan sponge/lap bersih.","Bilas bola dengan air bersih mengalir hingga tidak ada sisa busa sabun yang menempel.","Tiriskan dan angin-anginkan bola di tempat bersih beralaskan terpal/kain kering hingga benar-benar kering.","Seka dan sterilkan area kolam mandi bola sebelum bola dimasukkan kembali."],notes:"Jadwal pencucian berkala dilakukan minimal 1 kali seminggu atau sesuai kebutuhan kebersihan arena anak."},{id:8,title:"Membersihkan Filter Air Purifier",category:"Elektronik & Utility",icon:"\u{1F4A8}",color:"#0D9488",bg:"#F0FDFA",target:"Unit Air Purifier di ruang tunggu, ruang dokter, dan koridor",chemical_ratio:"-",tools:["Vacuum Cleaner (Nozzle Sikat Halus)","Kain Microfiber Kering","Kuas Cat Halus"],steps:["Matikan saklar power unit air purifier dan cabut kabel daya dari stop kontak dinding.","Buka penutup bodi (casing) depan atau belakang air purifier secara hati-hati.","Keluarkan pre-filter jaring nilon dari unit.","Bersihkan debu tebal pada pre-filter menggunakan vacuum cleaner nozzle sikat halus atau cuci dengan air mengalir (jika tipe wash-able).","Untuk filter HEPA dan Karbon Aktif: HANYA sedot debu permukaannya secara perlahan dengan vacuum, JANGAN DICUCI DENGAN AIR.","Lap bodi luar dan kisi-kisi udara masuk menggunakan lap microfiber kering.","Pastikan pre-filter benar-benar kering sebelum dipasang kembali ke posisinya semula."],notes:"Jangan pernah mencuci filter HEPA dengan air karena akan merusak serat mikro penyaring bakteri/virus."},{id:9,title:"Membersihkan Filter & Bodi AC",category:"Elektronik & Utility",icon:"\u2744\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Filter udara dan penutup indoor unit AC split di seluruh ruangan",chemical_ratio:"-",tools:["Tangga Besi Lipat","Kain Lap Microfiber Lembap","Kuas Cat Halus","Kran Air Mengalir"],steps:["Matikan AC melalui remote dan matikan MCB/saklar AC bila diperlukan.","Gunakan tangga besi lipat yang stabil di lantai datar.","Buka tutup cover depan unit indoor AC ke atas hingga terkunci.","Lepaskan kedua kisi filter saringan udara plastik secara perlahan agar debu tidak berjatuhan.","Cuci saringan filter di bawah kran air mengalir dari arah bagian dalam ke bagian luar.","Keringkan filter di tempat teduh (jangan dijemur langsung di bawah terik matahari).","Lap bodi casing dan louver/swing AC dengan kain microfiber lembap.","Pasang kembali filter yang sudah kering, tutup cover AC, dan nyalakan AC untuk pengetesan."],notes:"Pembersihan rutin mandiri dilakukan 2 minggu sekali. Deep cleaning / cuci evaporator dilakukan oleh teknisi AC vendor."},{id:10,title:"Membersihkan Dental Unit",category:"Klinis & Medis",icon:"\u{1F9B7}",color:"#E11D48",bg:"#FFF1F2",target:"Kursi dental pasien, spittoon bowl, meja instrumen, lampu dental, cover bawah",chemical_ratio:"Glass Cleaner (netral) / Karbol Disinfektan (mangkuk) / Bowl Cleaner (kerak)",tools:["Sarung Tangan Karet (Wajib)","Kain Lap Microfiber Biru (Khusus Bodi Luar)","Kain Lap Microfiber Merah (Khusus Mangkuk Spittoon)","Botol Spray Glass Cleaner (25ml + Air)","Botol Spray Karbol (20ml + Air)","Botol Spray Bowl Cleaner (100ml + Air, bila berkerak)"],steps:["Pastikan saklar utama dental unit dalam posisi OFF dan kursi dinaikkan ke posisi tertinggi untuk memudahkan akses.","Lap seluruh jok kulit, sandaran kepala, sandaran tangan, dan tiang lampu menggunakan lap microfiber biru dan cairan pembersih netral/glass cleaner.","Bersihkan mangkuk kumur (spittoon bowl) menggunakan kain lap merah dan cairan karbol disinfektan.","Jika ditemukan kerak sisa darah/plak membandel di mangkuk spittoon, gunakan sedikit bowl cleaner, sikat perlahan, lalu bilas hingga bersih.","Lap permukaan meja instrumen meja dokter. Bila ada instrumen steril, geser dengan hati-hati dan jangan mengotori instrumen.","Buka penutup cover bawah unit dental, lap debu yang terjangkau dan pastikan kondisi area selang dalam keadaan kering."],notes:"Kain lap warna MERAH hanya untuk mangkuk spittoon. Dilarang keras menggosok jok kulit dental unit dengan bahan kimia asam keras."},{id:11,title:"Membersihkan Wastafel & Keran Air",category:"Sanitasi Toilet",icon:"\u{1F6B0}",color:"#0284C7",bg:"#F0F9FF",target:"Wastafel keramik ruang tindakan, wastafel lorong, kran stainless steel, afur sink",chemical_ratio:"Bubuk Pembersih KIFA / Dishwashing Liquid Prokleen",tools:["Sponge Scotch Brite Hijau/Kuning","Kain Microfiber Hijau (Area Basah)","Bubuk Pembersih KIFA","Sikat Gigi Bekas / Sikat Kecil"],steps:["Basahi seluruh mangkuk wastafel keramik dan kran air stainless steel.","Taburkan sedikit Bubuk KIFA pada tapas sponge yang basah.","Gosok perlahan noda sabun dan kerak air putih pada kran stainless steel searah guratan logam hingga mengkilap kembali.","Gunakan sikat kecil untuk membersihkan sela leher kran dan sudut sambungan afur pembuangan wastafel.","Bilas mangkuk wastafel dan kran dengan air bersih mengalir hingga seluruh sisa residu bubuk KIFA hilang.","Keringkan kran air stainless steel dan cermin wastafel menggunakan lap microfiber hijau kering."],notes:"DILARANG KERAS menggunakan Toilet Bowl Cleaner pada kran air stainless steel karena menyebabkan karat dan warna menghitam permanen."},{id:12,title:"Membersihkan Kulkas Klinik & Pantry",category:"Elektronik & Utility",icon:"\u{1F9CA}",color:"#475569",bg:"#F8FAFC",target:"Kulkas pantry staf dan kulkas penyimpan obat/material medis (eksterior & rak)",chemical_ratio:"30 ml Dishwashing Liquid : 1 Liter Air Hangat + Lap Kering",tools:["Kain Microfiber Biru & Hijau","Dishwashing Liquid","Ember Air Hangat","Sponge Scotch Brite Halus"],steps:["Keluarkan makanan/minuman yang telah kedaluwarsa atau berbau tidak sedap.","Lepas rak-rak kaca/plastik interior kulkas yang dapat dilepas, cuci dengan sabun cuci piring dan bilas air bersih.","Seka dinding interior kulkas menggunakan kain microfiber yang dicelupkan ke dalam air hangat bersabun lembut.","Lap kering seluruh permukaan dalam kulkas dengan microfiber kering bersih.","Pasang kembali rak interior dan susun kembali bahan/makanan secara tertib dan rapi.","Bersihkan karet pintu kulkas dari noda hitam/jamur menggunakan sikat halus dan lap kering.","Lap bodi luar kulkas dan pegangan pintu menggunakan cairan pembersih netral."],notes:"Untuk kulkas penyimpanan obat/material dental, koordinasikan terlebih dahulu dengan Perawat/PIC Obat sebelum membersihkan rak interior."},{id:13,title:"Membersihkan Sofa & Kursi Ruang Tunggu",category:"Pembersihan Khusus",icon:"\u{1F6CB}\uFE0F",color:"#9333EA",bg:"#FAF5FF",target:"Sofa ruang tunggu pasien, kursi reception, kursi kerja tim",chemical_ratio:"Baking Soda (penghilang bau) / Pembersih Upholstery secukupnya",tools:["Vacuum Cleaner (Nozzle Kecil & Sikat Lembut)","Kain Microfiber Bersih","Sikat Bulu Halus","Ember Air Hangat","Cairan Pembersih Khusus Sofa (Upholstery Cleaner)","Baking Soda (Penghilang Bau Tidak Sedap)","Sarung Tangan Karet"],steps:["Kenali bahan sofa (kain fabric, beludru, kulit asli, atau kulit sintetis/oscar).","Sedot debu di seluruh permukaan, lipatan jahitan, dan celah bantal sofa menggunakan vacuum cleaner nozzle kecil.","Gunakan sikat bulu halus untuk mengangkat rambut atau kotoran yang menempel kuat di serat kain.","Jika ada noda bercak (kopi, makanan, tinta), semprotkan pembersih upholstery secukupnya, gosok memutar perlahan dengan microfiber.","Celupkan kain microfiber ke air hangat suam-suam kuku, peras hingga lembap, lalu seka seluruh permukaan sofa.","Lap kembali dengan kain kering bersih untuk menyerap sisa kelembapan.","Jika sofa berbau apek: taburkan baking soda tipis di permukaan, diamkan 20 menit, lalu sedot kembali dengan vacuum cleaner.","Biarkan sofa mengering secara alami di dalam ruangan berventilasi atau ber-AC."],notes:"Jangan menjemur sofa langsung di bawah terik matahari karena dapat menyebabkan bahan kulit pecah-pecah dan warna kain memudar."}],Nn=[{no:1,name:"Sapu Lidi Gagang",qty:"1 unit",area:"Area parkir mobil & motor klinik"},{no:2,name:"Sapu & Pengki Nagata 742s",qty:"1 set",area:"Area pantry, balkon, dak lantai atas, atau teras"},{no:3,name:"Lobby Duster Set 60 cm",qty:"1 set",area:"Lobby utama dan koridor/lorong klinik"},{no:4,name:"Lobby Duster Set 40 cm / Refill",qty:"1 unit",area:"Area ruang tindakan dokter"},{no:5,name:"Single Bucket Mop Wringer 20L",qty:"1 paket",area:"Ember pel dengan pemeras + tongkat pel"},{no:6,name:"Double Bucket Wringer Trolley 46L",qty:"1 paket",area:"Khusus klinik besar / klinik dengan fasilitas lift"},{no:7,name:"Kemoceng Bulu Lembut",qty:"1 unit",area:"Membersihkan debu kasar meja dan perabot"},{no:8,name:"Lap Microfiber Warna Merah",qty:"2 pcs",area:"Khusus area toilet, sanitasi, dan mangkuk spittoon"},{no:9,name:"Lap Microfiber Warna Biru",qty:"2 pcs",area:"Area umum kering (meja, kursi, bodi peralatan)"},{no:10,name:"Lap Microfiber Warna Hijau",qty:"2 pcs",area:"Area umum basah (pantry, wastafel luar)"},{no:11,name:"Kanebo / Synthetic Chamois",qty:"1 pcs",area:"Area umum, cermin, dan pengering kaca"},{no:12,name:"Caddy Bag Cleaning",qty:"1 unit",area:"Tas perlengkapan pinggang petugas cleaning"},{no:13,name:"Sponge Scotch Brite 3M",qty:"2 pcs",area:"Pembersihan noda lantai dan kerak ringan (spotting)"},{no:14,name:"Scotch-Brite 3M Sikat Toilet",qty:"1 unit",area:"Sikat khusus kloset dan leher angsa"},{no:15,name:"Kapi Gagang Karet / Scrapper",qty:"1 unit",area:"Pembersihan kerak lantai, noda permen karet, stiker"},{no:16,name:"Kuas Cat 4 Inch",qty:"2 pcs",area:"Pembersih sela keyboard, ventilasi, dan sudut sempit"},{no:17,name:"Warning Sign (Lantai Licin)",qty:"1 unit",area:"Digunakan saat mopping dan general cleaning"},{no:18,name:"Teleskopik + Pembersih Kaca/Langit",qty:"1 set",area:"Membersihkan jendela tinggi, plafon, dan sarang laba-laba"},{no:19,name:"Sarung Tangan Karet 3M",qty:"1 pasang",area:"Perlindungan tangan area toilet dan bahan kimia"},{no:20,name:"Botol Spray 500ml",qty:"5 unit",area:"Wadah larutan chemical (diberi label masing-masing)"},{no:21,name:"Pengharum Ruangan Stella Aerosol",qty:"1 kaleng",area:"Digunakan darurat apabila timbul bau menyengat"},{no:22,name:"Kleaner Gantungan Sapu & Pel",qty:"1 set",area:"Gantungan penyimpanan alat cleaning di janitor"},{no:23,name:"Floor Squeegee 75 cm",qty:"1 unit",area:"Pendorong genangan air lantai toilet/teras"},{no:24,name:"Sikat Botol Ikea",qty:"1 unit",area:"Di setiap ruang tindakan dokter gigi"},{no:25,name:"Scotch-Brite 3M Sikat Lantai ID-51",qty:"1 unit",area:"Sikat lantai kamar mandi dan teras"},{no:26,name:"Kamper Toilet / Urinal Block",qty:"1 pack",area:"Pengharum di toilet dan urinoir"},{no:27,name:"Plastik Sampah Hitam Sedang (60x100)",qty:"1 pack",area:"Tempat sampah indoor ruangan"},{no:28,name:"Plastik Sampah Hitam Besar (90x120)",qty:"1 pack",area:"Tempat sampah luar & penampung sampah akhir"},{no:29,name:"Plastik Sampah Medis Kuning (90x120)",qty:"1 pack",area:"Limbah medis padat B3 klinik"},{no:30,name:"Plastik Sampah Medis Kuning (40x60)",qty:"1 pack",area:"Tempat sampah medis kecil di ruang tindakan"},{no:31,name:"Tangga Besi Lipat",qty:"1 unit",area:"Alat bantu membersihkan plafon, lampu, dan AC"},{no:32,name:"Baskom Plastik Besar 55cm",qty:"1 unit",area:"Mencuci bola kids playground"}],qn=[{no:1,name:"Floor Cleaner Prokleen Lavender",size:"5 Liter",ratio:"25 ml : 1 Liter Air",function:"Pembersihan dan pengepelan rutin seluruh lantai ruangan"},{no:2,name:"Karbol Prokleen Pine Disinfektan",size:"5 Liter",ratio:"25 ml : 1 Liter Air (Mop) / 20 ml per botol spray",function:"Mengepel lantai toilet, disinfeksi area basah dan berbau tak sedap"},{no:3,name:"Glass Cleaner Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air / Siap pakai di botol spray",function:"Membersihkan pintu kaca, jendela, cermin, dan layar monitor"},{no:4,name:"Manual Dishwashing Liquid Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air",function:"Mencuci piring, gelas, sendok pantry, dan mencuci bola playground"},{no:5,name:"Hand Soap Prokleen Green Tea",size:"5 Liter",ratio:"Siap pakai (tanpa campur)",function:"Refill sabun pencuci tangan di wastafel pasien dan tim"},{no:6,name:"Liquid Detergent Prokleen",size:"5 Liter",ratio:"Sesuai kapasitas mesin",function:"Mencuci handuk, kain microfiber, dan seragam"},{no:7,name:"Toilet Bowl Cleaner Prokleen",size:"5 Liter",ratio:"100 ml per botol spray / Khusus kerak",function:"Pembersih kerak membandel pada kloset, wastafel keramik, dan nat keramik"},{no:8,name:"Bubuk Pembersih KIFA",size:"500 gram",ratio:"Taburkan tipis pada tapas basah",function:"Membersihkan kerak air pada kran air stainless steel, jet washer, dan sink"}],_t="fcms_sop_fc_ik_v1",At="fcms_sop_fc_tools_v1",Tt="fcms_sop_fc_chem_v1";function Ea(){try{let e=localStorage.getItem(_t);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK",e)}return JSON.parse(JSON.stringify(jn))}function Ft(e){localStorage.setItem(_t,JSON.stringify(e))}function Sa(){try{let e=localStorage.getItem(At);if(e)return JSON.parse(e)}catch(e){console.error("Error load tools",e)}return JSON.parse(JSON.stringify(Nn))}function Ct(e){localStorage.setItem(At,JSON.stringify(e))}function Fa(){try{let e=localStorage.getItem(Tt);if(e)return JSON.parse(e)}catch(e){console.error("Error load chem",e)}return JSON.parse(JSON.stringify(qn))}function $t(e){localStorage.setItem(Tt,JSON.stringify(e))}async function Ca(e){let t=Z(),a=ee(t,"sop","write"),i=ee(t,"sop","delete"),s="ik",b="",x="all",g="tools",o="",l=Ea(),A=Sa(),T=Fa();function C(){e.innerHTML=`
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
          <button id="btn-export-tools" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F4E5} Download Data Excel
          </button>
        </div>
      </div>

      <!-- Navigation Tabs -->
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto; flex-wrap:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none;">
        <button class="sop-tab-btn ${s==="ik"?"active":""}" data-tab="ik" style="${_(s==="ik")}">
          \u{1F4CB} Instruksi Kerja (IK) <span class="badge" style="${z(s==="ik")}">${l.length}</span>
        </button>
        <button class="sop-tab-btn ${s==="tools"?"active":""}" data-tab="tools" style="${_(s==="tools")}">
          \u{1F6E0}\uFE0F Alat & Chemical <span class="badge" style="${z(s==="tools")}">${A.length+T.length}</span>
        </button>
        <button class="sop-tab-btn ${s==="policy"?"active":""}" data-tab="policy" style="${_(s==="policy")}">
          \u{1F4D6} Pedoman & Ketentuan
        </button>
        <button class="sop-tab-btn ${s==="links"?"active":""}" data-tab="links" style="${_(s==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sop-content-area"></div>
    `,B(),P()}function _(d){return d?"border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function z(d){return d?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function B(){e.querySelectorAll(".sop-tab-btn").forEach(L=>{L.addEventListener("click",()=>{s=L.dataset.tab,C()})});let d=e.querySelector("#btn-reset-default");d&&d.addEventListener("click",()=>{ie("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja, Alat, dan Chemical ke data standar dokumen asli?",()=>{localStorage.removeItem(_t),localStorage.removeItem(At),localStorage.removeItem(Tt),l=Ea(),A=Sa(),T=Fa(),H("Data SOP Facility Care berhasil di-reset ke standar asli"),C()},"Konfirmasi Reset Standar")});let f=e.querySelector("#btn-print-sop");f&&f.addEventListener("click",()=>{window.print()});let w=e.querySelector("#btn-export-tools");w&&w.addEventListener("click",()=>{E()})}function P(){let d=e.querySelector("#sop-content-area");d&&(s==="ik"?c(d):s==="tools"?m(d):s==="policy"?F(d):s==="links"&&k(d))}function c(d){d.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="ik-search" class="form-control" placeholder="Cari pekerjaan, alat, atau chemical..." value="${b}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="cat-pills-wrap">
              <button class="cat-pill ${x==="all"?"active":""}" data-cat="all">Semua (${l.length})</button>
              <button class="cat-pill ${x==="Pembersihan Rutin"?"active":""}" data-cat="Pembersihan Rutin">Rutin (${l.filter(p=>p.category==="Pembersihan Rutin").length})</button>
              <button class="cat-pill ${x==="Sanitasi Toilet"?"active":""}" data-cat="Sanitasi Toilet">Toilet & Wastafel (${l.filter(p=>p.category==="Sanitasi Toilet").length})</button>
              <button class="cat-pill ${x==="Klinis & Medis"?"active":""}" data-cat="Klinis & Medis">Klinis Dental (${l.filter(p=>p.category==="Klinis & Medis").length})</button>
              <button class="cat-pill ${x==="Elektronik & Utility"?"active":""}" data-cat="Elektronik & Utility">Elektronik & AC (${l.filter(p=>p.category==="Elektronik & Utility").length})</button>
              <button class="cat-pill ${x==="Pembersihan Khusus"?"active":""}" data-cat="Pembersihan Khusus">Khusus (${l.filter(p=>p.category==="Pembersihan Khusus").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${l.length} dari ${l.length} Instruksi Kerja
            </div>
            ${a?`
            <button id="btn-add-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
              \u2795 Tambah IK Baru
            </button>
            `:""}
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="ik-cards-grid-container"></div>
    `;let f=document.createElement("style");f.innerHTML=`
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
    `,d.appendChild(f);let w=d.querySelector("#ik-cards-grid-container"),L=d.querySelector("#ik-count-label");function I(){let p=l.filter(S=>{let u=x==="all"||S.category===x,M=b.toLowerCase().trim(),R=!M||S.title.toLowerCase().includes(M)||S.target.toLowerCase().includes(M)||S.tools.some(K=>K.toLowerCase().includes(M))||S.steps.some(K=>K.toLowerCase().includes(M));return u&&R});if(L&&(L.textContent=`Menampilkan ${p.length} dari ${l.length} Instruksi Kerja`),p.length===0){w.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}w.innerHTML=`
        <div class="sop-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${p.map(S=>`
            <div class="card ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${S.bg||"#EFF6FF"}; border-radius:10px;">${S.icon||"\u{1F4CB}"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        ${S.id}. ${S.title}
                      </h3>
                      <span class="badge" style="background:${S.bg||"#EFF6FF"}; color:${S.color||"#2563EB"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${S.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    ${a?`
                    <button class="btn btn-sm btn-ghost btn-edit-ik" data-id="${S.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    `:""}
                    ${i?`
                    <button class="btn btn-sm btn-ghost btn-delete-ik" data-id="${S.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                    `:""}
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${S.color||"#2563EB"};">
                  <strong>Area:</strong> ${S.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Bahan Utama:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${S.tools.slice(0,3).map(u=>`<li style="margin-bottom:2px;">${u}</li>`).join("")}
                    ${S.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${S.tools.length-3} alat lainnya...</li>`:""}
                  </ul>
                </div>

                ${S.chemical_ratio&&S.chemical_ratio!=="-"?`
                  <div style="font-size:0.8rem; background:#FEF3C7; color:#92400E; padding:6px 10px; border-radius:6px; margin-bottom:14px; font-weight:600;">
                    \u{1F9EA} Takaran: ${S.chemical_ratio}
                  </div>
                `:""}
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${S.steps.length} Langkah Kerja
                </span>
                <div style="display:flex; gap:6px;">
                  <button class="btn btn-sm btn-primary btn-view-ik" data-id="${S.id}" style="padding:6px 14px; font-weight:600; border-radius:6px;">
                    Lihat Panduan \u2794
                  </button>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,w.querySelectorAll(".btn-view-ik").forEach(S=>{S.addEventListener("click",()=>{let u=parseInt(S.dataset.id,10),M=l.find(R=>R.id===u);M&&y(M)})}),w.querySelectorAll(".btn-edit-ik").forEach(S=>{S.addEventListener("click",()=>{let u=parseInt(S.dataset.id,10),M=l.find(R=>R.id===u);M&&v(M)})}),w.querySelectorAll(".btn-delete-ik").forEach(S=>{S.addEventListener("click",()=>{let u=parseInt(S.dataset.id,10),M=l.find(R=>R.id===u);M&&ie(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${M.title}"?`,()=>{l=l.filter(R=>R.id!==u),Ft(l),H("Instruksi Kerja berhasil dihapus"),C()},"Hapus Instruksi Kerja")})})}I();let j=d.querySelector("#ik-search");j&&j.addEventListener("input",p=>{b=p.target.value,I()}),d.querySelectorAll(".cat-pill").forEach(p=>{p.addEventListener("click",()=>{d.querySelectorAll(".cat-pill").forEach(S=>S.classList.remove("active")),p.classList.add("active"),x=p.dataset.cat,I()})});let $=d.querySelector("#btn-add-ik");$&&$.addEventListener("click",()=>{v()})}function y(d){X({title:`${d.icon||"\u{1F4CB}"} IK-${d.id}: ${d.title}`,size:"lg",content:`
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
                ${d.tools.map(f=>`<li>${f}</li>`).join("")}
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
                ${d.steps.map(f=>`<li style="margin-bottom:6px;">${f}</li>`).join("")}
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
      `,confirmText:"Tutup Panduan",onConfirm:(f,w)=>{w()}})}function v(d=null){if(!a)return;let f=!!d,w=d?d.tools.join(`
`):"",L=d?d.steps.join(`
`):"",I=`
      <form id="ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-ik-title" class="form-control" required value="${d?n(d.title):""}" placeholder="Contoh: Membersihkan Dental Unit" style="width:100%; border-radius:6px; height:36px;">
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
            <input type="text" id="form-ik-ratio" class="form-control" value="${d?n(d.chemical_ratio||""):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-ik-target" class="form-control" required value="${d?n(d.target):""}" placeholder="Contoh: Kursi pasien, meja instrumen, mangkuk spittoon" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Bahan yang Digunakan (1 item per baris) *</label>
          <textarea id="form-ik-tools" class="form-control" rows="4" placeholder="Kain Microfiber Biru&#10;Sarung Tangan Karet&#10;Botol Spray Glass Cleaner" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${n(w)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah-Langkah Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-ik-steps" class="form-control" rows="6" placeholder="1. Pastikan posisi saklar OFF.&#10;2. Lap jok dental unit dengan kain microfiber biru.&#10;3. Sikat mangkuk spittoon..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${n(L)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Penting / Hal yang Perlu Diperhatikan (Safety)</label>
          <textarea id="form-ik-notes" class="form-control" rows="2" placeholder="Catatan larangan atau peringatan keselamatan..." style="width:100%; border-radius:6px; font-size:0.85rem;">${d?n(d.notes||""):""}</textarea>
        </div>
      </form>
    `;X({title:f?`\u270F\uFE0F Edit IK-${d.id}: ${d.title}`:"\u2795 Tambah Instruksi Kerja Baru",content:I,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(j,$)=>{let p=j.querySelector("#form-ik-title").value.trim(),S=j.querySelector("#form-ik-icon").value.trim()||"\u{1F4CB}",u=j.querySelector("#form-ik-cat").value,M=j.querySelector("#form-ik-ratio").value.trim(),R=j.querySelector("#form-ik-target").value.trim(),K=j.querySelector("#form-ik-tools").value.trim(),O=j.querySelector("#form-ik-steps").value.trim(),D=j.querySelector("#form-ik-notes").value.trim();if(!p||!R||!K||!O){q("Harap lengkapi seluruh field wajib (*)");return}let Y=K.split(`
`).map(W=>W.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),J=O.split(`
`).map(W=>W.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),U={"Pembersihan Rutin":{color:"#2563EB",bg:"#EFF6FF"},"Sanitasi Toilet":{color:"#DC2626",bg:"#FEF2F2"},"Klinis & Medis":{color:"#E11D48",bg:"#FFF1F2"},"Elektronik & Utility":{color:"#0D9488",bg:"#F0FDFA"},"Pembersihan Khusus":{color:"#7C3AED",bg:"#F5F3FF"}}[u]||{color:"#2563EB",bg:"#EFF6FF"};if(f)d.title=p,d.icon=S,d.category=u,d.chemical_ratio=M,d.target=R,d.tools=Y,d.steps=J,d.notes=D,d.color=U.color,d.bg=U.bg,Ft(l),H(`Instruksi Kerja "${p}" berhasil diperbarui`);else{let W=l.length>0?Math.max(...l.map(V=>V.id))+1:1,Q={id:W,title:p,category:u,icon:S,color:U.color,bg:U.bg,target:R,chemical_ratio:M||"-",tools:Y,steps:J,notes:D};l.push(Q),Ft(l),H(`Instruksi Kerja baru IK-${W} berhasil ditambahkan`)}$(),C()}})}function m(d){d.innerHTML=`
      <!-- Sub-tab Switcher & Actions -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-tools-btn" class="btn btn-sm ${g==="tools"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9F9} Alat Kebersihan (${A.length} Item)
          </button>
          <button id="view-chem-btn" class="btn btn-sm ${g==="chemical"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9EA} Bahan Chemical (${T.length} Item)
          </button>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="tool-search" class="form-control" placeholder="Cari nama / peruntukan / rasio..." value="${o}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
            ${a?`
          <button id="btn-add-tool-chem" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u2795 ${g==="tools"?"Tambah Alat Baru":"Tambah Chemical Baru"}
          </button>
            `:""}
        </div>
      </div>

      <!-- Table Container -->
      <div id="tool-table-container"></div>
    `;let f=d.querySelector("#tool-table-container"),w=d.querySelector("#tool-search");function L(){let I=o.toLowerCase().trim();if(g==="tools"){let j=A.filter($=>!I||$.name.toLowerCase().includes(I)||$.area.toLowerCase().includes(I));f.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Standar Peralatan Kebersihan (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-primary">${j.length} Item</span>
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
                  ${j.map(($,p)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${p%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${$.no||p+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${n($.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#E0F2FE; color:#0369A1; font-weight:600;">${n($.qty)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${n($.area)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                    ${a?`
                        <button class="btn btn-sm btn-ghost btn-edit-tool" data-no="${$.no}" title="Edit Alat" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                    `:""}
                    ${i?`
                        <button class="btn btn-sm btn-ghost btn-delete-tool" data-no="${$.no}" title="Hapus Alat" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    `:""}
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,f.querySelectorAll(".btn-edit-tool").forEach($=>{$.addEventListener("click",()=>{let p=parseInt($.dataset.no,10),S=A.find(u=>u.no===p);S&&r(S)})}),f.querySelectorAll(".btn-delete-tool").forEach($=>{$.addEventListener("click",()=>{let p=parseInt($.dataset.no,10),S=A.find(u=>u.no===p);S&&ie(`Apakah Anda yakin ingin menghapus alat: "${S.name}"?`,()=>{A=A.filter(u=>u.no!==p),A.forEach((u,M)=>{u.no=M+1}),Ct(A),H("Alat kebersihan berhasil dihapus"),C()},"Hapus Alat Kebersihan")})})}else{let j=T.filter($=>!I||$.name.toLowerCase().includes(I)||$.function.toLowerCase().includes(I));f.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Bahan Kimia & Standar Takaran (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-success">${j.length} Chemical</span>
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
                  ${j.map(($,p)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${p%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${$.no||p+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${n($.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#FEF3C7; color:#92400E; font-weight:600;">${n($.size)}</span></td>
                      <td style="padding:10px 14px;"><span style="background:#EFF6FF; color:#1E40AF; padding:3px 8px; border-radius:6px; font-weight:600; font-size:0.82rem;">${n($.ratio)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${n($.function)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                    ${a?`
                        <button class="btn btn-sm btn-ghost btn-edit-chem" data-no="${$.no}" title="Edit Chemical" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                    `:""}
                    ${i?`
                        <button class="btn btn-sm btn-ghost btn-delete-chem" data-no="${$.no}" title="Hapus Chemical" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    `:""}
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,f.querySelectorAll(".btn-edit-chem").forEach($=>{$.addEventListener("click",()=>{let p=parseInt($.dataset.no,10),S=T.find(u=>u.no===p);S&&h(S)})}),f.querySelectorAll(".btn-delete-chem").forEach($=>{$.addEventListener("click",()=>{let p=parseInt($.dataset.no,10),S=T.find(u=>u.no===p);S&&ie(`Apakah Anda yakin ingin menghapus chemical: "${S.name}"?`,()=>{T=T.filter(u=>u.no!==p),T.forEach((u,M)=>{u.no=M+1}),$t(T),H("Bahan chemical berhasil dihapus"),C()},"Hapus Bahan Chemical")})})}}L(),d.querySelector("#view-tools-btn").addEventListener("click",()=>{g="tools",m(d)}),d.querySelector("#view-chem-btn").addEventListener("click",()=>{g="chemical",m(d)}),w&&w.addEventListener("input",I=>{o=I.target.value,L()}),d.querySelector("#btn-add-tool-chem").addEventListener("click",()=>{g==="tools"?r():h()})}function r(d=null){if(!a)return;let f=!!d,w=`
      <form id="tool-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Alat / Item *</label>
          <input type="text" id="form-tool-name" class="form-control" required value="${d?n(d.name):""}" placeholder="Contoh: Lobby Duster Set 60 cm" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kuantitas Standar *</label>
          <input type="text" id="form-tool-qty" class="form-control" required value="${d?n(d.qty):"1 unit"}" placeholder="Contoh: 1 set, 2 pcs, 1 paket" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Peruntukan Area & Kegunaan *</label>
          <textarea id="form-tool-area" class="form-control" rows="3" required placeholder="Contoh: Lobby utama dan koridor/lorong klinik" style="width:100%; border-radius:6px; font-size:0.85rem;">${d?n(d.area):""}</textarea>
        </div>
      </form>
    `;X({title:f?`\u270F\uFE0F Edit Alat: ${d.name}`:"\u2795 Tambah Alat Kebersihan Baru",content:w,confirmText:"Simpan",onConfirm:(L,I)=>{let j=L.querySelector("#form-tool-name").value.trim(),$=L.querySelector("#form-tool-qty").value.trim(),p=L.querySelector("#form-tool-area").value.trim();if(!j||!$||!p){q("Harap lengkapi semua kolom");return}if(f)d.name=j,d.qty=$,d.area=p,Ct(A),H(`Alat "${j}" berhasil diperbarui`);else{let S=A.length>0?Math.max(...A.map(u=>u.no))+1:1;A.push({no:S,name:j,qty:$,area:p}),Ct(A),H(`Alat baru "${j}" berhasil ditambahkan`)}I(),C()}})}function h(d=null){if(!a)return;let f=!!d,w=`
      <form id="chem-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Produk Chemical *</label>
          <input type="text" id="form-chem-name" class="form-control" required value="${d?n(d.name):""}" placeholder="Contoh: Floor Cleaner Prokleen Lavender" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kemasan Standar *</label>
            <input type="text" id="form-chem-size" class="form-control" required value="${d?n(d.size):"5 Liter"}" placeholder="Contoh: 5 Liter, 500 gram" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Rasio & Takaran Standar *</label>
            <input type="text" id="form-chem-ratio" class="form-control" required value="${d?n(d.ratio):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Fungsi & Cara Penggunaan *</label>
          <textarea id="form-chem-function" class="form-control" rows="3" required placeholder="Contoh: Pembersihan dan pengepelan rutin seluruh lantai ruangan" style="width:100%; border-radius:6px; font-size:0.85rem;">${d?n(d.function):""}</textarea>
        </div>
      </form>
    `;X({title:f?`\u270F\uFE0F Edit Chemical: ${d.name}`:"\u2795 Tambah Chemical Baru",content:w,confirmText:"Simpan",onConfirm:(L,I)=>{let j=L.querySelector("#form-chem-name").value.trim(),$=L.querySelector("#form-chem-size").value.trim(),p=L.querySelector("#form-chem-ratio").value.trim(),S=L.querySelector("#form-chem-function").value.trim();if(!j||!$||!p||!S){q("Harap lengkapi semua kolom");return}if(f)d.name=j,d.size=$,d.ratio=p,d.function=S,$t(T),H(`Chemical "${j}" berhasil diperbarui`);else{let u=T.length>0?Math.max(...T.map(M=>M.no))+1:1;T.push({no:u,name:j,size:$,ratio:p,function:S}),$t(T),H(`Chemical baru "${j}" berhasil ditambahkan`)}I(),C()}})}function F(d){d.innerHTML=`
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
    `}function k(d){d.innerHTML=`
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
    `}async function E(){try{let{downloadExcel:d}=await Promise.resolve().then(()=>(de(),fe)),f=l.map(L=>({"No IK":`IK-${L.id}`,"Judul Pekerjaan":L.title,Kategori:L.category,"Sasaran Target Area":L.target,"Rasio Chemical":L.chemical_ratio||"-","Alat & Bahan":L.tools.join(", "),"Langkah Kerja":L.steps.map((I,j)=>`${j+1}. ${I}`).join(" | "),"Catatan Safety":L.notes||"-"})),w=new Date().toISOString().slice(0,10);d(f,`SOP_Facility_Care_Instruksi_Kerja_${w}`),H("Data Excel Instruksi Kerja berhasil diunduh")}catch(d){console.error(d),q("Gagal mendownload data Excel: "+d.message)}}function n(d){return d?String(d).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}C()}ce();var Rn=[{id:1,title:"Pengamanan Area & Aset Perusahaan",category:"Pengamanan Fisik & Patroli",icon:"\u{1F6E1}\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Seluruh area dalam & luar Head Quarter dan Klinik Cabang Audy Dental",tools:["Google Form Checklist Monitoring Klinik: https://forms.gle/1EAvkqDJVdctXEGn7","Senter Patroli","Komputer dengan Akses Monitor CCTV (Live View)","Alat Komunikasi / Handphone / HT"],steps:["Lakukan Safety Patrol secara rutin mengelilingi seluruh area sesuai jadwal yang telah ditentukan (Patrol 1, 2, dan 3).","Catat hasil temuan dan kondisi keamanan secara berkala ke dalam Google Form Safety Patrol resmi.","Awasi kendaraan yang terparkir di area luar klinik; pastikan khusus sepeda motor dalam keadaan terkunci ganda.","Pantau situasi lingkungan dan area sensitif secara berkala melalui layar monitor CCTV saat berada di pos security.","Segera lakukan tindakan persuasif ramah apabila ditemukan potensi gangguan keamanan, dan segera hubungi pihak Kepolisian terdekat bila situasi darurat tidak dapat diatasi.","Bila ditemukan permasalahan fisik, kerusakan fasilitas, atau kendala di area klinik, segera dokumentasikan (foto) dan laporkan langsung ke SPV Klinik dan Leader Facility Care.","Disaat malam hari atau saat klinik telah tutup operasional, pastikan seluruh AC dan lampu ruangan dalam keadaan mati, kecuali lampu keamanan tertentu yang memang harus dibiarkan menyala terang."],notes:"Prioritaskan keselamatan diri dan tindakan persuasif. Dilarang meninggalkan area pengawasan tanpa personil pengganti yang jelas."},{id:2,title:"Penertiban & Pengaturan Parkir Kendaraan",category:"Operasional Parkir & Luar",icon:"\u{1F17F}\uFE0F",color:"#D97706",bg:"#FFFBEB",target:"Area parkir mobil, motor, akses keluar masuk, dan teras depan klinik",tools:["Lampu Lalu Lintas (Lampu Lalin)","Cone / Kerucut Pembatas Parkir","Peluit Parkir"],steps:["Berdiri di posisi strategis dan bantu mengarahkan setiap kendaraan pasien atau tamu yang datang secara tepat dan rapi ke slot parkir yang tersedia.","Berikan informasi dan solusi dengan sikap ramah serta santun apabila seluruh slot parkir klinik dalam kondisi penuh.","Dalam kondisi tertentu (bila ada mobil yang terhalang mobil lain dan harus dipindahkan, sementara pasien sedang dalam tindakan perawatan): bantu memindahkan kendaraan HANYA JIKA petugas security memiliki kemampuan mengemudi dan memiliki SIM aktif.","Apabila petugas security tidak memiliki kemampuan mengemudikan mobil: sampaikan permohonan maaf dengan tulus kepada pasien/tamu, minta untuk menunggu sejenak, dan segera minta bantuan kepada staf Resepsionis untuk menghubungi pemilik mobil yang menghalangi.","Bantu pasien atau tamu yang akan keluar parkir dengan memberikan aba-aba dan arahan lalu lintas jalan raya yang aman dan tepat."],notes:"Dilarang memaksakan memindahkan mobil pasien jika tidak memiliki lisensi mengemudi demi menghindari risiko gesekan atau kerusakan kendaraan."},{id:3,title:"Penerimaan Tamu & Pasien Klinik",category:"Layanan & Hospitalitas",icon:"\u{1F91D}",color:"#059669",bg:"#ECFDF5",target:"Pintu masuk utama klinik, lobby, area resepsionis, dan pos penerimaan tamu",tools:["Alat Tulis Kantor (Pulpen)","Buku Tamu / Log Book Keamanan","Alat Komunikasi / Handphone"],steps:['Berdiri dengan tegap dan sigap, bukakan pintu untuk setiap pasien/tamu yang datang dengan menerapkan budaya 5S (Senyum, Salam, Sapa, Sopan, Santun) dan menyapa: "Selamat pagi / siang / sore, ada yang bisa kami bantu?".',"Setelah pasien atau tamu memberitahukan tujuannya, persilakan dan arahkan pasien langsung menuju meja Resepsionis.","Apabila yang datang adalah tamu instansi, vendor, atau pihak Manajemen/BOD: persilakan mengisi Buku Tamu (Log Book) terlebih dahulu dan segera hubungi staf/karyawan yang ingin ditemui.",'Apabila pasien langsung melangkah ke arah tujuannya, berikan gestur tubuh hormat menyilangkan kedua tangan di depan, dan jika pasien/tamu mengucapkan terima kasih, jawab dengan ramah: "Dengan Senang Hati".',"Dalam kondisi khusus (pasien lansia, ibu hamil, atau pengguna kursi roda): petugas security wajib mendampingi, mengawal, dan membantu mendorong kursi roda sampai ke depan meja resepsionis.",'Buka kembali pintu untuk setiap pasien yang telah selesai melakukan perawatan saat hendak keluar klinik sambil mengucapkan ucapan hangat: "Terima Kasih, semoga lekas sehat kembali".'],notes:"Jaga senyum tulus dan keramahan sebagai garda terdepan representasi citra Audy Dental."}],On=[{no:1,type:"Aset Tetap",item:"Rumah & Bangunan",description:"Gedung klinik, pos satpam, pagar, kanopi, teras beton",pic:"Building Maintenance & Security"},{no:2,type:"Aset Tetap",item:"Genset & Lift",description:"Mesin generator listrik cadangan dan elevator penumpang tetap",pic:"Teknisi & Security"},{no:3,type:"Aset Tetap",item:"AC Central & Split Tetap",description:"Outdoor dan indoor AC yang terpasang permanen di dinding/plafon",pic:"Teknisi AC & Security"},{no:4,type:"Aset Tetap",item:"Dental Unit Permanen",description:"Kursi dental unit terpasang instalasi air dan kompresor",pic:"Tim Klinik & Security"},{no:5,type:"Aset Tetap",item:"Instalasi Listrik, Pipa & Jaringan IT Permanen",description:"Panel MCB, kabel fiber optik, pipa drainase, server rack",pic:"IT & Teknisi"},{no:6,type:"Aset Tetap",item:"Furniture Melekat (Built-in)",description:"Meja resepsionis cor/melekat, partisi kaca mati, kitchen set pantry",pic:"General Support"},{no:7,type:"Aset Bergerak",item:"Komputer, Laptop & Printer",description:"PC kasir/resepsionis, laptop staf, printer rontgen/dokumen",pic:"Staf Terkait & IT"},{no:8,type:"Aset Bergerak",item:"Kendaraan Operasional",description:"Mobil operasional dan sepeda motor inventaris perusahaan",pic:"Security & Driver"},{no:9,type:"Aset Bergerak",item:"Furniture Lepasan",description:"Meja kerja, kursi resepsionis, kursi tunggu pasien, lemari file",pic:"Facility Care & Staf"},{no:10,type:"Aset Bergerak",item:"Hand Tools, Alat Kerja & Alat Dental Portabel",description:"Handpiece dental, alat rontgen portable, toolbox teknisi",pic:"Dokter/Perawat & Teknisi"},{no:11,type:"Low Value Asset",item:"Peralatan Kantor (Telepon, ATK)",description:"Pesawat telepon, kalkulator, stapler, perlengkapan meja",pic:"Resepsionis"},{no:12,type:"Low Value Asset",item:"Alat-alat Kebersihan",description:"Lobby duster, ember wringer pel, sapu, tangga, botol spray",pic:"Facility Care"},{no:13,type:"Low Value Asset",item:"Perlengkapan Security",description:"Lampu lalin, cone parkir, peluit, senter patroli, jas hujan",pic:"Security"}],$a={cabang:[{zone:"Zona Hijau",level:"Akses Publik Bebas",color:"#059669",bg:"#ECFDF5",rooms:["Area Parkir Kendaraan","Lobby Lantai 1","Teras Luar Klinik"],desc:"Area dengan pengamanan umum yang bebas diakses pasien dan tamu setelah pengarahan security."},{zone:"Zona Kuning",level:"Akses Terbatas / Pendampingan",color:"#D97706",bg:"#FFFBEB",rooms:["Koridor Ruang Tindakan","Tangga & Lift Klinik","Ruang Tunggu Lantai 2 & 3","Kids Playground","Toilet Lantai 1 & 2","Mushola Staf & Pasien"],desc:"Area terbatas, tamu wajib didampingi petugas atau staf klinik terkait."},{zone:"Zona Merah",level:"Area Terlarang / Restricted",color:"#DC2626",bg:"#FEF2F2",rooms:["Area Meja Receptionist / Kasir","Seluruh Ruang Dental Unit (DU 101, 102, 201, 202, 301, 302)","Ruang Sterilisasi Alat","Ruang Rontgen Radiologi","Ruang Suction & Kompresor","Ruang Tim Staf & Gudang Material Medis"],desc:"Area sangat sensitif, hanya staf medis bertugas dan atasan pemilik area. Tamu wajib izin Manajemen/BOD dan didampingi penuh."}],hq:[{floor:"Lantai 1",hijau:["Area Parkir Mobil/Motor","Lobby Resepsionis Lantai 1"],kuning:["Seluruh Ruang Meeting Lantai 1","Ruang Training Lantai 1","Koridor Lift & Toilet"],merah:["Ruang Customer Services","Gudang Utama Lantai 1","Gudang Facility Care","Gudang Teknisi & Ruang Teknisi","Pos Satpam & Ruang Genset"]},{floor:"Lantai 2",hijau:["-"],kuning:["Ruang Meeting Lantai 2","Koridor Lift & Toilet","Pantry"],merah:["Ruang CEO","Ruang Sekretaris CEO","Ruang Human Capital","Ruang Dentist Relation","Ruang General Support","Ruang Eks. Patrel"]},{floor:"Lantai 3",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry","Mushola & Tempat Wudhu","Ruang Santai Karyawan"],merah:["Ruang Marketing","Studio Konten Marketing","Gudang Marketing","Ruang Patient Relation","Ruang OPS","Ruang Training Lantai 3","Gudang Lantai 3"]},{floor:"Lantai 4",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry"],merah:["Ruang Finance & Accounting","Ruang VP Finance & Accounting","Ruang Audit Internal","Ruang Training Lantai 4","Gudang Lantai 4"]},{floor:"Lantai 5",hijau:["-"],kuning:["Lobby Lantai 5","Pantry Lantai 5","Toilet Lantai 5","Ruang Training Besar Lantai 5"],merah:["Area Dak Toren Air & Rooftop"]}]},Pt="fcms_sop_sec_ik_v1",Lt="fcms_sop_sec_assets_v1";function _a(){try{let e=localStorage.getItem(Pt);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK Sec",e)}return JSON.parse(JSON.stringify(Rn))}function Bt(e){localStorage.setItem(Pt,JSON.stringify(e))}function Aa(){try{let e=localStorage.getItem(Lt);if(e)return JSON.parse(e)}catch(e){console.error("Error load assets sec",e)}return JSON.parse(JSON.stringify(On))}function Dt(e){localStorage.setItem(Lt,JSON.stringify(e))}async function Ta(e){let t=Z(),a=ee(t,"sop","write"),i=ee(t,"sop","delete"),s="ik",b="",x="all",g="",o="cabang",l=_a(),A=Aa();function T(){e.innerHTML=`
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
          <button id="btn-export-sec" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u{1F4E5} Download Data Excel
          </button>
        </div>
      </div>

      <!-- Navigation Tabs -->
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto; flex-wrap:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none;">
        <button class="sec-tab-btn ${s==="ik"?"active":""}" data-tab="ik" style="${C(s==="ik")}">
          \u{1F6E1}\uFE0F Instruksi Kerja (IK) <span class="badge" style="${_(s==="ik")}">${l.length}</span>
        </button>
        <button class="sec-tab-btn ${s==="zonasi"?"active":""}" data-tab="zonasi" style="${C(s==="zonasi")}">
          \u{1F5FA}\uFE0F Zonasi Akses & Denah
        </button>
        <button class="sec-tab-btn ${s==="assets"?"active":""}" data-tab="assets" style="${C(s==="assets")}">
          \u{1F4E6} Klasifikasi Aset <span class="badge" style="${_(s==="assets")}">${A.length}</span>
        </button>
        <button class="sec-tab-btn ${s==="policy"?"active":""}" data-tab="policy" style="${C(s==="policy")}">
          \u{1F4CB} Shift, Patrol & Insiden
        </button>
        <button class="sec-tab-btn ${s==="links"?"active":""}" data-tab="links" style="${C(s==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sec-content-area"></div>
    `,z(),B()}function C(n){return n?"border:none; background:#1E40AF; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function _(n){return n?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function z(){e.querySelectorAll(".sec-tab-btn").forEach(w=>{w.addEventListener("click",()=>{s=w.dataset.tab,T()})});let n=e.querySelector("#btn-reset-sec-default");n&&n.addEventListener("click",()=>{ie("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja dan Klasifikasi Aset Security ke standar dokumen asli?",()=>{localStorage.removeItem(Pt),localStorage.removeItem(Lt),l=_a(),A=Aa(),H("Data SOP Security berhasil di-reset ke standar dokumen"),T()},"Reset ke Standar Dokumen")});let d=e.querySelector("#btn-print-sec");d&&d.addEventListener("click",()=>{window.print()});let f=e.querySelector("#btn-export-sec");f&&f.addEventListener("click",()=>{k()})}function B(){let n=e.querySelector("#sec-content-area");n&&(s==="ik"?P(n):s==="zonasi"?v(n):s==="assets"?m(n):s==="policy"?h(n):s==="links"&&F(n))}function P(n){n.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="sec-ik-search" class="form-control" placeholder="Cari instruksi kerja security..." value="${b}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="sec-cat-pills-wrap">
              <button class="sec-cat-pill ${x==="all"?"active":""}" data-cat="all">Semua (${l.length})</button>
              <button class="sec-cat-pill ${x==="Pengamanan Fisik & Patroli"?"active":""}" data-cat="Pengamanan Fisik & Patroli">Patroli & Fisik (${l.filter($=>$.category==="Pengamanan Fisik & Patroli").length})</button>
              <button class="sec-cat-pill ${x==="Operasional Parkir & Luar"?"active":""}" data-cat="Operasional Parkir & Luar">Parkir & Luar (${l.filter($=>$.category==="Operasional Parkir & Luar").length})</button>
              <button class="sec-cat-pill ${x==="Layanan & Hospitalitas"?"active":""}" data-cat="Layanan & Hospitalitas">Tamu & Pasien (${l.filter($=>$.category==="Layanan & Hospitalitas").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="sec-ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${l.length} Instruksi Kerja
            </div>
            ${a?`
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
    `,n.appendChild(d);let f=n.querySelector("#sec-ik-cards-grid-container"),w=n.querySelector("#sec-ik-count-label");function L(){let $=l.filter(p=>{let S=x==="all"||p.category===x,u=b.toLowerCase().trim(),M=!u||p.title.toLowerCase().includes(u)||p.target.toLowerCase().includes(u)||p.tools.some(R=>R.toLowerCase().includes(u))||p.steps.some(R=>R.toLowerCase().includes(u));return S&&M});if(w&&(w.textContent=`Menampilkan ${$.length} dari ${l.length} Instruksi Kerja`),$.length===0){f.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}f.innerHTML=`
        <div class="sec-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${$.map(p=>`
            <div class="card sec-ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${p.bg||"#EFF6FF"}; border-radius:10px;">${p.icon||"\u{1F6E1}\uFE0F"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        IK-${String(p.id).padStart(2,"0")}. ${p.title}
                      </h3>
                      <span class="badge" style="background:${p.bg||"#EFF6FF"}; color:${p.color||"#1E40AF"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${p.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    ${a?`
                    <button class="btn btn-sm btn-ghost btn-edit-sec-ik" data-id="${p.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    `:""}
                    ${i?`
                    <button class="btn btn-sm btn-ghost btn-delete-sec-ik" data-id="${p.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                    `:""}
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${p.color||"#1E40AF"};">
                  <strong>Area Sasaran:</strong> ${p.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Perlengkapan:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${p.tools.slice(0,3).map(S=>`<li style="margin-bottom:2px;">${E(S)}</li>`).join("")}
                    ${p.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${p.tools.length-3} perlengkapan lainnya...</li>`:""}
                  </ul>
                </div>
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${p.steps.length} Langkah Prosedur
                </span>
                <button class="btn btn-sm btn-primary btn-view-sec-ik" data-id="${p.id}" style="padding:6px 14px; font-weight:600; border-radius:6px; background:#1E40AF; border-color:#1E40AF;">
                  Lihat Panduan Lengkap \u2794
                </button>
              </div>
            </div>
          `).join("")}
        </div>
      `,f.querySelectorAll(".btn-view-sec-ik").forEach(p=>{p.addEventListener("click",()=>{let S=parseInt(p.dataset.id,10),u=l.find(M=>M.id===S);u&&c(u)})}),f.querySelectorAll(".btn-edit-sec-ik").forEach(p=>{p.addEventListener("click",()=>{let S=parseInt(p.dataset.id,10),u=l.find(M=>M.id===S);u&&y(u)})}),f.querySelectorAll(".btn-delete-sec-ik").forEach(p=>{p.addEventListener("click",()=>{let S=parseInt(p.dataset.id,10),u=l.find(M=>M.id===S);u&&ie(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${u.title}"?`,()=>{l=l.filter(M=>M.id!==S),Bt(l),H("Instruksi Kerja Security berhasil dihapus"),T()},"Hapus Instruksi Kerja")})})}L();let I=n.querySelector("#sec-ik-search");I&&I.addEventListener("input",$=>{b=$.target.value,L()}),n.querySelectorAll(".sec-cat-pill").forEach($=>{$.addEventListener("click",()=>{n.querySelectorAll(".sec-cat-pill").forEach(p=>p.classList.remove("active")),$.classList.add("active"),x=$.dataset.cat,L()})});let j=n.querySelector("#btn-add-sec-ik");j&&j.addEventListener("click",()=>{y()})}function c(n){X({title:`${n.icon||"\u{1F6E1}\uFE0F"} IK-${String(n.id).padStart(2,"0")}: ${n.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${n.bg||"#EFF6FF"}; color:${n.color||"#1E40AF"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${n.category}</span>
            <span>Kode: IK-SEC-${String(n.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${n.color||"#1E40AF"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${n.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Perlengkapan yang Disiapkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${n.tools.map(d=>{if(d.includes("http")){let f=d.split(": ");return`<li>${f[0]}: <a href="${f.slice(1).join(": ")}" target="_blank" rel="noopener" style="color:#2563EB; font-weight:600; text-decoration:underline;">Buka Tautan Form \u2794</a></li>`}return`<li>${E(d)}</li>`}).join("")}
              </ul>
            </div>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah & Prosedur Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${n.steps.map(d=>`<li style="margin-bottom:6px;">${E(d)}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${n.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Khusus & Ketentuan Penting:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${E(n.notes)}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(d,f)=>{f()}})}function y(n=null){if(!a)return;let d=!!n,f=n?n.tools.join(`
`):"",w=n?n.steps.join(`
`):"",L=`
      <form id="sec-ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-sec-ik-title" class="form-control" required value="${n?E(n.title):""}" placeholder="Contoh: Pengamanan Area & Aset Perusahaan" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-sec-ik-icon" class="form-control" value="${n?n.icon:"\u{1F6E1}\uFE0F"}" placeholder="\u{1F6E1}\uFE0F" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
          <select id="form-sec-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Pengamanan Fisik & Patroli" ${n?.category==="Pengamanan Fisik & Patroli"?"selected":""}>Pengamanan Fisik & Patroli</option>
            <option value="Operasional Parkir & Luar" ${n?.category==="Operasional Parkir & Luar"?"selected":""}>Operasional Parkir & Luar</option>
            <option value="Layanan & Hospitalitas" ${n?.category==="Layanan & Hospitalitas"?"selected":""}>Layanan & Hospitalitas</option>
          </select>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-sec-ik-target" class="form-control" required value="${n?E(n.target):""}" placeholder="Contoh: Seluruh area luar klinik dan akses parkir" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Perlengkapan (1 per baris) *</label>
          <textarea id="form-sec-ik-tools" class="form-control" rows="3" placeholder="Senter Patroli&#10;Lampu Lalin&#10;Peluit" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${E(f)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah & Prosedur Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-sec-ik-steps" class="form-control" rows="5" placeholder="1. Lakukan Safety Patrol sesuai jadwal.&#10;2. Arahkan kendaraan pasien dengan ramah..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${E(w)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Khusus / Ketentuan Penting</label>
          <textarea id="form-sec-ik-notes" class="form-control" rows="2" placeholder="Peringatan keselamatan atau pedoman khusus..." style="width:100%; border-radius:6px; font-size:0.85rem;">${n?E(n.notes||""):""}</textarea>
        </div>
      </form>
    `;X({title:d?`\u270F\uFE0F Edit IK-${n.id}: ${n.title}`:"\u2795 Tambah Instruksi Kerja Security Baru",content:L,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(I,j)=>{let $=I.querySelector("#form-sec-ik-title").value.trim(),p=I.querySelector("#form-sec-ik-icon").value.trim()||"\u{1F6E1}\uFE0F",S=I.querySelector("#form-sec-ik-cat").value,u=I.querySelector("#form-sec-ik-target").value.trim(),M=I.querySelector("#form-sec-ik-tools").value.trim(),R=I.querySelector("#form-sec-ik-steps").value.trim(),K=I.querySelector("#form-sec-ik-notes").value.trim();if(!$||!u||!M||!R){q("Harap lengkapi semua kolom bertanda bintang (*)");return}let O=M.split(`
`).map(G=>G.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),D=R.split(`
`).map(G=>G.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),J={"Pengamanan Fisik & Patroli":{color:"#2563EB",bg:"#EFF6FF"},"Operasional Parkir & Luar":{color:"#D97706",bg:"#FFFBEB"},"Layanan & Hospitalitas":{color:"#059669",bg:"#ECFDF5"}}[S]||{color:"#1E40AF",bg:"#EFF6FF"};if(d)n.title=$,n.icon=p,n.category=S,n.target=u,n.tools=O,n.steps=D,n.notes=K,n.color=J.color,n.bg=J.bg,Bt(l),H(`Instruksi Kerja "${$}" berhasil diperbarui`);else{let G=l.length>0?Math.max(...l.map(W=>W.id))+1:1,U={id:G,title:$,category:S,icon:p,color:J.color,bg:J.bg,target:u,tools:O,steps:D,notes:K};l.push(U),Bt(l),H(`Instruksi Kerja baru IK-${G} berhasil ditambahkan`)}j(),T()}})}function v(n){n.innerHTML=`
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
    `;let d=n.querySelector("#zonasi-content-container");o==="cabang"?d.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
          ${$a.cabang.map(f=>`
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; border-bottom:1px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.05rem; font-weight:700; color:${f.color}; display:flex; align-items:center; gap:8px;">
                  <span style="width:12px; height:12px; border-radius:50%; background:${f.color}; display:inline-block;"></span>
                  ${f.zone}
                </h3>
                <span class="badge" style="background:${f.bg}; color:${f.color}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:6px;">
                  ${f.level}
                </span>
              </div>
              <p style="font-size:0.82rem; color:var(--text-3); margin-bottom:14px; line-height:1.5;">${f.desc}</p>
              <div>
                <strong style="font-size:0.8rem; text-transform:uppercase; color:var(--text-2); display:block; margin-bottom:6px;">Daftar Ruangan Terkait:</strong>
                <ul style="margin:0 0 0 16px; padding:0; font-size:0.85rem; color:var(--text-1); line-height:1.7;">
                  ${f.rooms.map(w=>`<li>${w}</li>`).join("")}
                </ul>
              </div>
            </div>
          `).join("")}
        </div>
      `:d.innerHTML=`
        <div style="display:flex; flex-direction:column; gap:16px;">
          ${$a.hq.map(f=>`
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; border-bottom:2px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1); display:flex; align-items:center; gap:8px;">
                  <span>\u{1F3E2}</span> Head Quarter \u2014 ${f.floor}
                </h3>
              </div>
              <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:14px;">
                <!-- Zona Hijau -->
                <div style="background:#ECFDF5; border:1px solid #A7F3D0; border-radius:8px; padding:12px;">
                  <strong style="color:#065F46; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E2} ZONA HIJAU (Publik):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#064E3B; line-height:1.6;">
                    ${f.hijau.map(w=>`<li>${w}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Kuning -->
                <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:8px; padding:12px;">
                  <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E1} ZONA KUNING (Terbatas):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#78350F; line-height:1.6;">
                    ${f.kuning.map(w=>`<li>${w}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Merah -->
                <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:8px; padding:12px;">
                  <strong style="color:#991B1B; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F534} ZONA MERAH (Restricted):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#7F1D1D; line-height:1.6;">
                    ${f.merah.map(w=>`<li>${w}</li>`).join("")}
                  </ul>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,n.querySelector("#view-cabang-btn").addEventListener("click",()=>{o="cabang",v(n)}),n.querySelector("#view-hq-btn").addEventListener("click",()=>{o="hq",v(n)})}function m(n){n.innerHTML=`
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
            ${a?`
          <button id="btn-add-sec-asset" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
            \u2795 Tambah Aset Baru
          </button>
            `:""}
        </div>
      </div>

      <div id="sec-asset-table-container"></div>
    `;let d=n.querySelector("#sec-asset-table-container"),f=n.querySelector("#sec-asset-search");function w(){let I=g.toLowerCase().trim(),j=A.filter($=>!I||$.item.toLowerCase().includes(I)||$.type.toLowerCase().includes(I)||$.description.toLowerCase().includes(I)||$.pic&&$.pic.toLowerCase().includes(I));d.innerHTML=`
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
            <strong style="color:var(--text-1); font-size:0.95rem;">Rekap Klasifikasi Aset Perusahaan</strong>
            <span class="badge badge-primary">${j.length} Aset</span>
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
                ${j.map(($,p)=>`
                  <tr style="border-bottom:1px solid #F1F5F9; background:${p%2===0?"#fff":"#FAFAFA"};">
                    <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${$.no||p+1}</td>
                    <td style="padding:10px 14px;">
                      <span class="badge" style="${L($.type)}">${E($.type)}</span>
                    </td>
                    <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${E($.item)}</td>
                    <td style="padding:10px 14px; color:var(--text-2);">${E($.description)}</td>
                    <td style="padding:10px 14px; font-weight:600; color:#1E40AF;">${E($.pic||"-")}</td>
                    <td style="padding:10px 14px; text-align:center;">
                    ${a?`
                      <button class="btn btn-sm btn-ghost btn-edit-asset" data-no="${$.no}" title="Edit Aset" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                    `:""}
                    ${i?`
                      <button class="btn btn-sm btn-ghost btn-delete-asset" data-no="${$.no}" title="Hapus Aset" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    `:""}
                    </td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </div>
      `,d.querySelectorAll(".btn-edit-asset").forEach($=>{$.addEventListener("click",()=>{let p=parseInt($.dataset.no,10),S=A.find(u=>u.no===p);S&&r(S)})}),d.querySelectorAll(".btn-delete-asset").forEach($=>{$.addEventListener("click",()=>{let p=parseInt($.dataset.no,10),S=A.find(u=>u.no===p);S&&ie(`Apakah Anda yakin ingin menghapus aset: "${S.item}"?`,()=>{A=A.filter(u=>u.no!==p),A.forEach((u,M)=>{u.no=M+1}),Dt(A),H("Aset berhasil dihapus"),T()},"Hapus Aset")})})}function L(I){return I==="Aset Tetap"?"background:#EFF6FF; color:#1E40AF; font-weight:700; padding:3px 8px; border-radius:6px;":I==="Aset Bergerak"?"background:#ECFDF5; color:#047857; font-weight:700; padding:3px 8px; border-radius:6px;":"background:#FEF3C7; color:#92400E; font-weight:700; padding:3px 8px; border-radius:6px;"}w(),f&&f.addEventListener("input",I=>{g=I.target.value,w()}),n.querySelector("#btn-add-sec-asset").addEventListener("click",()=>{r()})}function r(n=null){if(!a)return;let d=!!n,f=`
      <form id="asset-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Tipe / Klasifikasi Aset *</label>
          <select id="form-asset-type" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Aset Tetap" ${n?.type==="Aset Tetap"?"selected":""}>Aset Tetap</option>
            <option value="Aset Bergerak" ${n?.type==="Aset Bergerak"?"selected":""}>Aset Bergerak</option>
            <option value="Low Value Asset" ${n?.type==="Low Value Asset"?"selected":""}>Low Value Asset</option>
          </select>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Aset / Kelompok *</label>
          <input type="text" id="form-asset-item" class="form-control" required value="${n?E(n.item):""}" placeholder="Contoh: Genset & Elevator" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Deskripsi / Contoh Item *</label>
          <textarea id="form-asset-desc" class="form-control" rows="3" required placeholder="Contoh: Generator daya darurat dan panel listrik utama" style="width:100%; border-radius:6px; font-size:0.85rem;">${n?E(n.description):""}</textarea>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">PIC Pengawas / Penanggung Jawab *</label>
          <input type="text" id="form-asset-pic" class="form-control" required value="${n?E(n.pic):""}" placeholder="Contoh: Teknisi & Security" style="width:100%; border-radius:6px; height:36px;">
        </div>
      </form>
    `;X({title:d?`\u270F\uFE0F Edit Aset: ${n.item}`:"\u2795 Tambah Klasifikasi Aset Baru",content:f,confirmText:"Simpan",onConfirm:(w,L)=>{let I=w.querySelector("#form-asset-type").value,j=w.querySelector("#form-asset-item").value.trim(),$=w.querySelector("#form-asset-desc").value.trim(),p=w.querySelector("#form-asset-pic").value.trim();if(!j||!$||!p){q("Harap lengkapi semua kolom");return}if(d)n.type=I,n.item=j,n.description=$,n.pic=p,Dt(A),H(`Aset "${j}" berhasil diperbarui`);else{let S=A.length>0?Math.max(...A.map(u=>u.no))+1:1;A.push({no:S,type:I,item:j,description:$,pic:p}),Dt(A),H(`Aset baru "${j}" berhasil ditambahkan`)}L(),T()}})}function h(n){n.innerHTML=`
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
    `}function F(n){n.innerHTML=`
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
    `}async function k(){try{let{downloadExcel:n}=await Promise.resolve().then(()=>(de(),fe)),d=l.map(w=>({"Kode IK":`IK-SEC-${String(w.id).padStart(2,"0")}`,"Judul Prosedur":w.title,Kategori:w.category,"Sasaran Target":w.target,"Alat & Perlengkapan":w.tools.join(", "),"Langkah Pelaksanaan":w.steps.map((L,I)=>`${I+1}. ${L}`).join(" | "),"Catatan Khusus":w.notes||"-"})),f=new Date().toISOString().slice(0,10);n(d,`SOP_Security_Instruksi_Kerja_${f}`),H("Data Excel SOP Security berhasil diunduh")}catch(n){console.error(n),q("Gagal mendownload data Excel: "+n.message)}}function E(n){return n?String(n).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}T()}ce();var Ba=[{id:1,name:"Berita Acara Kerusakan Alat",category:"Berita Acara",document_link:"https://docs.google.com/document/d/1example_ba_kerusakan",description:"Format berita acara resmi pencatatan dan eskalasi kerusakan peralatan klinik dan aset fasilitas."},{id:2,name:"Checklist Apar",category:"Apar & K3",document_link:"https://docs.google.com/spreadsheets/d/1example_apar",description:"Lembar inspeksi berkala tabung pemadam api ringan (APAR), tekanan manometer, dan masa kadaluarsa."},{id:3,name:"Checklist Deep Cleaning",category:"Kegiatan Deep Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_deep_cleaning",description:"Panduan dan verifikasi checklist pembersihan mendalam (deep cleaning) area dental unit dan klinik."},{id:4,name:"Checklist Fulldraine",category:"Drain Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_fulldraine",description:"Checklist monitoring pembersihan saluran pembuangan air kotor, floor drain, dan sedotan suction."},{id:5,name:"Checklist General Cleaning",category:"Kegiatan General Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_general_cleaning",description:"Standar checklist kebersihan harian/mingguan seluruh ruangan, lobby, toilet, dan koridor."},{id:6,name:"Form Aktivity Reliefer",category:"Form Operasional",document_link:"https://forms.gle/1example_reliefer",description:"Formulir pencatatan absensi dan penugasan harian reliefer Facility Care di cabang pengganti."},{id:7,name:"Form Aktivity Security HQ",category:"Form Operasional",document_link:"https://forms.gle/1example_security_hq",description:"Formulir monitoring shift dan pengawasan keamanan harian di Kantor Pusat (Head Quarter)."}];function Da(e=""){let t=(e||"").toLowerCase();return t.includes("berita acara")?{label:e,bg:"#F3E8FF",color:"#7E22CE",border:"#D8B4FE",icon:"\u2696\uFE0F"}:t.includes("apar")||t.includes("k3")||t.includes("safety")?{label:e,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F9EF}"}:t.includes("deep cleaning")?{label:e,bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2728"}:t.includes("drain")?{label:e,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u{1F6BF}"}:t.includes("general cleaning")||t.includes("kebersihan")?{label:e,bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F9F9}"}:t.includes("form")?{label:e,bg:"#E0E7FF",color:"#4338CA",border:"#C7D2FE",icon:"\u{1F4DD}"}:{label:e||"Dokumen",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4C1}"}}function Pa(e=""){if(!e)return{type:"Tidak Ada Tautan",icon:"\u{1F4CE}",badgeClass:"bg-slate-100 text-slate-500"};let t=e.toLowerCase();return t.includes("forms.gle")||t.includes("google.com/forms")?{type:"Google Form",icon:"\u{1F7E3}",badgeClass:"badge-form",color:"#673AB7"}:t.includes("docs.google.com/spreadsheets")||t.includes("sheet")?{type:"Google Sheets",icon:"\u{1F7E2}",badgeClass:"badge-sheet",color:"#0F9D58"}:t.includes("docs.google.com/document")||t.includes("doc")?{type:"Google Docs",icon:"\u{1F535}",badgeClass:"badge-doc",color:"#4285F4"}:t.includes("drive.google.com")?{type:"Google Drive",icon:"\u{1F7E1}",badgeClass:"badge-drive",color:"#F4B400"}:t.endsWith(".pdf")?{type:"PDF Document",icon:"\u{1F534}",badgeClass:"badge-pdf",color:"#EA4335"}:{type:"Web Link",icon:"\u{1F310}",badgeClass:"badge-link",color:"#2563EB"}}async function La(e){let t=Z(),a=ee(t,"checklist","write"),i=ee(t,"checklist","delete"),s=[],b="all",x="",g="fcms-master-doc-styles";if(!document.getElementById(g)){let c=document.createElement("style");c.id=g,c.textContent=`
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
    `,document.head.appendChild(c)}async function o(){try{let c=await N("/api/checklist?limit=1000");if(c.ok&&c.data?.data&&Array.isArray(c.data.data)&&c.data.data.length>0)s=c.data.data,localStorage.setItem("fcms_master_doc_cache",JSON.stringify(s));else{let y=localStorage.getItem("fcms_master_doc_cache");y?s=JSON.parse(y):(s=[...Ba],localStorage.setItem("fcms_master_doc_cache",JSON.stringify(s)))}}catch(c){console.warn("Fallback loading documents from cache/seed:",c);let y=localStorage.getItem("fcms_master_doc_cache");s=y?JSON.parse(y):[...Ba]}B()}function l(){return s.filter(c=>{let y=b==="all"||c.category===b,v=x.toLowerCase().trim(),m=!v||(c.name||"").toLowerCase().includes(v)||(c.category||"").toLowerCase().includes(v)||(c.description||"").toLowerCase().includes(v);return y&&m})}function A(){let c={all:s.length};return s.forEach(y=>{let v=y.category||"Lainnya";c[v]=(c[v]||0)+1}),c}function T(c=null){if(!a)return;let y=!!c,v=Array.from(new Set(s.map(r=>r.category).filter(Boolean))),m=document.createElement("div");m.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Nama Dokumen / Checklist <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Checklist Pembersihan Wastafel & Toilet" value="${c?.name?String(c.name).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Kategori Dokumen <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-category" list="doc-categories-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Pilih atau ketik kategori baru..." value="${c?.category?String(c.category).replace(/"/g,"&quot;"):""}" required />
          <datalist id="doc-categories-list">
            ${v.map(r=>`<option value="${r}">${r}</option>`).join("")}
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
            placeholder="https://docs.google.com/..." value="${c?.document_link?String(c.document_link).replace(/"/g,"&quot;"):""}" />
          <span style="font-size:0.75rem; color:#64748B; margin-top:4px; display:block;">
            Bisa berupa tautan Google Forms untuk inspeksi harian, atau link Google Sheets / Docs sebagai acuan.
          </span>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Deskripsi / Petunjuk Penggunaan
          </label>
          <textarea id="doc-modal-desc" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan keterangan singkat mengenai tujuan dan instruksi pengisian dokumen ini...">${c?.description||""}</textarea>
        </div>
      </div>
    `,X({title:y?"\u270F\uFE0F Edit Master Dokumen":"\u2795 Tambah Master Dokumen Baru",content:m,size:"md",confirmText:y?"Simpan Perubahan":"Tambahkan Dokumen",confirmClass:"btn-primary",onConfirm:async(r,h)=>{let F=m.querySelector("#doc-modal-name"),k=m.querySelector("#doc-modal-category"),E=m.querySelector("#doc-modal-link"),n=m.querySelector("#doc-modal-desc"),d=(F.value||"").trim(),f=(k.value||"").trim(),w=(E.value||"").trim(),L=(n.value||"").trim();if(!d){q("Nama Dokumen wajib diisi"),F.focus();return}if(!f){q("Kategori Dokumen wajib diisi"),k.focus();return}let I={name:d,category:f,document_link:w,description:L};try{if(y)if((await N(`/api/checklist/${c.id}`,{method:"PUT",body:JSON.stringify(I)})).ok){let $=s.findIndex(p=>p.id===c.id);$!==-1&&(s[$]={...s[$],...I}),H("Dokumen berhasil diperbarui")}else{let $=s.findIndex(p=>p.id===c.id);$!==-1&&(s[$]={...s[$],...I}),H("Dokumen disimpan secara lokal")}else{let j=await N("/api/checklist",{method:"POST",body:JSON.stringify(I)});if(j.ok&&j.data?.id)s.unshift({id:j.data.id,...I}),H("Dokumen baru berhasil ditambahkan");else{let $=Date.now();s.unshift({id:$,...I}),H("Dokumen baru ditambahkan secara lokal")}}localStorage.setItem("fcms_master_doc_cache",JSON.stringify(s)),h(),B()}catch(j){console.error("Error saving document:",j),q("Gagal menyimpan dokumen: "+j.message)}}})}function C(c){ie(`Apakah Anda yakin ingin menghapus dokumen <strong>"${c.name}"</strong>?`,async()=>{try{let y=await N(`/api/checklist/${c.id}`,{method:"DELETE"});y.ok||console.warn("API delete warning:",y),s=s.filter(v=>v.id!==c.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(s)),H(`Dokumen "${c.name}" berhasil dihapus`),B()}catch{s=s.filter(v=>v.id!==c.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(s)),H(`Dokumen "${c.name}" berhasil dihapus`),B()}},"Hapus Dokumen")}function _(c,y){if(!c){ft("Dokumen belum memiliki tautan digital.");return}navigator.clipboard.writeText(c).then(()=>{H(`Tautan "${y}" disalin ke clipboard!`)}).catch(()=>{ft(`Tautan: ${c}`)})}async function z(){try{let{downloadExcel:c}=await Promise.resolve().then(()=>(de(),fe)),y=s.map((v,m)=>({No:m+1,"Nama Dokumen":v.name||"",Kategori:v.category||"","Platform / Format":Pa(v.document_link).type,"Link Dokumen":v.document_link||"","Deskripsi / Keterangan":v.description||""}));c(y,`Master_Dokumen_FCMS_${new Date().toISOString().slice(0,10)}`),H("Berhasil mengunduh data dokumen ke Excel.")}catch(c){q("Gagal mengekspor data: "+c.message)}}function B(){let c=l(),y=A(),v=Object.keys(y).filter(r=>r!=="all"),m=s.filter(r=>!!r.document_link).length;e.innerHTML=`
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
              ${a?`
              <button id="doc-btn-add" class="btn btn-primary" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap; border:none;">
                <span>\u2795</span>
                <span>Tambah Dokumen Baru</span>
              </button>
              `:""}
              <button id="doc-btn-excel" class="btn btn-ghost" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>
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
                <div class="doc-stat-value">${s.length}</div>
                <div class="doc-stat-label">Total Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon">\u{1F5C2}\uFE0F</div>
              <div>
                <div class="doc-stat-value">${v.length}</div>
                <div class="doc-stat-label">Kategori Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon" style="background:rgba(16,185,129,0.18); color:#34D399;">\u{1F310}</div>
              <div>
                <div class="doc-stat-value">${m} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">/ ${s.length}</span></div>
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
              <input type="text" id="doc-search-input" class="doc-search-input" placeholder="Cari nama dokumen, kategori, atau deskripsi..." value="${x?x.replace(/"/g,"&quot;"):""}" />
            </div>
          </div>

          <!-- Category Pills -->
          <div class="doc-filter-pills">
            <button class="doc-pill ${b==="all"?"active":""}" data-cat="all">
              <span>Semua Dokumen</span>
              <span class="doc-pill-count">${y.all||0}</span>
            </button>
            ${v.map(r=>{let h=Da(r);return`
                <button class="doc-pill ${b===r?"active":""}" data-cat="${r}">
                  <span>${h.icon}</span>
                  <span>${r}</span>
                  <span class="doc-pill-count">${y[r]||0}</span>
                </button>
              `}).join("")}
          </div>
        </div>

        <!-- Documents View Container -->
        <div id="doc-list-view">
          ${c.length===0?`
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
              ${c.map(r=>{let h=Da(r.category),F=Pa(r.document_link);return`
                  <div class="doc-card" data-id="${r.id}">
                    <div>
                      <div class="doc-card-top">
                        <span class="doc-category-badge" style="background:${h.bg}; color:${h.color}; border:1px solid ${h.border};">
                          <span>${h.icon}</span>
                          <span>${h.label}</span>
                        </span>
                        <span class="doc-platform-badge">
                          <span>${F.icon}</span>
                          <span>${F.type}</span>
                        </span>
                      </div>
                      
                      <h3 class="doc-card-title">${r.name}</h3>
                      <p class="doc-card-desc">${r.description||"Formulir acuan standar operasional dan administrasi fasilitas."}</p>
                    </div>

                    <div class="doc-card-footer">
                      ${r.document_link?`
                        <a href="${r.document_link}" target="_blank" rel="noopener noreferrer" class="doc-btn-open">
                          <span>Buka Dokumen</span>
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                        </a>
                      `:`
                        <button class="doc-btn-open disabled" title="Belum ada tautan yang dikaitkan">
                          <span>Tidak Ada Link</span>
                        </button>
                      `}

                      <div class="doc-card-actions">
                        ${r.document_link?`
                          <button class="doc-icon-btn copy-link-btn" title="Salin Tautan Dokumen" data-link="${r.document_link}" data-title="${r.name}">
                            <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                          </button>
                        `:""}
                        ${a?`
                        <button class="doc-icon-btn edit-doc-btn" title="Edit Dokumen" data-id="${r.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        `:""}
                        ${i?`
                        <button class="doc-icon-btn delete delete-doc-btn" title="Hapus Dokumen" data-id="${r.id}">
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
    `,P()}function P(){let c=e.querySelector("#doc-btn-add");c&&(c.onclick=()=>T());let y=e.querySelector("#doc-btn-excel");y&&(y.onclick=()=>z());let v=e.querySelector("#doc-btn-refresh");v&&(v.onclick=()=>o());let m=e.querySelector("#doc-btn-reset-filter");m&&(m.onclick=()=>{x="",b="all",B()});let r=e.querySelector("#doc-search-input");r&&(r.oninput=h=>{x=h.target.value;let F=l();B();let k=e.querySelector("#doc-search-input");k&&(k.focus(),k.setSelectionRange(k.value.length,k.value.length))}),e.querySelectorAll(".doc-pill").forEach(h=>{h.onclick=()=>{b=h.getAttribute("data-cat")||"all",B()}}),e.querySelectorAll(".copy-link-btn").forEach(h=>{h.onclick=F=>{F.stopPropagation();let k=h.getAttribute("data-link"),E=h.getAttribute("data-title");_(k,E)}}),e.querySelectorAll(".edit-doc-btn").forEach(h=>{h.onclick=F=>{F.stopPropagation();let k=Number(h.getAttribute("data-id")),E=s.find(n=>n.id===k);E&&T(E)}}),e.querySelectorAll(".delete-doc-btn").forEach(h=>{h.onclick=F=>{F.stopPropagation();let k=Number(h.getAttribute("data-id")),E=s.find(n=>n.id===k);E&&C(E)}})}await o()}ce();de();var Ne={"Standar Umum":{icon:"\u{1F4CC}",bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE",rooms:["Lantai & Sudut Ruangan","Dinding, Partisi & Plint","Kaca, Cermin & Partisi Kaca","Sanitair & Area Basah Toilet","Plafon, Lampu & Kisi Ventilasi","Mebel, Meja & Kursi Kerja/Pasien","Peralatan Khusus & Medis Non-Steril","Wadah Sampah & Utilitas"]},"Area Depan & Pelayanan":{icon:"\u{1F6CB}\uFE0F",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",rooms:["Teras","Lobby 1","Receptionist","Snack Corner"]},"Area Klinis & Medis":{icon:"\u{1FA7A}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD",rooms:["Ruang Tindakan","Ronsen","Ruang Dokter","Ruang Steril"]},"Toilet & Sanitasi":{icon:"\u{1F6BF}",bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4",rooms:["Toilet 1","Janitor","Ruang Limbah"]},Sirkulasi:{icon:"\u{1F6B6}",bg:"#F1F5F9",color:"#334155",border:"#CBD5E1",rooms:["Area Tangga","Lift","Area Tangga Exit"]},"Area Anak & Fasilitas":{icon:"\u{1F9F8}",bg:"#FEF3C7",color:"#92400E",border:"#FDE68A",rooms:["Lobby 2","Playground","Musholla","Tempat Wudhu","Ruang Team"]},"Gudang & Teknis":{icon:"\u2699\uFE0F",bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA",rooms:["Gudang Farmasi & Bahan","RoofTop","Ruang Kompressor","Area Genset"]},Eksterior:{icon:"\u{1F333}",bg:"#DCFCE7",color:"#15803D",border:"#86EFAC",rooms:["Taman","Pos Security","Toilet Security","Halaman Parkir"]}};function It(e){for(let[t,a]of Object.entries(Ne))if(a.rooms.includes(e))return{zoneName:t,...a};return{zoneName:"Lainnya",icon:"\u{1F4CD}",bg:"#F1F5F9",color:"#475569",border:"#E2E8F0"}}async function Ia(e){e.innerHTML=`
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
        ${Object.entries(Ne).map(([v,m])=>`
          <button class="zone-tab-btn" data-zone="${v}" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
            <span>${m.icon}</span> ${v}
            <span class="badge" style="background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-${v.replace(/[\s&]+/g,"-")}">-</span>
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
  `;let t="all",a="table",i=[],s=Z(),b=s&&["superadmin","admin","manager","spv"].includes(s.role);async function x(){try{let v=await N("/api/hygiene-standards?all=1");if(!v.ok)throw new Error("Gagal memuat standar hygiene");i=Array.isArray(v.data)?v.data:v.data?.data||[],o(),g(),A()}catch(v){console.error(v),q(v.message||"Terjadi kesalahan memuat data");let m=e.querySelector("#hygiene-view-container");m&&(m.innerHTML=`
          <div class="empty-state" style="text-align:center; padding:40px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:12px;">\u26A0\uFE0F</div>
            <h3 style="margin:0 0 8px 0; color:var(--text-1);">Gagal Memuat Data</h3>
            <p style="color:var(--text-3); margin-bottom:16px;">${v.message}</p>
            <button id="btn-retry-hygiene" class="btn btn-primary" style="padding:8px 16px; border-radius:8px;">Coba Lagi</button>
          </div>
        `,m.querySelector("#btn-retry-hygiene")?.addEventListener("click",x))}}function g(){let v=e.querySelector("#badge-count-all");v&&(v.textContent=i.length),Object.entries(Ne).forEach(([h,F])=>{let k=h.replace(/[\s&]+/g,"-"),E=e.querySelector(`#badge-count-${k}`);if(E){let n=i.filter(d=>F.rooms.includes(d.room_name)).length;E.textContent=n}});let m=new Set(i.map(h=>h.room_name)),r=e.querySelector("#card-total-rooms");r&&(r.textContent=m.size)}function o(){let v=e.querySelector("#hygiene-room-filter");if(!v)return;let m=v.value,r=[];if(t==="all")r=Array.from(new Set(i.map(h=>h.room_name))).filter(Boolean).sort();else{let h=Ne[t];h&&(r=h.rooms.filter(F=>i.some(k=>k.room_name===F)))}v.innerHTML=`<option value="">Semua Ruangan (${r.length})</option>`+r.map(h=>`<option value="${h}" ${h===m?"selected":""}>${h}</option>`).join("")}function l(){let v=(e.querySelector("#hygiene-search")?.value||"").toLowerCase().trim(),m=e.querySelector("#hygiene-room-filter")?.value||"";return i.filter(r=>{if(t!=="all"){let h=Ne[t];if(!h||!h.rooms.includes(r.room_name))return!1}if(m&&r.room_name!==m)return!1;if(v){let h=(r.room_name||"").toLowerCase(),F=(r.item_name||"").toLowerCase(),k=(r.cleanliness_standard||"").toLowerCase();if(!h.includes(v)&&!F.includes(v)&&!k.includes(v))return!1}return!0})}function A(){let v=l(),m=e.querySelector("#hygiene-stats-text");m&&(m.innerHTML=`Menampilkan <strong>${v.length}</strong> dari <strong>${i.length}</strong> parameter`);let r=e.querySelector("#hygiene-view-container");if(r){if(v.length===0){r.innerHTML=`
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
      `,r.querySelector("#btn-clear-search-empty")?.addEventListener("click",()=>{let h=e.querySelector("#hygiene-search"),F=e.querySelector("#hygiene-room-filter");h&&(h.value=""),F&&(F.value=""),t="all",e.querySelectorAll(".zone-tab-btn").forEach(k=>k.classList.toggle("active",k.dataset.zone==="all")),e.querySelectorAll(".zone-tab-btn").forEach(k=>{k.dataset.zone==="all"?(k.style.background="#2563EB",k.style.color="#fff"):(k.style.background="#F1F5F9",k.style.color="var(--text-2)")}),o(),A()});return}a==="table"?T(r,v):C(r,v)}}function T(v,m){v.innerHTML=`
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse:collapse; text-align:left; font-size:0.9rem;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:2px solid #E2E8F0; color:#475569; font-size:0.78rem; text-transform:uppercase; letter-spacing:0.6px;">
                <th style="padding:14px 16px; width:50px; text-align:center;">No</th>
                <th style="padding:14px 16px; width:220px;">Ruangan & Zona</th>
                <th style="padding:14px 16px; width:280px;">Item / Objek Inspeksi</th>
                <th style="padding:14px 16px;">Standar Kebersihan Fisik</th>
                ${b?'<th style="padding:14px 16px; width:100px; text-align:center;">Aksi</th>':""}
              </tr>
            </thead>
            <tbody>
              ${m.map((r,h)=>{let F=It(r.room_name);return`
                  <tr style="border-bottom:1px solid #F1F5F9; transition:background-color 0.15s ease;" onmouseover="this.style.backgroundColor='#F8FAFC'" onmouseout="this.style.backgroundColor='transparent'">
                    <td style="padding:14px 16px; text-align:center; color:var(--text-3); font-weight:600; font-size:0.85rem;">
                      ${h+1}
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="margin-bottom:4px;">
                        <span style="display:inline-flex; align-items:center; gap:6px; background:${F.bg}; color:${F.color}; border:1px solid ${F.border}; padding:4px 10px; border-radius:6px; font-weight:700; font-size:0.83rem;">
                          <span>${F.icon}</span> ${r.room_name}
                        </span>
                      </div>
                      <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">
                        ${F.zoneName}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="font-weight:700; color:var(--text-1); line-height:1.45; font-size:0.9rem;">
                        ${r.item_name}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-left:3px solid #10B981; border-radius:6px; padding:10px 12px; line-height:1.55; color:#334155; font-size:0.875rem;">
                        <span style="color:#059669; font-weight:700; margin-right:4px;">\u2713</span> ${r.cleanliness_standard}
                      </div>
                    </td>
                    ${b?`
                      <td style="padding:14px 16px; text-align:center; vertical-align:middle;">
                        <div style="display:inline-flex; gap:6px;">
                          <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${r.id}" title="Edit Parameter" style="padding:6px 8px; border-radius:6px; color:#2563EB; cursor:pointer;">
                            \u270F\uFE0F
                          </button>
                          <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${r.id}" data-name="${r.item_name}" title="Hapus Parameter" style="padding:6px 8px; border-radius:6px; color:#DC2626; cursor:pointer;">
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
    `,_(v)}function C(v,m){let r={};m.forEach(h=>{let F=It(h.room_name);r[F.zoneName]||(r[F.zoneName]={info:F,rooms:{}}),r[F.zoneName].rooms[h.room_name]||(r[F.zoneName].rooms[h.room_name]=[]),r[F.zoneName].rooms[h.room_name].push(h)}),v.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:24px;">
        ${Object.entries(r).map(([h,F])=>{let k=Object.values(F.rooms).reduce((E,n)=>E+n.length,0);return`
            <div class="card" style="background:#fff; border:1px solid ${F.info.border}; border-radius:14px; overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.03);">
              <!-- Zone Card Header -->
              <div style="background:${F.info.bg}; border-bottom:1px solid ${F.info.border}; padding:14px 20px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="font-size:1.4rem;">${F.info.icon}</span>
                  <div>
                    <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:${F.info.color};">
                      ${h}
                    </h3>
                    <div style="font-size:0.78rem; color:${F.info.color}; opacity:0.85;">
                      ${Object.keys(F.rooms).length} Ruangan \xB7 ${k} Parameter Standar
                    </div>
                  </div>
                </div>
                <span class="badge" style="background:#fff; color:${F.info.color}; border:1px solid ${F.info.border}; font-weight:700; padding:4px 10px; border-radius:12px; font-size:0.8rem;">
                  ${k} Parameter
                </span>
              </div>

              <!-- Zone Rooms List -->
              <div style="padding:16px 20px;">
                ${Object.entries(F.rooms).map(([E,n])=>`
                  <div style="margin-bottom:20px; last-child:margin-bottom:0;">
                    <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                      <h4 style="margin:0; font-size:0.95rem; font-weight:700; color:var(--text-1);">
                        \u{1F6AA} ${E}
                      </h4>
                      <span class="badge badge-sm" style="background:#F1F5F9; color:var(--text-3); font-size:0.75rem;">
                        ${n.length} item
                      </span>
                    </div>

                    <div style="border:1px solid #E2E8F0; border-radius:8px; overflow:hidden;">
                      <table style="width:100%; border-collapse:collapse; font-size:0.875rem;">
                        <thead>
                          <tr style="background:#F8FAFC; border-bottom:1px solid #E2E8F0; color:#475569; font-size:0.75rem; text-transform:uppercase;">
                            <th style="padding:8px 12px; width:40px; text-align:center;">No</th>
                            <th style="padding:8px 12px; width:260px;">Item / Objek Inspeksi</th>
                            <th style="padding:8px 12px;">Standar Kebersihan Fisik</th>
                            ${b?'<th style="padding:8px 12px; width:80px; text-align:center;">Aksi</th>':""}
                          </tr>
                        </thead>
                        <tbody>
                          ${n.map((d,f)=>`
                            <tr style="border-bottom:1px solid #F1F5F9;">
                              <td style="padding:10px 12px; text-align:center; color:var(--text-3); font-size:0.8rem;">
                                ${f+1}
                              </td>
                              <td style="padding:10px 12px; font-weight:600; color:var(--text-1); vertical-align:top;">
                                ${d.item_name}
                              </td>
                              <td style="padding:10px 12px; vertical-align:top; color:#334155;">
                                <div style="display:flex; align-items:flex-start; gap:6px;">
                                  <span style="color:#10B981; font-weight:700;">\u2713</span>
                                  <span>${d.cleanliness_standard}</span>
                                </div>
                              </td>
                              ${b?`
                                <td style="padding:10px 12px; text-align:center; vertical-align:middle;">
                                  <div style="display:inline-flex; gap:4px;">
                                    <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${d.id}" title="Edit" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                                    <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${d.id}" data-name="${d.item_name}" title="Hapus" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
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
    `,_(v)}function _(v){v.querySelectorAll(".btn-edit-item").forEach(m=>{m.addEventListener("click",()=>{let r=m.dataset.id,h=i.find(F=>String(F.id)===String(r));h&&z(h)})}),v.querySelectorAll(".btn-delete-item").forEach(m=>{m.addEventListener("click",()=>{let r=m.dataset.id,h=m.dataset.name;ie(`Yakin ingin menghapus parameter standar <strong>"${h}"</strong>?`,async()=>{try{let F=await N(`/api/hygiene-standards/${r}`,{method:"DELETE"});if(!F.ok)throw new Error(F.data?.error||"Gagal menghapus parameter");H("Parameter kebersihan berhasil dihapus"),x()}catch(F){q(F.message)}},"Hapus Parameter Standar")})})}function z(v=null){let m=!!v,r=Array.from(new Set(i.map(F=>F.room_name))).filter(Boolean).sort(),h=document.createElement("div");h.innerHTML=`
      <form id="form-hygiene-item" style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Ruangan / Lokasi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-room-name" list="room-suggestions" class="form-control" placeholder="Contoh: Ruang Tindakan, Lobby 1, Toilet 1..." value="${v?.room_name||""}" required style="width:100%; height:40px; border-radius:8px;">
          <datalist id="room-suggestions">
            ${r.map(F=>`<option value="${F}">`).join("")}
          </datalist>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Pilih dari daftar atau ketik nama ruangan baru</span>
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Item / Objek Inspeksi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-item-name" class="form-control" placeholder="Contoh: Eksterior Dental Unit & Kursi, Lantai & Sudut..." value="${v?.item_name||""}" required style="width:100%; height:40px; border-radius:8px;">
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Standar Kebersihan Fisik <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="modal-standard" class="form-control" rows="4" placeholder="Deskripsi kondisi fisik kebersihan yang dipersyaratkan..." required style="width:100%; border-radius:8px; line-height:1.5;">${v?.cleanliness_standard||""}</textarea>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Jelaskan kriteria fisik: bebas debu kasat mata, steril, tidak licin, tidak berbau, dll.</span>
        </div>
      </form>
    `,X({title:m?"\u270F\uFE0F Edit Parameter Kebersihan":"\u2795 Tambah Parameter Kebersihan",content:h,confirmText:m?"Simpan Perubahan":"Tambah Parameter",size:"md",onConfirm:async(F,k)=>{let E=h.querySelector("#modal-room-name").value.trim(),n=h.querySelector("#modal-item-name").value.trim(),d=h.querySelector("#modal-standard").value.trim();if(!E||!n||!d){q("Semua kolom wajib diisi");return}try{let f={room_name:E,item_name:n,cleanliness_standard:d},w=m?`/api/hygiene-standards/${v.id}`:"/api/hygiene-standards",I=await N(w,{method:m?"PUT":"POST",body:JSON.stringify(f)});if(!I.ok)throw new Error(I.data?.error||"Gagal menyimpan parameter");H(m?"Parameter berhasil diperbarui":"Parameter berhasil ditambahkan"),k(),x()}catch(f){q(f.message)}}})}function B(){let v=window.open("","_blank");if(!v){q("Pop-up printer diblokir oleh browser. Izinkan pop-up untuk mencetak SOP.");return}let r=`
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
            <div>Total: <strong>${i.length} Parameter Fisik</strong></div>
          </div>
        </div>

        <div class="title">
          <h1>STANDAR ACUAN KEBERSIHAN FISIK SELURUH CABANG</h1>
          <p>Buku Panduan Kriteria Kebersihan Fisik 8 Zona Area (Cleaning Service & Facility Care)</p>
        </div>

        ${Object.entries(Ne).map(([h,F])=>{let k=i.filter(E=>F.rooms.includes(E.room_name));return k.length===0?"":`
            <div class="zone-title">${F.icon} ${h.toUpperCase()} (${k.length} PARAMETER)</div>
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
                ${k.map((E,n)=>`
                  <tr>
                    <td class="td-no">${n+1}</td>
                    <td class="td-room">${E.room_name}</td>
                    <td class="td-obj">${E.item_name}</td>
                    <td>${E.cleanliness_standard}</td>
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
    `;v.document.open(),v.document.write(r),v.document.close()}function P(){let v=i.map((r,h)=>{let F=It(r.room_name);return{No:h+1,"Kategori Zona":F.zoneName,"Ruangan / Lokasi":r.room_name,"Item / Objek Inspeksi":r.item_name,"Standar Kebersihan Fisik":r.cleanliness_standard}}),m=new Date().toISOString().slice(0,10);ne(v,`Master_Standar_Kebersihan_FCMS_${m}`),H("Data master kebersihan berhasil diekspor ke Excel")}e.querySelector("#hygiene-search")?.addEventListener("input",()=>A()),e.querySelector("#hygiene-room-filter")?.addEventListener("change",()=>A()),e.querySelector("#btn-reset-filters")?.addEventListener("click",()=>{let v=e.querySelector("#hygiene-search"),m=e.querySelector("#hygiene-room-filter");v&&(v.value=""),m&&(m.value=""),A()}),e.querySelectorAll(".zone-tab-btn").forEach(v=>{v.addEventListener("click",()=>{t=v.dataset.zone,e.querySelectorAll(".zone-tab-btn").forEach(m=>{let r=m.dataset.zone===t;if(m.classList.toggle("active",r),r){m.style.background="#2563EB",m.style.color="#fff";let h=m.querySelector(".badge");h&&(h.style.background="rgba(255,255,255,0.25)",h.style.color="#fff")}else{m.style.background="#F1F5F9",m.style.color="var(--text-2)";let h=m.querySelector(".badge");h&&(h.style.background="#CBD5E1",h.style.color="#334155")}}),o(),A()})});let c=e.querySelector("#view-table-btn"),y=e.querySelector("#view-zone-btn");c?.addEventListener("click",()=>{a="table",c.style.background="#fff",c.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",y.style.background="transparent",y.style.boxShadow="none",A()}),y?.addEventListener("click",()=>{a="zone",y.style.background="#fff",y.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",c.style.background="transparent",c.style.boxShadow="none",A()}),e.querySelector("#btn-print-hygiene")?.addEventListener("click",B),e.querySelector("#btn-export-hygiene")?.addEventListener("click",P),e.querySelector("#btn-add-hygiene")?.addEventListener("click",()=>z(null)),x()}ce();de();async function za(e,t="supply"){return Hn(e)}async function Hn(e){let a=((await N("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}));ue({container:e,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:i=>i?new Date(i).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(i,s)=>s.branch_name_ref||s.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:i=>{try{let s=JSON.parse(i);return Array.isArray(s)?s.join(", "):i}catch{return i||"-"}}},{key:"chemical_items",label:"Chemical",render:i=>{try{let s=JSON.parse(i);return Array.isArray(s)?s.join(", "):i}catch{return i||"-"}}},{key:"additional_notes",label:"Catatan",render:i=>i?.length>40?i.slice(0,40)+"\u2026":i||"-"},{key:"status",label:"Status",render:i=>be(i)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:i=>{let s=i?.tools_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}let b=i?.chemical_items;try{b=Array.isArray(JSON.parse(b))?JSON.parse(b).join(", "):b}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:i?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:i?.branch_id&&!a.find(x=>x.value==i.branch_id)?[...a,{value:i.branch_id,label:i.branch_name||i.branch_id}]:a,createApi:{path:"/api/branches",field:"full_name"},value:i?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:s},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:i?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:b},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:i?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:i?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:i?.status||""},{name:"processed_by",label:"Diproses Oleh",value:i?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async i=>{let s=new URLSearchParams(i||{}).toString(),b=await N(`/api/reports/supply?limit=10000&${s}`);if(b.ok){let x=b.data.data.map(g=>{let o=g.tools_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}let l=g.chemical_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}return{Waktu:g.submitted_at||"",Pengirim:g.submitter_name||"",Cabang:g.branch_name_ref||g.branch_name||"","Alat/Barang":o||"",Chemical:l||"",Catatan:g.additional_notes||"",Status:g.status||"","Diproses Oleh":g.processed_by||""}});ne(x,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async i=>{let b=(await N("/api/branches?all=1")).data?.data||[],x=A=>{if(!A)return null;let T=String(A||"").toLowerCase(),C=b.find(_=>String(_.full_name||"").toLowerCase()===T||String(_.code||"").toLowerCase()===T||String(_.name||"").toLowerCase()===T);return C?C.id:null},g=A=>{if(A==null||A==="")return"";if(A instanceof Date&&!isNaN(A.getTime()))return A.toISOString().slice(0,10);let T=String(A).trim();if(T===""||T==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(T))return T.slice(0,10);if(/^\d{4,5}$/.test(T)){let _=Number(T);if(_>2e4&&_<99999){let z=new Date(Date.UTC(1899,11,30)+_*864e5);return isNaN(z.getTime())?"":z.toISOString().slice(0,10)}}let C=T.split(/[\/\-\.]/);if(C.length===3){let[_,z,B]=C.map(P=>P.trim());if(_.length===4&&z.length<=2&&B.length<=2)return`${_}-${z.padStart(2,"0")}-${B.padStart(2,"0")}`;if(B.length===4&&z.length<=2&&_.length<=2)return`${B}-${z.padStart(2,"0")}-${_.padStart(2,"0")}`}return T},o=i.map(A=>({submitted_at:g(A.Waktu||A.Tanggal),submitter_name:String(A.Pengirim||"").trim(),branch_id:x(String(A.Cabang||"").trim()),tools_items:String(A["Alat/Barang"]||A.Alat||"").trim(),chemical_items:String(A.Chemical||"").trim(),additional_notes:String(A.Catatan||A.Keterangan||"").trim(),status:String(A.Status||"").trim(),processed_by:String(A["Diproses Oleh"]||A.PIC||"").trim()})).filter(A=>A.submitted_at&&A.submitter_name&&A.branch_id),l=await N("/api/reports/supply/import",{method:"POST",body:JSON.stringify(o)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(i,s)=>{let b=X({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(x,g)=>{let o=x.querySelector("#supply-status").value,l=x.querySelector("#supply-processed-by").value;(await N(`/api/reports/supply/${i.id}`,{method:"PUT",body:JSON.stringify({status:o,processed_by:l})})).ok?(H("Status diperbarui."),g(),s()):q("Gagal update status.")}})}}]})}de();async function Ma(e){let t=Z();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}ue({container:e,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:a=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",editor_khusus:"badge-warning",input_lembur:"badge-info",viewer:"badge-neutral"}[a]||"badge-neutral"}">${a}</span>`},{key:"is_active",label:"Status",render:a=>a?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:a=>a?new Date(a).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:a=>{let i=!!a;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:a?.full_name},{name:"username",label:"Username",required:!0,placeholder:"username",value:a?.username,autocomplete:"off"}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!0,placeholder:"email@contoh.com",value:a?.email,autocomplete:"off"},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"editor_khusus",label:"Editor (Hanya Lembur & Masalah)"},{value:"input_lembur",label:"Input Lembur (Hanya Isi Lembur)"},{value:"viewer",label:"Viewer"}],value:a?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:i?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!i,placeholder:"Min. 6 karakter",autocomplete:"new-password"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:i?a?.is_active:1}]}]},onBeforeSubmit:(a,i)=>(a.full_name&&(a.full_name=String(a.full_name).trim()),a.username&&(a.username=String(a.username).trim()),a.email&&(a.email=String(a.email).trim().toLowerCase()),i&&(!a.password||String(a.password).trim()==="")&&delete a.password,a),exportOptions:{moduleName:"users",onExport:async()=>{let a=await N(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let i=a.data.data.map(s=>({"Nama Lengkap":s.full_name||"",Username:s.username||"",Email:s.email||"",Role:s.role||"",Status:s.is_active?"Aktif":"Nonaktif"}));ne(i,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async a=>{let i=a.map(b=>({full_name:String(b["Nama Lengkap"]||"").trim(),username:String(b.Username||"").trim(),email:String(b.Email||"").trim().toLowerCase(),role:String(b.Role||"").trim()||"viewer",password:String(b.Password||"").trim()})).filter(b=>b.username&&b.password&&b.email&&b.full_name),s=await N("/api/users/import",{method:"POST",body:JSON.stringify(i)});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}de();async function Ka(e){ue({container:e,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!1,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:t=>t?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:t=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:t?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:t?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:t?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:t?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:t?.is_active!==void 0?t.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let t=await N(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(t.ok)ne(t.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{ne([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async t=>{let a=t.map(s=>({code:String(s["Kode Cabang"]||"").trim(),name:String(s["Nama Pendek"]||"").trim(),full_name:String(s["Nama Lengkap"]||"").trim(),city:String(s.Kota||"").trim()})).filter(s=>s.code&&s.name),i=await N("/api/branches/import",{method:"POST",body:JSON.stringify(a)});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}async function ja(e){let t=new Date,a=[];e.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{t.setMonth(t.getMonth()-1),s()}),document.getElementById("cal-next").addEventListener("click",()=>{t.setMonth(t.getMonth()+1),s()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(x=>x.addEventListener("change",s));async function i(){try{let x=`${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,"0")}`;a=(await N(`/api/dashboard/calendar?month=${x}`)).data?.data||[]}catch(x){console.warn("[Calendar] Failed to load events, rendering empty grid:",x),a=[]}}async function s(){let x=document.getElementById("calendar-grid");if(x){x.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await i();try{let g=t.getFullYear(),o=t.getMonth(),l=t.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),A=document.getElementById("cal-month-label");A&&(A.textContent=l);let T=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(r=>r.value)),C=a.filter(r=>T.has(r.type)),_={};C.forEach(r=>{let h=Un(r.event_date);h&&(_[h]||(_[h]=[]),_[h].push(r))});let z=new Date(g,o,1).getDay(),B=new Date(g,o+1,0).getDate(),P=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],c=new Date().toISOString().slice(0,10),y='<div class="calendar-grid">';P.forEach(r=>{y+=`<div class="cal-day-header">${r}</div>`});for(let r=0;r<z;r++)y+='<div class="cal-cell cal-cell-empty"></div>';for(let r=1;r<=B;r++){let h=`${g}-${String(o+1).padStart(2,"0")}-${String(r).padStart(2,"0")}`,F=_[h]||[],k=h===c;y+=`
          <div class="cal-cell ${k?"cal-today":""} ${F.length?"cal-has-events":""}"
               data-date="${h}" tabindex="0" role="button" aria-label="${h}">
            <div class="cal-day-num ${k?"today-num":""}">${r}</div>
            <div class="cal-events-preview">
              ${F.slice(0,3).map(E=>`
                <div class="cal-event-dot cal-color-${E.color||"gray"}" title="${Ee(E.title||E.type)}">
                  <span class="cal-event-dot-label">${Gn(E.title||E.branch_name||E.type,18)}</span>
                </div>
              `).join("")}
              ${F.length>3?`<div class="cal-more">+${F.length-3} lagi</div>`:""}
            </div>
          </div>`}let m=(z+B)%7;if(m!==0)for(let r=0;r<7-m;r++)y+='<div class="cal-cell cal-cell-empty"></div>';y+="</div>",x.innerHTML=y,x.querySelectorAll(".cal-cell[data-date]").forEach(r=>{r.addEventListener("click",()=>{let h=r.dataset.date,F=_[h]||[];if(!F.length)return;let k=document.getElementById("cal-event-list"),E=new Date(h+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=E,document.getElementById("cal-event-items").innerHTML=F.map(n=>`
            <div class="cal-event-item cal-color-border-${n.color||"gray"}">
              <div class="cal-event-type">
                ${Jn(n.type)}
                ${n.date_type?` <span class="badge" style="font-size:0.75rem;padding:2px 6px;background:rgba(0,0,0,0.06);margin-left:4px;">${Ee(n.date_type)}</span>`:""}
              </div>
              <div class="cal-event-title" style="font-weight:600;font-size:0.95rem;margin-top:2px;">${Ee(n.title||"-")}</div>
              ${n.branch_name?`<div class="cal-event-branch" style="color:var(--text-2);font-size:0.85rem;margin-top:2px;">\u{1F4CD} ${Ee(n.branch_name)}</div>`:""}
              ${n.pic||n.period||n.opening_date||n.target_date||n.completion_date||n.notes?`
                <div class="cal-event-meta" style="margin-top:6px;padding-top:6px;border-top:1px dashed var(--border);font-size:0.8rem;color:var(--text-2);display:flex;flex-direction:column;gap:3px;">
                  ${n.pic?`<div>\u{1F464} <b>PIC:</b> ${Ee(n.pic)}</div>`:""}
                  ${n.period?`<div>\u{1F4C5} <b>Periode:</b> ${Ee(n.period)}</div>`:""}
                  ${n.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${Ee(zt(n.opening_date))}</div>`:""}
                  ${n.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${Ee(zt(n.target_date))}</div>`:""}
                  ${n.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${Ee(zt(n.completion_date))}</div>`:""}
                  ${n.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${Ee(n.notes)}</div>`:""}
                </div>
              `:""}
              ${n.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${n.status==="Done"?"badge-success":"badge-warning"}">${Ee(n.status)}</span></div>`:""}
              ${n.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${n.days_remaining} hari</div>`:""}
            </div>
          `).join(""),k.style.display="block"})})}catch(g){console.error("[Calendar] Render error:",g),x&&(x.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let b=()=>{i().then(()=>s())};ot.on("data:changed",b),s()}function Un(e){if(!e||e==="-"||String(e).trim()==="")return"";let t=String(e).trim();if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);let a=t.split(/[\/\-\.]/);if(a.length===3){let[i,s,b]=a.map(x=>x.trim());if(b.length===4)return`${b}-${s.padStart(2,"0")}-${i.padStart(2,"0")}`;if(i.length===4)return`${i}-${s.padStart(2,"0")}-${b.padStart(2,"0")}`}return t.slice(0,10)}function zt(e){if(!e||e==="-"||String(e).trim()==="")return"";let a=String(e).trim().split(/[\/\-\.]/);if(a.length===3){let[i,s,b]=a.map(x=>x.trim());if(i.length===4)return`${b.padStart(2,"0")}-${s.padStart(2,"0")}-${i}`;if(b.length===4)return`${i.padStart(2,"0")}-${s.padStart(2,"0")}-${b}`}return e}function Gn(e,t){return e?e.length>t?e.slice(0,t)+"\u2026":e:""}function Ee(e){return e?String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Jn(e){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging"}[e]||e}async function Na(e){let t=Z(),a=(t?.full_name||t?.username||"U")[0].toUpperCase(),s={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[t?.role]||"#64748B";e.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${s},${s}99)">
            ${a}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${t?.full_name||"\u2014"}</div>
            <div class="profile-username">@${t?.username||"\u2014"}</div>
            <span class="badge badge-info" style="background:${s}18;color:${s};margin-top:6px">
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
            <span class="info-value" style="color:${s};font-weight:700">${t?.role||"\u2014"}</span>
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
  `;let b=localStorage.getItem("fm_token"),x=document.getElementById("session-info");if(b&&x)try{let g=JSON.parse(atob(b.split(".")[1])),o=new Date(g.exp*1e3);x.textContent=`Berakhir: ${o.toLocaleString("id-ID")}`}catch{x.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async g=>{g.preventDefault();let o=document.getElementById("pwd-error"),l=document.getElementById("pwd-success"),A=document.getElementById("btn-save-pwd");o.style.display="none",l.style.display="none";let T=g.target,C=T.current_password.value,_=T.new_password.value,z=T.confirm_password.value;if(_!==z){o.textContent="\u274C Konfirmasi password tidak cocok.",o.style.display="block";return}if(_.length<6){o.textContent="\u274C Password baru minimal 6 karakter.",o.style.display="block";return}A.disabled=!0,A.textContent="\u23F3 Menyimpan...";let B=await N("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:C,new_password:_})});A.disabled=!1,A.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',B.ok?(l.textContent="\u2705 Password berhasil diubah.",l.style.display="block",T.reset(),H("Password berhasil diubah.")):(o.textContent=B.data?.error||"Gagal mengubah password.",o.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}var pt={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function ge(e){if(e==null||e==="")return null;if(e instanceof Date)return isNaN(e.getTime())?null:e.toISOString().slice(0,10);let t=String(e).trim();if(t===""||t==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);if(/^\d{4,5}$/.test(t)){let s=Number(t);if(s>2e4&&s<99999){let b=new Date(Date.UTC(1899,11,30)+s*864e5);return isNaN(b.getTime())?null:b.toISOString().slice(0,10)}}let a=t.split(/[\/\-\.]/);if(a.length===3){let[s,b,x]=a.map(A=>A.trim()),g=Number(s),o=Number(b),l=Number(x);if(s.length===4&&g>1900)return`${s}-${b.padStart(2,"0")}-${x.padStart(2,"0")}`;if(x.length===4&&l>1900)return g>12?`${x}-${b.padStart(2,"0")}-${s.padStart(2,"0")}`:o>12?`${x}-${s.padStart(2,"0")}-${b.padStart(2,"0")}`:`${x}-${b.padStart(2,"0")}-${s.padStart(2,"0")}`;if(x.length===2&&!isNaN(l)){let A=l>=50?`19${x}`:`20${x}`;return g>12?`${A}-${b.padStart(2,"0")}-${s.padStart(2,"0")}`:`${A}-${b.padStart(2,"0")}-${s.padStart(2,"0")}`}}let i=new Date(t);return isNaN(i.getTime())?null:i.toISOString().slice(0,10)}function qa(e){return Object.values(e).every(t=>t==null||String(t).trim()==="")}var Yn={validation:{required:[],map:e=>({cabang:e.CABANG,pic:e.PIC,kegiatan:e.KEGIATAN,quartal:e.QUARTAL,masa_pkwt:e["MASA PKWT"],pic_pelapor:e["PIC PELAPOR"],kontrak:e.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:e=>({name:e["Nama SOP"],category:e.Kategori||"Umum",document_link:e["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({full_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||e.Divisi||"FACILITY CARE",phone:e["No. Hp"]||e["No. HP"],join_date:ge(e["Tanggal Masuk"]||e["Tgl Masuk"]),status:e.Status||"",target_pindah_os:ge(e["Target Pindah OS"]),target_selesai_os:ge(e["Target Selesai OS"]),notes:e.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({employee_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||"FACILITY CARE",start_date:ge(e["Tanggal Mulai"]),end_date:ge(e["Tanggal Selesai"]),contract_type:e["Tipe Kontrak"]||"",pkwt_number:e.PKWT||"",status:e.Status||"",notes:e.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:e=>({report_date:ge(e["Tanggal Info"]),branch_name:e.Cabang,category:e.Kategori,source:e["Sumber Laporan"],complaint:e.Keluhan,employee_name:e["Nama FC"],fc_specialist:e["FC Spesialis"],solution:e.Solusi,status:e.Status||"",completion_date:ge(e["Tanggal Selesai"])})},one_on_one:{required:[],map:e=>({meeting_date:ge(e.Tanggal),branch_name:e.Cabang,employee_name:e["Nama Karyawan"],pic:e.Pic,problem:e.Masalah,solution:e.Solusi,status:e.Status||"",completion_date:ge(e["Tanggal Selesai"]),document_link:e["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:e=>({branch_name:e.Cabang,activity_type:e.Kegiatan,period:e.Periode,pic:e.Pic||e.PIC,opening_date:ge(e["Tanggal Opening"]||e["Tgl Opening"]),target_date:ge(e["Tanggal Target"]||e["Tgl Target"]),completion_date:ge(e["Tanggal Selesai"]||e["Tgl Selesai"]),status:e.Status||"",notes:e.Keterangan||e.Catatan})},inspection:{required:[],map:e=>({inspection_date:ge(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",fc_score:e["Point FC SP"]!==void 0&&e["Point FC SP"]!==null?parseFloat(String(e["Point FC SP"]).replace(",",".")):null,spv_score:e["Point SPV"]!==void 0&&e["Point SPV"]!==null?parseFloat(String(e["Point SPV"]).replace(",",".")):null,document_link:e.Link,notes:""})},cleaning:{required:[],map:e=>({activity_date:ge(e.Tanggal),branch_name:e.Cabang,activity_type:e["Jenis Kegiatan"]||"General Cleaning",period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},fogging:{required:[],map:e=>({activity_date:ge(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:e=>({info_date:ge(e["Tgl Info"]),branch_name:e.Cabang,problem:e.Permasalahan,pic:e.PIC,done_date:ge(e["Tgl Done"]),status:e.Status||"",notes:e.Ket})},relievers:{required:[],map:e=>({branch_name:e.Cabang,original_fc_name:e["Nama Facility care"],period:e.Periode,reliever_name:e.Relifer,backup_date:ge(e["Tanggal Back Up"]),completion_date:ge(e["Tanggal Selesai"]),reason:e.Keterangan,shift:e.Shift,status:e.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:e=>({training_date:ge(e.Tanggal),batch:e.Batch,subject:e.Materi,participants:e.Peserta,branch_name:e.Cabang,trainer:e.Trainer,score:e.Nilai!==void 0&&e.Nilai!==null?parseFloat(String(e.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:e=>({name:e["Master Checklist"],category:"Umum",document_link:e["Link Document"],description:""})},supply:{required:[],map:e=>({submitted_at:ge(e.Timestamp),submitter_name:e["Nama Lengkap"],branch_name:e["Kebutuhan Untuk Cabang"],tools_items:e["Alat - Alat / Barang"],tools_quantity:e["Jumlah Permintaan Alat / Barang"],chemical_items:e.Chemical,chemical_quantity:e["Jumlah Permintaan Chemical"],additional_notes:e["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:e.Status||""})}};function Wn(e,t){let a=pt[e];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let i=Yn[a.module];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let s=[],b=[],x=[];return t.filter(o=>!qa(o)).forEach((o,l)=>{let A=t.indexOf(o)+2,T=[];i.required.forEach(({key:_,label:z})=>{let B=o[_];if(B==null||String(B).trim()===""){let P=Object.keys(o).filter(c=>c.trim()).join(", ");T.push({column:z,originalValue:B||"",reason:`Kolom "${z}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${P.slice(0,120)}`})}});let C=i.map(o);T.length>0?b.push({row:A,data:C,raw:o,errors:T}):(s.push(o),x.push(C))}),{valid:s,errors:b,mapped:x}}function Ra(e){let t=[];return e.SheetNames.forEach(a=>{let i=pt[a];if(!i)return;let s=e.Sheets[a],b=window.XLSX.utils.sheet_to_json(s,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),x=Wn(a,b),g=b.filter(o=>!qa(o));t.push({sheetName:a,module:i.module,label:i.label,total:g.length,valid:x.mapped.length,errorCount:x.errors.length,errors:x.errors,mapped:x.mapped,skipped:!1})}),t}function Oa(){let e=window.XLSX,t=e.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([i,s])=>{e.utils.book_append_sheet(t,e.utils.json_to_sheet(s),i)}),e.writeFile(t,"Template_Import_Data_Awal_FCMS.xlsx")}function Ha(e){let t=window.XLSX,a=t.utils.book_new(),i=!1;return e.forEach(s=>{if(!s.errors||s.errors.length===0)return;i=!0;let b=s.errors.map(g=>({"No. Baris":g.row,"Kolom Gagal":(g.errors||[]).map(o=>o.column||o).join("; "),"Alasan Error":(g.errors||[]).map(o=>o.reason||o).join("; "),...Object.fromEntries(Object.entries(g.data||{}).map(([o,l])=>[o,l??""]))})),x=t.utils.json_to_sheet(b);t.utils.book_append_sheet(a,x,s.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),i?(t.writeFile(a,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Vn=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","sop","inspection","cleaning","fogging","basecamp","supply"];function Ua(e){e.innerHTML=`
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
              ${Object.entries(pt).map(([B,{label:P}])=>`<span class="import-sheet-tag">\u{1F4C4} ${B} \u2192 ${P}</span>`).join("")}
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
  `;let t=null,a=null,i=0,s={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function b(B){Object.entries(s).forEach(([P,c])=>{c.style.display=P===B?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let B=document.getElementById("btn-backup-db");B.disabled=!0,B.textContent="\u23F3 Memproses Backup...";try{let P=await N("/api/import/backup");if(P.ok){if(!window.XLSX){q("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let c=window.XLSX,y=c.utils.book_new();Object.entries(P.data.database).forEach(([v,m])=>{let r=m.length>0?m:[{}],h=c.utils.json_to_sheet(r);c.utils.book_append_sheet(y,h,v.substring(0,31))}),c.writeFile(y,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),H("Backup berhasil diunduh!")}else q("Gagal memproses backup: "+(P.data?.error||"Unknown error"))}catch(P){q("Gagal memproses backup: "+P.message)}finally{B.disabled=!1,B.textContent="\u{1F4E6} Backup Database"}});let x=document.getElementById("btn-sync-google");x&&x.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let B=x.innerHTML;x.innerHTML='<span class="spinner"></span> Menyinkronkan...',x.disabled=!0;try{let P=await N("/api/sync/google-sheets",{method:"POST"});P.ok?alert("Sinkronisasi Berhasil: "+(P.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(P.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{x.innerHTML=B,x.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Oa(),H("Template Excel berhasil didownload!")});let g=document.getElementById("file-input"),o=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",B=>{B.stopPropagation(),g.click()}),g.addEventListener("change",B=>{B.target.files[0]&&l(B.target.files[0])}),o.addEventListener("dragover",B=>{B.preventDefault(),o.classList.add("drag-over")}),o.addEventListener("dragleave",()=>o.classList.remove("drag-over")),o.addEventListener("drop",B=>{B.preventDefault(),o.classList.remove("drag-over");let P=B.dataTransfer.files[0];P&&P.name.match(/\.xlsx?$/i)?l(P):q("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{t=null,g.value="",document.getElementById("file-info").style.display="none",o.style.display="",b("upload")});async function l(B){t=B,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${B.name} (${(B.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",o.style.display="none",await A(B)}async function A(B){b("validating");let P=document.getElementById("validation-status"),c=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");P.textContent="Membaca file Excel...",c.style.width="20%",await We(200);let y=await B.arrayBuffer(),v=window.XLSX.read(y,{type:"array",cellDates:!0});P.textContent=`Memvalidasi ${v.SheetNames.length} sheet...`,c.style.width="50%",await We(100),a=Ra(v),c.style.width="100%",P.textContent="Validasi selesai!",await We(300),T()}catch(y){b("upload"),q("Gagal memproses file: "+y.message),document.getElementById("file-info").style.display="flex",o.style.display="none"}}function T(){b("preview");let B=a.filter(k=>!k.skipped).length,P=a.reduce((k,E)=>k+E.total,0),c=a.reduce((k,E)=>k+E.valid,0),y=a.reduce((k,E)=>k+E.errorCount,0),v=P>0?Math.round(c/P*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${B} sheet</span>
      <span class="badge badge-secondary">${P} baris</span>
      <span class="badge badge-success">${c} valid (${v}%)</span>
      ${y>0?`<span class="badge badge-danger">${y} error</span>`:""}
    `;let m=document.getElementById("preview-table-container");m.innerHTML=`
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
          ${a.map((k,E)=>`
            <tr class="${k.errorCount>0?"row-error":k.skipped?"row-skipped":"row-ok"}">
              <td><strong>${k.sheetName}</strong></td>
              <td>${k.label}</td>
              <td style="text-align:center">${k.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${k.valid}</span></td>
              <td style="text-align:center">${k.errorCount>0?`<span class="badge badge-danger">${k.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${k.skipped?'<span class="badge badge-neutral">Dilewati</span>':k.errorCount>0&&k.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':k.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':k.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${k.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${E}">\u{1F50D} ${k.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,m.querySelectorAll(".btn-detail-error").forEach(k=>{k.addEventListener("click",()=>{let E=a[Number(k.dataset.idx)];C(E)})});let r=document.getElementById("error-detail-section"),h=document.getElementById("error-detail-container");h.innerHTML="",r.style.display="none";let F=document.getElementById("btn-start-import");c===0?(F.disabled=!0,F.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(F.disabled=!1,y>0?(F.innerHTML=`\u{1F680} Import ${c} Data Valid (${y} dilewati)`,F.title="Baris error akan dilewati, baris valid tetap diimport"):F.innerHTML=`\u{1F680} Mulai Import ${c} Data`)}function C(B){let P=document.getElementById("error-detail-section"),c=document.getElementById("error-detail-container");P.style.display="";let y=B.errors.slice(0,100).map(v=>(Array.isArray(v.errors)?v.errors:[]).map(r=>{let h=typeof r=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${v.row}</span></td>
            <td><strong>${h?r.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${h&&r.originalValue!==void 0?r.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${h?r.reason:r}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${h&&r.aliases?`Gunakan salah satu nama kolom:<br><em>${r.aliases}</em>`:h&&r.hint?r.hint:""}
            </td>
          </tr>
        `}).join("")).join("");c.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${B.sheetName} \u2014 ${B.errorCount} baris error dari ${B.total} total
          ${B.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
        ${B.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,P.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{b("upload"),document.getElementById("file-info").style.display="none",o.style.display="",t=null,g.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!a)return;Ha(a)?H("Log error berhasil didownload."):H("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let B=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";_(B)});async function _(B){b("importing"),i=Date.now();let P=[];Vn.forEach(r=>{let h=a?.find(F=>F.module===r&&F.mapped?.length>0);h&&P.push(h)});let c=document.getElementById("import-steps-list");c.innerHTML=P.map(r=>`
      <div class="import-step-item" id="step-item-${r.module}">
        <span class="step-item-icon" id="step-icon-${r.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${r.label} <span class="step-item-count">(${r.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${r.module}"></span>
      </div>
    `).join("");let y=document.getElementById("import-bar"),v=document.getElementById("import-current-status"),m={totalSheets:P.length,totalRows:P.reduce((r,h)=>r+h.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let r=0;r<P.length;r++){let h=P[r],F=document.getElementById(`step-icon-${h.module}`),k=document.getElementById(`step-status-${h.module}`);F.textContent="\u{1F504}",k.textContent="Mengimport...",v.textContent=`Mengimport ${h.label}...`,y.style.width=`${Math.round(r/P.length*100)}%`;try{let E=await N(`/api/import/${h.module}`,{method:"POST",body:JSON.stringify({rows:h.mapped,onDuplicate:B})});if(E.ok){let n=E.data;m.inserted+=n.inserted||0,m.skipped+=n.skipped||0,m.moduleResults.push({label:h.label,inserted:n.inserted||0,skipped:n.skipped||0,status:"ok"}),F.textContent="\u2705",k.innerHTML=`<span class="badge badge-success">${n.inserted||0} berhasil</span>${n.skipped>0?` <span class="badge badge-neutral">${n.skipped} skip</span>`:""}`}else m.failed++,m.moduleResults.push({label:h.label,inserted:0,skipped:0,status:"error",error:E.data?.error}),F.textContent="\u274C",k.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(E){m.failed++,m.moduleResults.push({label:h.label,inserted:0,skipped:0,status:"error",error:E.message}),F.textContent="\u274C",k.innerHTML='<span class="badge badge-danger">Gagal</span>'}await We(150)}y.style.width="100%",v.textContent="Selesai!",await We(400),z(m)}function z(B){b("summary");let P=((Date.now()-i)/1e3).toFixed(1),c=B.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${c?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${c?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${B.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${B.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${B.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${B.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${B.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${B.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
      <div class="summary-stat-card info">
        <div class="stat-value">${P}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${B.moduleResults.map(y=>`
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{t=null,a=null,g.value="",document.getElementById("file-info").style.display="none",o.style.display="",b("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function We(e){return new Promise(t=>setTimeout(t,e))}ce();var Ga=[{id:1,employee_name:"Teguh Ariyandi",division:"FACILITY CARE",branch_id:18,branch_name:"018. Harapan Indah",tanggal:"2025-08-19",akhir_sp:"2026-02-18",sp_type:"SP 1",status:"Aktif",document_link:"",notes:"Keterlambatan berulang dan ketidaklengkapan checklist harian lantai dasar."},{id:2,employee_name:"Andri",division:"FACILITY CARE",branch_id:1,branch_name:"001. Pondok Bambu",tanggal:"2025-04-10",akhir_sp:"2025-10-09",sp_type:"SP 1",status:"Selesai",document_link:"",notes:"Pelanggaran grooming dan standar pembersihan area tindakan tanpa APD lengkap."},{id:3,employee_name:"M. FIRMANSYAH",division:"FACILITY CARE",branch_id:25,branch_name:"025. Bogor Baru",tanggal:"2024-03-21",akhir_sp:"2024-09-20",sp_type:"SP 2",status:"Selesai",document_link:"",notes:"Meninggalkan area tugas tanpa konfirmasi supervisor selama jam operasional klinik."}];function De(e){if(!e)return{status:"Tidak Ditentukan",label:"Tidak Ada Tanggal",bg:"#F1F5F9",color:"#64748B",border:"#CBD5E1",icon:"\u26AA",daysLeft:0,isExpired:!1,isExpiringSoon:!1};let t=new Date;t.setHours(0,0,0,0);let a=new Date(e);a.setHours(0,0,0,0);let i=a.getTime()-t.getTime(),s=Math.ceil(i/(1e3*60*60*24));return s<0?{status:"Selesai",label:`Masa Sanksi Selesai (${Math.abs(s)} hari lalu)`,shortLabel:"Selesai",bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2705",daysLeft:s,isExpired:!0,isExpiringSoon:!1}:s<=30?{status:"Mendekati Berakhir",label:`Segera Berakhir (Sisa ${s} hari)`,shortLabel:`Sisa ${s} hari`,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F",daysLeft:s,isExpired:!1,isExpiringSoon:!0}:{status:"Aktif",label:`Aktif Berlaku (Sisa ${s} hari)`,shortLabel:`Aktif (${s} hari)`,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F534}",daysLeft:s,isExpired:!1,isExpiringSoon:!1}}function Mt(e=""){let t=(e||"").toUpperCase().trim();return t==="SP 3"?{label:"SP 3 (Terakhir)",bg:"#7F1D1D",color:"#FFFFFF",border:"#991B1B",icon:"\u{1F6A8}"}:t==="SP 2"?{label:"SP 2",bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F6D1}"}:t==="SP 1"?{label:"SP 1",bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F"}:t.includes("TEGURAN")?{label:"Teguran Lisan",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F4AC}"}:{label:e||"SP",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4DC}"}}function qe(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Ja(e){let t=Z(),a=ee(t,"sp","write"),i=ee(t,"sp","write"),s=ee(t,"sp","delete"),b=[],x=[],g=[],o="all",l="all",A="all",T="",C=localStorage.getItem("fcms_sp_view_mode")||"card",_="fcms-sp-page-styles",z=document.getElementById(_);z||(z=document.createElement("style"),z.id=_,document.head.appendChild(z)),z.textContent=`
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
    `;async function B(){try{let[k,E]=await Promise.all([re(),me()]);x=k||[],g=E||[];let n=await N("/api/sp?limit=1000");if(n.ok&&n.data?.data&&Array.isArray(n.data.data)&&n.data.data.length>0)b=n.data.data,localStorage.setItem("fcms_sp_data_cache",JSON.stringify(b));else{let d=localStorage.getItem("fcms_sp_data_cache");d?b=JSON.parse(d):(b=[...Ga],localStorage.setItem("fcms_sp_data_cache",JSON.stringify(b)))}}catch(k){console.warn("Fallback loading SP data from cache/seed:",k);let E=localStorage.getItem("fcms_sp_data_cache");b=E?JSON.parse(E):[...Ga]}h()}function P(){return b.filter(k=>{let E=T.toLowerCase().trim(),n=!E||(k.employee_name||"").toLowerCase().includes(E)||(k.branch_name||"").toLowerCase().includes(E)||(k.division||"").toLowerCase().includes(E)||(k.sp_type||"").toLowerCase().includes(E),d=l==="all"||String(k.branch_id)===String(l),f=A==="all"||k.division===A,w=De(k.akhir_sp),L=!0;return o==="active"?L=!w.isExpired:o==="expiring"?L=w.isExpiringSoon:o==="expired"?L=w.isExpired:o==="sp1"?L=(k.sp_type||"").toUpperCase()==="SP 1":o==="sp2"?L=(k.sp_type||"").toUpperCase()==="SP 2":o==="sp3"?L=(k.sp_type||"").toUpperCase()==="SP 3":o==="teguran"&&(L=(k.sp_type||"").toUpperCase().includes("TEGURAN")),n&&d&&f&&L})}function c(){let k=0,E=0,n=0,d=0,f=0,w=0,L=0,I=new Set;return b.forEach(j=>{j.employee_name&&I.add(j.employee_name.toLowerCase().trim());let $=De(j.akhir_sp);$.isExpired?n++:k++,$.isExpiringSoon&&E++;let p=(j.sp_type||"").toUpperCase();p==="SP 1"?d++:p==="SP 2"?f++:p==="SP 3"?w++:p.includes("TEGURAN")&&L++}),{total:b.length,active:k,expiring:E,expired:n,sp1:d,sp2:f,sp3:w,teguran:L,uniqueEmployees:I.size}}function y(k=null){let E=!!k,n=k?.tanggal?k.tanggal.slice(0,10):new Date().toISOString().slice(0,10),d=k?.akhir_sp?k.akhir_sp.slice(0,10):"";if(!d){let I=new Date(n);I.setMonth(I.getMonth()+6),d=I.toISOString().slice(0,10)}let f=document.createElement("div");f.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <input type="text" id="sp-modal-emp" list="sp-emp-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih atau ketik nama..." value="${k?.employee_name?String(k.employee_name).replace(/"/g,"&quot;"):""}" required />
            <datalist id="sp-emp-list">
              ${g.map(I=>`<option value="${I.label}">${I.label}</option>`).join("")}
            </datalist>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-div" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="FACILITY CARE" ${k?.division==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${k?.division==="SECURITY"?"selected":""}>SECURITY</option>
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
              ${x.map(I=>`<option value="${I.value}" ${String(k?.branch_id)===String(I.value)?"selected":""}>${I.label}</option>`).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tingkat / Jenis SP <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="SP 1" ${k?.sp_type==="SP 1"?"selected":""}>SP 1 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 2" ${k?.sp_type==="SP 2"?"selected":""}>SP 2 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 3" ${k?.sp_type==="SP 3"?"selected":""}>SP 3 (Masa Berlaku 6 Bulan / Terakhir)</option>
              <option value="Teguran Lisan" ${k?.sp_type==="Teguran Lisan"?"selected":""}>Teguran Lisan (Pembinaan)</option>
            </select>
          </div>
        </div>

        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Terbit SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${n}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Akhir Masa Berlaku SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-akhir" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${d}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Link Berkas / Foto BAP / Scan Surat SP
          </label>
          <input type="url" id="sp-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://drive.google.com/..." value="${k?.document_link?String(k.document_link).replace(/"/g,"&quot;"):""}" />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Kronologi Pelanggaran / Catatan Pembinaan
          </label>
          <textarea id="sp-modal-notes" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan ringkasan pelanggaran disiplin kerja yang terjadi...">${k?.notes||k?.description||""}</textarea>
        </div>
      </div>
    `;let w=f.querySelector("#sp-modal-tanggal"),L=f.querySelector("#sp-modal-akhir");w.addEventListener("change",()=>{if(w.value){let I=new Date(w.value);I.setMonth(I.getMonth()+6),L.value=I.toISOString().slice(0,10)}}),X({title:E?"\u270F\uFE0F Edit Data Surat Peringatan (SP)":"\u2795 Terbitkan Surat Peringatan (SP) Baru",content:f,size:"md",confirmText:E?"Simpan Perubahan":"Terbitkan SP",confirmClass:"btn-primary",onConfirm:async(I,j)=>{let $=(f.querySelector("#sp-modal-emp").value||"").trim(),p=f.querySelector("#sp-modal-div").value,S=f.querySelector("#sp-modal-branch").value,u=f.querySelector("#sp-modal-type").value,M=f.querySelector("#sp-modal-tanggal").value,R=f.querySelector("#sp-modal-akhir").value,K=(f.querySelector("#sp-modal-link").value||"").trim(),O=(f.querySelector("#sp-modal-notes").value||"").trim();if(!$){q("Nama Karyawan wajib diisi");return}if(!S){q("Cabang penempatan wajib dipilih");return}if(!M||!R){q("Tanggal SP dan Akhir SP wajib diisi");return}let D=x.find(G=>String(G.value)===String(S)),Y=D?D.label:"Cabang",J={employee_name:$,division:p,branch_id:Number(S),branch_name:Y,sp_type:u,tanggal:M,akhir_sp:R,document_link:K,notes:O,status:De(R).status};try{if(E){let G=await N(`/api/sp/${k.id}`,{method:"PUT",body:JSON.stringify(J)}),U=b.findIndex(W=>W.id===k.id);U!==-1&&(b[U]={...b[U],...J}),H("Data SP berhasil diperbarui")}else{let G=await N("/api/sp",{method:"POST",body:JSON.stringify(J)}),U=G.ok&&G.data?.id?G.data.id:Date.now();b.unshift({id:U,...J}),H("Surat Peringatan baru berhasil diterbitkan")}localStorage.setItem("fcms_sp_data_cache",JSON.stringify(b)),j(),h()}catch(G){console.error("Error saving SP:",G),q("Gagal menyimpan SP: "+G.message)}}})}function v(k){ie(`Hapus data Surat Peringatan untuk karyawan <strong>${k.employee_name}</strong> (${k.sp_type})?`,async()=>{try{await N(`/api/sp/${k.id}`,{method:"DELETE"})}catch(E){console.warn("API delete error fallback to local:",E)}b=b.filter(E=>E.id!==k.id),localStorage.setItem("fcms_sp_data_cache",JSON.stringify(b)),H(`Data SP ${k.employee_name} berhasil dihapus`),h()},"Hapus Data SP")}function m(k){let E=De(k.akhir_sp),n=Mt(k.sp_type),d=document.createElement("div");d.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div style="display:flex; align-items:center; gap:14px; padding-bottom:14px; border-bottom:1px solid #E2E8F0;">
          <div class="sp-avatar-box" style="width:54px; height:54px; font-size:1.2rem;">
            ${k.employee_name.slice(0,2).toUpperCase()}
          </div>
          <div>
            <h3 style="margin:0 0 4px 0; font-size:1.15rem; font-weight:700; color:#0F172A;">
              ${k.employee_name}
            </h3>
            <div style="display:flex; align-items:center; gap:8px;">
              <span class="badge" style="background:#EFF6FF; color:#1D4ED8; font-weight:700;">${k.division||"FACILITY CARE"}</span>
              <span style="color:#64748B; font-size:0.85rem;">\u{1F4CD} ${k.branch_name||"Cabang"}</span>
            </div>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Tingkat Sanksi</div>
            <span class="sp-type-badge" style="background:${n.bg}; color:${n.color}; border:1px solid ${n.border}; font-size:0.85rem;">
              <span>${n.icon}</span>
              <span>${n.label}</span>
            </span>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Status Masa Berlaku</div>
            <span class="sp-status-badge" style="background:${E.bg}; color:${E.color}; border:1px solid ${E.border}; font-size:0.85rem;">
              <span>${E.icon}</span>
              <span>${E.label}</span>
            </span>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Tanggal Terbit SP</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${qe(k.tanggal)}</div>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Akhir Masa Berlaku</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${qe(k.akhir_sp)}</div>
          </div>
        </div>

        ${k.notes||k.description?`
          <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:10px; padding:14px;">
            <div style="font-size:0.78rem; font-weight:700; color:#92400E; margin-bottom:4px;">\u{1F4DD} Catatan Pelanggaran / Pembinaan:</div>
            <div style="font-size:0.88rem; color:#78350F; line-height:1.5;">${k.notes||k.description}</div>
          </div>
        `:""}

        ${k.document_link?`
          <div>
            <a href="${k.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-primary" style="width:100%; justify-content:center; padding:10px; font-weight:600; display:flex; align-items:center; gap:8px;">
              <span>\u{1F4C4}</span> Buka Dokumen BAP / Scan SP Fisik
            </a>
          </div>
        `:`
          <div style="text-align:center; padding:8px; font-size:0.82rem; color:#94A3B8;">
            (Belum ada scan dokumen / BAP yang diunggah)
          </div>
        `}
      </div>
    `,X({title:"\u{1F4CB} Lembar Riwayat Surat Peringatan (SP)",content:d,size:"md",cancelText:"Tutup"})}async function r(){try{let{downloadExcel:k}=await Promise.resolve().then(()=>(de(),fe)),E=b.map((n,d)=>{let f=De(n.akhir_sp);return{No:d+1,"Nama Karyawan":n.employee_name||"",Divisi:n.division||"",Cabang:n.branch_name||"","Jenis SP":n.sp_type||"","Tanggal SP":n.tanggal||"","Akhir SP":n.akhir_sp||"","Status Masa Berlaku":f.status,"Sisa Hari":f.daysLeft>0?f.daysLeft:0,"Link BAP / Foto":n.document_link||"","Catatan / Kronologi":n.notes||n.description||""}});k(E,`Data_Surat_Peringatan_FCMS_${new Date().toISOString().slice(0,10)}`),H("Berhasil mengunduh data SP ke format Excel.")}catch(k){q("Gagal ekspor: "+k.message)}}function h(){let k=P(),E=c();e.innerHTML=`
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
              ${a?`
                <button id="sp-btn-add" class="btn btn-primary sp-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Terbitkan SP</span>
                </button>
              `:""}
              
              <button id="sp-btn-excel" class="btn btn-ghost sp-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

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
                <div class="sp-stat-value">${E.active}</div>
                <div class="sp-stat-label">SP Aktif Berlaku</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">\u26A0\uFE0F</div>
              <div>
                <div class="sp-stat-value">${E.expiring}</div>
                <div class="sp-stat-label">Segera Berakhir (&lt;30 Hari)</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">\u2705</div>
              <div>
                <div class="sp-stat-value">${E.expired}</div>
                <div class="sp-stat-label">Masa Sanksi Selesai</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">\u{1F465}</div>
              <div>
                <div class="sp-stat-value">${E.uniqueEmployees} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">karyawan</span></div>
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
              <input type="text" id="sp-search-input" class="sp-search-input" placeholder="Cari nama karyawan, cabang, atau divisi..." value="${T?T.replace(/"/g,"&quot;"):""}" />
            </div>

            <!-- Branch Dropdown -->
            <select id="sp-filter-branch" class="sp-filter-select">
              <option value="all">Semua Cabang (${x.length})</option>
              ${x.map(n=>`<option value="${n.value}" ${String(l)===String(n.value)?"selected":""}>${n.label}</option>`).join("")}
            </select>

            <!-- Division Dropdown -->
            <select id="sp-filter-division" class="sp-filter-select">
              <option value="all" ${A==="all"?"selected":""}>Semua Divisi</option>
              <option value="FACILITY CARE" ${A==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${A==="SECURITY"?"selected":""}>SECURITY</option>
            </select>

            <div class="sp-view-toggle">
              <button type="button" class="sp-view-btn ${C==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="sp-view-btn ${C==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Filter Pills -->
          <div class="sp-filter-pills">
            <button class="sp-pill ${o==="all"?"active":""}" data-filter="all">
              <span>Semua SP</span>
              <span class="sp-pill-count">${E.total}</span>
            </button>
            <button class="sp-pill ${o==="active"?"active":""}" data-filter="active">
              <span>\u{1F534} SP Aktif</span>
              <span class="sp-pill-count">${E.active}</span>
            </button>
            <button class="sp-pill ${o==="expiring"?"active":""}" data-filter="expiring">
              <span>\u26A0\uFE0F &lt;30 Hari</span>
              <span class="sp-pill-count">${E.expiring}</span>
            </button>
            <button class="sp-pill ${o==="expired"?"active":""}" data-filter="expired">
              <span>\u2705 Selesai</span>
              <span class="sp-pill-count">${E.expired}</span>
            </button>
            <button class="sp-pill ${o==="sp1"?"active":""}" data-filter="sp1">
              <span>SP 1</span>
              <span class="sp-pill-count">${E.sp1}</span>
            </button>
            <button class="sp-pill ${o==="sp2"?"active":""}" data-filter="sp2">
              <span>SP 2</span>
              <span class="sp-pill-count">${E.sp2}</span>
            </button>
            <button class="sp-pill ${o==="sp3"?"active":""}" data-filter="sp3">
              <span>SP 3</span>
              <span class="sp-pill-count">${E.sp3}</span>
            </button>
          </div>
        </div>

        <!-- Content Area -->
        <div id="sp-list-container">
          ${k.length===0?`
            <div class="sp-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data SP Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="sp-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:C==="table"?`
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
                    ${k.map((n,d)=>{let f=De(n.akhir_sp),w=Mt(n.sp_type);return`
                        <tr data-id="${n.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${d+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="sp-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${(n.employee_name||"SP").slice(0,2).toUpperCase()}
                              </div>
                              <div style="font-weight:700; color:#0F172A; white-space:nowrap;">
                                ${n.employee_name}
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#334155; font-size:0.84rem;">\u{1F4CD} ${n.branch_name||"Cabang"}</div>
                            <div style="font-size:0.74rem; color:#2563EB; font-weight:600; margin-top:2px;">${n.division||"FACILITY CARE"}</div>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="sp-type-badge" style="background:${w.bg}; color:${w.color}; border:1px solid ${w.border};">
                              <span>${w.icon}</span>
                              <span>${w.label}</span>
                            </span>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            ${qe(n.tanggal)}
                          </td>
                          <td style="white-space:nowrap; font-weight:700; color:#0F172A;">
                            ${qe(n.akhir_sp)}
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="sp-status-badge" style="background:${f.bg}; color:${f.color}; border:1px solid ${f.border};">
                              <span>${f.icon}</span>
                              <span>${f.shortLabel}</span>
                            </span>
                          </td>
                          <td style="max-width:240px;">
                            <div style="font-size:0.82rem; color:#64748B; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; margin-bottom:4px;" title="${String(n.notes||n.description||"").replace(/"/g,"&quot;")}">
                              ${n.notes||n.description||"Pencatatan sanksi peringatan disiplin operasional."}
                            </div>
                            <button class="btn btn-sm btn-ghost sp-detail-btn" style="border:1px solid #CBD5E1; border-radius:6px; font-weight:600; font-size:0.74rem; padding:2px 8px;" data-id="${n.id}">
                              \u{1F441}\uFE0F Rincian BAP
                            </button>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <div style="display:inline-flex; align-items:center; gap:6px;">
                              ${n.document_link?`
                                <a href="${n.document_link}" target="_blank" rel="noopener noreferrer" class="sp-icon-btn" title="Buka Dokumen">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                                </a>
                              `:""}
                              ${i?`
                                <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${n.id}">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              `:""}
                              ${s?`
                                <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${n.id}">
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
              ${k.map(n=>{let d=De(n.akhir_sp),f=Mt(n.sp_type);return`
                  <div class="sp-card" data-id="${n.id}">
                    <div>
                      <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px;">
                        <span class="sp-type-badge" style="background:${f.bg}; color:${f.color}; border:1px solid ${f.border};">
                          <span>${f.icon}</span>
                          <span>${f.label}</span>
                        </span>
                        <span class="sp-status-badge" style="background:${d.bg}; color:${d.color}; border:1px solid ${d.border};">
                          <span>${d.icon}</span>
                          <span>${d.shortLabel}</span>
                        </span>
                      </div>

                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="sp-avatar-box" style="width:48px; height:48px; font-size:1.05rem;">
                          ${n.employee_name.slice(0,2).toUpperCase()}
                        </div>
                        <div>
                          <h3 style="margin:0 0 3px 0; font-size:1.1rem; font-weight:700; color:#0F172A;">
                            ${n.employee_name}
                          </h3>
                          <div style="font-size:0.8rem; color:#64748B;">
                            \u{1F4CD} ${n.branch_name||"Cabang"} \u2022 <span style="color:#2563EB; font-weight:600;">${n.division||"FC"}</span>
                          </div>
                        </div>
                      </div>

                      <div style="background:#F8FAFC; border-radius:10px; padding:10px 12px; margin-bottom:14px; font-size:0.82rem; display:flex; justify-content:space-between;">
                        <div>
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Tanggal Terbit</div>
                          <div style="font-weight:700; color:#1E293B;">${qe(n.tanggal)}</div>
                        </div>
                        <div style="text-align:right;">
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Akhir SP</div>
                          <div style="font-weight:700; color:#0F172A;">${qe(n.akhir_sp)}</div>
                        </div>
                      </div>

                      <p style="font-size:0.83rem; color:#64748B; line-height:1.45; margin:0 0 16px 0; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">
                        ${n.notes||n.description||"Pencatatan sanksi peringatan disiplin operasional."}
                      </p>
                    </div>

                    <div style="padding-top:12px; border-top:1px solid #F1F5F9; display:flex; align-items:center; justify-content:space-between;">
                      <button class="btn btn-sm btn-ghost sp-detail-btn" style="border:1px solid #CBD5E1; border-radius:8px; font-weight:600; font-size:0.82rem; padding:6px 12px;" data-id="${n.id}">
                        \u{1F441}\uFE0F Rincian BAP
                      </button>

                      <div style="display:flex; align-items:center; gap:6px;">
                        ${n.document_link?`
                          <a href="${n.document_link}" target="_blank" rel="noopener noreferrer" class="sp-icon-btn" title="Buka Dokumen">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                          </a>
                        `:""}
                        ${i?`
                          <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${n.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                        `:""}
                        ${s?`
                          <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${n.id}">
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
    `,F()}function F(){let k=e.querySelector("#sp-btn-add");k&&(k.onclick=()=>y());let E=e.querySelector("#sp-btn-excel");E&&(E.onclick=()=>r());let n=e.querySelector("#sp-btn-refresh");n&&(n.onclick=()=>B());let d=e.querySelector("#sp-btn-reset-filter");d&&(d.onclick=()=>{T="",l="all",A="all",o="all",h()});let f=e.querySelector("#sp-search-input");f&&(f.oninput=I=>{T=I.target.value,h();let j=e.querySelector("#sp-search-input");j&&(j.focus(),j.setSelectionRange(j.value.length,j.value.length))});let w=e.querySelector("#sp-filter-branch");w&&(w.onchange=I=>{l=I.target.value,h()});let L=e.querySelector("#sp-filter-division");L&&(L.onchange=I=>{A=I.target.value,h()}),e.querySelectorAll(".sp-pill").forEach(I=>{I.onclick=()=>{o=I.getAttribute("data-filter")||"all",h()}}),e.querySelectorAll(".sp-view-btn").forEach(I=>{I.onclick=()=>{let j=I.getAttribute("data-view");j&&j!==C&&(C=j,localStorage.setItem("fcms_sp_view_mode",j),h())}}),e.querySelectorAll(".sp-detail-btn, .sp-view-detail-link").forEach(I=>{I.onclick=j=>{j.stopPropagation();let $=Number(I.getAttribute("data-id")),p=b.find(S=>S.id===$);p&&m(p)}}),e.querySelectorAll(".sp-edit-btn").forEach(I=>{I.onclick=j=>{j.stopPropagation();let $=Number(I.getAttribute("data-id")),p=b.find(S=>S.id===$);p&&y(p)}}),e.querySelectorAll(".sp-delete-btn").forEach(I=>{I.onclick=j=>{j.stopPropagation();let $=Number(I.getAttribute("data-id")),p=b.find(S=>S.id===$);p&&v(p)}})}await B()}ce();var Kt=[{id:1,tanggal:"2026-10-02",employee_name:"TAUFIK HIDAYAT",from_branch_id:12,from_branch_name:"012. BSD",to_branch_id:0,to_branch_name:"000. Management",status:"Selesai",document_link:""},{id:2,tanggal:"2026-09-30",employee_name:"ELVAN JANGGA PRAMANA",from_branch_id:55,from_branch_name:"055. Gajah Mada",to_branch_id:0,to_branch_name:"000. Management",status:"Proses",document_link:""},{id:3,tanggal:"2026-06-23",employee_name:"Widya Astuti",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:7,to_branch_name:"007. Bekasi",status:"Selesai",document_link:""},{id:4,tanggal:"2026-06-22",employee_name:"Gema Azan Saputra",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:21,to_branch_name:"021. Kelapa Gading Boulevard",status:"Selesai",document_link:""},{id:5,tanggal:"2026-06-22",employee_name:"Ahmad Farhan",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:62,to_branch_name:"062. Summarecon Bekasi",status:"Selesai",document_link:""}];function Ya(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Wa(e){let t=[],a=[],i=[],s="all",b="all",x="all",g="",o=!1,l=localStorage.getItem("fcms_mutasi_view_mode")||"card",A=Z(),T=ee(A,"mutasi","write"),C=ee(A,"mutasi","write"),_=ee(A,"mutasi","delete");try{[a,i]=await Promise.all([re(),me()])}catch(r){console.error("Error fetching cache:",r)}async function z(){o=!0,v();try{let r=await N("/api/mutasi?limit=10000");if(r.ok){let h=r.data?.data||r.data||[];t=Array.isArray(h)&&h.length>0?h:Kt}else t=Kt}catch(r){console.warn("Gagal load mutasi, menggunakan data lokal:",r),t=Kt}finally{o=!1,v()}}async function B(){try{let{downloadExcel:r}=await Promise.resolve().then(()=>(de(),fe)),h=t.map((F,k)=>({No:k+1,Tanggal:F.tanggal||"","Nama Karyawan":F.employee_name||"","Cabang Asal":F.from_branch_name||"","Cabang Tujuan":F.to_branch_name||"",Status:F.status||"","Link Dokumen":F.document_link||""}));r(h,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`),H("Data mutasi berhasil diekspor ke Excel!")}catch(r){q("Gagal ekspor Excel: "+r.message)}}async function P(r){if(!C)return;let h=t.find(F=>F.id===r);ie({title:"Hapus Data Mutasi",message:`Apakah Anda yakin ingin menghapus data mutasi untuk "${h?.employee_name||"karyawan ini"}"?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let F=await N(`/api/mutasi/${r}`,{method:"DELETE"});F.ok?(H("Data mutasi berhasil dihapus"),t=t.filter(k=>k.id!==r),v()):q(F.data?.error||"Gagal menghapus mutasi")}catch(F){q("Gagal menghapus data mutasi: "+F.message)}}})}function c(r=null){if(!C)return;let h=!!r,F=r?.tanggal?r.tanggal.slice(0,10):new Date().toISOString().slice(0,10),k=document.createElement("div");k.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="mutasi-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${F}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Proses" ${r?.status==="Proses"?"selected":""}>\u23F3 Proses</option>
              <option value="Selesai" ${r?.status==="Selesai"||!r?"selected":""}>\u2705 Selesai</option>
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Karyawan <span style="color:#DC2626;">*</span>
          </label>
          <input list="emp-list-options" id="mutasi-modal-emp" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Ketik atau pilih nama karyawan..." value="${r?.employee_name?String(r.employee_name).replace(/"/g,"&quot;"):""}" required />
          <datalist id="emp-list-options">
            ${i.map(E=>`<option value="${typeof E=="string"?E:E.label||E.name||""}">`).join("")}
          </datalist>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Asal <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-from" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang Asal --</option>
              ${a.map(E=>`
                <option value="${E.value}" ${String(r?.from_branch_id)===String(E.value)?"selected":""}>
                  ${E.label}
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
              ${a.map(E=>`
                <option value="${E.value}" ${String(r?.to_branch_id)===String(E.value)?"selected":""}>
                  ${E.label}
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
            placeholder="https://drive.google.com/file/d/..." value="${r?.document_link?String(r.document_link).replace(/"/g,"&quot;"):""}" />
        </div>
      </div>
    `,X({title:h?"\u270F\uFE0F Edit Data Mutasi Karyawan":"\u2795 Tambah Data Mutasi Karyawan",content:k,size:"md",confirmText:h?"Simpan Perubahan":"Simpan Mutasi",confirmClass:"btn-primary",onConfirm:async(E,n)=>{let d=k.querySelector("#mutasi-modal-tanggal").value,f=(k.querySelector("#mutasi-modal-emp").value||"").trim(),w=k.querySelector("#mutasi-modal-from").value,L=k.querySelector("#mutasi-modal-to").value,I=k.querySelector("#mutasi-modal-status").value,j=(k.querySelector("#mutasi-modal-link").value||"").trim();if(!d){q("Tanggal mutasi wajib diisi");return}if(!f){q("Nama karyawan wajib diisi");return}if(!w){q("Cabang asal wajib dipilih");return}if(!L){q("Cabang tujuan wajib dipilih");return}let $=a.find(u=>String(u.value)===String(w)),p=a.find(u=>String(u.value)===String(L)),S={tanggal:d,employee_name:f,from_branch_id:Number(w),from_branch_name:$?$.label:"Cabang Asal",to_branch_id:Number(L),to_branch_name:p?p.label:"Cabang Tujuan",status:I,document_link:j};try{if(h){let u=await N(`/api/mutasi/${r.id}`,{method:"PUT",body:JSON.stringify(S)});if(u.ok){H("Data mutasi berhasil diperbarui");let M=t.findIndex(R=>R.id===r.id);M!==-1&&(t[M]={...t[M],...S}),n(),v()}else q(u.data?.error||"Gagal memperbarui data mutasi")}else{let u=await N("/api/mutasi",{method:"POST",body:JSON.stringify(S)});if(u.ok){H("Data mutasi baru berhasil ditambahkan");let M=u.data?.id||Date.now();t.unshift({id:M,...S}),n(),v()}else q(u.data?.error||"Gagal menambahkan mutasi")}}catch(u){q("Terjadi kesalahan: "+u.message)}}})}function y(){return t.filter(r=>{if(g){let h=g.toLowerCase(),F=(r.employee_name||"").toLowerCase().includes(h),k=(r.from_branch_name||"").toLowerCase().includes(h),E=(r.to_branch_name||"").toLowerCase().includes(h);if(!F&&!k&&!E)return!1}return!(b!=="all"&&String(r.from_branch_id)!==String(b)||x!=="all"&&String(r.to_branch_id)!==String(x)||s!=="all"&&r.status!==s)})}function v(){let r=y(),h=t.length,F=t.filter(w=>w.status==="Proses").length,k=t.filter(w=>w.status==="Selesai").length,E=new Date,n=E.getFullYear(),d=E.getMonth(),f=t.filter(w=>{if(!w.tanggal)return!1;let L=new Date(w.tanggal);return!isNaN(L.getTime())&&L.getFullYear()===n&&L.getMonth()===d}).length;e.innerHTML=`
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
              ${T?`
                <button id="mutasi-btn-add" class="btn btn-primary mutasi-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Mutasi</span>
                </button>
              `:""}
              
              <button id="mutasi-btn-excel" class="btn btn-ghost mutasi-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

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
                <div class="mutasi-stat-val">${h}</div>
                <div class="mutasi-stat-lbl">Total Mutasi</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u23F3
              </div>
              <div>
                <div class="mutasi-stat-val">${F}</div>
                <div class="mutasi-stat-lbl">Dalam Proses</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(34,197,94,0.2); color:#4ADE80;">
                \u2705
              </div>
              <div>
                <div class="mutasi-stat-val">${k}</div>
                <div class="mutasi-stat-lbl">Mutasi Selesai</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(168,85,247,0.2); color:#C084FC;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="mutasi-stat-val">${f}</div>
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
              <option value="all">Semua Cabang Asal (${a.length})</option>
              ${a.map(w=>`<option value="${w.value}" ${String(b)===String(w.value)?"selected":""}>${w.label}</option>`).join("")}
            </select>

            <select id="mutasi-filter-to" class="mutasi-select">
              <option value="all">Semua Cabang Tujuan (${a.length})</option>
              ${a.map(w=>`<option value="${w.value}" ${String(x)===String(w.value)?"selected":""}>${w.label}</option>`).join("")}
            </select>

            <div class="mutasi-view-toggle">
              <button type="button" class="mutasi-view-btn ${l==="card"?"active":""}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="mutasi-view-btn ${l==="table"?"active":""}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Filter Pills (Swipeable on Mobile) -->
          <div class="mutasi-filter-pills">
            <div class="mutasi-pill ${s==="all"?"active":""}" data-status="all">
              <span>Semua Mutasi</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${h}</span>
            </div>
            <div class="mutasi-pill ${s==="Proses"?"active":""}" data-status="Proses">
              <span style="color:#D97706;">\u23F3</span>
              <span>Proses</span>
              <span class="badge" style="background:rgba(217,119,6,0.15); color:#B45309; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${F}</span>
            </div>
            <div class="mutasi-pill ${s==="Selesai"?"active":""}" data-status="Selesai">
              <span style="color:#16A34A;">\u2705</span>
              <span>Selesai</span>
              <span class="badge" style="background:rgba(22,163,74,0.15); color:#15803D; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${k}</span>
            </div>
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="mutasi-list-container">
          ${r.length===0?`
            <div class="mutasi-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data Mutasi Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="mutasi-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:l==="table"?`
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
                      ${C||_?'<th style="width:90px; text-align:center;">Aksi</th>':""}
                    </tr>
                  </thead>
                  <tbody>
                    ${r.map((w,L)=>{let I=w.status==="Selesai",j=(w.employee_name||"Karyawan").split(" ").map($=>$[0]).filter(Boolean).slice(0,2).join("").toUpperCase();return`
                        <tr data-id="${w.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${L+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="mutasi-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${j}
                              </div>
                              <div>
                                <div style="font-weight:700; color:#0F172A; white-space:nowrap;">${w.employee_name}</div>
                                <div style="font-size:0.72rem; color:#64748B;">ID Mutasi: #${w.id}</div>
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            \u{1F5D3}\uFE0F ${Ya(w.tanggal)}
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#334155; font-size:0.85rem; display:inline-flex; align-items:center; gap:4px;">
                              <span>\u{1F4CD}</span>
                              <span>${w.from_branch_name||"-"}</span>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:700; color:#2563EB; font-size:0.85rem; display:inline-flex; align-items:center; gap:5px;">
                              <span>\u2794</span>
                              <span>${w.to_branch_name||"-"}</span>
                            </div>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="mutasi-status-badge ${I?"selesai":"proses"}">
                              <span>${I?"\u2705":"\u23F3"}</span>
                              <span>${w.status||"Proses"}</span>
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            ${w.document_link?`
                              <a href="${w.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.76rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                                <span>\u{1F4C4} Buka SK</span>
                              </a>
                            `:`
                              <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                            `}
                          </td>
                          ${C||_?`
                            <td style="text-align:center; white-space:nowrap;">
                              <div style="display:inline-flex; align-items:center; gap:6px;">
                                ${C?`
                                  <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${w.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                  </button>
                                `:""}
                                ${_?`
                                  <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${w.id}">
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
              ${r.map(w=>{let L=w.status==="Selesai",I=(w.employee_name||"Karyawan").split(" ").map(j=>j[0]).filter(Boolean).slice(0,2).join("").toUpperCase();return`
                  <div class="mutasi-card" data-id="${w.id}">
                    <div>
                      <!-- Card Header: Status & Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="mutasi-status-badge ${L?"selesai":"proses"}">
                          <span>${L?"\u2705":"\u23F3"}</span>
                          <span>${w.status||"Proses"}</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B; display:inline-flex; align-items:center; gap:4px;">
                          <span>\u{1F5D3}\uFE0F</span>
                          <span>${Ya(w.tanggal)}</span>
                        </span>
                      </div>

                      <!-- Employee Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="mutasi-avatar-box">
                          ${I}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${w.employee_name}
                          </h4>
                          <div style="font-size:0.75rem; color:#64748B;">
                            ID Mutasi: #${w.id}
                          </div>
                        </div>
                      </div>

                      <!-- Route Visual Box -->
                      <div class="mutasi-route-box">
                        <div class="mutasi-route-node">
                          <div class="mutasi-route-label">Cabang Asal</div>
                          <div class="mutasi-route-branch" title="${w.from_branch_name||"-"}">
                            ${w.from_branch_name||"-"}
                          </div>
                        </div>

                        <div class="mutasi-route-arrow">
                          \u2794
                        </div>

                        <div class="mutasi-route-node" style="text-align:right;">
                          <div class="mutasi-route-label">Cabang Tujuan</div>
                          <div class="mutasi-route-branch" style="color:#2563EB;" title="${w.to_branch_name||"-"}">
                            ${w.to_branch_name||"-"}
                          </div>
                        </div>
                      </div>
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:8px;">
                      <div>
                        ${w.document_link?`
                          <a href="${w.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>\u{1F4C4} Buka SK</span>
                          </a>
                        `:`
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${C||_?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${C?`
                            <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${w.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                          `:""}
                          ${_?`
                            <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${w.id}">
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
    `,m()}function m(){let r=e.querySelector("#mutasi-btn-add");r&&(r.onclick=()=>c());let h=e.querySelector("#mutasi-btn-excel");h&&(h.onclick=()=>B());let F=e.querySelector("#mutasi-btn-refresh");F&&(F.onclick=()=>z());let k=e.querySelector("#mutasi-btn-reset-filter");k&&(k.onclick=()=>{g="",b="all",x="all",s="all",v()});let E=e.querySelector("#mutasi-search-input");E&&(E.oninput=f=>{g=f.target.value,v();let w=e.querySelector("#mutasi-search-input");w&&(w.focus(),w.setSelectionRange(w.value.length,w.value.length))});let n=e.querySelector("#mutasi-filter-from");n&&(n.onchange=f=>{b=f.target.value,v()});let d=e.querySelector("#mutasi-filter-to");d&&(d.onchange=f=>{x=f.target.value,v()}),e.querySelectorAll(".mutasi-pill").forEach(f=>{f.onclick=()=>{s=f.getAttribute("data-status")||"all",v()}}),e.querySelectorAll(".mutasi-view-btn").forEach(f=>{f.onclick=()=>{let w=f.getAttribute("data-view");w&&w!==l&&(l=w,localStorage.setItem("fcms_mutasi_view_mode",w),v())}}),e.querySelectorAll(".mutasi-edit-btn").forEach(f=>{f.onclick=w=>{w.stopPropagation();let L=Number(f.getAttribute("data-id")),I=t.find(j=>j.id===L);I&&c(I)}}),e.querySelectorAll(".mutasi-delete-btn").forEach(f=>{f.onclick=w=>{w.stopPropagation();let L=Number(f.getAttribute("data-id"));P(L)}})}await z()}ce();var jt=[{id:1,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:1,employee_name:"ADE SURAHMAN",start_time:"23:00",end_time:"08:00",break_hours:1,total_hours:8,reason:"General Cleaning"},{id:2,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:2,employee_name:"TAUFIK HIDAYAT",start_time:"00:00",end_time:"08:00",break_hours:1,total_hours:7,reason:"General Cleaning"},{id:3,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:3,employee_name:"YUDA WILDAN WIGUNA",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"},{id:4,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:4,employee_name:"REZA RIZKI",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"},{id:5,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:5,employee_name:"ADITYA PRATAMA",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"}],Va=["General Cleaning","Deep Cleaning","Cover manpower","Longshift","Pekerjaan urgent"];function Qa(e,t,a){if(!e||!t)return 0;try{let[i,s]=e.split(":").map(Number),[b,x]=t.split(":").map(Number),g=b*60+x-(i*60+s);g<0&&(g+=24*60);let o=g/60-(parseFloat(a)||0);return o<0&&(o=0),Number.isInteger(o)?o:parseFloat(o.toFixed(2))}catch{return 0}}function Nt(e=""){let t=(e||"").toLowerCase();return t.includes("general")?{bg:"#EFF6FF",color:"#1D4ED8",border:"#BFDBFE",icon:"\u{1F9F9}"}:t.includes("deep")?{bg:"#F5F3FF",color:"#6D28D9",border:"#DDD6FE",icon:"\u2728"}:t.includes("cover")?{bg:"#ECFDF5",color:"#047857",border:"#A7F3D0",icon:"\u{1F465}"}:t.includes("longshift")?{bg:"#FFFBEB",color:"#B45309",border:"#FDE68A",icon:"\u23F3"}:t.includes("urgent")?{bg:"#FFF1F2",color:"#BE123C",border:"#FECDD3",icon:"\u{1F6A8}"}:{bg:"#F8FAFC",color:"#475569",border:"#CBD5E1",icon:"\u23F1\uFE0F"}}function qt(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Xa(e){let t=[],a=[],i=[],s="all",b="all",x="",g=!1,o=localStorage.getItem("fcms_overtime_view_mode")||"card",l=Z(),A=ee(l,"overtime","write"),T=ee(l,"overtime","write"),C=ee(l,"overtime","delete");try{[a,i]=await Promise.all([re(),Te()])}catch(m){console.error("Error fetching cache:",m)}async function _(){g=!0,y();try{let m=await N("/api/overtime?limit=10000");if(m.ok){let r=m.data?.data||m.data||[];t=Array.isArray(r)&&r.length>0?r:jt}else t=jt}catch(m){console.warn("Gagal load lembur, menggunakan data lokal:",m),t=jt}finally{g=!1,y()}}async function z(){try{let{downloadExcel:m}=await Promise.resolve().then(()=>(de(),fe)),r=t.map((h,F)=>({No:F+1,Tanggal:h.date||"","Cabang Lembur":h.branch_name||"","Nama Karyawan":h.employee_name||"","Jam Mulai":h.start_time||"","Jam Selesai":h.end_time||"","Istirahat (Jam)":h.break_hours!=null?h.break_hours:0,"Total Jam":h.total_hours!=null?h.total_hours:0,"Alasan / Keterangan":h.reason||""}));m(r,`Data_Lembur_${new Date().toISOString().slice(0,10)}`),H("Data lembur berhasil diekspor ke Excel!")}catch(m){q("Gagal ekspor Excel: "+m.message)}}async function B(m){if(!C)return;let r=t.find(h=>h.id===m);ie({title:"Hapus Catatan Lembur",message:`Apakah Anda yakin ingin menghapus catatan lembur untuk "${r?.employee_name||"karyawan ini"}" pada tanggal ${qt(r?.date)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let h=await N(`/api/overtime/${m}`,{method:"DELETE"});h.ok?(H("Catatan lembur berhasil dihapus"),t=t.filter(F=>F.id!==m),y()):q(h.data?.error||"Gagal menghapus lembur")}catch(h){q("Gagal menghapus catatan lembur: "+h.message)}}})}function P(m=null){if(!T)return;let r=!!m,h=m?.date?m.date.slice(0,10):new Date().toISOString().slice(0,10),F=m?.start_time||"22:00",k=m?.end_time||"08:00",E=m?.break_hours!=null?m.break_hours:1,n=m?.total_hours!=null?m.total_hours:Qa(F,k,E),d=document.createElement("div");d.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="ot-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${h}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penugasan <span style="color:#DC2626;">*</span>
            </label>
            <select id="ot-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${a.map($=>`
                <option value="${$.value}" ${String(m?.branch_id)===String($.value)?"selected":""}>
                  ${$.label}
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
            ${i.map($=>{let p=typeof $=="string"?$:$.label||$.full_name||$.name||"",S=typeof $=="object"&&$!==null?$.value!=null?$.value:$.id!=null?$.id:p:$,u=String(m?.employee_id)===String(S)||m?.employee_name===p;return`<option value="${S}" ${u?"selected":""}>${p}</option>`}).join("")}
          </select>
        </div>

        <!-- Shift & Waktu Lembur -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Mulai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${F}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Selesai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${k}" required />
          </div>
        </div>

        <!-- Durasi & Istirahat -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Istirahat (Jam) <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.5" min="0" max="10" id="ot-modal-break" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${E}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Total Jam Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.1" min="0" id="ot-modal-total" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; background:#F8FAFC; font-weight:700; color:#2563EB;" 
              value="${n}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Alasan / Keterangan Penugasan Lembur <span style="color:#DC2626;">*</span>
          </label>
          <select id="ot-modal-reason" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
            ${Va.map($=>`
              <option value="${$}" ${m?.reason===$?"selected":""}>${$}</option>
            `).join("")}
          </select>
        </div>
      </div>
    `;let f=d.querySelector("#ot-modal-start"),w=d.querySelector("#ot-modal-end"),L=d.querySelector("#ot-modal-break"),I=d.querySelector("#ot-modal-total");function j(){let $=f.value,p=w.value,S=L.value;I.value=Qa($,p,S)}f.addEventListener("input",j),w.addEventListener("input",j),L.addEventListener("input",j),X({title:r?"\u270F\uFE0F Edit Data Lembur Karyawan":"\u2795 Catat Penugasan Lembur Baru",content:d,size:"md",confirmText:r?"Simpan Perubahan":"Catat Lembur",confirmClass:"btn-primary",onConfirm:async($,p)=>{let S=d.querySelector("#ot-modal-date").value,u=d.querySelector("#ot-modal-branch").value,M=d.querySelector("#ot-modal-emp").value,R=d.querySelector("#ot-modal-start").value,K=d.querySelector("#ot-modal-end").value,O=parseFloat(d.querySelector("#ot-modal-break").value)||0,D=parseFloat(d.querySelector("#ot-modal-total").value)||0,Y=d.querySelector("#ot-modal-reason").value;if(!S){q("Tanggal lembur wajib diisi");return}if(!u){q("Cabang penugasan wajib dipilih");return}if(!M){q("Nama karyawan wajib dipilih");return}if(!R||!K){q("Jam mulai dan jam selesai wajib diisi");return}let J=a.find(Q=>String(Q.value)===String(u)),G=i.find(Q=>String(Q.value)===String(M)||String(Q.id)===String(M)),U=G?G.label||G.full_name||G.name||"Karyawan":typeof M=="string"&&isNaN(Number(M))?M:"Karyawan",W={date:S,branch_id:Number(u),branch_name:J?J.label:"Cabang Lembur",employee_id:Number(M),employee_name:G?G.label:"Karyawan",start_time:R,end_time:K,break_hours:O,total_hours:D,reason:Y};try{if(r){let Q=await N(`/api/overtime/${m.id}`,{method:"PUT",body:JSON.stringify(W)});if(Q.ok){H("Data lembur berhasil diperbarui");let V=t.findIndex(te=>te.id===m.id);V!==-1&&(t[V]={...t[V],...W}),p(),y()}else q(Q.data?.error||"Gagal memperbarui data lembur")}else{let Q=await N("/api/overtime",{method:"POST",body:JSON.stringify(W)});if(Q.ok){H("Data lembur berhasil ditambahkan");let V=Q.data?.id||Date.now();t.unshift({id:V,...W}),p(),y()}else q(Q.data?.error||"Gagal menambahkan lembur")}}catch(Q){q("Terjadi kesalahan: "+Q.message)}}})}function c(){return t.filter(m=>{if(x){let r=x.toLowerCase(),h=(m.employee_name||"").toLowerCase().includes(r),F=(m.branch_name||"").toLowerCase().includes(r),k=(m.reason||"").toLowerCase().includes(r);if(!h&&!F&&!k)return!1}return!(b!=="all"&&String(m.branch_id)!==String(b)||s!=="all"&&m.reason!==s)})}function y(){let m=c(),r=t.length,h=t.reduce((f,w)=>f+(parseFloat(w.total_hours)||0),0),F=r>0?(h/r).toFixed(1):0,k=new Date,E=k.getFullYear(),n=k.getMonth(),d=t.filter(f=>{if(!f.date)return!1;let w=new Date(f.date);return!isNaN(w.getTime())&&w.getFullYear()===E&&w.getMonth()===n}).length;e.innerHTML=`
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
              ${A?`
                <button id="lembur-btn-add" class="btn btn-primary lembur-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Lembur</span>
                </button>
              `:""}
              
              <button id="lembur-btn-excel" class="btn btn-ghost lembur-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

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
                <div class="lembur-stat-val">${h} Jam</div>
                <div class="lembur-stat-lbl">Total Jam Lembur</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F4DD}
              </div>
              <div>
                <div class="lembur-stat-val">${r}</div>
                <div class="lembur-stat-lbl">Catatan Penugasan</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4CA}
              </div>
              <div>
                <div class="lembur-stat-val">${F} Jam</div>
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
              <input type="text" id="lembur-search-input" class="lembur-search-input" placeholder="Cari nama karyawan, cabang, atau alasan lembur..." value="${x}" />
            </div>

            <select id="lembur-filter-branch" class="lembur-select">
              <option value="all">Semua Cabang (${a.length})</option>
              ${a.map(f=>`<option value="${f.value}" ${String(b)===String(f.value)?"selected":""}>${f.label}</option>`).join("")}
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
            <div class="lembur-pill ${s==="all"?"active":""}" data-reason="all">
              <span>Semua Lembur</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${r}</span>
            </div>
            ${Va.map(f=>{let w=t.filter(I=>I.reason===f).length,L=Nt(f);return`
                <div class="lembur-pill ${s===f?"active":""}" data-reason="${f}">
                  <span>${L.icon}</span>
                  <span>${f}</span>
                  ${w>0?`<span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${w}</span>`:""}
                </div>
              `}).join("")}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="lembur-list-container">
          ${m.length===0?`
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
                      ${T||C?'<th style="width:90px; text-align:center;">Aksi</th>':""}
                    </tr>
                  </thead>
                  <tbody>
                    ${m.map((f,w)=>{let L=(f.employee_name||"Karyawan").split(" ").map(j=>j[0]).filter(Boolean).slice(0,2).join("").toUpperCase(),I=Nt(f.reason);return`
                        <tr data-id="${f.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${w+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="lembur-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${L}
                              </div>
                              <div style="font-weight:700; color:#0F172A; white-space:nowrap;">
                                ${f.employee_name}
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#334155; font-size:0.85rem; display:inline-flex; align-items:center; gap:4px;">
                              <span>\u{1F4CD}</span>
                              <span>${f.branch_name||"Cabang"}</span>
                            </div>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            \u{1F4C5} ${qt(f.date)}
                          </td>
                          <td style="white-space:nowrap;">
                            <span style="font-weight:700; color:#2563EB;">\u23F0 ${f.start_time||"00:00"} - ${f.end_time||"00:00"}</span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="lembur-hours-badge" style="padding:3px 8px; font-size:0.8rem;">
                              <span>\u23F1\uFE0F</span>
                              <span>${f.total_hours!=null?f.total_hours:0} Jam</span>
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap; color:#64748B; font-weight:600; font-size:0.82rem;">
                            \u2615 ${f.break_hours!=null?f.break_hours:0} Jam
                          </td>
                          <td style="white-space:nowrap;">
                            <span class="lembur-reason-badge" style="background:${I.bg}; color:${I.color}; border:1px solid ${I.border};">
                              <span>${I.icon}</span>
                              <span>${f.reason||"Lembur Operasional"}</span>
                            </span>
                          </td>
                          ${T||C?`
                            <td style="text-align:center; white-space:nowrap;">
                              <div style="display:inline-flex; align-items:center; gap:6px;">
                                ${T?`
                                  <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${f.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                  </button>
                                `:""}
                                ${C?`
                                  <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${f.id}">
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
              ${m.map(f=>{let w=(f.employee_name||"Karyawan").split(" ").map(I=>I[0]).filter(Boolean).slice(0,2).join("").toUpperCase(),L=Nt(f.reason);return`
                  <div class="lembur-card" data-id="${f.id}">
                    <div>
                      <!-- Card Header: Jam & Istirahat -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="lembur-hours-badge">
                          <span>\u23F1\uFE0F</span>
                          <span>${f.total_hours!=null?f.total_hours:0} Jam Lembur</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                          \u2615 Istirahat: ${f.break_hours!=null?f.break_hours:0} Jam
                        </span>
                      </div>

                      <!-- Employee & Branch -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="lembur-avatar-box">
                          ${w}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${f.employee_name}
                          </h4>
                          <div style="font-size:0.8rem; color:#64748B; display:flex; align-items:center; gap:4px;">
                            <span>\u{1F4CD}</span>
                            <span style="font-weight:600; color:#334155;">${f.branch_name||"Cabang"}</span>
                          </div>
                        </div>
                      </div>

                      <!-- Shift Times & Date Box -->
                      <div class="lembur-time-box">
                        <div>
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Tanggal</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#1E293B;">
                            \u{1F5D3}\uFE0F ${qt(f.date)}
                          </div>
                        </div>

                        <div style="text-align:right;">
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Waktu Shift</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#2563EB;">
                            \u23F0 ${f.start_time||"00:00"} - ${f.end_time||"00:00"}
                          </div>
                        </div>
                      </div>

                      <!-- Reason Badge -->
                      <div>
                        <span class="lembur-reason-badge" style="background:${L.bg}; color:${L.color}; border:1px solid ${L.border};">
                          <span>${L.icon}</span>
                          <span>${f.reason||"Lembur Operasional"}</span>
                        </span>
                      </div>
                    </div>

                    <!-- Card Footer: Aksi -->
                    ${T||C?`
                      <div style="display:flex; justify-content:flex-end; align-items:center; gap:6px; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:12px;">
                        ${T?`
                          <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${f.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                        `:""}
                        ${C?`
                          <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${f.id}">
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
    `,v()}function v(){let m=e.querySelector("#lembur-btn-add");m&&(m.onclick=()=>P());let r=e.querySelector("#lembur-btn-excel");r&&(r.onclick=()=>z());let h=e.querySelector("#lembur-btn-refresh");h&&(h.onclick=()=>_());let F=e.querySelector("#lembur-btn-reset-filter");F&&(F.onclick=()=>{x="",b="all",s="all",y()});let k=e.querySelector("#lembur-search-input");k&&(k.oninput=n=>{x=n.target.value,y();let d=e.querySelector("#lembur-search-input");d&&(d.focus(),d.setSelectionRange(d.value.length,d.value.length))});let E=e.querySelector("#lembur-filter-branch");E&&(E.onchange=n=>{b=n.target.value,y()}),e.querySelectorAll(".lembur-pill").forEach(n=>{n.onclick=()=>{s=n.getAttribute("data-reason")||"all",y()}}),e.querySelectorAll(".lembur-view-btn").forEach(n=>{n.onclick=()=>{let d=n.getAttribute("data-view");d&&d!==o&&(o=d,localStorage.setItem("fcms_overtime_view_mode",d),y())}}),e.querySelectorAll(".lembur-edit-btn").forEach(n=>{n.onclick=d=>{d.stopPropagation();let f=Number(n.getAttribute("data-id")),w=t.find(L=>L.id===f);w&&P(w)}}),e.querySelectorAll(".lembur-delete-btn").forEach(n=>{n.onclick=d=>{d.stopPropagation();let f=Number(n.getAttribute("data-id"));B(f)}})}await _()}async function Za(e){let t=Z();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}ue({container:e,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:a=>new Date(a).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(a,i)=>`<strong>${a||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${i.user_role||""}</small>`},{key:"action",label:"Aksi",render:a=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[a]||"badge-neutral"}">${a}</span>`},{key:"module",label:"Modul",render:a=>`<span style="text-transform:capitalize">${(a||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(a,i)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${a}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async a=>{try{let b=(((await N(`/api/audit-logs?search=${a}`)).data||{}).data||[]).find(l=>String(l.id)===String(a));if(!b)return alert("Data tidak ditemukan");let x=l=>{if(!l)return"Tidak ada data";try{return JSON.stringify(JSON.parse(l),null,2)}catch{return l}},g=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${x(b.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${x(b.new_data)}</pre>
           </div>
         </div>
       `,{createModal:o}=await Promise.resolve().then(()=>(ce(),Wt));o({title:`Detail Audit Log #${a}`,content:g,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}ce();de();var Rt={"2026-09":"September 2026","2026-10":"Oktober 2026","2026-11":"November 2026","2026-12":"Desember 2026"},Ve={scheduled:{label:"Terjadwal",badgeClass:"badge-warning",icon:"\u{1F7E1}",bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},in_progress:{label:"Sedang Proses",badgeClass:"badge-info",icon:"\u{1F535}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},completed:{label:"Selesai",badgeClass:"badge-success",icon:"\u{1F7E2}",bg:"#DCFCE7",color:"#166534",border:"#86EFAC"}};async function en(e){e.innerHTML=`
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
        <button id="btn-export-migrations" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; height:38px; padding:0 14px; border-radius:10px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; box-sizing:border-box;">
          \u{1F4E5} Export Excel
        </button>
        <button id="btn-add-migration" class="btn btn-primary" style="font-weight:700; height:38px; padding:0 16px; border-radius:10px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; box-sizing:border-box; background:#2563EB; box-shadow:0 4px 12px rgba(37,99,235,0.3);">
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
  `;let t="board",a=null,i=[],s=Z(),b=s&&["superadmin","admin","manager","spv"].includes(s.role);N("/api/branches?all=1").then(c=>{c.ok&&c.data?.data&&(i=c.data.data)}).catch(()=>{});async function x(){try{let c=await N("/api/os-migrations/summary");if(!c.ok)throw new Error("Gagal memuat data migrasi");a=c.data?.data||c.data,g(a),l()}catch(c){document.getElementById("migration-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${c.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function g(c){let y=document.getElementById("migration-summary-cards");if(!y||!c)return;let v=c.total_branches||0,m=c.vendors?.AEON?.total||0,r=c.vendors?.["Triple S"]?.total||0,h=c.by_status?.completed||0,F=c.by_status?.in_progress||0,k=c.by_status?.scheduled||0,E=v>0?Math.round(h/v*100):0;y.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${v} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.8rem; font-weight:800; color:#1E40AF; line-height:1.2;">${m} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 9 \u2022 Nov: 3 \u2022 Des: 8</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #0D9488; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#0F766E; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor Triple S</div>
        <div style="font-size:1.8rem; font-weight:800; color:#0F766E; line-height:1.2;">${r} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 5 \u2022 Okt: 5 \u2022 Nov: 5 \u2022 Des: 27</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #16A34A; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <span style="font-size:0.8rem; font-weight:600; color:#15803D; text-transform:uppercase; letter-spacing:0.5px;">Progres Realisasi</span>
          <span style="font-size:0.85rem; font-weight:700; color:#15803D;">${E}%</span>
        </div>
        <div style="width:100%; height:8px; background:#E2E8F0; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${E}%; height:100%; background:#16A34A; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.75rem; color:var(--text-2); font-weight:500;">
          <span>\u{1F7E2} Selesai: <strong>${h}</strong></span>
          <span>\u{1F535} Proses: <strong>${F}</strong></span>
          <span>\u{1F7E1} Terjadwal: <strong>${k}</strong></span>
        </div>
      </div>
    `}function o(){if(!a||!a.matrix)return[];let c=(document.getElementById("migration-search")?.value||"").toLowerCase().trim(),y=document.getElementById("filter-vendor")?.value||"",v=document.getElementById("filter-month")?.value||"",m=document.getElementById("filter-status")?.value||"";return a.matrix.filter(r=>!(y&&r.vendor!==y||v&&r.target_month!==v||m&&r.status!==m||c&&!`${r.branch_code} ${r.branch_name} ${r.branch_full_name} ${r.vendor} ${r.notes||""} ${r.pic_vendor||""}`.toLowerCase().includes(c)))}function l(){let c=o(),y=document.getElementById("filter-stats-text");y&&(y.innerHTML=`Menampilkan <strong>${c.length}</strong> dari <strong>${a?.total_branches||0}</strong> cabang`);let v=document.getElementById("migration-view-container");t==="board"?A(v,c):T(v,c)}function A(c,y){if(!y.length){c.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;return}let v=["AEON","Triple S"],m=["2026-09","2026-10","2026-11","2026-12"],r="";v.forEach(h=>{let F=y.filter(w=>w.vendor===h);if(document.getElementById("filter-vendor")?.value&&document.getElementById("filter-vendor").value!==h||F.length===0)return;let k=F.length,E=h==="AEON"?"#1E40AF":"#0F766E",n=h==="AEON"?"#EFF6FF":"#F0FDFA",d=h==="AEON"?"#BFDBFE":"#99F6E4",f=m.filter(w=>F.some(L=>L.target_month===w));r+=`
        <div class="vendor-board-section" style="margin-bottom:32px; background:#fff; border:1px solid var(--border); border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <!-- Vendor Section Header -->
          <div style="padding:14px 20px; background:${n}; border-bottom:1px solid ${d}; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.1rem; font-weight:800; color:${E};">${h==="AEON"?"\u{1F3E2} AEON":"\u{1F6E1}\uFE0F Triple S"}</span>
              <span class="badge" style="background:#fff; color:${E}; border:1px solid ${d}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px;">
                ${k} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
              Format Timeline Berdasarkan Dokumen Migrasi
            </div>
          </div>

          <!-- Months Grid (Hanya menampilkan bulan yang memiliki jadwal) -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px; padding:18px;">
            ${f.map(w=>{let L=F.filter($=>$.target_month===w);return`
                <div class="month-column" style="background:#F8FAFC; border:1px solid var(--border); border-radius:10px; display:flex; flex-direction:column; min-height:220px;">
                  <!-- Month Header -->
                  <div style="padding:12px 14px; background:#F1F5F9; border-bottom:1px solid var(--border); border-radius:9px 9px 0 0; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">
                      \u{1F4C5} ${(Rt[w]||w).split(" ")[0]}
                    </div>
                    <span class="badge" style="background:#E2E8F0; color:var(--text-2); font-weight:700; font-size:0.75rem; padding:2px 7px; border-radius:4px;">
                      ${L.length}
                    </span>
                  </div>

                  <!-- Branch Cards Container -->
                  <div style="padding:10px; display:flex; flex-direction:column; gap:8px; flex:1; max-height:480px; overflow-y:auto;">
                    ${L.map($=>{let p=Ve[$.status]||Ve.scheduled;return`
                        <div class="branch-card" style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 12px; box-shadow:0 1px 2px rgba(0,0,0,0.03); transition:transform 0.15s, box-shadow 0.15s;">
                          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px; margin-bottom:6px;">
                            <div style="font-weight:700; font-size:0.9rem; color:var(--text-1); line-height:1.3;">
                              <span style="color:#2563EB; font-family:monospace; margin-right:4px;">${$.branch_code}.</span>${$.branch_name}
                            </div>
                            ${b?`
                              <div style="display:flex; gap:2px;">
                                <button class="btn-icon-mini edit-card-btn" data-id="${$.id}" title="Edit" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--text-3);">\u270F\uFE0F</button>
                                <button class="btn-icon-mini delete-card-btn" data-id="${$.id}" data-name="${$.branch_name}" title="Hapus" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--danger);">\u{1F5D1}\uFE0F</button>
                              </div>
                            `:""}
                          </div>

                          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px; margin-top:8px;">
                            <!-- Status Selector Dropdown -->
                            <div style="position:relative;">
                              <button class="status-toggle-btn" data-id="${$.id}" data-status="${$.status}" style="background:${p.bg}; color:${p.color}; border:1px solid ${p.border}; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; display:inline-flex; align-items:center; gap:4px; cursor:pointer;" title="Klik untuk ubah status">
                                ${p.icon} ${p.label} \u25BE
                              </button>
                            </div>

                            ${$.employee_count>0?`
                              <span style="font-size:0.75rem; color:var(--text-3); font-weight:600; display:inline-flex; align-items:center; gap:3px;" title="${$.employee_count} personil Facility Care aktif di cabang ini">
                                \u{1F464} ${$.employee_count} FC
                              </span>
                            `:""}
                          </div>

                          ${$.actual_date?`
                            <div style="font-size:0.7rem; color:var(--text-3); margin-top:6px;">
                              Realisasi: <strong>${$.actual_date}</strong>
                            </div>
                          `:""}
                        </div>
                      `}).join("")}
                  </div>
                </div>
              `}).join("")}
          </div>
        </div>
      `}),c.innerHTML=r,C(c)}function T(c,y){if(!y.length){c.innerHTML=`
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
                ${b?'<th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${y.map((v,m)=>{let r=Ve[v.status]||Ve.scheduled;return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3);">${m+1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">${v.branch_full_name}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${v.vendor==="AEON"?"#DBEAFE":"#CCFBF1"}; color:${v.vendor==="AEON"?"#1E40AF":"#0F766E"}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${v.vendor}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:var(--text-2);">
                      ${Rt[v.target_month]||v.target_month}
                    </td>
                    <td style="padding:12px 16px;">
                      <button class="status-toggle-btn" data-id="${v.id}" data-status="${v.status}" style="background:${r.bg}; color:${r.color}; border:1px solid ${r.border}; padding:4px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; cursor:pointer;" title="Klik untuk ubah status">
                        ${r.icon} ${r.label} \u25BE
                      </button>
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-2);">
                      ${v.actual_date||"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; font-weight:600; color:var(--text-2);">
                      ${v.employee_count>0?`\u{1F464} ${v.employee_count}`:"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3); max-width:200px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${v.notes||"-"}
                    </td>
                    ${b?`
                      <td style="padding:12px 16px; text-align:right;">
                        <button class="btn btn-outline btn-sm edit-card-btn" data-id="${v.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-card-btn" data-id="${v.id}" data-name="${v.branch_name}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,C(c)}function C(c){c.querySelectorAll(".status-toggle-btn").forEach(y=>{y.addEventListener("click",v=>{v.stopPropagation();let m=y.dataset.id,r=y.dataset.status;_(y,m,r)})}),c.querySelectorAll(".edit-card-btn").forEach(y=>{y.addEventListener("click",()=>{let v=y.dataset.id,m=a?.matrix?.find(r=>String(r.id)===String(v));m&&z(m)})}),c.querySelectorAll(".delete-card-btn").forEach(y=>{y.addEventListener("click",()=>{let v=y.dataset.id,m=y.dataset.name||"Cabang";ie(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${m}</strong>?`,async()=>{try{if(!(await N(`/api/os-migrations/${v}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus");H(`Jadwal migrasi ${m} berhasil dihapus`),x()}catch(r){q(r.message)}})})})}function _(c,y,v){document.querySelectorAll(".status-popup-menu").forEach(k=>k.remove());let m=c.getBoundingClientRect(),r=document.createElement("div");r.className="status-popup-menu",r.style.cssText=`
      position: fixed;
      top: ${m.bottom+window.scrollY+4}px;
      left: ${m.left+window.scrollX}px;
      background: #fff;
      border: 1px solid var(--border);
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 9999;
      padding: 6px 0;
      min-width: 160px;
    `;let h=[{key:"scheduled",label:"\u{1F7E1} Terjadwal"},{key:"in_progress",label:"\u{1F535} Sedang Proses"},{key:"completed",label:"\u{1F7E2} Selesai"}];r.innerHTML=h.map(k=>`
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
        color: ${k.key===v?"#2563EB":"var(--text-1)"};
        font-weight: ${k.key===v?"700":"500"};
      ">
        ${k.label} ${k.key===v?"\u2713":""}
      </button>
    `).join("");let F=k=>{!r.contains(k.target)&&k.target!==c&&(r.remove(),document.removeEventListener("click",F))};setTimeout(()=>document.addEventListener("click",F),50),r.querySelectorAll(".status-option-item").forEach(k=>{k.addEventListener("click",async()=>{let E=k.dataset.key;if(r.remove(),document.removeEventListener("click",F),E!==v)try{if(!(await N(`/api/os-migrations/${y}/status`,{method:"PATCH",body:JSON.stringify({status:E})})).ok)throw new Error("Gagal memperbarui status");H("Status berhasil diubah"),x()}catch(n){q(n.message)}})}),document.body.appendChild(r)}function z(c=null){let y=!!c,v=y?"Edit Jadwal Migrasi OS":"Tambah Jadwal Migrasi OS",r=`
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${i.map(h=>`
      <option value="${h.id}" ${c&&c.branch_id===h.id?"selected":""}>
        ${h.full_name}
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
    `;X({title:v,content:r,confirmText:y?"Simpan Perubahan":"Tambah Cabang",onConfirm:async(h,F)=>{let k=h.querySelector("#migration-form");if(!k.checkValidity()){k.reportValidity();return}let E=new FormData(k),n={branch_id:parseInt(E.get("branch_id"),10),vendor:E.get("vendor"),target_month:E.get("target_month"),status:E.get("status"),actual_date:E.get("actual_date")||null,pic_vendor:E.get("pic_vendor")||null,notes:E.get("notes")||null};try{let d=y?`/api/os-migrations/${c.id}`:"/api/os-migrations",w=await N(d,{method:y?"PUT":"POST",body:JSON.stringify(n)});if(!w.ok){let L=w.data?.error||w.data?.message||"Gagal menyimpan data";throw new Error(L)}H(y?"Jadwal migrasi berhasil diupdate":"Jadwal migrasi berhasil ditambahkan"),F(),x()}catch(d){q(d.message)}}})}let B=document.getElementById("view-board-btn"),P=document.getElementById("view-table-btn");B.addEventListener("click",()=>{t="board",B.style.background="#fff",B.classList.remove("btn-ghost"),P.style.background="transparent",P.classList.add("btn-ghost"),l()}),P.addEventListener("click",()=>{t="table",P.style.background="#fff",P.classList.remove("btn-ghost"),B.style.background="transparent",B.classList.add("btn-ghost"),l()}),["migration-search","filter-vendor","filter-month","filter-status"].forEach(c=>{let y=document.getElementById(c);y&&(y.addEventListener("input",l),y.addEventListener("change",l))}),document.getElementById("btn-add-migration").addEventListener("click",()=>{z()}),document.getElementById("btn-export-migrations").addEventListener("click",()=>{let c=o();if(!c.length){q("Tidak ada data untuk diexport");return}let y=c.map((v,m)=>({No:m+1,"Kode Cabang":v.branch_code,"Nama Cabang":v.branch_name,"Cabang Lengkap":v.branch_full_name,"Vendor OS":v.vendor,"Target Bulan":Rt[v.target_month]||v.target_month,Status:Ve[v.status]?.label||v.status,"Tanggal Realisasi":v.actual_date||"","Jml Personil FC":v.employee_count||0,"PIC Vendor":v.pic_vendor||"",Catatan:v.notes||""}));ne(y,`Timeline_Migrasi_OS_${new Date().toISOString().slice(0,10)}`),H("File Excel berhasil diunduh")}),x()}ce();de();var tn={Persiapan:{bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE"},"Area Umum":{bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD"},"Ruang Tindakan":{bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},"Ruang Penunjang":{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},Toilet:{bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4"},"Area Luar":{bg:"#FEF3C7",color:"#92400E",border:"#FDE68A"},"Kontrol & Check":{bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA"},Istirahat:{bg:"#E2E8F0",color:"#475569",border:"#CBD5E1"},"Layanan Pasien":{bg:"#DCFCE7",color:"#15803D",border:"#86EFAC"},"Sampah Umum":{bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},"Sampah Medis":{bg:"#FEE2E2",color:"#991B1B",border:"#FCA5A5"},Penutup:{bg:"#E0E7FF",color:"#3730A3",border:"#C7D2FE"}};async function an(e){e.innerHTML=`
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
  `;let t="pagi",a="timeline",i=[],s=Z(),b=s&&["superadmin","admin","manager","spv"].includes(s.role);async function x(){try{let c=await N("/api/dac?all=1");if(!c.ok)throw new Error("Gagal memuat standar DAC");i=c.data?.data||c.data||[];let y=i.filter(h=>h.shift==="pagi").length,v=i.filter(h=>h.shift==="siang").length;document.getElementById("badge-count-pagi").textContent=y,document.getElementById("badge-count-siang").textContent=v,document.getElementById("badge-count-all").textContent=i.length;let m=document.getElementById("dac-category-filter"),r=Array.from(new Set(i.map(h=>h.category))).sort();m.innerHTML='<option value="">Semua Kategori</option>'+r.map(h=>`<option value="${h}">${h}</option>`).join(""),o(),l()}catch(c){document.getElementById("dac-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${c.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function g(){let c=(document.getElementById("dac-search")?.value||"").toLowerCase().trim(),y=document.getElementById("dac-category-filter")?.value||"",v=document.getElementById("dac-period-filter")?.value||"";return i.filter(m=>{if(t!=="all"&&m.shift!==t||y&&m.category!==y)return!1;if(v==="morning"){let r=parseInt(m.start_time.split(":")[0],10);if(m.shift==="pagi"&&r>=12||m.shift==="siang"&&r>=17)return!1}else if(v==="afternoon"){let r=parseInt(m.start_time.split(":")[0],10);if(m.shift==="pagi"&&r<12||m.shift==="siang"&&r<17)return!1}return!(c&&!`${m.start_time} ${m.end_time} ${m.category} ${m.location} ${m.activity}`.toLowerCase().includes(c))})}function o(){let c=document.getElementById("dac-summary-cards");if(!c)return;let y=i.filter(F=>F.shift==="pagi"),v=i.filter(F=>F.shift==="siang"),m=t==="pagi"?y:t==="siang"?v:i,r=t==="pagi"?"08:00 \u2013 17:00":t==="siang"?"13:00 \u2013 22:00":"08:00 \u2013 22:00",h=t==="pagi"?"Shift Pagi":t==="siang"?"Shift Siang":"Semua Shift";c.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Aktivitas (${h})</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${m.length} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Aktivitas</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Urutan Terjadwal dari Awal s/d Penutup</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Jam Kerja Standar</div>
        <div style="font-size:1.5rem; font-weight:800; color:#1E40AF; line-height:1.2;">${r}</div>
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
    `}function l(){let c=g(),y=document.getElementById("dac-stats-text");y&&(y.innerHTML=`Menampilkan <strong>${c.length}</strong> aktivitas`);let v=document.getElementById("dac-view-container");a==="timeline"?A(v,c):T(v,c)}function A(c,y){if(!y.length){c.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;return}c.innerHTML=`
      <div class="timeline-container" style="display:flex; flex-direction:column; gap:12px; max-width:960px; margin:0 auto;">
        ${y.map(v=>{let m=tn[v.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},r=v.shift==="pagi"?"\u{1F305} Pagi":"\u{1F307} Siang",h=v.shift==="pagi"?"#FEF3C7":"#E0E7FF",F=v.shift==="pagi"?"#92400E":"#3730A3";return`
            <div class="dac-timeline-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:16px 20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; gap:16px; align-items:flex-start; position:relative;">
              <!-- Time Column -->
              <div style="min-width:130px; text-align:center; background:#F8FAFC; border:1px solid var(--border); border-radius:8px; padding:10px 8px;">
                <div style="font-size:1.1rem; font-weight:800; color:var(--text-1); font-family:monospace;">
                  ${v.start_time} - ${v.end_time}
                </div>
                <div style="font-size:0.75rem; color:var(--text-3); font-weight:600; margin-top:4px;">
                  Item #${v.item_no}
                </div>
                ${t==="all"?`
                  <div style="margin-top:6px;">
                    <span class="badge" style="background:${h}; color:${F}; font-size:0.7rem; font-weight:700; padding:2px 6px; border-radius:4px;">
                      ${r}
                    </span>
                  </div>
                `:""}
              </div>

              <!-- Content Column -->
              <div style="flex:1;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                  <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span class="badge" style="background:${m.bg}; color:${m.color}; border:1px solid ${m.border}; font-weight:700; font-size:0.75rem; padding:3px 10px; border-radius:6px;">
                      ${v.category}
                    </span>
                    <span style="font-weight:700; font-size:0.95rem; color:var(--text-1);">
                      \u{1F4CD} ${v.location}
                    </span>
                  </div>
                  ${b?`
                    <div style="display:flex; gap:4px;">
                      <button class="btn btn-ghost btn-sm edit-dac-btn" data-id="${v.id}" title="Edit" style="padding:4px 8px; font-size:0.8rem; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-ghost btn-sm delete-dac-btn" data-id="${v.id}" data-name="${v.location}" title="Hapus" style="padding:4px 8px; font-size:0.8rem; color:var(--danger); cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </div>
                  `:""}
                </div>

                <div style="font-size:0.9rem; color:var(--text-2); line-height:1.5; background:#FAFBFD; padding:10px 14px; border-radius:8px; border-left:3px solid ${m.color}; margin-top:8px;">
                  ${v.activity}
                </div>
              </div>
            </div>
          `}).join("")}
      </div>
    `,C(c)}function T(c,y){if(!y.length){c.innerHTML=`
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
                ${b?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right; width:100px;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${y.map(v=>{let m=tn[v.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"};return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 14px; font-size:0.85rem; color:var(--text-3); font-weight:600;">${v.item_no}</td>
                    ${t==="all"?`
                      <td style="padding:12px 14px;">
                        <span class="badge" style="background:${v.shift==="pagi"?"#FEF3C7":"#E0E7FF"}; color:${v.shift==="pagi"?"#92400E":"#3730A3"}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:4px;">
                          ${v.shift==="pagi"?"Pagi":"Siang"}
                        </span>
                      </td>
                    `:""}
                    <td style="padding:12px 14px; font-family:monospace; font-weight:700; font-size:0.85rem; color:var(--text-1); white-space:nowrap;">
                      ${v.start_time} - ${v.end_time}
                    </td>
                    <td style="padding:12px 14px;">
                      <span class="badge" style="background:${m.bg}; color:${m.color}; border:1px solid ${m.border}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px; white-space:nowrap;">
                        ${v.category}
                      </span>
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; font-weight:600; color:var(--text-1);">
                      ${v.location}
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; color:var(--text-2); line-height:1.4;">
                      ${v.activity}
                    </td>
                    ${b?`
                      <td style="padding:12px 14px; text-align:right; white-space:nowrap;">
                        <button class="btn btn-outline btn-sm edit-dac-btn" data-id="${v.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-dac-btn" data-id="${v.id}" data-name="${v.location}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,C(c)}function C(c){c.querySelectorAll(".edit-dac-btn").forEach(y=>{y.addEventListener("click",()=>{let v=y.dataset.id,m=i.find(r=>String(r.id)===String(v));m&&_(m)})}),c.querySelectorAll(".delete-dac-btn").forEach(y=>{y.addEventListener("click",()=>{let v=y.dataset.id,m=y.dataset.name||"item";ie(`Apakah Anda yakin ingin menghapus aktivitas pada lokasi <strong>${m}</strong>?`,async()=>{try{if(!(await N(`/api/dac/${v}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus item");H("Aktivitas DAC berhasil dihapus"),x()}catch(r){q(r.message)}})})})}function _(c=null){let y=!!c,v=y?"Edit Aktivitas DAC":"Tambah Aktivitas DAC",m=["Persiapan","Area Umum","Ruang Tindakan","Ruang Penunjang","Toilet","Area Luar","Kontrol & Check","Istirahat","Layanan Pasien","Sampah Umum","Sampah Medis","Penutup"],r=`
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
            ${m.map(h=>`<option value="${h}">`).join("")}
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
    `;X({title:v,content:r,confirmText:y?"Simpan Perubahan":"Tambah Aktivitas",onConfirm:async(h,F)=>{let k=h.querySelector("#dac-form");if(!k.checkValidity()){k.reportValidity();return}let E=new FormData(k),n={shift:E.get("shift"),item_no:E.get("item_no")?parseInt(E.get("item_no"),10):void 0,start_time:E.get("start_time"),end_time:E.get("end_time"),category:E.get("category"),location:E.get("location"),activity:E.get("activity")};try{let d=y?`/api/dac/${c.id}`:"/api/dac",w=await N(d,{method:y?"PUT":"POST",body:JSON.stringify(n)});if(!w.ok){let L=w.data?.error||w.data?.message||"Gagal menyimpan aktivitas";throw new Error(L)}H(y?"Aktivitas berhasil diperbarui":"Aktivitas berhasil ditambahkan"),F(),x()}catch(d){q(d.message)}}})}function z(){let c=g();if(!c.length){q("Tidak ada data untuk dicetak");return}let y=t==="pagi"?"SHIFT PAGI (08:00 - 17:00)":t==="siang"?"SHIFT SIANG (13:00 - 22:00)":"SELURUH SHIFT (PAGI & SIANG)",v=window.open("","_blank");if(!v){q("Popup diblokir browser, mohon izinkan popup untuk mencetak");return}v.document.write(`
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
            ${c.map(m=>`
              <tr>
                <td style="text-align:center; font-weight:bold;">${m.item_no}</td>
                ${t==="all"?`<td>${m.shift==="pagi"?"Pagi":"Siang"}</td>`:""}
                <td class="jam">${m.start_time} - ${m.end_time}</td>
                <td><strong>${m.category}</strong></td>
                <td><strong>${m.location}</strong></td>
                <td>${m.activity}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
        <div style="margin-top:20px; text-align:right; font-size:9pt; color:#64748b;">
          Dicetak pada: ${new Date().toLocaleString("id-ID")} | Standar Resmi Operasional Facility Care Audy Dental
        </div>
      </body>
      </html>
    `),v.document.close(),v.focus(),setTimeout(()=>{v.print()},400)}document.querySelectorAll(".shift-tab-btn").forEach(c=>{c.addEventListener("click",()=>{document.querySelectorAll(".shift-tab-btn").forEach(y=>{y.style.background="#F1F5F9",y.style.color="var(--text-2)"}),c.style.background="#2563EB",c.style.color="#fff",t=c.dataset.shift,o(),l()})});let B=document.getElementById("view-timeline-btn"),P=document.getElementById("view-table-btn");B.addEventListener("click",()=>{a="timeline",B.style.background="#fff",B.classList.remove("btn-ghost"),P.style.background="transparent",P.classList.add("btn-ghost"),l()}),P.addEventListener("click",()=>{a="table",P.style.background="#fff",P.classList.remove("btn-ghost"),B.style.background="transparent",B.classList.add("btn-ghost"),l()}),["dac-search","dac-category-filter","dac-period-filter"].forEach(c=>{let y=document.getElementById(c);y&&(y.addEventListener("input",l),y.addEventListener("change",l))}),document.getElementById("btn-add-dac").addEventListener("click",()=>{_()}),document.getElementById("btn-print-dac").addEventListener("click",()=>{z()}),document.getElementById("btn-export-dac").addEventListener("click",()=>{let c=g();if(!c.length){q("Tidak ada data untuk diexport");return}let y=c.map(v=>({No:v.item_no,Shift:v.shift==="pagi"?"Shift Pagi":"Shift Siang","Jam Mulai":v.start_time,"Jam Selesai":v.end_time,Kategori:v.category,"Lokasi / Area":v.location,"Rincian Aktivitas":v.activity}));ne(y,`DAC_Daily_Activity_Cleaning_${t}_${new Date().toISOString().slice(0,10)}`),H("File Excel berhasil diunduh")}),x()}window.parseFlexibleDate=e=>{if(!e||e==="-")return"";if(e=String(e).trim(),/^\d{5}$/.test(e)){let a=Math.floor(Number(e)-25569);return new Date(a*86400*1e3).toISOString().split("T")[0]}if(e.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let a=e.split(/[\/\-]/);return`${a[2]}-${a[1]}-${a[0]}`}let t=e.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(t){let a=t[1],i=parseInt(t[2],10),s=parseInt(t[3],10);if(i>12&&s<=12)return`${a}-${t[3]}-${t[2]}`}return e.split("T")[0]};window.formatDate=e=>{let t=window.parseFlexibleDate(e);if(!t)return"";let a=t.split("-");if(a.length===3&&a[0].length===4){let i=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],s=parseInt(a[2],10),b=i[parseInt(a[1],10)-1];return`${s} ${b} ${a[0]}`}return t};function pe(e){return async t=>{if(!Xe()){Ae("/login");return}return e(t)}}var Qe=null;function Qn(){Qe&&clearInterval(Qe);let e=()=>{let t=new Date,a=t.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),i=t.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),s=document.getElementById("header-clock-time"),b=document.getElementById("header-clock-date");s&&(s.textContent=a),b&&(b.textContent=i)};e(),Qe=setInterval(e,1e3)}async function Xn(){try{let e=await N("/api/dashboard/kpi");if(!e.ok)return;let t=e.data?.data||e.data||{},a=(i,s)=>{let b=document.getElementById(i);b&&(b.textContent=s>0?s:"",b.style.display=s>0?"inline-flex":"none")};a("badge-issues",t.issues?.current||0),a("badge-contracts",t.expiring30?.current||0),a("badge-oo1",t.one_on_one?.current||0),a("badge-schedule",t.schedule?.current||0),a("badge-supply",t.supply?.current||0)}catch{}}var Re=[];async function Zn(){try{let e=await N("/api/dashboard/notifications");if(!e.ok)return;Re=e.data?.data||e.data||[];let t=document.getElementById("notif-dot");t&&(t.style.display=Re.length>0?"block":"none",t.textContent=Re.length)}catch{}}function ei(){if(!Re.length){X({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(t,a)=>a()});return}let e=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${Re.map(t=>`
        <div class="notif-item notif-severity-${t.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${t.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${t.date}</span>
            <span class="badge badge-${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"info"}">${t.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;X({title:`Notifikasi (${Re.length})`,content:e,confirmText:"Tutup",onConfirm:(t,a)=>a()})}function nn(){let e=Z(),t=(e?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
                <span class="topbar-greeting-time">${(()=>{let l=new Date().getHours();return l>=4&&l<11?"Selamat Pagi":l>=11&&l<15?"Selamat Siang":l>=15&&l<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${e?.full_name||e?.username||"Guest"}</span> \u{1F44B}
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
  `;let a=document.getElementById("sidebar"),i=document.getElementById("sidebar-overlay"),s=document.getElementById("topbar-menu-btn"),b=document.getElementById("sidebar-close"),x=()=>{a.classList.add("open"),i.classList.add("show")},g=()=>{a.classList.remove("open"),i.classList.remove("show")};s?.addEventListener("click",x),b?.addEventListener("click",g),i?.addEventListener("click",g),document.querySelectorAll(".nav-item").forEach(l=>l.addEventListener("click",g)),e&&e.role==="input_lembur"&&(document.querySelectorAll(".nav-item").forEach(l=>{l.dataset.route!=="/overtime"&&(l.style.display="none")}),document.querySelectorAll(".nav-section").forEach(l=>{Array.from(l.querySelectorAll(".nav-item")).some(T=>T.style.display!=="none")||(l.style.display="none")}));function o(){let l=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(C=>{let _=C.dataset.route;C.classList.toggle("active",l===_||_!=="/dashboard"&&l.startsWith(_))});let A=document.getElementById("topbar-title"),T=document.querySelector(".nav-item.active .nav-label");A&&T&&(A.textContent=T.textContent)}window.addEventListener("hashchange",o),o(),Qn(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await N("/api/auth/logout",{method:"POST"}),Ze(),Qe&&clearInterval(Qe),Ae("/login")}),Xn(),Zn(),document.getElementById("btn-notif")?.addEventListener("click",l=>{l.preventDefault(),ei()})}async function ti(){oe("/login",({main:t})=>na(t)),oe("/dashboard",pe(({main:t})=>Xt(t))),oe("/calendar",pe(({main:t})=>ja(t))),oe("/employees",pe(({main:t,params:a})=>oa(t,a))),oe("/contracts",pe(({main:t,params:a})=>la(t,a))),oe("/sp",pe(({main:t})=>Ja(t))),oe("/mutasi",pe(({main:t})=>Wa(t))),oe("/overtime",pe(({main:t})=>Xa(t))),oe("/timeline",pe(({main:t,params:a})=>da(t,a))),oe("/os-migrations",pe(({main:t})=>en(t))),oe("/issues",pe(({main:t,params:a})=>ua(t,a))),oe("/one-on-one",pe(({main:t,params:a})=>ga(t,a))),oe("/training",pe(({main:t})=>ha(t))),oe("/relievers",pe(({main:t,params:a})=>xa(t,a))),oe("/reports/inspection",pe(({main:t})=>ya(t))),oe("/reports/cleaning",pe(({main:t})=>va(t))),oe("/reports/fogging",pe(({main:t})=>ka(t))),oe("/reports/basecamp",pe(({main:t})=>wa(t))),oe("/reports/supply",pe(({main:t})=>za(t,"supply"))),oe("/sop",()=>{window.location.hash="#/sop-fc"}),oe("/sop-fc",pe(({main:t})=>Ca(t))),oe("/sop-security",pe(({main:t})=>Ta(t))),oe("/checklist",pe(({main:t})=>La(t))),oe("/hygiene-standards",pe(({main:t})=>Ia(t))),oe("/dac",pe(({main:t})=>an(t))),oe("/users",pe(({main:t})=>Ma(t))),oe("/branches",pe(({main:t})=>Ka(t))),oe("/profile",pe(({main:t})=>Na(t))),oe("/settings/import",pe(({main:t})=>Ua(t))),oe("/audit-logs",pe(({main:t})=>Za(t)));let e=Xe();if(!e&&window.location.hash!=="#/login"&&Ae("/login"),e){let t=await N("/api/auth/me");if(t.ok){et(t.data.data),nn();let a=Z(),i=window.location.hash;a&&a.role==="input_lembur"&&(i==="#/dashboard"||i===""||i==="#/")&&Ae("/overtime")}else Ze(),Ae("/login")}window.addEventListener("fm:login",()=>{nn();let t=Z();t&&t.role==="input_lembur"?Ae("/overtime"):Ae("/dashboard")}),Yt()}ti();
