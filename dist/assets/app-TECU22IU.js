var tn=Object.defineProperty;var jt=(e,t)=>()=>(e&&(t=e(e=0)),t);var Kt=(e,t)=>{for(var i in t)tn(e,i,{get:t[i],enumerable:!0})};var Rt={};Kt(Rt,{confirmDialog:()=>X,createModal:()=>G});function G({title:e,content:t,onConfirm:i,onCancel:n,confirmText:c="Simpan",cancelText:x="Batal",size:m="md",confirmClass:u="btn-primary"}){let o={sm:"400px",md:"560px",lg:"720px",xl:"900px"},b=document.createElement("div");b.className="modal-overlay",b.innerHTML=`
    <div class="modal" style="max-width:${o[m]||o.md}">
      <div class="modal-header">
        <h3 class="modal-title">${e}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof t=="string"?t:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${x}</button>
        ${i?`<button class="btn ${u} modal-confirm">${c}</button>`:""}
      </div>
    </div>
  `,t instanceof HTMLElement&&b.querySelector(".modal-body").appendChild(t);let B=()=>{b.classList.remove("show"),setTimeout(()=>b.remove(),250)};return b.querySelector(".modal-close").addEventListener("click",()=>{n&&n(),B()}),b.querySelector(".modal-cancel").addEventListener("click",()=>{n&&n(),B()}),i&&b.querySelector(".modal-confirm").addEventListener("click",()=>i(b,B)),b.addEventListener("click",D=>{D.target===b&&(n&&n(),B())}),document.body.appendChild(b),requestAnimationFrame(()=>b.classList.add("show")),{overlay:b,close:B}}function X(e,t,i="Konfirmasi"){if(typeof e=="object"&&e!==null){let n=e;return G({title:n.title||"Konfirmasi",content:typeof n.message=="string"?`<p>${n.message}</p>`:n.message,onConfirm:async(c,x)=>{n.onConfirm&&await n.onConfirm(),x()},onCancel:n.onCancel,confirmText:n.confirmText||"Ya, Lanjutkan",cancelText:n.cancelText||"Batal",confirmClass:n.confirmClass||"btn-danger"})}return G({title:i,content:typeof e=="string"?`<p>${e}</p>`:e,onConfirm:async(n,c)=>{t&&await t(),c()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var le=jt(()=>{});var he={};Kt(he,{downloadExcel:()=>Q,parseExcel:()=>qe,renderExcelButtons:()=>on});function qe(e){return new Promise((t,i)=>{let n=new FileReader;n.onload=c=>{try{let x=new Uint8Array(c.target.result),m=XLSX.read(x,{type:"array"}),u=m.SheetNames[0],o=m.Sheets[u];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${e.name}`),console.log(`File Size: ${(e.size/1024).toFixed(2)} KB`),console.log(`File Type: ${e.type||"unknown"}`),console.log(`Sheets Found: ${m.SheetNames.join(", ")}`),console.log(`Sheet Used: ${u}`);let b=XLSX.utils.decode_range(o["!ref"]||"A1:A1"),B=b.e.r-b.s.r+1,D=b.e.c-b.s.c+1;console.log(`Total Rows (including empty): ${B}`),console.log(`Total Columns: ${D}`);let S=[];for(let $=b.s.c;$<=b.e.c;++$){let E=o[XLSX.utils.encode_cell({c:$,r:b.s.r})];E&&E.v&&S.push(E.v)}console.log(`Headers Found: ${S.join(", ")}`),console.log("---------------------------");let F=XLSX.utils.sheet_to_json(o,{defval:""});Object.defineProperty(F,"__worksheet",{value:o,enumerable:!1}),Object.defineProperty(F,"__headers",{value:S,enumerable:!1}),t(F)}catch(x){i(x)}},n.onerror=c=>i(c),n.readAsArrayBuffer(e)})}function Q(e,t){try{let i=XLSX.utils.json_to_sheet(e),n=XLSX.utils.book_new();XLSX.utils.book_append_sheet(n,i,"Data"),XLSX.writeFile(n,`${t}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function on(e){return`
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
  `}var re=jt(()=>{});var an="https://fm-operations-api.facilitycare-audydental.workers.dev",nn=an,Be=1e4;function Ye(){return localStorage.getItem("fm_token")}function Nt(e){localStorage.setItem("fm_token",e)}function Ve(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ee(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function We(e){localStorage.setItem("fm_user",JSON.stringify(e))}async function K(e,t={}){let i=Ye(),n={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...t.headers||{}};try{let c=`cb=${Date.now()}`,x=e.includes("?")?"&":"?",m=`${nn}${e}${x}${c}`,u=await fetch(m,{...t,headers:n}),o;try{let b=await u.text();try{o=JSON.parse(b)}catch{o={error:`Server Error (${u.status}): ${b.substring(0,80)}...`}}}catch{o={error:"Gagal membaca respon dari server"}}return u.status===401&&(Ve(),window.location.hash="#/login"),{ok:u.ok,status:u.status,data:o}}catch(c){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${c.message})`}}}}var pt={},Qe=null;function ne(e,t){pt[e]=t}function $e(e){window.location.hash=e}function qt(){async function e(){let t=window.location.hash.replace("#","")||"/dashboard",[i,...n]=t.split("?"),c=pt[i];if(!c){for(let[m,u]of Object.entries(pt))if(m.endsWith("/*")&&i.startsWith(m.slice(0,-2))){c=u;break}}Qe&&(Qe(),Qe=null);let x=document.getElementById("main-content");if(x&&(x.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),c){let m=new URLSearchParams(n.join("?")),u=i.split("/").filter(Boolean),o=await c({path:i,params:m,segments:u,main:x});o&&(Qe=o)}else{let m=x||document.getElementById("app");m&&(m.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",e),e()}var Ke;function rn(){return Ke||(Ke=document.createElement("div"),Ke.id="toast-container",document.body.appendChild(Ke)),Ke}function ct(e,t="info",i=3500){let n=rn(),c=document.createElement("div");c.className=`toast toast-${t}`;let x={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};c.innerHTML=`<span class="toast-icon">${x[t]||"\u2139"}</span><span class="toast-msg">${e}</span>`,n.appendChild(c),requestAnimationFrame(()=>c.classList.add("show")),setTimeout(()=>{c.classList.remove("show"),setTimeout(()=>c.remove(),350)},i)}var R=e=>ct(e,"success"),N=e=>ct(e,"error");var ut=e=>ct(e,"info");le();le();function Xe({columns:e,data:t,onEdit:i,onDelete:n,onView:c,actions:x=[],emptyText:m="Tidak ada data",bulkSelect:u=null}){let o=document.createElement("div");if(o.className="table-wrapper",!t||t.length===0)return o.innerHTML=`<div class="empty-state"><p>${m}</p></div>`,o;let b=document.createElement("table");b.className="data-table";let B=document.createElement("thead"),D=document.createElement("tr");if(u){let F=document.createElement("th");F.style.width="40px",F.style.textAlign="center";let $=document.createElement("input");$.type="checkbox",$.id="select-all-checkbox",$.title="Pilih semua",$.addEventListener("change",()=>{t.forEach(E=>{$.checked?u.selectedIds.add(E.id):u.selectedIds.delete(E.id)}),o.querySelectorAll(".row-checkbox").forEach(E=>E.checked=$.checked),u.onToggle()}),F.appendChild($),D.appendChild(F)}if(e.forEach(F=>{let $=document.createElement("th");$.textContent=F.label,F.width&&($.style.width=F.width),D.appendChild($)}),i||n||c||x.length>0){let F=document.createElement("th");F.textContent="Aksi",F.style.width="120px",D.appendChild(F)}B.appendChild(D),b.appendChild(B);let S=document.createElement("tbody");return t.forEach(F=>{let $=document.createElement("tr");if(u){let E=document.createElement("td");E.style.textAlign="center",E.style.width="40px";let k=document.createElement("input");k.type="checkbox",k.className="row-checkbox",k.checked=u.selectedIds.has(F.id),k.addEventListener("change",()=>{if(k.checked)u.selectedIds.add(F.id);else{u.selectedIds.delete(F.id);let h=document.getElementById("select-all-checkbox");h&&(h.checked=!1)}u.onToggle()}),E.appendChild(k),$.appendChild(E)}if(e.forEach(E=>{let k=document.createElement("td");if(E.render){let h=E.render(F[E.key],F);h instanceof HTMLElement?k.appendChild(h):k.innerHTML=h||""}else k.textContent=F[E.key]!==null&&F[E.key]!==void 0&&F[E.key]!==""?F[E.key]:"";E.nowrap&&(k.style.whiteSpace="nowrap"),$.appendChild(k)}),i||n||c||x.length>0){let E=document.createElement("td");E.className="actions-cell";let k=document.createElement("div");if(k.className="btn-group",c){let h=document.createElement("button");h.className="btn btn-xs btn-ghost",h.innerHTML="\u{1F441}",h.title="Lihat",h.addEventListener("click",()=>c(F)),k.appendChild(h)}if(i){let h=document.createElement("button");h.className="btn btn-xs btn-secondary",h.innerHTML="\u270F\uFE0F",h.title="Edit",h.addEventListener("click",()=>i(F)),k.appendChild(h)}x.forEach(h=>{let a=document.createElement("button");a.className=`btn btn-xs ${h.class||"btn-ghost"}`,a.innerHTML=h.icon||h.label,a.title=h.label,a.addEventListener("click",()=>h.handler(F)),k.appendChild(a)}),E.appendChild(k),$.appendChild(E)}S.appendChild($)}),b.appendChild(S),o.appendChild(b),o}function Ze({page:e,pages:t,total:i,limit:n,onPage:c}){if(t<=1)return null;let x=document.createElement("div");x.className="pagination";let m=document.createElement("span");m.className="pagination-info",m.textContent=`Total: ${i} data`,x.appendChild(m);let u=document.createElement("div");u.className="pagination-btns";let o=(D,S,F=!1,$=!1)=>{let E=document.createElement("button");E.className=`btn btn-sm ${$?"btn-primary":"btn-ghost"} pagination-btn`,E.textContent=D,E.disabled=F,E.addEventListener("click",()=>c(S)),u.appendChild(E)};o("\xAB",1,e===1),o("\u2039",e-1,e===1);let b=Math.max(1,e-2),B=Math.min(t,e+2);for(let D=b;D<=B;D++)o(D,D,!1,D===e);return o("\u203A",e+1,e===t),o("\xBB",t,e===t),x.appendChild(u),x}le();function Ne(e){return e.map(t=>{if(t.type==="hidden")return`<input type="hidden" name="${t.name}" value="${t.value||""}">`;if(t.type==="html")return t.html||"";if(t.type==="row")return`<div class="form-row">${Ne(t.fields)}</div>`;let i=t.required?"required":"",n=t.label?`<label class="form-label">${t.label}${t.required?' <span class="required">*</span>':""}</label>`:"",c="";switch(t.type){case"textarea":c=`<textarea name="${t.name}" class="form-control" placeholder="${t.placeholder||""}" ${i} rows="${t.rows||3}">${t.value||""}</textarea>`;break;case"select":let m=(t.options||[]).map(F=>{let $=typeof F=="object"?F.value:F,E=typeof F=="object"?F.label:F,k=t.value==$?"selected":"";return`<option value="${$}" ${k}>${E}</option>`}).join("");c=`<select name="${t.name}" class="form-control" ${i}><option value="">-- Pilih ${t.label||""} --</option>${m}</select>`;break;case"combobox":let u=`dl-${t.name}-${Math.random().toString(36).substring(7)}`,o=(t.options||[]).map(F=>{let $=typeof F=="object"?F.value:F,E=typeof F=="object"?F.label||F.value||"":F||"";return(E==="undefined"||E==="[object Object]"||E==="null")&&(E=""),E?`<option value="${E}"></option>`:""}).join(""),b=t.value||"";if(t.value){let F=(t.options||[]).find($=>(typeof $=="object"?$.value:$)==t.value);if(F){let $=typeof F=="object"?F.label||F.value||"":F||"";$&&$!=="undefined"&&$!=="[object Object]"&&$!=="null"&&(b=$)}}c=`
          <input type="text" name="${t.name}" list="${u}" class="form-control" value="${b}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${u}">${o}</datalist>
        `;break;case"checkbox":c=`<label class="checkbox-label"><input type="checkbox" name="${t.name}" value="1" ${t.value?"checked":""}> ${t.checkLabel||t.label}</label>`;break;case"date":let B=window.parseFlexibleDate&&t.value?window.parseFlexibleDate(t.value):t.value||"";c=`<input type="date" name="${t.name}" class="form-control" value="${B}" ${i}>`;break;case"month":let D=t.value?String(t.value).slice(0,7):"";c=`<input type="month" name="${t.name}" class="form-control" value="${D}" ${i}>`;break;case"number":let S=t.readonly?"readonly":"";c=`<input type="number" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" min="${t.min||""}" max="${t.max||""}" step="${t.step||"1"}" ${i} ${S}>`;break;case"email":c=`<input type="email" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${i}>`;break;case"url":c=`<input type="url" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||"https://..."}" ${i}>`;break;default:c=`<input type="${t.type||"text"}" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${i} autocomplete="off">`}let x=t.hint?`<div class="form-hint">${t.hint}</div>`:"";return`<div class="form-group ${t.class||""}">${n}${c}${x}</div>`}).join("")}function et(e){let t={},i=new FormData(e);for(let[n,c]of i.entries())t[n]=c===""?null:c;return e.querySelectorAll("input[type=checkbox]").forEach(n=>{n.checked||(t[n.name]=null)}),t}function tt(e,t){t&&Object.entries(t).forEach(([i,n])=>{let c=e.querySelector(`[name="${i}"]`);c&&(c.hasAttribute("list")||(c.type==="checkbox"?c.checked=!!n:c.type==="date"&&n&&window.parseFlexibleDate?c.value=window.parseFlexibleDate(n):c.type==="month"&&n?c.value=String(n).slice(0,7):c.value=n??""))})}re();var Se={},at={on(e,t){Se[e]||(Se[e]=new Set),Se[e].add(t)},off(e,t){Se[e]&&Se[e].delete(t)},emit(e,t){Se[e]&&Se[e].forEach(i=>{try{i(t)}catch(n){console.warn("[calendarBus] Handler error:",n)}})},clear(){Object.keys(Se).forEach(e=>delete Se[e])}},sn=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function be(e){if(!e){at.emit("data:changed",{module:"unknown"});return}let t=String(e).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");at.emit("data:changed",{module:t,relevant:sn.has(t)})}function de({container:e,title:t,icon:i,apiPath:n,columns:c,formFields:x,filterFields:m,defaultFilters:u={},itemLabel:o="Data",canCreate:b=!0,canEdit:B=!0,canDelete:D=!0,onBeforeSubmit:S,onAfterLoad:F,onDataLoaded:$,extraActions:E=[],initialSearch:k="",exportOptions:h=null,bulkDelete:a=!1,paginationMode:p="server"}){let d=ee();d&&d.role==="viewer"&&(b=!1,B=!1,D=!1,a=!1,h=null),d&&d.role==="editor_khusus"&&n!=="/api/issues"&&n!=="/api/overtime"&&(b=!1,B=!1,D=!1,a=!1,h=null),d&&d.role==="input_lembur"&&(n==="/api/overtime"?(D=!1,a=!1,h=null):(b=!1,B=!1,D=!1,a=!1,h=null));let g=1,s={...u};k&&(s.search=k);let r=new Set;e.innerHTML=`
    ${a?`
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
        ${b?`<button class="btn btn-primary" id="btn-create">+ Tambah ${o}</button>`:""}
        ${h?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${h.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${h.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${h.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${h.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
    

    ${m&&m.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${m.filter(y=>y.type==="search").map(y=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${y.placeholder||"Cari..."}" id="filter-search" value="${s.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${m.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let _=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select" name="${y.name}" id="filter-${y.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${_}</option>${(y.options||[]).map(T=>`<option value="${typeof T=="object"?T.value:T}" ${s[y.name]===(typeof T=="object"?T.value:T)?"selected":""}>${typeof T=="object"?T.label:T}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${m.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let _=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select filter-select-sheet" name="${y.name}-sheet" id="filter-sheet-${y.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${_}</option>${(y.options||[]).map(T=>`<option value="${typeof T=="object"?T.value:T}" ${s[y.name]===(typeof T=="object"?T.value:T)?"selected":""}>${typeof T=="object"?T.label:T}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function l(){let y=document.getElementById("bulk-toolbar");if(!y)return;let _=document.getElementById("bulk-count"),T=document.getElementById("btn-bulk-delete"),j=document.getElementById("btn-bulk-cancel");_.textContent=`${r.size} item dipilih`,r.size>0?(y.style.display="flex",T.disabled=!1,j.disabled=!1):(y.style.display="none",T.disabled=!0,j.disabled=!0);let O=document.getElementById("select-all-checkbox");if(O){let M=document.querySelectorAll(".row-checkbox");if(M.length>0){let q=[...M].every(Y=>Y.checked),H=[...M].some(Y=>Y.checked);O.checked=q,O.indeterminate=H&&!q}else O.checked=!1,O.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{r.clear(),document.querySelectorAll(".row-checkbox").forEach(_=>_.checked=!1);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1),l()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(r.size===0)return;let y=[...r],_=document.createElement("div");_.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",_.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${y.length} ${o}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${y.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(_),_.querySelector("#bulk-cancel-btn").addEventListener("click",()=>_.remove()),_.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let T=_.querySelector("#bulk-confirm-btn");T.disabled=!0,T.textContent="Menghapus...";let j=await K(`${n}/bulk`,{method:"DELETE",body:JSON.stringify({ids:y})});_.remove(),j.ok?(R(`${y.length} ${o} berhasil dihapus.`),r.clear(),l(),be(n),w()):N(j.data?.error||"Gagal menghapus data.")})});let f=document.getElementById("filter-search"),v;f?.addEventListener("input",y=>{clearTimeout(v),v=setTimeout(()=>{s.search=y.target.value,g=1,r.clear(),l(),w()},400)}),m?.forEach(y=>{(y.type==="select"||y.type==="combobox")&&(document.getElementById(`filter-${y.name}`)?.addEventListener("change",_=>{s[y.name]=_.target.value;let T=document.getElementById(`filter-sheet-${y.name}`);T&&(T.value=_.target.value),g=1,r.clear(),l(),w()}),document.getElementById(`filter-sheet-${y.name}`)?.addEventListener("change",_=>{s[y.name]=_.target.value;let T=document.getElementById(`filter-${y.name}`);T&&(T.value=_.target.value),g=1,r.clear(),l(),w(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{s={},f&&(f.value=""),m?.forEach(y=>{let _=document.getElementById(`filter-${y.name}`);_&&(_.value="");let T=document.getElementById(`filter-sheet-${y.name}`);T&&(T.value="")}),g=1,r.clear(),l(),w()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{s={},f&&(f.value=""),m?.forEach(y=>{let _=document.getElementById(`filter-${y.name}`);_&&(_.value="");let T=document.getElementById(`filter-sheet-${y.name}`);T&&(T.value="")}),g=1,r.clear(),l(),w(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>z(null)),h&&document.addEventListener("click",function(y){let _=document.getElementById("aksi-menu-main"),T=document.getElementById("btn-aksi-main");_&&T&&!T.contains(y.target)&&!_.contains(y.target)&&_.classList.remove("show-aksi-menu")});let L=document.getElementById("btn-mobile-filter"),A=document.getElementById("filter-options-wrapper"),P=document.getElementById("btn-close-filter-sheet");if(L&&A&&(L.addEventListener("click",y=>{y.preventDefault(),A.classList.add("sheet-open")}),P&&P.addEventListener("click",y=>{y.preventDefault(),A.classList.remove("sheet-open")})),h){document.getElementById(`btn-export-${h.moduleName}`)?.addEventListener("click",async _=>{let T=_.target,j=T.innerHTML;T.innerHTML="\u23F3 Loading...",T.disabled=!0;try{await h.onExport()}catch{N("Gagal export data")}finally{T.innerHTML=j,T.disabled=!1}}),document.getElementById(`btn-template-${h.moduleName}`)?.addEventListener("click",()=>{h.onTemplate()});let y=document.getElementById(`input-import-${h.moduleName}`);y?.addEventListener("change",async _=>{let T=_.target.files[0];if(!T)return;let j=document.getElementById(`label-import-${h.moduleName}`),O=j?j.querySelector(".import-text"):null,M=O?O.innerText:"";O&&(O.innerText="\u231B Memproses..."),j&&(j.style.pointerEvents="none"),y.disabled=!0;try{let q=await qe(T);if(q.length===0)throw new Error("File kosong atau format salah");await h.onImport(q),R("Import berhasil!"),be(n),w()}catch(q){N(q.message||"Gagal import data")}finally{O&&(O.innerText=M),j&&(j.style.pointerEvents="auto"),y.disabled=!1,y.value=""}})}async function w(){l();let y=document.getElementById("table-container");if(!y)return;y.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let _=p==="client",T=_?1:g,j=_?Be:20,O=new URLSearchParams({page:T,limit:j,...Object.fromEntries(Object.entries(s).filter(([,U])=>U))}),M=await K(`${n}?${O}`);if(!M.ok){y.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${M.data?.error||"Error"}</p></div>`;return}let q=M.data?.data||M.data||[],H=M.data?.pagination,Y=q.length;if(_){q=$(q);let U=q.length,V=20,ae=Math.ceil(U/V);g>ae&&ae>0&&(g=ae);let J=(g-1)*V,ie=g*V;q=q.slice(J,ie),H={page:g,limit:V,total:U,pages:ae}}!1,F&&F(q);let W=Xe({columns:c,data:q,onEdit:B?U=>z(U):null,actions:E.map(U=>({...U,handler:V=>U.handler(V,w)})),emptyText:`Tidak ada ${String(o||"").toLowerCase()}`,bulkSelect:a?{selectedIds:r,onToggle:l}:null});y.innerHTML="",y.appendChild(W);let Z=document.getElementById("pagination-container");if(Z&&(Z.innerHTML="",H&&H.pages>1)){let U=Ze({page:H.page,pages:H.pages,total:H.total,limit:H.limit,onPage:V=>{g=V,w()}});U&&Z.appendChild(U)}}function I(y){let _=typeof x=="function"?x(y):x;return Ne(_)}function z(y){let _=!!y,T=document.createElement("form");if(T.noValidate=!0,T.innerHTML=I(y),_){let O=typeof x=="function"?x(y):x;tt(T,y)}let{close:j}=G({title:_?`Edit ${o}`:`Tambah ${o}`,content:T,size:"lg",confirmText:_?"Simpan Perubahan":`Tambah ${o}`,onConfirm:async(O,M)=>{if(!T.reportValidity())return;let q=O.querySelector(".modal-confirm");q.disabled=!0,q.textContent="Menyimpan...";let H=et(T),Y=typeof x=="function"?x(y):x,W=async ae=>{for(let J of ae)if(J.type==="row")await W(J.fields);else if(J.type==="combobox"&&H[J.name]){let ie=H[J.name],me=(J.options||[]).find(se=>{let ce=String(typeof se=="object"?se.value:se),Je=String(typeof se=="object"?se.label:se);return ce===ie||Je===ie});if(me)H[J.name]=typeof me=="object"?me.value:me;else if(J.createApi){let se={};se[J.createApi.field]=ie,J.createApi.extra&&Object.assign(se,J.createApi.extra);let ce=await K(J.createApi.path,{method:"POST",body:JSON.stringify(se)});if(ce.ok&&ce.data?.id)H[J.name]=ce.data.id;else if(ce.ok&&!ce.data?.id)H[J.name]=ie;else throw new Error(`Gagal membuat master data: ${ce.data?.error||"Unknown error"}`)}}};try{await W(Y)}catch(ae){N(ae.message),q.disabled=!1,q.textContent=_?"Simpan Perubahan":`Tambah ${o}`;return}S&&(H=await S(H,y));let Z=_?"PUT":"POST",U=_?`${n}/${y.id}`:n,V=await K(U,{method:Z,body:JSON.stringify(H)});V.ok?(R(_?`${o} berhasil diperbarui.`:`${o} berhasil ditambahkan.`),M(),be(n),w()):(N(V.data?.error||"Gagal menyimpan data."),q.disabled=!1,q.textContent=_?"Simpan Perubahan":`Tambah ${o}`)}})}function C(y){X(`Hapus ${o} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let _=await K(`${n}/${y.id}`,{method:"DELETE"});_.ok?(R(`${o} berhasil dihapus.`),be(n),w()):N(_.data?.error||"Gagal menghapus.")},`Hapus ${o}`)}return w(),w}var Te=null,nt=null;async function _e(e=!1){if(Te&&!e)return console.log("Employees Raw (Cache Hit)",Te.slice(0,5)),Te;let t=await K(`/api/employees?limit=${Be}&status=Aktif`);return Te=(t.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",t.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",Te.slice(0,5)),Te}async function ue(e=!1){let i=(await _e(e)).map(n=>({value:n.label,label:n.label}));return console.log("Employee Options",i.slice(0,5)),i}async function te(e=!1){return nt&&!e||(nt=((await K("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),nt}function ge(e){let t={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!e||e==="-"||String(e).trim()===""?"":`<span class="badge ${t[e]||"badge-neutral"}">${e}</span>`}function mt(e){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[e]||"badge-neutral"}">${e||"-"}</span>`}function xe(e){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[e]||"badge-neutral"}">${e||"-"}</span>`}re();le();le();re();le();var Ce={};function Le(e){if(Ce[e]){try{Ce[e].destroy()}catch{}delete Ce[e]}}function ln(){Object.keys(Ce).forEach(Le)}var fe=(e,t=0)=>{let i=Number(e);return isNaN(i)||e===null||e===void 0?t:i},De=(e,t="\u2014")=>{if(e==null||e==="")return t;let i=String(e).trim();return i===""||i==="[object Object]"?t:i};function it(e,t,i=900){if(!e)return;e._animFrame&&cancelAnimationFrame(e._animFrame);let n=Math.max(0,Math.round(fe(t))),c=e.textContent?e.textContent.replace(/[^\d]/g,""):"",x=c!==""&&parseInt(c)||0;if(n===x){e.textContent=n.toLocaleString("id-ID"),e._animFrame=null;return}let m=Date.now(),u=()=>{let o=Math.min((Date.now()-m)/i,1),b=1-Math.pow(1-o,3);e.textContent=Math.round(x+b*(n-x)).toLocaleString("id-ID"),o<1?e._animFrame=requestAnimationFrame(u):(e.textContent=n.toLocaleString("id-ID"),e._animFrame=null)};e._animFrame=requestAnimationFrame(u)}var dn={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},pn=e=>{let t=De(e,"\u2014");return`<span class="status-pill ${dn[t]||"pill-neutral"}">${t}</span>`};var ve={family:"Inter",size:11},Ee="#94A3B8",Pe="#F1F5F9",bt=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],cn=()=>window.innerWidth<768;function rt(e={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:cn()?"bottom":"top",labels:{font:ve,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ve,titleFont:{...ve,weight:"700"}}},scales:{x:{grid:{color:Pe},ticks:{font:ve,color:Ee,maxRotation:0}},y:{grid:{color:Pe},ticks:{font:ve,color:Ee},beginAtZero:!0}},...e}}var un=()=>Array(6).fill(0).map(()=>`
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
  </div>`).join(""),gn=()=>Array(6).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Ot(e=3){return Array(e).fill(0).map((t,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<e-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function ye(e,t,i=8e3){try{let n=new AbortController,c=setTimeout(()=>n.abort(),i),x=await K(e,{signal:n.signal}).catch(()=>null);if(clearTimeout(c),!x||!x.ok)return t;let m=x.data;return m?m.data!==void 0?m.data??t:m:t}catch{return t}}function mn(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(n=>{let c=document.getElementById(n);c&&(c.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(n=>{let c=document.getElementById(n);if(c&&c.style.display==="none"){c.style.display="block";let x=c.parentElement;if(x&&!x.querySelector(".chart-empty")){let m=document.createElement("div");m.className="chart-empty",m.textContent="Belum ada data",c.style.display="none",x.appendChild(m)}}});let e=document.getElementById("kpi-row");e&&e.querySelector(".skeleton")&&Gt({});let t=document.getElementById("mini-stats-row");t&&t.querySelector(".skeleton")&&Jt({}),["table-contracts","table-issues"].forEach(n=>{let c=document.getElementById(n);c&&c.querySelector(".skeleton")&&(c.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Ut(e){ln(),e._dashRefresh&&clearInterval(e._dashRefresh),e._skelTimeout&&clearTimeout(e._skelTimeout);let t=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});e.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">
      
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h2 style="margin:0; font-size:1.2rem; color:var(--text-1);">Dashboard Overview</h2>
        <button id="btn-dash-refresh" class="btn-primary" style="padding:6px 12px; font-size:0.8rem; display:flex; align-items:center; gap:6px; cursor:pointer; background:var(--primary); color:#fff; border:none; border-radius:6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
          Refresh Data
        </button>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${un()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${gn()}</div>

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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${Ot(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${Ot(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>Ht(e)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async i=>{let n=i.target.value,c=document.getElementById("jadwal-year-label");c&&(c.textContent=n);let x=document.getElementById("skel-jadwal"),m=document.getElementById("chart-jadwal");x&&(x.style.display="block",x.style.position="absolute"),m&&(m.style.display="none");let u=await ye(`/api/dashboard/schedule-chart?year=${n}`,{},8e3);try{Yt(u)}catch(o){console.warn("ScheduleChart render:",o),ke("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let n=i.target.value,c=n?`/api/dashboard/inspection-bar?month=${n}`:"/api/dashboard/inspection-bar",x=document.getElementById("skel-insp"),m=document.getElementById("chart-insp");x&&(x.style.display="block",x.style.position="absolute"),m&&(m.style.display="none");let u=await ye(c,{},8e3);try{Vt(u)}catch(o){console.warn("InspBar render:",o),ke("skel-insp","chart-insp")}}),e._skelTimeout=setTimeout(()=>mn(),5e3),await Ht(e)}async function Ht(e){e._skelTimeout&&(clearTimeout(e._skelTimeout),e._skelTimeout=null);let[t,i,n,c,x,m,u,o]=await Promise.all([ye("/api/dashboard/kpi",{},8e3),ye("/api/dashboard/issues-trend",{},8e3),ye("/api/dashboard/issues-summary",{},8e3),ye("/api/dashboard/stats",{},8e3),ye("/api/dashboard/calendar",[],8e3),ye("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),ye(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),ye("/api/dashboard/dropdown-aggregates",{},8e3)]),b=document.getElementById("filter-insp-month"),B=b?b.value:"",D=B?`/api/dashboard/inspection-bar?month=${B}`:"/api/dashboard/inspection-bar",S=await ye(D,{},8e3);window.dashboardAggregates=o||{};try{Gt(t)}catch(F){console.warn("KPI render:",F)}try{Jt(t)}catch(F){console.warn("MiniStats render:",F)}try{Yt(u)}catch(F){console.warn("ScheduleChart render:",F),ke("skel-jadwal","chart-jadwal")}try{bn(Array.isArray(n?.by_category)?n.by_category:[])}catch(F){console.warn("Donut render:",F),ke("skel-donut","chart-donut")}try{hn(i)}catch(F){console.warn("Trend render:",F),ke("skel-trend","chart-trend")}try{Vt(S)}catch(F){console.warn("InspBar render:",F),ke("skel-insp","chart-insp")}try{let F=Array.isArray(c)?c:Array.isArray(c?.recent_issues)?c.recent_issues:[];xn(F)}catch(F){console.warn("IssuesTable render:",F)}try{let F=Array.isArray(c?.expiring_contracts)?c.expiring_contracts:[];fn(m)}catch(F){console.warn("ContractsTable render:",F)}try{yn(Array.isArray(x)?x:[])}catch(F){console.warn("Agenda render:",F)}try{vn()}catch(F){console.warn("Quick Actions render:",F)}}function Gt(e){let t=document.getElementById("kpi-row");if(!t)return;e=e||{};let i=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,n=window.dashboardAggregates||{},c=o=>{let b=window.dashboardAggregates?.schedule_quarter_details?.[o];return b?`Insp: ${b.insp} \u2022 GC: ${b.gc} \u2022 DC: ${b.dc} \u2022 Fog: ${b.fog}`:""},x=o=>{let b=window.dashboardAggregates?.schedule_quarter_details?.[o];return b?`Rincian ${o}: Inspeksi: ${b.insp} | General Cleaning: ${b.gc} | Deep Cleaning: ${b.dc} | Fogging: ${b.fog}`:`Total Kegiatan ${o}`},m=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
          <option value="Q1" ${i==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${i==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${i==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${i==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:n.schedule_by_quarter?.[i]??e.schedule?.current??0,href:`#/timeline?dash_filter=period_${i.toLowerCase()}`,color:"kpi-blue",sub:c(i),subId:"mini-jadwal-detail",tooltip:x(i)}];t.innerHTML=m.map(o=>{let b=o.val!==void 0?fe(o.val,0):fe(e[o.key]?.current,0);return o.dropdown?`
        <div class="kpi-card ${o.color}" ${o.id?`id="${o.id}"`:""} ${o.tooltip?`title="${o.tooltip}"`:""} data-href="${o.href}" tabindex="0" role="link" style="cursor:pointer;padding:10px 12px">
          <div style="display:flex; gap:10px; align-items:center;">
            <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${o.icon}</span></div>
            <div style="flex:1;min-width:0;">
              <div style="display:flex; align-items:baseline; gap:3px;">
                <div class="kpi-value" data-target="${b}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${b}</div>
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
              <div class="kpi-value" data-target="${b}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${b}</div>
              ${o.dropdown?o.dropdown:""}
            </div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${o.label}</div>
            <div class="kpi-subtitle" ${o.subId?`id="${o.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${o.tooltip||""}">${o.sub}</div>
          </div>
        </div>
      </a>`}).join(""),t.querySelectorAll(".kpi-card[data-href]").forEach(o=>{o.addEventListener("click",b=>{b.target.closest("select")||o.dataset.href&&(window.location.hash=o.dataset.href)}),o.addEventListener("keydown",b=>{b.key==="Enter"&&!b.target.closest("select")&&o.dataset.href&&(window.location.hash=o.dataset.href)})}),t.querySelectorAll(".kpi-value").forEach(o=>{it(o,parseInt(o.dataset.target)||0)});let u=document.getElementById("dash-jadwal-period");if(u){let o=b=>{let D=(window.dashboardAggregates?.schedule_by_quarter||{})[b]||0,S=document.querySelector("#mini-jadwal .kpi-value");S&&(S.dataset.target=D,it(S,D,400));let F=document.getElementById("mini-jadwal-detail");F&&(F.textContent=c(b),F.title=x(b));let $=document.getElementById("mini-jadwal");if($){let E=`#/timeline?dash_filter=period_${b.toLowerCase()}`;$.tagName==="A"?$.href=E:$.dataset.href=E,$.title=x(b)}};u.addEventListener("change",b=>o(b.target.value))}}function Jt(e){let t=document.getElementById("mini-stats-row");if(!t)return;e=e||{};let i=new Date().getFullYear(),n=String(new Date().getMonth()+1).padStart(2,"0"),c=`${i}-${n}`,x=b=>`
    <select id="${b}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
      <option value="${i}-01" ${c===`${i}-01`?"selected":""}>Jan</option>
      <option value="${i}-02" ${c===`${i}-02`?"selected":""}>Feb</option>
      <option value="${i}-03" ${c===`${i}-03`?"selected":""}>Mar</option>
      <option value="${i}-04" ${c===`${i}-04`?"selected":""}>Apr</option>
      <option value="${i}-05" ${c===`${i}-05`?"selected":""}>Mei</option>
      <option value="${i}-06" ${c===`${i}-06`?"selected":""}>Jun</option>
      <option value="${i}-07" ${c===`${i}-07`?"selected":""}>Jul</option>
      <option value="${i}-08" ${c===`${i}-08`?"selected":""}>Agu</option>
      <option value="${i}-09" ${c===`${i}-09`?"selected":""}>Sep</option>
      <option value="${i}-10" ${c===`${i}-10`?"selected":""}>Okt</option>
      <option value="${i}-11" ${c===`${i}-11`?"selected":""}>Nov</option>
      <option value="${i}-12" ${c===`${i}-12`?"selected":""}>Des</option>
    </select>
  `,m=window.dashboardAggregates||{},u=[{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:x("dash-inspeksi-month"),val:m.inspeksi_by_month?.[c]??e.inspection_month?.current??0,href:`#/timeline?dash_filter=inspeksi&month=${c}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:x("dash-gcdc-month"),val:m.gcdc_by_month?.[c]??e.cleaning_month?.current??0,href:`#/timeline?dash_filter=gcdc&month=${c}`,color:"mini-green"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:x("dash-reliefer-month"),val:m.relievers_by_month?.[c]??e.reliever_completed?.current??0,href:`#/relievers?dash_filter=reliever&month=${c}`,color:"mini-teal"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:x("dash-fogging-month"),val:m.fogging_by_month?.[c]??e.fogging_month?.current??0,href:`#/reports/fogging?dash_filter=fogging&month=${c}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:e.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:e.branches?.current,href:"#/branches",color:"mini-teal"}];t.innerHTML=u.map(b=>b.dropdown?`
        <div class="mini-stat ${b.color}" style="cursor:pointer;" id="${b.id||""}" data-href="${b.href}" tabindex="0" role="link">
          <div class="mini-stat-icon">${b.icon}</div>
          <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="mini-stat-value" data-target="${fe(b.val)}">0</div>
              ${b.dropdown}
            </div>
            <div class="mini-stat-text">${b.label}</div>
            ${b.extra||""}
          </div>
        </div>`:`
      <a href="${b.href}" class="mini-stat ${b.color}" style="text-decoration:none" id="${b.id||""}">
        <div class="mini-stat-icon">${b.icon}</div>
        <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
          <div style="display:flex; align-items:baseline; gap:3px;">
            <div class="mini-stat-value" data-target="${fe(b.val)}">0</div>
            ${b.dropdown?b.dropdown:""}
          </div>
          <div class="mini-stat-text">${b.label}</div>
          ${b.extra||""}
        </div>
      </a>`).join(""),t.querySelectorAll(".mini-stat[data-href]").forEach(b=>{b.addEventListener("click",B=>{B.target.closest("select")||b.dataset.href&&(window.location.hash=b.dataset.href)}),b.addEventListener("keydown",B=>{B.key==="Enter"&&!B.target.closest("select")&&b.dataset.href&&(window.location.hash=b.dataset.href)})}),t.querySelectorAll(".mini-stat-value").forEach(b=>it(b,parseInt(b.dataset.target)||0,700));let o=(b,B,D,S)=>{let F=document.getElementById(b);if(F){let $=E=>{let h=(window.dashboardAggregates?.[D]||{})[E]||0,a=document.querySelector(`#${B} .mini-stat-value`);a&&(a.dataset.target=h,it(a,h,400));let p=document.getElementById(B);if(p){let d=`${S}&month=${E}`;p.tagName==="A"?p.href=d:p.dataset.href=d}};F.addEventListener("change",E=>$(E.target.value))}};o("dash-reliefer-month","mini-reliefer","relievers_by_month","#/relievers?dash_filter=reliever"),o("dash-inspeksi-month","mini-inspeksi","inspeksi_by_month","#/timeline?dash_filter=inspeksi"),o("dash-gcdc-month","mini-gcdc","gcdc_by_month","#/timeline?dash_filter=gcdc"),o("dash-fogging-month","mini-fogging","fogging_by_month","#/reports/fogging?dash_filter=fogging")}function bn(e){ke("skel-donut","chart-donut");let t=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!t||!i)return;Le("donut");let n=(e||[]).filter(o=>fe(o.count)>0);if(!n.length){Re(t,"Belum ada data permasalahan");return}let c=n.map(o=>`${De(o.category,"Lainnya")}`),x=n.map(o=>fe(o.count)),m=x.reduce((o,b)=>o+b,0);i.innerHTML=n.map((o,b)=>{let B=bt[b%bt.length],D=m>0?Math.round(o.count/m*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${B}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${o.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${D}%)</span></div>
          <div class="donut-legend-label">${c[b]}</div>
        </div>
      </div>
    `}).join("");let u={id:"centerText",beforeDraw:function(o){let b=o.width,B=o.height,D=o.ctx;D.restore();let S=(B/80).toFixed(2);D.font="bold "+S+"em Inter",D.textBaseline="middle",D.fillStyle="#1E293B";let F=m.toString(),$=Math.round((b-D.measureText(F).width)/2),E=B/2;D.fillText(F,$,E-4),D.font="600 "+(S*.35).toFixed(2)+"em Inter",D.fillStyle="#64748B";let k="Total",h=Math.round((b-D.measureText(k).width)/2);D.fillText(k,h,E+10),D.save()}};Ce.donut=new Chart(t,{type:"doughnut",data:{labels:c,datasets:[{data:x,backgroundColor:bt,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ve,titleFont:{...ve,weight:"700"},callbacks:{label:o=>` ${o.label}: ${o.parsed} kasus`}}},cutout:"75%"},plugins:[u]})}function hn(e){ke("skel-trend","chart-trend");let t=document.getElementById("chart-trend");if(!t)return;Le("trend"),e=e||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],n=(e.labels||[]).map(m=>{if(!m||typeof m!="string")return"";try{let[u,o]=m.split("-");return(i[Number(o)-1]||o)+" "+String(u).slice(-2)}catch{return m}}),c=(e.open||[]).map(m=>fe(m)),x=(e.closed||[]).map(m=>fe(m));if(!n.length){Re(t,"Belum ada data trend");return}Ce.trend=new Chart(t,{type:"line",data:{labels:n,datasets:[{label:"Open",data:c,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:x,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:rt({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Ee,maxRotation:45,autoSkip:!0}},y:{grid:{color:Pe},ticks:{font:{family:"Inter",size:9},color:Ee},beginAtZero:!0}}})})}function Yt(e){ke("skel-jadwal","chart-jadwal");let t=document.getElementById("chart-jadwal");if(!t)return;Le("jadwal"),e=e||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(e).some(o=>Array.isArray(o)&&o.some(b=>b>0))){Re(t,"Belum ada data jadwal");return}let c=e["Inspeksi Hygiene"]||Array(12).fill(0),x=e["General Cleaning"]||Array(12).fill(0),m=e["Deep Cleaning"]||Array(12).fill(0),u=e.Fogging||Array(12).fill(0);Ce.jadwal=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Inspeksi",data:c,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:x,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:m,backgroundColor:"#F59E0B"},{label:"Fogging",data:u,backgroundColor:"#EF4444"}]},options:rt({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Ee,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Pe},ticks:{font:{family:"Inter",size:9},color:Ee},min:0}}})})}function Vt(e){ke("skel-insp","chart-insp");let t=document.getElementById("chart-insp");if(!t)return;Le("inspBar"),e=e||{};let i=e.labels||[],n=(e.fc||[]).map(x=>fe(x)),c=(e.spv||[]).map(x=>fe(x));if(!i.length){Re(t,"Belum ada data inspeksi");return}Ce.inspBar=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:n,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:c,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:rt({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ve,color:Ee,maxRotation:45,minRotation:45}},y:{grid:{color:Pe},ticks:{font:ve,color:Ee},min:0,max:100}}})})}function fn(e){ke("skel-contract-mini","chart-contract-mini");let t=document.getElementById("chart-contract-mini");if(!t)return;Le("contractMiniBar"),e=e||{};let i={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},n=(e.labels||[]).map(m=>{let u=m.split("-")[1];return i[u]||m}),c=(e.data||[]).map(m=>fe(m));if(!n.length){Re(t,"Belum ada data");return}let x=t.getContext("2d");Ce.contractMiniBar=new Chart(t,{type:"bar",data:{labels:n,datasets:[{label:"Kontrak Habis",data:c,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:rt({onClick:(m,u)=>{if(u&&u.length>0){let o=u[0].index,b=(e.labels||[])[o];b&&(window.location.hash="#/contracts?month_expiry="+b)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ve,color:Ee,maxRotation:0,autoSkip:!1}},y:{grid:{color:Pe,borderDash:[4,4],drawBorder:!1},ticks:{font:ve,color:Ee,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function xn(e){let t=document.getElementById("table-issues");if(!t)return;let i=(e||[]).slice(0,8);if(!i.length){t.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}t.innerHTML=`
    <div class="dash-list">
      ${i.map(n=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${pn(n.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${De(n.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${De(n.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function yn(e){let t=document.getElementById("widget-agenda");if(!t)return;let i=new Date,n=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,x=(e||[]).filter(m=>(m.event_date||"").startsWith(n)).slice(0,10);if(!x.length){t.innerHTML="";return}t.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${x.map(m=>{let u="#3B82F6",o="#EFF6FF",b="Agenda",B=(m.title||"").toLowerCase();return B.includes("inspeksi")?(u="#10B981",o="#ECFDF5",b="Inspeksi"):B.includes("cleaning")||B.includes("gcdc")?(u="#3B82F6",o="#EFF6FF",b="Cleaning"):B.includes("reliefer")?(u="#F59E0B",o="#FFFBEB",b="Reliefer"):B.includes("fogging")&&(u="#8B5CF6",o="#F5F3FF",b="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(m.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${u};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${De(m.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${De(m.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${o};color:${u}">${b}</div>
        </div>
      `}).join("")}
    </div>
  `}function vn(){let e=document.getElementById("quick-actions");if(!e)return;let t=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];e.innerHTML=t.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function ke(e,t){let i=document.getElementById(e),n=document.getElementById(t);if(i&&(i.style.display="none",i.style.position=""),n){n.style.display="block";let c=n.parentElement;if(c){let x=c.querySelector(".chart-empty");x&&x.remove()}}}function Re(e,t="Belum ada data"){if(!e)return;e.style.display="none";let i=e.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let c=document.createElement("div");c.className="chart-empty",c.textContent=t,i.appendChild(c)}}async function Wt(e){document.getElementById("app").innerHTML=`
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
  `;let t=document.getElementById("login-form"),i=document.getElementById("login-error"),n=document.getElementById("login-btn"),c=document.getElementById("toggle-password"),x=document.getElementById("login-password");c?.addEventListener("click",()=>{let m=x.type==="text";x.type=m?"password":"text",c.style.color=m?"":"var(--primary)"}),t?.addEventListener("submit",async m=>{m.preventDefault(),i.style.display="none";let u=t.username.value.trim(),o=t.password.value;if(!u||!o){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}n.querySelector(".btn-text").style.display="none",n.querySelector(".btn-spinner").style.display="",n.disabled=!0;try{let b=await K("/api/auth/login",{method:"POST",body:JSON.stringify({username:u,password:o})});b.ok&&b.data.success?(Nt(b.data.data.token),We(b.data.data.user),R("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=b.data.error||"Username atau password salah.",i.style.display="block",n.classList.add("shake"),setTimeout(()=>n.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{n.querySelector(".btn-text").style.display="",n.querySelector(".btn-spinner").style.display="none",n.disabled=!1}})}le();var ht=[{id:1,full_name:"ACHMAD NURYAINI",branch_id:7,branch_name:"007. Bekasi",division:"FACILITY CARE",phone:"081298765431",join_date:"2023-10-09",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:2,full_name:"ADAM JHORDY",branch_id:32,branch_name:"032. Cawang",division:"FACILITY CARE",phone:"085712345678",join_date:"2024-03-18",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:3,full_name:"ADE MUSTIKA REZDAFITRA",branch_id:16,branch_name:"016. Cibubur",division:"FACILITY CARE",phone:"081345678901",join_date:"2022-06-21",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:4,full_name:"ADE SURAHMAN",branch_id:0,branch_name:"000. Management",division:"FACILITY CARE",phone:"081223344556",join_date:"2021-01-15",status:"Aktif",notes:"Trainer & Supervisor Operasional",target_pindah_os:"",target_selesai_os:""},{id:5,full_name:"ADITYA",branch_id:26,branch_name:"026. Tangerang City",division:"FACILITY CARE",phone:"087811223344",join_date:"2024-10-19",status:"Aktif",notes:"",target_pindah_os:"",target_selesai_os:""},{id:6,full_name:"ADITYA PRATAMA",branch_id:54,branch_name:"054. Bogor Gunung Batu",division:"FACILITY CARE",phone:"089699887766",join_date:"2025-09-04",status:"Aktif",notes:"Lulusan Training Batch 5 nilai tertinggi",target_pindah_os:"",target_selesai_os:""},{id:7,full_name:"AGUNG SEPTIADI",branch_id:0,branch_name:"000. Management",division:"FC - RELIEFER",phone:"082155667788",join_date:"2026-02-12",status:"Aktif",notes:"Personil cover cabang mobile",target_pindah_os:"",target_selesai_os:""}],Qt=["FACILITY CARE","SECURITY","FC - RELIEFER"],kn=["Aktif","Tidak Aktif","Resign","Cut"];function wn(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function En(e){if(!e||e==="-")return"\u{1F464}";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function Xt(e,t){let i=[],n=[],c="all",x="all",m="all",u="",o=!1,b=t?t.get("dash_filter"):null;b==="active"?c="Aktif":b==="reliefer"&&(x="FC - RELIEFER");let D=ee()?.role!=="viewer";try{n=await te()}catch(a){console.error("Error fetching branches:",a)}async function S(){o=!0,h();try{let a=await K("/api/employees?limit=10000");if(a.ok){let p=a.data?.data||a.data||[];i=Array.isArray(p)&&p.length>0?p:ht}else i=ht}catch(a){console.warn("Gagal load employees, menggunakan data lokal:",a),i=ht}finally{o=!1,h()}}async function F(){try{let{downloadExcel:a}=await Promise.resolve().then(()=>(re(),he)),p=i.map((d,g)=>({No:g+1,"Nama Lengkap":d.full_name||"",Cabang:d.branch_name||"",Divisi:d.division||"","No. HP":d.phone||"","Tgl Masuk":d.join_date||"",Status:d.status||"","Target Pindah OS":d.target_pindah_os||"","Target Selesai OS":d.target_selesai_os||"",Catatan:d.notes||""}));a(p,`Data_Karyawan_FCMS_${new Date().toISOString().split("T")[0]}`),R("Berhasil mengekspor data karyawan ke Excel")}catch(a){N("Gagal ekspor Excel: "+a.message)}}function $(a=null){let p=!!a,d=a?.division||"FACILITY CARE",g=a?.status||"Aktif",s=document.createElement("div");s.className="emp-modal-body",s.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Lengkap Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <input type="text" id="emp-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: Achmad Nuryaini" value="${a?.full_name?String(a.full_name).replace(/"/g,"&quot;"):""}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              No. Handphone / WhatsApp
            </label>
            <input type="tel" id="emp-modal-phone" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="08xx-xxxx-xxxx" value="${a?.phone?String(a.phone).replace(/"/g,"&quot;"):""}" />
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${n.map(r=>`
                <option value="${r.value}" ${String(a?.branch_id)===String(r.value)?"selected":""}>
                  ${r.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi / Bagian <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-division" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${Qt.map(r=>`
                <option value="${r}" ${d===r?"selected":""}>${r}</option>
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
              value="${a?.join_date||""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${kn.map(r=>`
                <option value="${r}" ${g===r?"selected":""}>${r}</option>
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
              value="${a?.target_pindah_os?String(a.target_pindah_os).slice(0,7):""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Target Selesai OS (Opsional)
            </label>
            <input type="month" id="emp-modal-target-selesai" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${a?.target_selesai_os?String(a.target_selesai_os).slice(0,7):""}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan / Keterangan Karyawan
          </label>
          <textarea id="emp-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Catatan keahlian, tugas khusus, riwayat cover...">${a?.notes||""}</textarea>
        </div>
      </div>
    `,G({title:p?"\u270F\uFE0F Edit Data Karyawan":"\u2795 Tambah Data Karyawan",content:s,size:"md",confirmText:p?"Simpan Perubahan":"Simpan Karyawan",confirmClass:"btn-primary",onConfirm:async(r,l)=>{let f=(s.querySelector("#emp-modal-name").value||"").trim(),v=(s.querySelector("#emp-modal-phone").value||"").trim(),L=s.querySelector("#emp-modal-branch").value,A=s.querySelector("#emp-modal-division").value,P=s.querySelector("#emp-modal-joindate").value||null,w=s.querySelector("#emp-modal-status").value,I=s.querySelector("#emp-modal-target-pindah").value||null,z=s.querySelector("#emp-modal-target-selesai").value||null,C=(s.querySelector("#emp-modal-notes").value||"").trim();if(!f||!L||!A||!w){N("Harap lengkapi Nama Lengkap, Cabang, Divisi, dan Status!");return}let y={full_name:f,phone:v,branch_id:Number(L),division:A,join_date:P,status:w,target_pindah_os:I,target_selesai_os:z,notes:C};try{let _;p?_=await K(`/api/employees/${a.id}`,{method:"PUT",body:JSON.stringify(y)}):_=await K("/api/employees",{method:"POST",body:JSON.stringify(y)}),_.ok?(R(p?"Data karyawan berhasil diperbarui!":"Karyawan baru berhasil ditambahkan!"),l(),await S()):N(_.data?.error||"Gagal menyimpan data karyawan")}catch(_){N("Terjadi kesalahan: "+_.message)}}})}async function E(a){X({title:"Nonaktifkan Karyawan?",message:`Apakah Anda yakin ingin menonaktifkan personil <strong>${a.full_name}</strong> dari cabang <em>${a.branch_name||"cabang"}</em>?`,confirmText:"Ya, Nonaktifkan",confirmClass:"btn-danger",onConfirm:async()=>{try{let p=await K(`/api/employees/${a.id}`,{method:"DELETE"});p.ok?(R("Status karyawan berhasil diubah menjadi Tidak Aktif"),await S()):N(p.data?.error||"Gagal menonaktifkan data")}catch(p){N("Terjadi kesalahan: "+p.message)}}})}function k(){return i.filter(a=>{if(u){let p=u.toLowerCase(),d=(a.full_name||"").toLowerCase().includes(p),g=(a.branch_name||"").toLowerCase().includes(p),s=(a.phone||"").toLowerCase().includes(p),r=(a.division||"").toLowerCase().includes(p),l=(a.notes||"").toLowerCase().includes(p);if(!d&&!g&&!s&&!r&&!l)return!1}return!(c!=="all"&&(c==="Aktif"&&a.status!=="Aktif"||c==="Nonaktif"&&a.status==="Aktif")||x!=="all"&&a.division!==x||m!=="all"&&String(a.branch_id)!==String(m))})}function h(){let a=k(),p=i.length,d=i.filter(C=>C.status==="Aktif").length,g=i.filter(C=>C.division==="FACILITY CARE"&&C.status==="Aktif").length,s=i.filter(C=>(C.division==="SECURITY"||C.division==="FC - RELIEFER")&&C.status==="Aktif").length,r=i.length,l=i.filter(C=>C.status==="Aktif").length,f=i.filter(C=>C.status!=="Aktif").length;e.innerHTML=`
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
              ${D?`
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
                <div class="emp-stat-val">${p}</div>
                <div class="emp-stat-lbl">Total Karyawan</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">\u{1F7E2}</div>
              <div>
                <div class="emp-stat-val">${d}</div>
                <div class="emp-stat-lbl">Personil Aktif</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(14, 165, 233, 0.2); color:#7DD3FC;">\u{1F3E2}</div>
              <div>
                <div class="emp-stat-val">${g}</div>
                <div class="emp-stat-lbl">Facility Care</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(124, 58, 237, 0.2); color:#C4B5FD;">\u{1F6E1}\uFE0F</div>
              <div>
                <div class="emp-stat-val">${s}</div>
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
              <option value="all">Semua Cabang (${n.length})</option>
              ${n.map(C=>`
                <option value="${C.value}" ${String(m)===String(C.value)?"selected":""}>
                  ${C.label}
                </option>
              `).join("")}
            </select>

            <select id="emp-division-filter" class="emp-select">
              <option value="all">Semua Divisi</option>
              ${Qt.map(C=>`
                <option value="${C}" ${x===C?"selected":""}>
                  ${C}
                </option>
              `).join("")}
            </select>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="emp-filter-pills">
            <div class="emp-pill ${c==="all"&&x==="all"?"active":""}" data-status="all" data-div="all">
              <span>Semua Karyawan</span>
              <span class="emp-pill-count">${r}</span>
            </div>
            <div class="emp-pill ${c==="Aktif"?"active":""}" data-status="Aktif" data-div="all">
              <span class="emp-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="emp-pill-count">${l}</span>
            </div>
            <div class="emp-pill ${x==="FC - RELIEFER"?"active":""}" data-status="all" data-div="FC - RELIEFER">
              <span class="emp-pill-dot" style="background:#8B5CF6;"></span>
              <span>FC - Reliefer</span>
              <span class="emp-pill-count">${i.filter(C=>C.division==="FC - RELIEFER").length}</span>
            </div>
            <div class="emp-pill ${c==="Nonaktif"?"active":""}" data-status="Nonaktif" data-div="all">
              <span class="emp-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="emp-pill-count">${f}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive Cards Grid -->
        <div class="emp-cards-grid">
          ${a.length===0?`
            <div class="emp-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada karyawan yang cocok</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian, filter cabang, atau filter divisi.</p>
              <button id="emp-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
          `:a.map(C=>{let y=C.division||"FACILITY CARE",_="emp-div-fc",T="";y==="SECURITY"?(_="emp-div-sec",T="sec"):y==="FC - RELIEFER"&&(_="emp-div-rel",T="rel");let j=C.status==="Aktif",O=j?"emp-status-aktif":"emp-status-nonaktif",M=j?"\u{1F7E2} Aktif":`\u26AA ${C.status||"Nonaktif"}`;return`
              <div class="emp-card" data-id="${C.id}">
                <!-- Top Badges Row -->
                <div class="emp-card-top">
                  <span class="emp-div-badge ${_}">
                    ${y==="SECURITY"?"\u{1F6E1}\uFE0F":y==="FC - RELIEFER"?"\u{1F504}":"\u{1F9F9}"} ${y}
                  </span>
                  <div style="display:flex; align-items:center; gap:8px;">
                    <span class="emp-status-badge ${O}">
                      ${M}
                    </span>
                  </div>
                </div>

                <!-- Profile Box -->
                <div class="emp-profile-box">
                  <div class="emp-avatar ${T}">
                    ${En(C.full_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="emp-name" style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${C.full_name||"-"}
                    </div>
                    <div class="emp-subtitle">
                      Personil ${y}
                    </div>
                  </div>
                </div>

                <!-- Information Grid -->
                <div class="emp-info-card">
                  <div class="emp-info-row">
                    <span class="emp-info-label">\u{1F4CD} Cabang:</span>
                    <span class="emp-info-val">${C.branch_name||"Cabang Tidak Ditentukan"}</span>
                  </div>
                  <div class="emp-info-row">
                    <span class="emp-info-label">\u{1F4DE} No. HP:</span>
                    <span class="emp-info-val">
                      ${C.phone?`
                        <a href="tel:${C.phone}" class="emp-phone-link">
                          <span>${C.phone}</span>
                        </a>
                      `:'<span style="color:#94A3B8;">-</span>'}
                    </span>
                  </div>
                  <div class="emp-info-row">
                    <span class="emp-info-label">\u{1F5D3}\uFE0F Tgl Masuk:</span>
                    <span class="emp-info-val">${wn(C.join_date)}</span>
                  </div>
                  ${C.target_pindah_os||C.target_selesai_os?`
                    <div class="emp-info-row target-os" style="border-top:1px dashed #E2E8F0; padding-top:6px; margin-top:2px; flex-direction:column; align-items:flex-start; gap:3px; width:100%;">
                      <span class="emp-info-label" style="font-weight:600; color:#475569;">\u{1F504} Target Migrasi OS:</span>
                      <span class="emp-info-val" style="font-size:0.78rem; color:#4F46E5; text-align:left; width:100%; word-break:break-word; overflow-wrap:anywhere;">
                        ${C.target_pindah_os?"Pindah: "+C.target_pindah_os:""}
                        ${C.target_selesai_os?(C.target_pindah_os?" \u2192 ":"")+"Selesai: "+C.target_selesai_os:""}
                      </span>
                    </div>
                  `:""}
                </div>

                <!-- Notes box (if any) -->
                ${C.notes?`
                  <div class="emp-notes-box">
                    \u{1F4DD} <strong>Catatan:</strong> ${C.notes}
                  </div>
                `:""}

                <!-- Footer Actions -->
                <div class="emp-card-footer">
                  <div style="font-size:0.75rem; color:#94A3B8;">
                    ID #${C.id}
                  </div>

                  <div class="emp-action-btns">
                    ${C.phone?`
                      <a href="https://wa.me/${C.phone.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="emp-icon-btn" title="Chat WhatsApp" style="color:#16A34A;">
                        <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                      </a>
                    `:""}

                    ${D?`
                      <button class="emp-icon-btn edit" data-action="edit" data-id="${C.id}" title="Edit Data Karyawan">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                      <button class="emp-icon-btn delete" data-action="delete" data-id="${C.id}" title="Nonaktifkan Karyawan">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                      </button>
                    `:""}
                  </div>
                </div>
              </div>
            `}).join("")}
        </div>

      </div>
    `;let v=e.querySelector("#emp-search-input");v&&v.addEventListener("input",C=>{u=C.target.value,h();let y=e.querySelector("#emp-search-input");y&&(y.focus(),y.setSelectionRange(y.value.length,y.value.length))});let L=e.querySelector("#emp-branch-filter");L&&L.addEventListener("change",C=>{m=C.target.value,h()});let A=e.querySelector("#emp-division-filter");A&&A.addEventListener("change",C=>{x=C.target.value,h()}),e.querySelectorAll(".emp-pill").forEach(C=>{C.addEventListener("click",()=>{c=C.dataset.status;let y=C.dataset.div;y&&y!=="all"?x=y:C.dataset.status==="all"&&(x="all"),h()})});let P=e.querySelector("#emp-reset-filter");P&&P.addEventListener("click",()=>{u="",m="all",x="all",c="all",h()});let w=e.querySelector("#emp-btn-add");w&&w.addEventListener("click",()=>$());let I=e.querySelector("#emp-btn-excel");I&&I.addEventListener("click",()=>F());let z=e.querySelector("#emp-btn-refresh");z&&z.addEventListener("click",()=>S()),e.querySelectorAll(".emp-card").forEach(C=>{let y=Number(C.dataset.id),_=i.find(O=>O.id===y);if(!_)return;let T=C.querySelector('[data-action="edit"]');T&&T.addEventListener("click",()=>$(_));let j=C.querySelector('[data-action="delete"]');j&&j.addEventListener("click",()=>E(_))})}await S()}le();re();function ft(e){if(!e||e==="-")return"-";if(String(e).startsWith("2099"))return"Tetap / PKWTT";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Fe(e){return e?String(e).startsWith("2099"):!1}function Sn(e){if(!e||e==="-")return"\u{1F464}";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function Zt(e,t){let i=[],n=[],c=[],x={},m="",u="all",o="all",b="all",B="all",D=!1,S=t?t.get("dash_filter"):null;S==="expiring30"?(B="30",b="Aktif"):S==="active"&&(b="Aktif");let F=ee(),$=F&&F.role!=="viewer";try{let[s,r]=await Promise.all([te(),_e()]);n=Array.isArray(s)?s:[],c=Array.isArray(r)?r:[],c.forEach(l=>{l.id&&l.phone&&(x[l.id]=l.phone),l.full_name&&l.phone&&(x[l.full_name.toLowerCase()]=l.phone)})}catch(s){console.error("Error fetching cache options for contracts:",s)}async function E(){D=!0,g();try{let s=await K("/api/contracts?limit=10000");if(s.ok){let r=s.data?.data||s.data||[];i=Array.isArray(r)?r:[]}else i=[],N("Gagal memuat data kontrak dari server.")}catch(s){console.error("Error load contracts:",s),i=[],N("Terjadi kesalahan koneksi saat memuat kontrak.")}finally{D=!1,g()}}function k(){return i.filter(s=>{if(m.trim()){let r=m.toLowerCase().trim(),l=String(s.employee_name||"").toLowerCase(),f=String(s.branch_name||"").toLowerCase(),v=String(s.pkwt_number||"").toLowerCase(),L=String(s.contract_type||"").toLowerCase(),A=String(s.notes||"").toLowerCase();if(!(l.includes(r)||f.includes(r)||v.includes(r)||L.includes(r)||A.includes(r)))return!1}if(u!=="all"&&String(s.branch_id)!==String(u)||o!=="all"&&s.division!==o||b==="Aktif"&&s.status!=="Aktif")return!1;if(b==="Expiring30"){let r=Fe(s.end_date),l=Number(s.days_remaining);if(s.status!=="Aktif"||r||isNaN(l)||l<0||l>30)return!1}if(b==="Tetap"&&(!Fe(s.end_date)||s.status!=="Aktif")||b==="Nonaktif"&&s.status==="Aktif")return!1;if(B!=="all"){let r=parseInt(B,10),l=Fe(s.end_date),f=Number(s.days_remaining);if(l||s.status!=="Aktif"||isNaN(f)||f<0||f>r)return!1}return!0})}async function h(){G({title:"\u{1F50D} Analisis Selisih Kontrak Karyawan",content:`
        <div style="text-align:center; padding:24px 10px;">
          <div style="font-size:2rem; margin-bottom:10px;">\u23F3</div>
          <p style="color:#64748B; font-size:0.9rem;">Sedang menyinkronkan data karyawan aktif dengan kontrak aktif di database...</p>
        </div>
      `,cancelText:"Tutup"});try{let[s,r]=await Promise.all([K("/api/employees?limit=10000&status=Aktif"),K("/api/contracts?limit=10000")]),l=s.ok?s.data?.data||s.data||[]:[],f=r.ok?r.data?.data||r.data||[]:[],v=new Set(f.filter(P=>P.status==="Aktif").map(P=>P.employee_id)),L=l.filter(P=>!v.has(P.id)),A="";L.length===0?A=`
          <div style="text-align:center; padding:20px 10px;">
            <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F389}</div>
            <h4 style="font-size:1.1rem; font-weight:800; color:#10B981; margin:0 0 8px 0;">Semua Karyawan Aktif Memiliki Kontrak!</h4>
            <p style="font-size:0.86rem; color:#64748B; margin:0;">
              Total <strong>${l.length}</strong> karyawan aktif tercatat memiliki perikatan kontrak aktif yang sah di sistem.
            </p>
          </div>
        `:A=`
          <div style="display:flex; flex-direction:column; gap:12px;">
            <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:10px; padding:12px; font-size:0.84rem; color:#991B1B;">
              Ditemukan <strong>${L.length}</strong> dari <strong>${l.length}</strong> karyawan aktif yang <strong>belum memiliki Kontrak Aktif</strong>.
            </div>
            <div style="max-height:360px; overflow-y:auto; display:flex; flex-direction:column; gap:8px; padding-right:4px;">
              ${L.map(P=>{let w=f.filter(z=>z.employee_id===P.id),I='<span style="color:#D97706; font-weight:600;">Belum pernah ada data kontrak</span>';if(w.length>0){let z=w[0];I=`Riwayat: Status <strong>${z.status}</strong> (Selesai: ${ft(z.end_date)})`}return`
                  <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:8px; padding:10px 12px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">
                    <div>
                      <div style="font-weight:700; font-size:0.88rem; color:#0F172A;">${P.full_name}</div>
                      <div style="font-size:0.76rem; color:#64748B;">
                        Cabang: ${P.branch_name||"-"} \u2022 Divisi: ${P.division||"FACILITY CARE"}
                      </div>
                      <div style="font-size:0.75rem; color:#475569; margin-top:2px;">
                        ${I}
                      </div>
                    </div>
                    ${$?`
                      <button class="btn btn-primary btn-add-contract-for-emp" data-empid="${P.id}" data-empname="${String(P.full_name).replace(/"/g,"&quot;")}" data-branchid="${P.branch_id||""}" data-division="${P.division||"FACILITY CARE"}" style="font-size:0.76rem; padding:5px 10px; border-radius:6px; background:#2563EB;">
                        + Buat Kontrak
                      </button>
                    `:""}
                  </div>
                `}).join("")}
            </div>
          </div>
        `,G({title:"Analisis Selisih Kontrak Karyawan",content:A,cancelText:"Tutup"}),setTimeout(()=>{document.querySelectorAll(".btn-add-contract-for-emp").forEach(P=>{P.addEventListener("click",w=>{let I=w.currentTarget,z=I.getAttribute("data-empid"),C=I.getAttribute("data-empname"),y=I.getAttribute("data-branchid"),_=I.getAttribute("data-division"),T=document.querySelector(".modal-close");T&&T.click(),a({employee_id:z,employee_name:C,branch_id:y,division:_,status:"Aktif"})})})},100)}catch(s){console.error(s),N("Gagal melakukan analisis selisih data kontrak.")}}function a(s=null){let r=!!(s&&s.id),l=r?`\u270F\uFE0F Edit Kontrak: ${s.employee_name||"Karyawan"}`:"\u2795 Tambah Kontrak Karyawan Baru",f=s?.employee_id||"",v=s?.branch_id||"",L=s?.division||"FACILITY CARE",A=s?.status||"Aktif",P=s?.start_date?String(s.start_date).slice(0,10):"",w=s?.end_date&&!Fe(s.end_date)?String(s.end_date).slice(0,10):"",I=s?Fe(s.end_date):!1,z=s?.contract_type||"KONTRAK 1 TAHUN",C=s?.pkwt_number||"PKWT 1",y=s?.notes||"",_=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-employee" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Karyawan --</option>
              ${c.map(T=>`
                <option value="${T.id}" ${String(f)===String(T.id)?"selected":""}>
                  ${T.full_name} (${T.division||"FC"} - ${T.branch_name||"Cabang -"})
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${n.map(T=>`
                <option value="${T.value}" ${String(v)===String(T.value)?"selected":""}>
                  ${T.label}
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
              <option value="Aktif" ${A==="Aktif"?"selected":""}>\u{1F7E2} Aktif</option>
              <option value="Tidak Aktif" ${A==="Tidak Aktif"?"selected":""}>\u26AA Tidak Aktif</option>
              <option value="Resign" ${A==="Resign"?"selected":""}>\u{1F534} Resign</option>
              <option value="Cut" ${A==="Cut"?"selected":""}>\u2702\uFE0F Cut / Pemutusan</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tipe / Durasi Kontrak
            </label>
            <select id="con-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="KONTRAK 6 BULAN" ${z==="KONTRAK 6 BULAN"?"selected":""}>KONTRAK 6 BULAN</option>
              <option value="KONTRAK 1 TAHUN" ${z==="KONTRAK 1 TAHUN"?"selected":""}>KONTRAK 1 TAHUN</option>
              <option value="KONTRAK 2 TAHUN" ${z==="KONTRAK 2 TAHUN"?"selected":""}>KONTRAK 2 TAHUN</option>
              <option value="PKWTT (TETAP)" ${z==="PKWTT (TETAP)"?"selected":""}>PKWTT (TETAP)</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tingkat / No. PKWT
            </label>
            <select id="con-modal-pkwt" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="PKWT 1" ${C==="PKWT 1"?"selected":""}>PKWT 1</option>
              <option value="PKWT 2" ${C==="PKWT 2"?"selected":""}>PKWT 2</option>
              <option value="PKWT 3" ${C==="PKWT 3"?"selected":""}>PKWT 3</option>
              <option value="PKWT 4" ${C==="PKWT 4"?"selected":""}>PKWT 4</option>
              <option value="PKWT 5" ${C==="PKWT 5"?"selected":""}>PKWT 5</option>
              <option value="PKWT 6" ${C==="PKWT 6"?"selected":""}>PKWT 6</option>
              <option value="TETAP" ${C==="TETAP"?"selected":""}>PKWTT / TETAP</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Mulai Kontrak <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="con-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${P}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Selesai Kontrak
            </label>
            <input type="date" id="con-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; ${I?"opacity:0.5;":""}" 
              value="${w}" ${I?"disabled":""} />
            <div style="margin-top:6px; display:flex; align-items:center; gap:6px;">
              <input type="checkbox" id="con-modal-is-tetap" ${I?"checked":""} style="cursor:pointer;" />
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
            placeholder="Keterangan perpanjangan, klausul khusus, atau nomor SK...">${y}</textarea>
        </div>
      </div>
    `;G({title:l,content:_,confirmText:r?"\u{1F4BE} Simpan Perubahan":"\u2795 Tambah Kontrak",cancelText:"Batal",onConfirm:async(T,j)=>{let O=T.querySelector("#con-modal-employee"),M=T.querySelector("#con-modal-branch"),q=T.querySelector("#con-modal-division"),H=T.querySelector("#con-modal-status"),Y=T.querySelector("#con-modal-start"),W=T.querySelector("#con-modal-end"),Z=T.querySelector("#con-modal-is-tetap"),U=T.querySelector("#con-modal-type"),V=T.querySelector("#con-modal-pkwt"),ae=T.querySelector("#con-modal-notes"),J=O?.value,ie=M?.value,me=q?.value||"FACILITY CARE",se=H?.value||"Aktif",ce=Y?.value,Va=Z?.checked?"2099-12-31":W?.value||"2099-12-31",Wa=U?.value,Qa=V?.value,Xa=ae?.value?.trim()||null;if(!J){N("Harap pilih nama karyawan!"),O?.focus();return}if(!ie){N("Harap pilih cabang penempatan!"),M?.focus();return}if(!ce){N("Harap tentukan tanggal mulai kontrak!"),Y?.focus();return}let Za={employee_id:parseInt(J,10),branch_id:parseInt(ie,10),division:me,status:se,start_date:ce,end_date:Va,contract_type:Wa,pkwt_number:Qa,notes:Xa};try{let lt=r?`/api/contracts/${s.id}`:"/api/contracts",dt=await K(lt,{method:r?"PUT":"POST",body:JSON.stringify(Za)});if(dt.ok)R(r?"Data kontrak berhasil diperbarui!":"Kontrak baru berhasil ditambahkan!"),j(),await E();else{let en=dt.data?.error||dt.data?.message||"Gagal menyimpan data kontrak.";N(en)}}catch(lt){console.error(lt),N("Terjadi kesalahan koneksi ke server.")}}}),setTimeout(()=>{let T=document.getElementById("con-modal-is-tetap"),j=document.getElementById("con-modal-end");T&&j&&T.addEventListener("change",O=>{O.target.checked?(j.disabled=!0,j.style.opacity="0.5"):(j.disabled=!1,j.style.opacity="1")})},100)}function p(s){G({title:"\u{1F5D1}\uFE0F Hapus Data Kontrak",content:`
        <div style="font-size:0.92rem; color:#334155; line-height:1.5;">
          Apakah Anda yakin ingin menghapus catatan kontrak untuk <strong>${s.employee_name||"karyawan ini"}</strong> (ID #${s.id})?
          <div style="margin-top:10px; padding:10px; background:#FEF2F2; border-radius:8px; border:1px solid #FECACA; font-size:0.8rem; color:#991B1B;">
            \u26A0\uFE0F Tindakan ini akan menghapus riwayat masa berlaku kontrak terkait dari sistem.
          </div>
        </div>
      `,confirmText:"Ya, Hapus",cancelText:"Batal",onConfirm:async(r,l)=>{try{let f=await K(`/api/contracts/${s.id}`,{method:"DELETE"});f.ok?(R("Data kontrak berhasil dihapus."),l(),await E()):N("Gagal menghapus kontrak: "+(f.data?.message||"Terjadi kesalahan"))}catch(f){console.error(f),N("Terjadi kendala koneksi.")}}})}function d(){if(!i.length){N("Belum ada data kontrak untuk diekspor.");return}let s=i.map(r=>{let l=Fe(r.end_date);return{"Nama Lengkap":r.employee_name||"",Cabang:r.branch_name||"",Divisi:r.division||"","No. PKWT":r.pkwt_number||"","Tipe Kontrak":r.contract_type||"","Tanggal Mulai":r.start_date||"","Tanggal Selesai":l?"Tetap / PKWTT":r.end_date||"","Sisa Hari":l?"Tetap":r.days_remaining!==null&&r.days_remaining!==void 0?`${r.days_remaining} Hari`:"",Status:r.status||"",Catatan:r.notes||""}});Q(s,`Data_Kontrak_Karyawan_${new Date().toISOString().slice(0,10)}`),R("File Excel Data Kontrak berhasil diunduh!")}function g(){let s=k(),r=i.length,l=i.filter(j=>j.status==="Aktif").length,f=i.filter(j=>{let O=Fe(j.end_date),M=Number(j.days_remaining);return j.status==="Aktif"&&!O&&!isNaN(M)&&M>=0&&M<=30}).length,v=i.filter(j=>Fe(j.end_date)&&j.status==="Aktif").length,L=i.filter(j=>j.status!=="Aktif").length;e.innerHTML=`
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
              ${$?`
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
                <div class="con-stat-val">${r}</div>
                <div class="con-stat-lbl">Total Kontrak</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#34D399;">
                \u{1F7E2}
              </div>
              <div>
                <div class="con-stat-val" style="color:#34D399;">${l}</div>
                <div class="con-stat-lbl">Kontrak Aktif</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FBBF24;">
                \u26A0\uFE0F
              </div>
              <div>
                <div class="con-stat-val" style="color:#FBBF24;">${f}</div>
                <div class="con-stat-lbl">Habis &lt; 30 Hari</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(99, 102, 241, 0.2); color:#A5B4FC;">
                \u{1F3DB}\uFE0F
              </div>
              <div>
                <div class="con-stat-val" style="color:#A5B4FC;">${v}</div>
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
                placeholder="Cari nama karyawan, cabang, no PKWT, catatan..." value="${m?String(m).replace(/"/g,"&quot;"):""}" />
            </div>

            <select id="con-branch-filter" class="con-select">
              <option value="all">Semua Cabang (${n.length})</option>
              ${n.map(j=>`
                <option value="${j.value}" ${String(u)===String(j.value)?"selected":""}>
                  ${j.label}
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
              <option value="all" ${B==="all"?"selected":""}>Semua Durasi Masa Habis</option>
              <option value="7" ${B==="7"?"selected":""}>Akan Habis \u2264 7 Hari</option>
              <option value="14" ${B==="14"?"selected":""}>Akan Habis \u2264 14 Hari</option>
              <option value="30" ${B==="30"?"selected":""}>Akan Habis \u2264 30 Hari</option>
              <option value="60" ${B==="60"?"selected":""}>Akan Habis \u2264 60 Hari</option>
            </select>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="con-filter-pills">
            <div class="con-pill ${b==="all"&&B==="all"?"active":""}" data-status="all">
              <span>Semua Kontrak</span>
              <span class="con-pill-count">${r}</span>
            </div>
            <div class="con-pill ${b==="Aktif"&&B==="all"?"active":""}" data-status="Aktif">
              <span class="con-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="con-pill-count">${l}</span>
            </div>
            <div class="con-pill ${b==="Expiring30"?"active":""}" data-status="Expiring30">
              <span class="con-pill-dot" style="background:#F59E0B;"></span>
              <span>Akan Habis (&lt;30 Hari)</span>
              <span class="con-pill-count">${f}</span>
            </div>
            <div class="con-pill ${b==="Tetap"?"active":""}" data-status="Tetap">
              <span class="con-pill-dot" style="background:#6366F1;"></span>
              <span>PKWTT (Tetap)</span>
              <span class="con-pill-count">${v}</span>
            </div>
            <div class="con-pill ${b==="Nonaktif"?"active":""}" data-status="Nonaktif">
              <span class="con-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="con-pill-count">${L}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive Cards Grid -->
        <div class="con-cards-grid">
          ${s.length===0?`
            <div class="con-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada kontrak yang sesuai kriteria</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba atur ulang kata kunci pencarian, filter cabang, atau filter masa berlaku kontrak.</p>
              <button id="con-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
          `:s.map(j=>{let O=j.division||"FACILITY CARE",M="con-div-fc",q="";O==="SECURITY"?(M="con-div-sec",q="sec"):O==="FC - RELIEFER"&&(M="con-div-rel",q="rel");let H=Fe(j.end_date),Y=j.status==="Aktif",W=Number(j.days_remaining),Z="con-status-aktif",U="\u{1F7E2} Aktif";H?(Z="con-status-tetap",U="\u{1F3DB}\uFE0F PKWTT (Tetap)"):Y?W<=0?(Z="con-status-danger",U="\u26A0\uFE0F Habis Hari Ini"):W<=30?(Z="con-status-warning",U=`\u26A0\uFE0F Sisa ${W} Hari`):(Z="con-status-aktif",U=`\u{1F7E2} Aktif (${W} Hari)`):(Z="con-status-danger",U=`\u26AA ${j.status||"Nonaktif"}`);let V=x[j.employee_id]||x[String(j.employee_name||"").toLowerCase()]||"";return`
              <div class="con-card" data-id="${j.id}">
                <!-- Top Badges Row -->
                <div class="con-card-top">
                  <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
                    <span class="con-div-badge ${M}">
                      ${O==="SECURITY"?"\u{1F6E1}\uFE0F":O==="FC - RELIEFER"?"\u{1F504}":"\u{1F9F9}"} ${O}
                    </span>
                    ${j.pkwt_number?`
                      <span class="con-pkwt-badge">
                        \u{1F4DC} ${j.pkwt_number}
                      </span>
                    `:""}
                  </div>
                  <span class="con-status-badge ${Z}">
                    ${U}
                  </span>
                </div>

                <!-- Profile Box -->
                <div class="con-profile-box">
                  <div class="con-avatar ${q}">
                    ${Sn(j.employee_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="con-name">
                      ${j.employee_name||"Karyawan Tanpa Nama"}
                    </div>
                    <div class="con-subtitle">
                      Personil ${O} \u2022 ${j.contract_type||"Kontrak Kerja"}
                    </div>
                  </div>
                </div>

                <!-- Information Grid -->
                <div class="con-info-card">
                  <div class="con-info-row">
                    <span class="con-info-label">\u{1F4CD} Cabang:</span>
                    <span class="con-info-val">${j.branch_name||"Cabang Tidak Ditentukan"}</span>
                  </div>
                  <div class="con-info-row">
                    <span class="con-info-label">\u{1F5D3}\uFE0F Mulai Kontrak:</span>
                    <span class="con-info-val">${ft(j.start_date)}</span>
                  </div>
                  <div class="con-info-row">
                    <span class="con-info-label">\u{1F3C1} Selesai Kontrak:</span>
                    <span class="con-info-val" style="${H?"color:#4338CA; font-weight:700;":""}">
                      ${H?"Tetap / PKWTT":ft(j.end_date)}
                    </span>
                  </div>
                  <div class="con-info-row">
                    <span class="con-info-label">\u23F3 Masa Berlaku:</span>
                    <span class="con-info-val" style="${H?"color:#15803D;":W<=30&&Y?"color:#D97706; font-weight:700;":""}">
                      ${H?"Karyawan Tetap (Permanen)":W<=0?"Masa Kontrak Telah Habis":`${W} Hari Tersisa`}
                    </span>
                  </div>
                </div>

                <!-- Notes box (if any) -->
                ${j.notes?`
                  <div class="con-notes-box">
                    \u{1F4DD} <strong>Catatan:</strong> ${j.notes}
                  </div>
                `:""}

                <!-- Footer Actions -->
                <div class="con-card-footer">
                  <div style="font-size:0.75rem; color:#94A3B8;">
                    ID #${j.id} \u2022 ${j.pkwt_number||"PKWT"}
                  </div>

                  <div class="con-action-btns">
                    ${V?`
                      <a href="https://wa.me/${V.replace(/[^0-9]/g,"").replace(/^0/,"62")}" target="_blank" rel="noopener" class="con-icon-btn" title="Chat WhatsApp Karyawan" style="color:#16A34A;">
                        <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                      </a>
                    `:""}

                    ${$?`
                      <button class="con-icon-btn edit" data-action="edit" data-id="${j.id}" title="Edit Data Kontrak">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                      <button class="con-icon-btn delete" data-action="delete" data-id="${j.id}" title="Hapus Kontrak">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                      </button>
                    `:""}
                  </div>
                </div>
              </div>
            `}).join("")}
        </div>

      </div>
    `;let A=e.querySelector("#con-search-input");A&&A.addEventListener("input",j=>{m=j.target.value,g()});let P=e.querySelector("#con-branch-filter");P&&P.addEventListener("change",j=>{u=j.target.value,g()});let w=e.querySelector("#con-division-filter");w&&w.addEventListener("change",j=>{o=j.target.value,g()});let I=e.querySelector("#con-expiring-filter");I&&I.addEventListener("change",j=>{B=j.target.value,g()}),e.querySelectorAll(".con-pill").forEach(j=>{j.addEventListener("click",O=>{let M=O.currentTarget.getAttribute("data-status");b=M,M==="Expiring30"?B="30":B="all",g()})});let z=e.querySelector("#con-reset-filter");z&&z.addEventListener("click",()=>{m="",u="all",o="all",b="all",B="all",g()});let C=e.querySelector("#con-btn-add");C&&C.addEventListener("click",()=>a());let y=e.querySelector("#con-btn-excel");y&&y.addEventListener("click",d);let _=e.querySelector("#con-btn-missing");_&&_.addEventListener("click",h);let T=e.querySelector("#con-btn-refresh");T&&T.addEventListener("click",()=>E()),e.querySelectorAll(".con-icon-btn.edit").forEach(j=>{j.addEventListener("click",O=>{let M=parseInt(O.currentTarget.getAttribute("data-id"),10),q=i.find(H=>H.id===M);q&&a(q)})}),e.querySelectorAll(".con-icon-btn.delete").forEach(j=>{j.addEventListener("click",O=>{let M=parseInt(O.currentTarget.getAttribute("data-id"),10),q=i.find(H=>H.id===M);q&&p(q)})})}await E()}re();var xt=[],Oe=[];function Cn(e){if(!Array.isArray(e))return"Q3";let t=["Q4","Q3","Q2","Q1"];for(let i of t)if(e.some(n=>n.period===i))return i;return"Q3"}async function ea(e,t){xt=await te();let i=await ue();Oe=["Berlin Ariansyah","Ade Surahman"];let n=k=>k&&!Oe.find(h=>String(typeof h=="object"?h.value:h).toLowerCase()===String(k).toLowerCase())?[...Oe,k]:Oe,c=await K(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),x=k=>{if(!k||k==="-"||String(k).trim()==="")return"";let h=String(k).split("-");return h.length===3&&h[0].length===4?`${h[2]}-${h[1]}-${h[0]}`:k},m=c.data?.data||[],u=Cn(m),o=t?t.get("dash_filter"):null,b=new Date,B=`${b.getFullYear()}-${String(b.getMonth()+1).padStart(2,"0")}`,D={},S=t&&t.get("month")?t.get("month"):null;o==="inspeksi"?D={status:"Done",activity_type:"Inspeksi Hygiene",month:S}:o==="gcdc"?D={status:"Done",activity_type:"GCDC",month:S}:o&&o.startsWith("period_")&&(D={period:o.replace("period_","").toUpperCase()});let F=new Date().getFullYear(),E=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((k,h)=>{let a=String(h+1).padStart(2,"0");return{value:`${F}-${a}`,label:`${k} ${F}`}});de({container:e,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:D,onDataLoaded:k=>k.sort((h,a)=>{let p=h.opening_date?new Date(h.opening_date).getTime():0;return(a.opening_date?new Date(a.opening_date).getTime():0)-p}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:k=>mt(k)},{key:"period",label:"Periode",render:k=>xe(k)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:k=>x(k)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:k=>x(k)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:k=>x(k)},{key:"status",label:"Status",render:k=>ge(k)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:xt},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:E},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Oe}],formFields:k=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:xt,value:k?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:k?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:k?.period},{name:"pic",label:"PIC",type:"combobox",options:n(k?.pic),value:k?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:k?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:k?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:k?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:k?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:k?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let k=await K(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(k.ok){let h=k.data.data.map(a=>({Cabang:a.branch_name||"",Kegiatan:a.activity_type||"",Periode:a.period||"",PIC:a.pic||"","Tgl Opening":a.opening_date||"","Tgl Target":a.target_date||"","Tgl Selesai":a.completion_date||"",Status:a.status||""}));Q(h,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Q([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async k=>{let a=(await K("/api/branches?all=1")).data?.data||[],p=r=>{if(!r)return null;let l=String(r||"").toLowerCase().trim(),f=l.match(/^(\d{3})/);if(f){let L=a.find(A=>String(A.code||"")===f[1]);if(L)return L.id}let v=a.find(L=>String(L.full_name||"").toLowerCase()===l||String(L.code||"").toLowerCase()===l||String(L.name||"").toLowerCase()===l);return v?v.id:null},d=r=>{if(r==null||r==="")return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let l=String(r).trim();if(l===""||l==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);if(/^\d{4,5}$/.test(l)){let v=Number(l);if(v>2e4&&v<99999){let L=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(L.getTime())?"":L.toISOString().slice(0,10)}}let f=l.split(/[\/\-\.]/);if(f.length===3){let[v,L,A]=f.map(P=>P.trim());if(v.length===4&&L.length<=2&&A.length<=2)return`${v}-${L.padStart(2,"0")}-${A.padStart(2,"0")}`;if(A.length===4&&L.length<=2&&v.length<=2)return`${A}-${L.padStart(2,"0")}-${v.padStart(2,"0")}`}return l},g=k.map(r=>({branch_id:p(String(r.Cabang||"").trim()),branch_name:String(r.Cabang||"").trim(),activity_type:String(r.Kegiatan||"").trim(),period:String(r.Periode||"").trim(),pic:String(r.PIC||r.Pic||"").trim(),opening_date:d(r["Tgl Opening"]||r["Tanggal Opening"]||r["Tgl Openir"]),target_date:d(r["Tgl Target"]||r["Tanggal Target"]),completion_date:d(r["Tgl Selesai"]||r["Tanggal Selesai"]),status:String(r.Status||"").trim(),notes:String(r.Catatan||r.Keterangan||"").trim()})).filter(r=>r.activity_type&&r.period),s=await K("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:g,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return be("schedule"),s.data}}})}le();var yt=[{id:1,report_date:"2026-08-27",branch_id:19,branch_name:"019. Cinere",category:"SDM",source:"AM",complaint:"Keluhan AM fc atas nama Malik bau badan, dan sudah dibilangin berulang kali terkait kebersihan seragam dan deodoran.",employee_name:"MALIK",fc_specialist:"Mizwar",solution:"Diberikan teguran lisan, edukasi kebersihan diri, dan pendampingan seragam cadangan di klinik.",status:"Done",completion_date:"2026-08-28",day_count:1},{id:2,report_date:"2026-08-26",branch_id:0,branch_name:"000. Management",category:"SDM",source:"Berlin",complaint:"Security atas nama Reza suhelmi meninggalkan area tugas jaga tanpa personil pengganti saat jam operasional klinik.",employee_name:"REZA SUHELMI",fc_specialist:"Ade Surahman",solution:"Pemanggilan investigasi dan penerbitan Surat Peringatan (SP) sesuai SOP security.",status:"Done",completion_date:"2026-08-27",day_count:1},{id:3,report_date:"2026-07-27",branch_id:30,branch_name:"030. Ciputat",category:"SDM",source:"AM",complaint:"Terinfo kembali permasalahan FC Ciputat, terkait standar pembersihan area dental unit yang belum memenuhi checklist hygiene.",employee_name:"HENDRA",fc_specialist:"Berlin Ariansyah",solution:"Re-training modul pembersihan dental unit dan monitoring harian oleh supervisor.",status:"In Progress",completion_date:null,day_count:null},{id:4,report_date:"2026-06-24",branch_id:30,branch_name:"030. Ciputat",category:"SDM",source:"AM",complaint:"FC atas nama Hendra mengintimidasi rcp pada saat briefing pagi mengenai pembagian area kerja.",employee_name:"HENDRA",fc_specialist:"Berlin Ariansyah",solution:"Mediasi bersama AM dan HR, diberikan pembinaan konseling perilaku kerja.",status:"Done",completion_date:"2026-06-25",day_count:1},{id:5,report_date:"2026-06-05",branch_id:0,branch_name:"000. Management",category:"SDM",source:"Berlin",complaint:"Kinerja kurang dan ada komplain dari cabang mengenai ketidakhadiran berkala personil pengganti.",employee_name:"Indra",fc_specialist:"Ade",solution:"Evaluasi penempatan dan pembinaan kedisiplinan kerja.",status:"Done",completion_date:"2026-06-08",day_count:3},{id:6,report_date:"2026-06-05",branch_id:0,branch_name:"000. Management",category:"SDM",source:"Ade",complaint:"Sering tidak masuk kerja dan memalsukan surat dokter saat izin sakit.",employee_name:"Kristian",fc_specialist:"Ade Surahman",solution:"Investigasi ke klinik dokter terkait dan eskalasi ke HRD untuk tindakan tegas.",status:"Done",completion_date:"2026-06-07",day_count:2},{id:7,report_date:"2026-05-25",branch_id:19,branch_name:"019. Cinere",category:"SDM",source:"SPV",complaint:"Tidak masuk kerja tanpa kabar (alpa), dan sudah di-one on one oleh SPV terkait absensi.",employee_name:"Malik",fc_specialist:"Mizwar",solution:"Pemberian SP 1 dan penandatanganan komitmen kehadiran kerja.",status:"Done",completion_date:"2026-05-26",day_count:1}],ta=["SDM","Cleaning","Aset","K3","Lainnya"],Fn=["Open","In Progress","Done"],$n=["SPV","AM","RCP","Perawat","FC","Security","Berlin","Ade","Pattrel","Dentrel"];function aa(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function _n(e){if(!e||e==="-")return"\u26A0\uFE0F";let t=String(e).trim().split(/\s+/);return t.length===1?t[0].slice(0,2).toUpperCase():(t[0][0]+t[1][0]).toUpperCase()}async function na(e,t){let i=[],n=[],c=[],x="all",m="all",u="all",o="",b=!1;(t?t.get("dash_filter"):null)==="open"&&(x="Open");let S=ee()?.role!=="viewer";try{[n,c]=await Promise.all([te(),ue()])}catch(d){console.error("Error fetching cache:",d)}async function F(){b=!0,p();try{let d=await K("/api/issues?limit=10000");if(d.ok){let g=d.data?.data||d.data||[];i=Array.isArray(g)&&g.length>0?g:yt}else i=yt}catch(d){console.warn("Gagal load issues, menggunakan data lokal:",d),i=yt}finally{b=!1,p()}}async function $(){try{let{downloadExcel:d}=await Promise.resolve().then(()=>(re(),he)),g=i.map((s,r)=>({No:r+1,"Tanggal Info":s.report_date||"",Cabang:s.branch_name||"",Kategori:s.category||"","Sumber Laporan":s.source||"",Keluhan:s.complaint||"","Nama FC / Security":s.employee_name||"","FC Spesialis":s.fc_specialist||"","Solusi / Tindak Lanjut":s.solution||"",Status:s.status||"","Tanggal Selesai":s.completion_date||"","Durasi Hari":s.day_count!=null?s.day_count:""}));d(g,`Data_Permasalahan_FCMS_${new Date().toISOString().split("T")[0]}`),R("Berhasil mengekspor data permasalahan ke Excel")}catch(d){N("Gagal ekspor Excel: "+d.message)}}async function E(d,g){try{let s=g==="Done"?d.completion_date||new Date().toISOString().split("T")[0]:g==="Open"?null:d.completion_date,r=d.day_count;if(d.report_date&&s){let v=new Date(d.report_date),L=new Date(s);r=Math.max(0,Math.floor((L-v)/(1e3*60*60*24)))}else g==="Open"&&(r=null);let l={...d,status:g,completion_date:s,day_count:r},f=await K(`/api/issues/${d.id}`,{method:"PUT",body:JSON.stringify(l)});f.ok?(R(`Status masalah diubah ke "${g}"`),await F()):N(f.data?.error||"Gagal memperbarui status")}catch(s){N("Terjadi kesalahan: "+s.message)}}function k(d=null){let g=!!d,s=new Date().toISOString().split("T")[0],r=d?.report_date||s,l=d?.category||"SDM",f=d?.source||"AM",v=d?.status||"Open",L=document.createElement("div");L.className="issue-modal-body",L.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Info <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="iss-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${r}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${n.map(C=>`
                <option value="${C.value}" ${String(d?.branch_id)===String(C.value)?"selected":""}>
                  ${C.label}
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
              ${ta.map(C=>`
                <option value="${C}" ${l===C?"selected":""}>${C}</option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Sumber Laporan
            </label>
            <input list="iss-source-options" id="iss-modal-source" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: AM, SPV, Berlin..." value="${f}" />
            <datalist id="iss-source-options">
              ${$n.map(C=>`<option value="${C}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Detail Keluhan / Permasalahan <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="iss-modal-complaint" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Jelaskan kendala, kronologi, atau komplain..." required>${d?.complaint||""}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Personil FC / Security
            </label>
            <input list="iss-emp-options" id="iss-modal-employee" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Ketik nama karyawan terkait..." value="${d?.employee_name?String(d.employee_name).replace(/"/g,"&quot;"):""}" />
            <datalist id="iss-emp-options">
              ${c.map(C=>`<option value="${C.label}">`).join("")}
            </datalist>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              FC Spesialis / Penanggung Jawab
            </label>
            <input list="iss-spec-options" id="iss-modal-spec" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: Mizwar, Ade, Berlin..." value="${d?.fc_specialist?String(d.fc_specialist).replace(/"/g,"&quot;"):""}" />
            <datalist id="iss-spec-options">
              ${c.map(C=>`<option value="${C.label}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Solusi / Tindak Lanjut Penanganan
          </label>
          <textarea id="iss-modal-solution" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tindakan yang telah atau akan diambil untuk menyelesaikan masalah...">${d?.solution||""}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Masalah <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${Fn.map(C=>`
                <option value="${C}" ${v===C?"selected":""}>${C}</option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Selesai
            </label>
            <input type="date" id="iss-modal-completion" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${d?.completion_date||""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Durasi (Hari)
            </label>
            <input type="number" id="iss-modal-daycount" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; background:#F8FAFC;" 
              value="${d?.day_count!=null?d.day_count:""}" readonly placeholder="Otomatis" />
          </div>
        </div>
      </div>
    `;let A=L.querySelector("#iss-modal-date"),P=L.querySelector("#iss-modal-completion"),w=L.querySelector("#iss-modal-daycount"),I=L.querySelector("#iss-modal-status");function z(){if(A.value&&P.value){let C=new Date(A.value),y=new Date(P.value),_=Math.floor((y-C)/(1e3*60*60*24));w.value=isNaN(_)?"":Math.max(0,_)}else w.value=""}A.addEventListener("change",z),P.addEventListener("change",z),I.addEventListener("change",()=>{I.value==="Done"&&!P.value?(P.value=new Date().toISOString().split("T")[0],z()):I.value==="Open"&&(P.value="",w.value="")}),G({title:g?"\u270F\uFE0F Edit Data Permasalahan":"\u2795 Tambah Data Permasalahan",content:L,size:"md",confirmText:g?"Simpan Perubahan":"Simpan Permasalahan",confirmClass:"btn-primary",onConfirm:async(C,y)=>{let _=L.querySelector("#iss-modal-date").value,T=L.querySelector("#iss-modal-branch").value,j=L.querySelector("#iss-modal-category").value,O=(L.querySelector("#iss-modal-source").value||"").trim(),M=(L.querySelector("#iss-modal-complaint").value||"").trim(),q=(L.querySelector("#iss-modal-employee").value||"").trim(),H=(L.querySelector("#iss-modal-spec").value||"").trim(),Y=(L.querySelector("#iss-modal-solution").value||"").trim(),W=L.querySelector("#iss-modal-status").value,Z=L.querySelector("#iss-modal-completion").value||null;if(!_||!T||!M||!j){N("Harap lengkapi Tanggal, Cabang, Kategori, dan Keluhan!");return}let U=null;if(_&&Z){let ae=new Date(_),J=new Date(Z);U=Math.max(0,Math.floor((J-ae)/(1e3*60*60*24)))}let V={report_date:_,branch_id:Number(T),category:j,source:O,complaint:M,employee_name:q,fc_specialist:H,solution:Y,status:W,completion_date:Z,day_count:U};try{let ae;g?ae=await K(`/api/issues/${d.id}`,{method:"PUT",body:JSON.stringify(V)}):ae=await K("/api/issues",{method:"POST",body:JSON.stringify(V)}),ae.ok?(R(g?"Permasalahan berhasil diperbarui!":"Permasalahan baru berhasil ditambahkan!"),y(),await F()):N(ae.data?.error||"Gagal menyimpan permasalahan")}catch(ae){N("Terjadi kesalahan: "+ae.message)}}})}async function h(d){X({title:"Hapus Permasalahan?",message:`Apakah Anda yakin ingin menghapus tiket permasalahan dari <strong>${d.branch_name||"cabang"}</strong>:<br><br><em>"${d.complaint}"</em>?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let g=await K(`/api/issues/${d.id}`,{method:"DELETE"});g.ok?(R("Permasalahan berhasil dihapus"),await F()):N(g.data?.error||"Gagal menghapus data")}catch(g){N("Terjadi kesalahan: "+g.message)}}})}function a(){return i.filter(d=>{if(o){let g=o.toLowerCase(),s=(d.complaint||"").toLowerCase().includes(g),r=(d.employee_name||"").toLowerCase().includes(g),l=(d.branch_name||"").toLowerCase().includes(g),f=(d.source||"").toLowerCase().includes(g),v=(d.category||"").toLowerCase().includes(g),L=(d.solution||"").toLowerCase().includes(g);if(!s&&!r&&!l&&!f&&!v&&!L)return!1}return!(x!=="all"&&d.status!==x||m!=="all"&&d.category!==m||u!=="all"&&String(d.branch_id)!==String(u))})}function p(){let d=a(),g=i.length,s=i.filter(M=>M.status==="Open"||M.status==="In Progress").length,r=i.filter(M=>M.status==="Done").length,l=new Date,f=l.getFullYear(),v=l.getMonth(),L=i.filter(M=>{if(!M.report_date)return!1;let q=new Date(M.report_date);return!isNaN(q.getTime())&&q.getFullYear()===f&&q.getMonth()===v}).length,A=i.length,P=i.filter(M=>M.status==="Open").length,w=i.filter(M=>M.status==="In Progress").length,I=i.filter(M=>M.status==="Done").length;e.innerHTML=`
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
              ${S?`
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
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${b?"spin-icon":""}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="issues-stats-row">
            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(239, 68, 68, 0.2); color:#FCA5A5;">\u26A0\uFE0F</div>
              <div>
                <div class="issues-stat-val">${g}</div>
                <div class="issues-stat-lbl">Total Permasalahan</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FCD34D;">\u23F3</div>
              <div>
                <div class="issues-stat-val">${s}</div>
                <div class="issues-stat-lbl">Open & In Progress</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">\u2705</div>
              <div>
                <div class="issues-stat-val">${r}</div>
                <div class="issues-stat-lbl">Tuntas (Done)</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(59, 130, 246, 0.2); color:#93C5FD;">\u{1F4C5}</div>
              <div>
                <div class="issues-stat-val">${L}</div>
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
              <option value="all">Semua Cabang (${n.length})</option>
              ${n.map(M=>`
                <option value="${M.value}" ${String(u)===String(M.value)?"selected":""}>
                  ${M.label}
                </option>
              `).join("")}
            </select>

            <select id="issues-category-filter" class="issues-select">
              <option value="all">Semua Kategori</option>
              ${ta.map(M=>`
                <option value="${M}" ${m===M?"selected":""}>
                  Kategori: ${M}
                </option>
              `).join("")}
            </select>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="issues-filter-pills">
            <div class="issues-pill ${x==="all"?"active":""}" data-status="all">
              <span>Semua Status</span>
              <span class="issues-pill-count">${A}</span>
            </div>
            <div class="issues-pill ${x==="Open"?"active":""}" data-status="Open">
              <span class="issues-pill-dot" style="background:#F59E0B;"></span>
              <span>Open</span>
              <span class="issues-pill-count">${P}</span>
            </div>
            <div class="issues-pill ${x==="In Progress"?"active":""}" data-status="In Progress">
              <span class="issues-pill-dot" style="background:#3B82F6;"></span>
              <span>In Progress</span>
              <span class="issues-pill-count">${w}</span>
            </div>
            <div class="issues-pill ${x==="Done"?"active":""}" data-status="Done">
              <span class="issues-pill-dot" style="background:#10B981;"></span>
              <span>Done</span>
              <span class="issues-pill-count">${I}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive Cards Grid -->
        <div class="issues-cards-grid">
          ${d.length===0?`
            <div class="issues-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">\u{1F50D}</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada permasalahan yang cocok</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian atau filter status yang dipilih.</p>
              <button id="issues-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
          `:d.map(M=>{let q=M.category||"Lainnya",H=`issues-category-${q.toLowerCase().replace(/[^a-z0-9]/g,"")}`||"issues-category-lainnya",Y=M.status||"Open",W="issues-status-open",Z="\u26A0\uFE0F Open";Y==="In Progress"?(W="issues-status-inprogress",Z="\u{1F504} In Progress"):Y==="Done"&&(W="issues-status-done",Z="\u2705 Done");let U=Y==="Done"?"status-done":Y==="In Progress"?"status-inprogress":"";return`
              <div class="issues-card" data-id="${M.id}">
                <!-- Top Badges Row -->
                <div class="issues-card-top">
                  <span class="issues-category-badge ${H}">
                    \u{1F3F7}\uFE0F ${q}
                  </span>
                  <div style="display:flex; align-items:center; gap:8px;">
                    <span class="issues-status-badge ${W}">
                      ${Z}
                    </span>
                    <span style="font-size:0.8rem; color:#64748B; display:flex; align-items:center; gap:4px;">
                      \u{1F5D3}\uFE0F ${aa(M.report_date)}
                    </span>
                  </div>
                </div>

                <!-- Complaint Box -->
                <div class="issues-complaint-box ${U}">
                  "${M.complaint||"-"}"
                </div>

                <!-- Person & Location Box -->
                <div class="issues-person-box">
                  <div class="issues-avatar-box">
                    ${_n(M.employee_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="issues-person-sub">PERSONIL FC / SECURITY</div>
                    <div class="issues-person-name" style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${M.employee_name||"Tidak ada personil spesifik"}
                    </div>
                  </div>
                </div>

                <!-- Details Grid -->
                <div class="issues-details-grid">
                  <div class="issues-detail-item" style="grid-column: 1 / -1;">
                    <span style="color:#64748B;">\u{1F4CD} Cabang:</span>
                    <strong style="color:#0F172A;">${M.branch_name||"Cabang Tidak Diketahui"}</strong>
                  </div>
                  <div class="issues-detail-item">
                    <span style="color:#64748B;">\u{1F5E3}\uFE0F Sumber:</span>
                    <strong>${M.source||"-"}</strong>
                  </div>
                  <div class="issues-detail-item">
                    <span style="color:#64748B;">\u{1F6E1}\uFE0F Spesialis:</span>
                    <strong>${M.fc_specialist||"-"}</strong>
                  </div>
                </div>

                <!-- Solution & Completion Box (if available) -->
                ${M.solution||M.completion_date?`
                  <div class="issues-solution-box">
                    ${M.solution?`
                      <div class="issues-solution-title">
                        <span>\u{1F4A1} Tindak Lanjut / Solusi:</span>
                      </div>
                      <div style="margin-bottom:${M.completion_date?"6px":"0"};">
                        ${M.solution}
                      </div>
                    `:""}
                    ${M.completion_date?`
                      <div style="font-size:0.76rem; color:#15803D; font-weight:600; display:flex; align-items:center; gap:4px;">
                        <span>\u23F1\uFE0F Selesai: ${aa(M.completion_date)}</span>
                        ${M.day_count!=null?`<span>(Durasi: ${M.day_count} Hari)</span>`:""}
                      </div>
                    `:""}
                  </div>
                `:""}

                <!-- Footer Actions -->
                <div class="issues-card-footer">
                  <div>
                    ${S?Y==="Open"?`
                        <button class="issues-quick-btn start" data-action="quick-progress" data-id="${M.id}" title="Ubah status ke In Progress">
                          <span>\u{1F504} Mulai Tangani</span>
                        </button>
                      `:Y==="In Progress"?`
                        <button class="issues-quick-btn complete" data-action="quick-done" data-id="${M.id}" title="Tandai selesai">
                          <span>\u2705 Selesaikan</span>
                        </button>
                      `:`
                        <button class="issues-quick-btn reopen" data-action="quick-reopen" data-id="${M.id}" title="Buka kembali tiket masalah">
                          <span>\u21A9\uFE0F Buka Kembali</span>
                        </button>
                      `:""}
                  </div>

                  <div class="issues-action-btns">
                    ${S?`
                      <button class="issues-icon-btn edit" data-action="edit" data-id="${M.id}" title="Edit Masalah">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                      <button class="issues-icon-btn delete" data-action="delete" data-id="${M.id}" title="Hapus Masalah">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                      </button>
                    `:""}
                  </div>
                </div>
              </div>
            `}).join("")}
        </div>

      </div>
    `;let z=e.querySelector("#issues-search-input");z&&z.addEventListener("input",M=>{o=M.target.value,p();let q=e.querySelector("#issues-search-input");q&&(q.focus(),q.setSelectionRange(q.value.length,q.value.length))});let C=e.querySelector("#issues-branch-filter");C&&C.addEventListener("change",M=>{u=M.target.value,p()});let y=e.querySelector("#issues-category-filter");y&&y.addEventListener("change",M=>{m=M.target.value,p()}),e.querySelectorAll(".issues-pill").forEach(M=>{M.addEventListener("click",()=>{x=M.dataset.status,p()})});let _=e.querySelector("#issues-reset-filter");_&&_.addEventListener("click",()=>{o="",u="all",m="all",x="all",p()});let T=e.querySelector("#issues-btn-add");T&&T.addEventListener("click",()=>k());let j=e.querySelector("#issues-btn-excel");j&&j.addEventListener("click",()=>$());let O=e.querySelector("#issues-btn-refresh");O&&O.addEventListener("click",()=>F()),e.querySelectorAll(".issues-card").forEach(M=>{let q=Number(M.dataset.id),H=i.find(ae=>ae.id===q);if(!H)return;let Y=M.querySelector('[data-action="edit"]');Y&&Y.addEventListener("click",()=>k(H));let W=M.querySelector('[data-action="delete"]');W&&W.addEventListener("click",()=>h(H));let Z=M.querySelector('[data-action="quick-progress"]');Z&&Z.addEventListener("click",()=>E(H,"In Progress"));let U=M.querySelector('[data-action="quick-done"]');U&&U.addEventListener("click",()=>E(H,"Done"));let V=M.querySelector('[data-action="quick-reopen"]');V&&V.addEventListener("click",()=>E(H,"Open"))})}await F()}var Ie=[];function An(e,t){let i=String(e.status||"").toLowerCase();return t==="pending"?i==="pending":!1}async function ia(e,t){let i=t?t.get("dash_filter"):null;Ie=await te();let n=await ue(),c=["Ade","Berlin"],x=u=>u&&!n.find(o=>o.value===u)?[...n,{value:u,label:u}]:n,m=u=>u&&!c.find(o=>(typeof o=="object"?o.value:o)===u)?[...c,u]:c;de({container:e,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:u=>i?u.filter(o=>An(o,i)):u,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:u=>window.formatDate(u)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:u=>`<span title="${u||""}">${u?.length>50?u.slice(0,50)+"\u2026":u||"-"}</span>`},{key:"solution",label:"Solusi",render:u=>`<span title="${u||""}">${u?.length>40?u.slice(0,40)+"\u2026":u||"-"}</span>`},{key:"status",label:"Status",render:u=>ge(u)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:u=>window.formatDate(u)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:u=>u?`<a href="${u}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:Ie},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async u=>{let o=new URLSearchParams(u||{}).toString(),b=await K(`/api/one-on-one?limit=10000&${o}`);if(b.ok){let B=b.data.data.map(S=>({Tanggal:S.meeting_date||"",Cabang:S.branch_name||"","Nama Karyawan":S.employee_name||"",PIC:S.pic||"",Masalah:S.problem||"",Solusi:S.solution||"",Status:S.status||"","Tgl Selesai":S.completion_date||"",Dokumen:S.document_link||""})),{downloadExcel:D}=await Promise.resolve().then(()=>(re(),he));D(B,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let u=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:o}=await Promise.resolve().then(()=>(re(),he));o(u,"Template_Import_OneOnOne")},onImport:async u=>{let o=S=>{if(!S)return null;let F=String(S||"").toLowerCase(),$=Ie.find(E=>String(E.label||"").toLowerCase()===F);return $?$.value:null},b=S=>{if(!S)return"";if(S instanceof Date&&!isNaN(S.getTime()))return S.toISOString().slice(0,10);let F=String(S).trim();if(/^\d{4,5}$/.test(F)){let E=Number(F);if(E>2e4&&E<99999){let k=new Date(Date.UTC(1899,11,30)+E*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(F))return F.slice(0,10);let $=F.split(/[\/\-\.]/);if($.length===3){let[E,k,h]=$.map(a=>a.trim());if(E.length===4&&k.length<=2&&h.length<=2)return`${E}-${k.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&k.length<=2&&E.length<=2)return`${h}-${k.padStart(2,"0")}-${E.padStart(2,"0")}`}return F},B=u.map(S=>({meeting_date:b(S.Tanggal),employee_name:String(S["Nama Karyawan"]||"").trim(),branch_id:o(String(S.Cabang||"").trim()),pic:String(S.PIC||"").trim(),problem:String(S.Masalah||"").trim(),solution:String(S.Solusi||"").trim(),status:String(S.Status||"").trim(),completion_date:b(S["Tgl Selesai"]),document_link:String(S.Dokumen||"").trim()})).filter(S=>S.meeting_date&&S.employee_name&&S.branch_id),D=await K("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:B,onDuplicate:"update"})});if(!D.ok)throw new Error(D.data?.error||"Import gagal");return D.data}},formFields:u=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:u?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:u?.branch_id&&!Ie.find(o=>o.value==u.branch_id)?[...Ie,{value:u.branch_id,label:u.branch_name||u.branch_id}]:Ie,createApi:{path:"/api/branches",field:"full_name"},value:u?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:x(u?.employee_name),value:u?.employee_name},{name:"pic",label:"PIC",type:"select",options:m(u?.pic),createApi:{path:"/api/pic",field:"name"},value:u?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:u?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:u?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:u?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:u?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:u?.document_link}]})}le();var vt=[{id:1,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:6,branch_name:"006. Kelapa Gading MOI",trainer:"Berlin Ariansyah",participants:"MUHAMMAD PRAKAS AGAZI",score:90,document_link:""},{id:2,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:11,branch_name:"011. Kuningan",trainer:"Berlin Ariansyah",participants:"KAMALLUDIN",score:88,document_link:""},{id:3,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:42,branch_name:"042. Muara Karang",trainer:"Berlin Ariansyah",participants:"SAHRUL ICHSANI",score:85,document_link:""},{id:4,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:51,branch_name:"051. Bintaro Sektor 9",trainer:"Berlin Ariansyah",participants:"MUHAMAD VICKRI FATHURROHMAN",score:92,document_link:""},{id:5,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:45,branch_name:"045. FX Mall",trainer:"Berlin Ariansyah",participants:"DEDE AHDORI",score:86,document_link:""},{id:6,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:54,branch_name:"054. Bogor Gunung Batu",trainer:"Berlin Ariansyah",participants:"ADITYA PRATAMA",score:95,document_link:""},{id:7,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:55,branch_name:"055. Gajah Mada",trainer:"Berlin Ariansyah",participants:"ELVAN JANGGA PRAMANA",score:89,document_link:""}],ra=["Batch 5","Batch 4","Batch 3","Batch 2","Batch 1"],oa=["Berlin Ariansyah","Ade Surahman","Mizwar","Fajar","Ade","Berlin"];function Bn(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Ae(e){if(!e)return"-";try{let t=JSON.parse(e);if(Array.isArray(t))return t.join(", ")}catch{}return String(e)}async function sa(e){let t=[],i=[],n=[],c="all",x="all",m="all",u="",o=!1,B=ee()?.role!=="viewer";try{[i,n]=await Promise.all([te(),ue()])}catch(a){console.error("Error fetching cache:",a)}async function D(){o=!0,k();try{let a=await K("/api/training?limit=10000");if(a.ok){let p=a.data?.data||a.data||[];t=Array.isArray(p)&&p.length>0?p:vt}else t=vt}catch(a){console.warn("Gagal load training, menggunakan data lokal:",a),t=vt}finally{o=!1,k()}}async function S(){try{let{downloadExcel:a}=await Promise.resolve().then(()=>(re(),he)),p=t.map((d,g)=>({No:g+1,Tanggal:d.training_date||"",Batch:d.batch||"",Materi:d.subject||"",Cabang:d.branch_name||"",Trainer:d.trainer||"",Peserta:Ae(d.participants),Nilai:d.score!=null?d.score:"",Dokumen:d.document_link||"",Catatan:d.notes||""}));a(p,`Data_Training_${new Date().toISOString().slice(0,10)}`),R("Data training berhasil diekspor ke Excel!")}catch(a){N("Gagal ekspor Excel: "+a.message)}}async function F(a){if(!B)return;let p=t.find(d=>d.id===a);X({title:"Hapus Data Training",message:`Apakah Anda yakin ingin menghapus data sesi training "${p?.subject||"ini"}" untuk ${Ae(p?.participants)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let d=await K(`/api/training/${a}`,{method:"DELETE"});d.ok?(R("Data training berhasil dihapus"),t=t.filter(g=>g.id!==a),k()):N(d.data?.error||"Gagal menghapus data training")}catch(d){N("Gagal menghapus training: "+d.message)}}})}function $(a=null){if(!B)return;let p=!!a,d=a?.training_date?a.training_date.slice(0,10):new Date().toISOString().slice(0,10),g=a?.batch||"Batch 5",s=a?.trainer||"Berlin Ariansyah",r=Ae(a?.participants)!=="-"?Ae(a?.participants):"",l=document.createElement("div");l.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Training <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="tr-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${d}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Batch Pelatihan <span style="color:#DC2626;">*</span>
            </label>
            <input list="tr-batch-options" id="tr-modal-batch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih atau ketik batch..." value="${g}" required />
            <datalist id="tr-batch-options">
              ${ra.map(f=>`<option value="${f}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Materi / Topik Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="tr-modal-subject" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Basic Knowledge & Practical Skills" value="${a?.subject?String(a.subject).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="tr-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${i.map(f=>`
                <option value="${f.value}" ${String(a?.branch_id)===String(f.value)?"selected":""}>
                  ${f.label}
                </option>
              `).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Trainer / Instruktur <span style="color:#DC2626;">*</span>
            </label>
            <input list="tr-trainer-options" id="tr-modal-trainer" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih trainer..." value="${s}" required />
            <datalist id="tr-trainer-options">
              ${oa.map(f=>`<option value="${f}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Peserta Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="tr-modal-participants" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Ketik nama karyawan peserta training..." required>${r}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nilai / Score Evaluasi (0 - 100)
            </label>
            <input type="number" min="0" max="100" step="1" id="tr-modal-score" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: 90" value="${a?.score!=null?a.score:""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Link Dokumen / Sertifikat (Opsional)
            </label>
            <input type="url" id="tr-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="https://drive.google.com/..." value="${a?.document_link?String(a.document_link).replace(/"/g,"&quot;"):""}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan Pelatihan
          </label>
          <textarea id="tr-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Catatan hasil pelatihan, rekomendasi materi lanjutan...">${a?.notes||""}</textarea>
        </div>
      </div>
    `,G({title:p?"\u270F\uFE0F Edit Data Sesi Training":"\u2795 Tambah Data Sesi Training",content:l,size:"md",confirmText:p?"Simpan Perubahan":"Simpan Training",confirmClass:"btn-primary",onConfirm:async(f,v)=>{let L=l.querySelector("#tr-modal-date").value,A=(l.querySelector("#tr-modal-batch").value||"").trim(),P=(l.querySelector("#tr-modal-subject").value||"").trim(),w=l.querySelector("#tr-modal-branch").value,I=(l.querySelector("#tr-modal-trainer").value||"").trim(),z=(l.querySelector("#tr-modal-participants").value||"").trim(),C=l.querySelector("#tr-modal-score").value,y=C!==""?Number(C):null,_=(l.querySelector("#tr-modal-link").value||"").trim(),T=(l.querySelector("#tr-modal-notes").value||"").trim();if(!L){N("Tanggal training wajib diisi");return}if(!P){N("Materi / topik pelatihan wajib diisi");return}if(!w){N("Cabang klinik wajib dipilih");return}if(!I){N("Trainer wajib diisi");return}if(!z){N("Nama peserta training wajib diisi");return}let j=i.find(M=>String(M.value)===String(w)),O={training_date:L,batch:A,subject:P,branch_id:Number(w),branch_name:j?j.label:"Cabang",trainer:I,participants:z,score:y,document_link:_,notes:T};try{if(p){let M=await K(`/api/training/${a.id}`,{method:"PUT",body:JSON.stringify(O)});if(M.ok){R("Data training berhasil diperbarui");let q=t.findIndex(H=>H.id===a.id);q!==-1&&(t[q]={...t[q],...O}),v(),k()}else N(M.data?.error||"Gagal memperbarui data training")}else{let M=await K("/api/training",{method:"POST",body:JSON.stringify(O)});if(M.ok){R("Sesi training baru berhasil ditambahkan");let q=M.data?.id||Date.now();t.unshift({id:q,...O}),v(),k()}else N(M.data?.error||"Gagal menambahkan training")}}catch(M){N("Terjadi kesalahan: "+M.message)}}})}function E(){return t.filter(a=>{if(u){let p=u.toLowerCase(),d=(a.subject||"").toLowerCase().includes(p),g=(a.trainer||"").toLowerCase().includes(p),s=Ae(a.participants).toLowerCase().includes(p),r=(a.branch_name||"").toLowerCase().includes(p),l=(a.batch||"").toLowerCase().includes(p);if(!d&&!g&&!s&&!r&&!l)return!1}return!(c!=="all"&&a.batch!==c||x!=="all"&&String(a.branch_id)!==String(x)||m!=="all"&&a.trainer!==m)})}function k(){let a=E(),p=t.length,d=new Set;t.forEach(A=>{Ae(A.participants).split(",").forEach(w=>{let I=w.trim();I&&I!=="-"&&d.add(I)})});let g=d.size,r=new Set(t.map(A=>A.batch).filter(Boolean)).size,l=new Date,f=l.getFullYear(),v=l.getMonth(),L=t.filter(A=>{if(!A.training_date)return!1;let P=new Date(A.training_date);return!isNaN(P.getTime())&&P.getFullYear()===f&&P.getMonth()===v}).length;e.innerHTML=`
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
              ${B?`
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
                <div class="training-stat-val">${p}</div>
                <div class="training-stat-lbl">Sesi Training</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F465}
              </div>
              <div>
                <div class="training-stat-val">${g}</div>
                <div class="training-stat-lbl">Peserta Terlatih</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4DA}
              </div>
              <div>
                <div class="training-stat-val">${r} Batch</div>
                <div class="training-stat-lbl">Total Angkatan</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="training-stat-val">${L}</div>
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
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(A=>`<option value="${A.value}" ${String(x)===String(A.value)?"selected":""}>${A.label}</option>`).join("")}
            </select>

            <select id="tr-filter-trainer" class="training-select">
              <option value="all">Semua Trainer</option>
              ${oa.map(A=>`<option value="${A}" ${m===A?"selected":""}>${A}</option>`).join("")}
            </select>
          </div>

          <!-- Batch Filter Pills (Swipeable on Mobile) -->
          <div class="training-filter-pills">
            <div class="training-pill ${c==="all"?"active":""}" data-batch="all">
              <span>Semua Batch</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${p}</span>
            </div>
            ${ra.map(A=>{let P=t.filter(w=>w.batch===A).length;return`
                <div class="training-pill ${c===A?"active":""}" data-batch="${A}">
                  <span>\u{1F393}</span>
                  <span>${A}</span>
                  ${P>0?`<span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${P}</span>`:""}
                </div>
              `}).join("")}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="training-list-container">
          ${a.length===0?`
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
              ${a.map(A=>{let P=Ae(A.participants),I=P.split(",")[0].trim().split(" ").map(z=>z[0]).filter(Boolean).slice(0,2).join("").toUpperCase()||"TR";return`
                  <div class="training-card" data-id="${A.id}">
                    <div>
                      <!-- Card Header: Batch & Nilai/Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="training-batch-badge">
                          <span>\u{1F393}</span>
                          <span>${A.batch||"Reguler"}</span>
                        </span>
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${A.score!=null?`
                            <span class="training-score-badge">
                              <span>\u2B50</span>
                              <span>${A.score}</span>
                            </span>
                          `:""}
                          <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                            \u{1F5D3}\uFE0F ${Bn(A.training_date)}
                          </span>
                        </div>
                      </div>

                      <!-- Subject / Materi Title -->
                      <h4 style="margin:0 0 10px 0; font-size:1.05rem; font-weight:700; color:#0F172A; line-height:1.4;">
                        ${A.subject||"Materi Pelatihan"}
                      </h4>

                      <!-- Participant & Branch Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:10px;">
                        <div class="training-avatar-box">
                          ${I}
                        </div>
                        <div style="min-width:0;">
                          <div style="font-size:0.72rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:1px;">Peserta Training</div>
                          <div style="font-size:0.95rem; font-weight:700; color:#1E293B; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${P}">
                            ${P}
                          </div>
                        </div>
                      </div>

                      <!-- Detail Box: Cabang & Trainer -->
                      <div class="training-info-box">
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#334155;">
                          <span>\u{1F4CD}</span>
                          <span style="font-weight:600;">${A.branch_name||"Cabang Belum Ditentukan"}</span>
                        </div>
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#64748B;">
                          <span>\u{1F468}\u200D\u{1F3EB}</span>
                          <span>Trainer: <strong style="color:#0F172A;">${A.trainer||"-"}</strong></span>
                        </div>
                      </div>

                      ${A.notes?`
                        <div style="font-size:0.78rem; color:#64748B; font-style:italic; line-height:1.4; margin-bottom:8px;">
                          \u{1F4DD} ${A.notes}
                        </div>
                      `:""}
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:8px;">
                      <div>
                        ${A.document_link?`
                          <a href="${A.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>\u{1F4C4} Materi / Sertifikat</span>
                          </a>
                        `:`
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${B?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          <button class="training-icon-btn training-edit-btn" title="Edit Data Training" data-id="${A.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                          <button class="training-icon-btn delete training-delete-btn" title="Hapus Training" data-id="${A.id}">
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
    `,h()}function h(){let a=e.querySelector("#tr-btn-add");a&&(a.onclick=()=>$());let p=e.querySelector("#tr-btn-excel");p&&(p.onclick=()=>S());let d=e.querySelector("#tr-btn-refresh");d&&(d.onclick=()=>D());let g=e.querySelector("#tr-btn-reset-filter");g&&(g.onclick=()=>{u="",x="all",c="all",m="all",k()});let s=e.querySelector("#tr-search-input");s&&(s.oninput=f=>{u=f.target.value,k();let v=e.querySelector("#tr-search-input");v&&(v.focus(),v.setSelectionRange(v.value.length,v.value.length))});let r=e.querySelector("#tr-filter-branch");r&&(r.onchange=f=>{x=f.target.value,k()});let l=e.querySelector("#tr-filter-trainer");l&&(l.onchange=f=>{m=f.target.value,k()}),e.querySelectorAll(".training-pill").forEach(f=>{f.onclick=()=>{c=f.getAttribute("data-batch")||"all",k()}}),e.querySelectorAll(".training-edit-btn").forEach(f=>{f.onclick=v=>{v.stopPropagation();let L=Number(f.getAttribute("data-id")),A=t.find(P=>P.id===L);A&&$(A)}}),e.querySelectorAll(".training-delete-btn").forEach(f=>{f.onclick=v=>{v.stopPropagation();let L=Number(f.getAttribute("data-id"));F(L)}})}await D()}le();re();function la({container:e,title:t,icon:i,apiPath:n,columns:c,formFields:x,filterFields:m,defaultFilters:u={},itemLabel:o="Data",canCreate:b=!0,canEdit:B=!0,canDelete:D=!0,onBeforeSubmit:S,onAfterLoad:F,onDataLoaded:$,extraActions:E=[],initialSearch:k="",exportOptions:h=null,bulkDelete:a=!1,paginationMode:p="server"}){let d=ee();d&&d.role==="viewer"&&(b=!1,B=!1,D=!1,a=!1,h=null),d&&d.role==="editor_khusus"&&n!=="/api/issues"&&n!=="/api/overtime"&&(b=!1,B=!1,D=!1,a=!1,h=null),d&&d.role==="input_lembur"&&(n==="/api/overtime"?(D=!1,a=!1,h=null):(b=!1,B=!1,D=!1,a=!1,h=null));let g=1,s={...u};k&&(s.search=k);let r=new Set;e.innerHTML=`
    ${a?`
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
        ${b?`<button class="btn btn-primary" id="btn-create">+ Tambah ${o}</button>`:""}
        ${h?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${h.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${h.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${h.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${h.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
    

    ${m&&m.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${m.filter(y=>y.type==="search").map(y=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${y.placeholder||"Cari..."}" id="filter-search" value="${s.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${m.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let _=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select" name="${y.name}" id="filter-${y.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${_}</option>${(y.options||[]).map(T=>`<option value="${typeof T=="object"?T.value:T}" ${s[y.name]===(typeof T=="object"?T.value:T)?"selected":""}>${typeof T=="object"?T.label:T}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${m.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let _=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select filter-select-sheet" name="${y.name}-sheet" id="filter-sheet-${y.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${_}</option>${(y.options||[]).map(T=>`<option value="${typeof T=="object"?T.value:T}" ${s[y.name]===(typeof T=="object"?T.value:T)?"selected":""}>${typeof T=="object"?T.label:T}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function l(){let y=document.getElementById("bulk-toolbar");if(!y)return;let _=document.getElementById("bulk-count"),T=document.getElementById("btn-bulk-delete"),j=document.getElementById("btn-bulk-cancel");_.textContent=`${r.size} item dipilih`,r.size>0?(y.style.display="flex",T.disabled=!1,j.disabled=!1):(y.style.display="none",T.disabled=!0,j.disabled=!0);let O=document.getElementById("select-all-checkbox");if(O){let M=document.querySelectorAll(".row-checkbox");if(M.length>0){let q=[...M].every(Y=>Y.checked),H=[...M].some(Y=>Y.checked);O.checked=q,O.indeterminate=H&&!q}else O.checked=!1,O.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{r.clear(),document.querySelectorAll(".row-checkbox").forEach(_=>_.checked=!1);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1),l()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(r.size===0)return;let y=[...r],_=document.createElement("div");_.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",_.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${y.length} ${o}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${y.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(_),_.querySelector("#bulk-cancel-btn").addEventListener("click",()=>_.remove()),_.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let T=_.querySelector("#bulk-confirm-btn");T.disabled=!0,T.textContent="Menghapus...";let j=await K(`${n}/bulk`,{method:"DELETE",body:JSON.stringify({ids:y})});_.remove(),j.ok?(R(`${y.length} ${o} berhasil dihapus.`),r.clear(),l(),be(n),w()):N(j.data?.error||"Gagal menghapus data.")})});let f=document.getElementById("filter-search"),v;f?.addEventListener("input",y=>{clearTimeout(v),v=setTimeout(()=>{s.search=y.target.value,g=1,r.clear(),l(),w()},400)}),m?.forEach(y=>{(y.type==="select"||y.type==="combobox")&&(document.getElementById(`filter-${y.name}`)?.addEventListener("change",_=>{s[y.name]=_.target.value;let T=document.getElementById(`filter-sheet-${y.name}`);T&&(T.value=_.target.value),g=1,r.clear(),l(),w()}),document.getElementById(`filter-sheet-${y.name}`)?.addEventListener("change",_=>{s[y.name]=_.target.value;let T=document.getElementById(`filter-${y.name}`);T&&(T.value=_.target.value),g=1,r.clear(),l(),w(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{s={},f&&(f.value=""),m?.forEach(y=>{let _=document.getElementById(`filter-${y.name}`);_&&(_.value="");let T=document.getElementById(`filter-sheet-${y.name}`);T&&(T.value="")}),g=1,r.clear(),l(),w()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{s={},f&&(f.value=""),m?.forEach(y=>{let _=document.getElementById(`filter-${y.name}`);_&&(_.value="");let T=document.getElementById(`filter-sheet-${y.name}`);T&&(T.value="")}),g=1,r.clear(),l(),w(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>z(null)),h&&document.addEventListener("click",function(y){let _=document.getElementById("aksi-menu-main"),T=document.getElementById("btn-aksi-main");_&&T&&!T.contains(y.target)&&!_.contains(y.target)&&_.classList.remove("show-aksi-menu")});let L=document.getElementById("btn-mobile-filter"),A=document.getElementById("filter-options-wrapper"),P=document.getElementById("btn-close-filter-sheet");if(L&&A&&(L.addEventListener("click",y=>{y.preventDefault(),A.classList.add("sheet-open")}),P&&P.addEventListener("click",y=>{y.preventDefault(),A.classList.remove("sheet-open")})),h){document.getElementById(`btn-export-${h.moduleName}`)?.addEventListener("click",async _=>{let T=_.target,j=T.innerHTML;T.innerHTML="\u23F3 Loading...",T.disabled=!0;try{await h.onExport()}catch{N("Gagal export data")}finally{T.innerHTML=j,T.disabled=!1}}),document.getElementById(`btn-template-${h.moduleName}`)?.addEventListener("click",()=>{h.onTemplate()});let y=document.getElementById(`input-import-${h.moduleName}`);y?.addEventListener("change",async _=>{let T=_.target.files[0];if(!T)return;let j=document.getElementById(`label-import-${h.moduleName}`),O=j?j.querySelector(".import-text"):null,M=O?O.innerText:"";O&&(O.innerText="\u231B Memproses..."),j&&(j.style.pointerEvents="none"),y.disabled=!0;try{let q=await qe(T);if(q.length===0)throw new Error("File kosong atau format salah");await h.onImport(q),R("Import berhasil!"),be(n),w()}catch(q){N(q.message||"Gagal import data")}finally{O&&(O.innerText=M),j&&(j.style.pointerEvents="auto"),y.disabled=!1,y.value=""}})}async function w(){l();let y=document.getElementById("table-container");if(!y)return;y.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let _=p==="client",T=_?1:g,j=_?Be:20,O=new URLSearchParams({page:T,limit:j,...Object.fromEntries(Object.entries(s).filter(([,U])=>U))}),M=await K(`${n}?${O}`);if(!M.ok){y.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${M.data?.error||"Error"}</p></div>`;return}let q=M.data?.data||M.data||[],H=M.data?.pagination,Y=q.length;if(_){q=$(q);let U=q.length,V=20,ae=Math.ceil(U/V);g>ae&&ae>0&&(g=ae);let J=(g-1)*V,ie=g*V;q=q.slice(J,ie),H={page:g,limit:V,total:U,pages:ae}}!1,F&&F(q);let W=Xe({columns:c,data:q,onEdit:B?U=>z(U):null,actions:E.map(U=>({...U,handler:V=>U.handler(V,w)})),emptyText:`Tidak ada ${String(o||"").toLowerCase()}`,bulkSelect:a?{selectedIds:r,onToggle:l}:null});y.innerHTML="",y.appendChild(W);let Z=document.getElementById("pagination-container");if(Z&&(Z.innerHTML="",H&&H.pages>1)){let U=Ze({page:H.page,pages:H.pages,total:H.total,limit:H.limit,onPage:V=>{g=V,w()}});U&&Z.appendChild(U)}}function I(y){let _=typeof x=="function"?x(y):x;return Ne(_)}function z(y){let _=!!y,T=document.createElement("form");if(T.noValidate=!0,T.innerHTML=I(y),_){let O=typeof x=="function"?x(y):x;tt(T,y)}let{close:j}=G({title:_?`Edit ${o}`:`Tambah ${o}`,content:T,size:"lg",confirmText:_?"Simpan Perubahan":`Tambah ${o}`,onConfirm:async(O,M)=>{if(!T.reportValidity())return;let q=O.querySelector(".modal-confirm");q.disabled=!0,q.textContent="Menyimpan...";let H=et(T),Y=typeof x=="function"?x(y):x,W=async ae=>{for(let J of ae)if(J.type==="row")await W(J.fields);else if(J.type==="combobox"&&H[J.name]){let ie=H[J.name],me=(J.options||[]).find(se=>{let ce=String(typeof se=="object"?se.value:se),Je=String(typeof se=="object"?se.label:se);return ce===ie||Je===ie});if(me)H[J.name]=typeof me=="object"?me.value:me;else if(J.createApi){let se={};se[J.createApi.field]=ie,J.createApi.extra&&Object.assign(se,J.createApi.extra);let ce=await K(J.createApi.path,{method:"POST",body:JSON.stringify(se)});if(ce.ok&&ce.data?.id)H[J.name]=ce.data.id;else if(ce.ok&&!ce.data?.id)H[J.name]=ie;else throw new Error(`Gagal membuat master data: ${ce.data?.error||"Unknown error"}`)}}};try{await W(Y)}catch(ae){N(ae.message),q.disabled=!1,q.textContent=_?"Simpan Perubahan":`Tambah ${o}`;return}S&&(H=await S(H,y));let Z=_?"PUT":"POST",U=_?`${n}/${y.id}`:n,V=await K(U,{method:Z,body:JSON.stringify(H)});V.ok?(R(_?`${o} berhasil diperbarui.`:`${o} berhasil ditambahkan.`),M(),be(n),w()):(N(V.data?.error||"Gagal menyimpan data."),q.disabled=!1,q.textContent=_?"Simpan Perubahan":`Tambah ${o}`)}})}function C(y){X(`Hapus ${o} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let _=await K(`${n}/${y.id}`,{method:"DELETE"});_.ok?(R(`${o} berhasil dihapus.`),be(n),w()):N(_.data?.error||"Gagal menghapus.")},`Hapus ${o}`)}return w(),w}re();async function da(e,t){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await te(),n=await ue(),c=t?t.get("dash_filter"):null,x={};if(c==="reliever"){let S=new Date,F=`${S.getFullYear()}-${String(S.getMonth()+1).padStart(2,"0")}`;x={status:"Done",month:t&&t.get("month")?t.get("month"):F}}console.log("RAW",await _e()),console.log("OPTIONS",n);let m=S=>S&&!n.find(F=>F.value===S)?[...n,{value:S,label:S}]:n,u=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],o=S=>S&&!u.includes(S)?[...u,S]:u,b=new Date().getFullYear(),D=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((S,F)=>{let $=String(F+1).padStart(2,"0");return{value:`${b}-${$}`,label:`${S} ${b}`}});la({container:e,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:x,onDataLoaded:S=>S.sort((F,$)=>{let E=F.backup_date?new Date(F.backup_date).getTime():0;return($.backup_date?new Date($.backup_date).getTime():0)-E}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:S=>xe(S)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:S=>window.formatDate(S)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:S=>window.formatDate(S)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:S=>S?`<span class="badge badge-info">${S}</span>`:"-"},{key:"status",label:"Status",render:S=>ge(S)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:u},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"month",label:"Bulan",options:D},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:S=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:S?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:S?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:m(S?.original_fc_name),value:S?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:o(S?.reliever_name),value:S?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:S?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:S?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:S?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:S?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:S?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let S=await K(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(S.ok){let F=S.data.data.map($=>({Cabang:$.branch_name||"","Nama Facility care":$.original_fc_name||"",Periode:$.period||"",Relifer:$.reliever_name||"","Tanggal Back Up":$.backup_date||"","Tanggal Selesai":$.completion_date||"",Keterangan:$.reason||"",Shift:$.shift||"",Status:$.status||""}));F.length===0&&F.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),Q(F,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Q([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async S=>{let $=(await K("/api/branches?all=1")).data?.data||[],E=a=>{if(!a)return null;let p=String(a||"").toLowerCase(),d=$.find(g=>String(g.full_name||"").toLowerCase()===p||String(g.code||"").toLowerCase()===p||String(g.name||"").toLowerCase()===p);return d?d.id:null},k=S.map(a=>({branch_name:String(a.Cabang||"").trim(),backup_date:String(a["Tanggal Back Up"]||a["Tanggal Backup"]||"").trim(),original_fc_name:String(a["Nama Facility care"]||a["FC Digantikan"]||"").trim(),reliever_name:String(a.Relifer||a.Reliefer||"").trim(),period:String(a.Periode||"").trim(),reason:String(a.Keterangan||"").trim(),shift:String(a.Shift||"").trim(),completion_date:String(a["Tanggal Selesai"]||"").trim(),status:String(a.Status||"").trim()})).filter(a=>a.reliever_name&&a.backup_date),h=await K("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:k,onDuplicate:"update"})});if(!h.ok)throw new Error(h.data?.error||"Import gagal");return h.data}}})}re();async function pa(e){let t=await te(),i=Array.from({length:4},(n,c)=>String(new Date().getFullYear()-c));de({container:e,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:n=>xe(n)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"fc_score",label:"Point FC",render:n=>n!=null?`<strong class="${n>=80?"text-success":n>=60?"text-warning":"text-danger"}">${n}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"status",label:"Status",render:n=>ge(n)},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:n?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:n?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:n?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:n?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:n?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async n=>{let c=new URLSearchParams(n||{}).toString(),x=await K(`/api/reports/inspection?limit=10000&${c}`);if(x.ok){let m=x.data.data.map(u=>({Cabang:u.branch_name||"",Periode:u.period||"",Tanggal:u.inspection_date||"","Point FC":u.fc_score!==null&&u.fc_score!==void 0?u.fc_score:"","Point SPV":u.spv_score!==null&&u.spv_score!==void 0?u.spv_score:"",Status:u.status||"","Link Dokumen":u.document_link||""}));Q(m,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Q([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async n=>{let c=o=>{if(!o)return null;let b=String(o||"").toLowerCase(),B=t.find(D=>String(D.label||"").toLowerCase()===b);return B?B.value:null},x=o=>{if(o==null||o==="")return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let b=String(o).trim();if(b===""||b==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(b))return b.slice(0,10);if(/^\d{4,5}$/.test(b)){let D=Number(b);if(D>2e4&&D<99999){let S=new Date(Date.UTC(1899,11,30)+D*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}let B=b.split(/[\/\-\.]/);if(B.length===3){let[D,S,F]=B.map($=>$.trim());if(D.length===4&&S.length<=2&&F.length<=2)return`${D}-${S.padStart(2,"0")}-${F.padStart(2,"0")}`;if(F.length===4&&S.length<=2&&D.length<=2)return`${F}-${S.padStart(2,"0")}-${D.padStart(2,"0")}`}return b},m=n.map(o=>({branch_id:c(String(o.Cabang||"").trim()),period:String(o.Periode||"").trim(),inspection_date:x(o.Tanggal),fc_score:o["Point FC"]!==void 0&&o["Point FC"]!==""?Number(o["Point FC"]):null,spv_score:o["Point SPV"]!==void 0&&o["Point SPV"]!==""?Number(o["Point SPV"]):null,status:String(o.Status||"").trim(),document_link:String(o["Link Dokumen"]||"").trim(),notes:String(o.Catatan||o.Keterangan||"").trim()})).filter(o=>o.branch_id&&o.period&&o.inspection_date),u=await K("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}re();async function ca(e){let t=await te(),i=Array.from({length:4},(n,c)=>String(new Date().getFullYear()-c));de({container:e,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:n=>`<span class="badge ${n==="Deep Cleaning"?"badge-purple":"badge-success"}">${n}</span>`},{key:"period",label:"Periode",render:n=>xe(n)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"status",label:"Status",render:n=>ge(n)},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t,value:n?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:n?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:n?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:n?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:n?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async n=>{let c=new URLSearchParams(n||{}).toString(),x=await K(`/api/reports/cleaning?limit=10000&${c}`);if(x.ok){let m=x.data.data.map(u=>({Cabang:u.branch_name||"",Jenis:u.activity_type||"",Periode:u.period||"",Tanggal:u.activity_date||"",Status:u.status||"","Link Dokumen":u.document_link||""}));Q(m,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Q([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async n=>{let c=o=>{if(!o)return null;let b=String(o||"").toLowerCase(),B=t.find(D=>String(D.label||"").toLowerCase()===b);return B?B.value:null},x=o=>{if(o==null||o==="")return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let b=String(o).trim();if(b===""||b==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(b))return b.slice(0,10);if(/^\d{4,5}$/.test(b)){let D=Number(b);if(D>2e4&&D<99999){let S=new Date(Date.UTC(1899,11,30)+D*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}let B=b.split(/[\/\-\.]/);if(B.length===3){let[D,S,F]=B.map($=>$.trim());if(D.length===4&&S.length<=2&&F.length<=2)return`${D}-${S.padStart(2,"0")}-${F.padStart(2,"0")}`;if(F.length===4&&S.length<=2&&D.length<=2)return`${F}-${S.padStart(2,"0")}-${D.padStart(2,"0")}`}return b},m=n.map(o=>({branch_id:c(String(o.Cabang||"").trim()),activity_type:String(o.Jenis||o.Kegiatan||"").trim(),period:String(o.Periode||"").trim(),activity_date:x(o.Tanggal),status:String(o.Status||"").trim(),document_link:String(o["Link Dokumen"]||"").trim(),notes:String(o.Catatan||o.Keterangan||"").trim()})).filter(o=>o.branch_id&&o.activity_type&&o.period&&o.activity_date),u=await K("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}re();async function ua(e,t){let i=await te(),n=Array.from({length:4},(m,u)=>String(new Date().getFullYear()-u)),c=t?t.get("dash_filter"):null,x={};if(c==="fogging"){let m=new Date,u=String(m.getMonth()+1).padStart(2,"0"),o=String(m.getFullYear()),b=t?t.get("month"):null;b&&b.length===7&&(o=b.split("-")[0],u=b.split("-")[1]),x={status:"Done",month:u,year:o}}de({container:e,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:x,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:m=>`<span class="badge badge-warning">${m}</span>`},{key:"period",label:"Periode",render:m=>xe(m)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:m=>window.formatDate(m)},{key:"status",label:"Status",render:m=>ge(m)},{key:"document_link",label:"Dokumen",render:m=>m?`<a href="${m}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:m=>m||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:m=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:i,value:m?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:m?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:m?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:m?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:m?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:m?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async m=>{let u=new URLSearchParams(m||{}).toString(),o=await K(`/api/reports/fogging?limit=10000&${u}`);if(o.ok){let b=o.data.data.map(B=>({Cabang:B.branch_name||"",Jenis:B.activity_type||"Fogging",Periode:B.period||"",Tanggal:B.activity_date||"",Status:B.status||"","Link Dokumen":B.document_link||""}));Q(b,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Q([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async m=>{let u=D=>{if(!D)return null;let S=String(D||"").toLowerCase(),F=i.find($=>String($.label||"").toLowerCase()===S);return F?F.value:null},o=D=>{if(D==null||D==="")return"";if(D instanceof Date&&!isNaN(D.getTime()))return D.toISOString().slice(0,10);let S=String(D).trim();if(S===""||S==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(S))return S.slice(0,10);if(/^\d{4,5}$/.test(S)){let $=Number(S);if($>2e4&&$<99999){let E=new Date(Date.UTC(1899,11,30)+$*864e5);return isNaN(E.getTime())?"":E.toISOString().slice(0,10)}}let F=S.split(/[\/\-\.]/);if(F.length===3){let[$,E,k]=F.map(h=>h.trim());if($.length===4&&E.length<=2&&k.length<=2)return`${$}-${E.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&E.length<=2&&$.length<=2)return`${k}-${E.padStart(2,"0")}-${$.padStart(2,"0")}`}return S},b=m.map(D=>({branch_id:u(String(D.Cabang||"").trim()),activity_type:String(D.Jenis||D.Kegiatan||"Fogging").trim(),period:String(D.Periode||"").trim(),activity_date:o(D.Tanggal),status:String(D.Status||"").trim(),document_link:String(D["Link Dokumen"]||"").trim(),notes:String(D.Catatan||D.Keterangan||"").trim()})).filter(D=>D.branch_id&&D.period&&D.activity_date),B=await K("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(b)});if(!B.ok)throw new Error(B.data?.error||"Import gagal");return B.data}}})}re();async function ga(e){let t=await te(),i=await ue(),n=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],c=Array.from({length:4},(u,o)=>String(new Date().getFullYear()-o)),x=u=>u&&!i.find(o=>o.value===u)?[...i,{value:u,label:u}]:i,m=u=>u&&!n.find(o=>o.value===u)?[...n,{value:u,label:u}]:n;de({container:e,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:u=>window.formatDate(u)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:u=>`<span title="${u||""}">${u?.length>60?u.slice(0,60)+"\u2026":u||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:u=>window.formatDate(u)},{key:"status",label:"Status",render:u=>ge(u)},{key:"notes",label:"Keterangan",render:u=>u?.length>40?u.slice(0,40)+"\u2026":u||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:c}],formFields:u=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:u?.branch_id},{name:"pic",label:"PIC",type:"select",options:m(u?.pic),value:u?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:u?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:u?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:u?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:u?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:u?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async u=>{let o=new URLSearchParams(u||{}).toString(),b=await K(`/api/reports/basecamp?limit=10000&${o}`);if(b.ok){let B=b.data.data.map(D=>({"Tgl Info":D.info_date||"",Cabang:D.branch_name||"",Permasalahan:D.problem||"",PIC:D.pic||"","Tgl Done":D.done_date||"",Status:D.status||"",Keterangan:D.notes||""}));Q(B,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Q([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async u=>{let o=S=>{if(!S)return null;let F=String(S||"").toLowerCase(),$=t.find(E=>String(E.label||"").toLowerCase()===F);return $?$.value:null},b=S=>{if(S==null||S==="")return"";if(S instanceof Date&&!isNaN(S.getTime()))return S.toISOString().slice(0,10);let F=String(S).trim();if(F===""||F==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(F))return F.slice(0,10);if(/^\d{4,5}$/.test(F)){let E=Number(F);if(E>2e4&&E<99999){let k=new Date(Date.UTC(1899,11,30)+E*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}let $=F.split(/[\/\-\.]/);if($.length===3){let[E,k,h]=$.map(a=>a.trim());if(E.length===4&&k.length<=2&&h.length<=2)return`${E}-${k.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&k.length<=2&&E.length<=2)return`${h}-${k.padStart(2,"0")}-${E.padStart(2,"0")}`}return F},B=u.map(S=>({info_date:b(S["Tgl Info"]||S["Tanggal Info"]),branch_id:o(String(S.Cabang||"").trim()),problem:String(S.Permasalahan||"").trim(),pic:String(S.PIC||"").trim(),done_date:b(S["Tgl Done"]||S["Tanggal Done"]),status:String(S.Status||"").trim(),notes:String(S.Keterangan||S.Catatan||"").trim()})).filter(S=>S.info_date&&S.branch_id&&S.problem),D=await K("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(B)});if(!D.ok)throw new Error(D.data?.error||"Import gagal");return D.data}}})}le();var Tn=[{id:1,title:"Dusting (Pembersihan Debu)",category:"Pembersihan Rutin",icon:"\u{1F9F9}",color:"#2563EB",bg:"#EFF6FF",target:"Meja resepsionis, meja pasien, meja dokter, furniture, kaca, monitor & keyboard",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Kain Lap Microfiber Biru (Area Umum)","Kanebo (Kaca & Layar Monitor)","Kemoceng","Botol Spray Glass Cleaner","Kuas 4 Inch (Keyboard & Sela Sempit)"],steps:["Siapkan larutan glass cleaner di botol spray khusus kaca dan perabot.","Lipat kain microfiber menjadi 8 bidang kerja agar efisien dan bersih merata.","Semprotkan cairan pembersih ke permukaan kain lap, bukan langsung ke monitor atau bodi elektronik.","Usap permukaan perabot dari bagian paling atas ke bagian bawah secara searah atau pola S.","Gunakan kuas cat 4 inch untuk membersihkan sela-sela keyboard dan lubang ventilasi.","Keringkan kaca atau monitor menggunakan lap microfiber kering atau kanebo bersih."],notes:"Khusus layar monitor medis & dental unit, dilarang menyemprotkan cairan langsung ke layar."},{id:2,title:"Sweeping (Lobby Duster & Sapu Nylon)",category:"Pembersihan Rutin",icon:"\u{1F33E}",color:"#D97706",bg:"#FFFBEB",target:"Lobby utama, lorong, ruang tindakan, balkon, pantry, teras",chemical_ratio:"-",tools:["Lobby Duster Set 60cm (Area Lobby & Lorong)","Lobby Duster Set 40cm (Area Ruang Tindakan)","Sapu & Pengki Nagata 742s","Kapi Scrapper (Kerak/Permen Karet)"],steps:["Pastikan kain lobby duster dalam keadaan kering, bersih, dan bebas pasir kasar.","Dorong lobby duster menyusuri tepi dinding ruangan terlebih dahulu membentuk perimeter.","Lanjutkan mendorong dari sudut terjauh menuju arah pintu keluar tanpa mengangkat kain pel lobi dari lantai.","Kumpulkan debu dan kotoran di dekat pintu keluar, lalu angkat menggunakan sapu dan pengki Nagata.","Jika ada noda kerak atau bekas stiker/permen karet menempel di lantai, kikis perlahan menggunakan kapi scrapper.","Kibaskan kain duster di luar ruangan atau sedot sisa debu pada kain duster."],notes:"Jangan pernah menggunakan lobby duster di area lantai yang masih basah atau tergenang air."},{id:3,title:"Mopping (Pengepelan Lantai)",category:"Pembersihan Rutin",icon:"\u{1FAA3}",color:"#059669",bg:"#ECFDF5",target:"Seluruh lantai ruangan, koridor, dan area tunggu klinik",chemical_ratio:"25 ml Floor Cleaner per 1 Liter air bersih",tools:["Single Bucket Mop Wringer 20L / Double Bucket 46L","Mopping Stick / Tongkat & Kain Pel","Chemical Floor Cleaner Prokleen Lavender","Warning Sign Lantai Licin"],steps:['Pasang papan peringatan "Caution Wet Floor" di titik strategis yang mudah terlihat.',"Campurkan cairan Floor Cleaner ke dalam ember air bersih dengan takaran 25 ml per 1 Liter air.","Celupkan kain pel, masukkan ke dalam alat pemeras (wringer), dan peras hingga kondisi kain pel lembap (tidak menetes).","Mulai mengepel dari sudut terjauh ruangan mundur ke arah pintu keluar dengan gerakan membentuk pola angka delapan (8).","Bilas kain pel secara berkala apabila air pada ember bilas sudah terlihat keruh.","Biarkan lantai mengering sempurna secara alami sebelum papan peringatan wet floor dilepas."],notes:'Wajib pasang papan peringatan "Caution Wet Floor" sebelum mulai mengepel untuk mencegah risiko terpeleset.'},{id:4,title:"Membersihkan Toilet & Wastafel",category:"Sanitasi Toilet",icon:"\u{1F6BD}",color:"#DC2626",bg:"#FEF2F2",target:"Toilet pasien, toilet tim, closet, wastafel, dinding keramik, floor drain",chemical_ratio:"Karbol 25 ml:1 L Air / Bowl Cleaner 100 ml spray (kerak) / Bubuk KIFA",tools:["Sarung Tangan Karet (Wajib)","Sikat Kloset Scotch-Brite 3M","Kain Lap Microfiber Merah (Khusus Toilet)","Sponge Scotch Brite & Sikat Lantai ID-51","Floor Squeegee 75cm","Chemical Karbol Pine, Bowl Cleaner, Bubuk KIFA"],steps:["Kenakan sarung tangan karet dan siapkan perlengkapan sanitasi toilet.","Flush kloset, tuangkan Toilet Bowl Cleaner ke bagian dalam dinding kloset dan leher angsa, diamkan 5\u201310 menit.","Sikat bagian dalam kloset secara menyeluruh dengan sikat kloset, lalu bilas (flush) kembali hingga bersih.","Seka dudukan, penutup, dan tabung luar kloset menggunakan kain lap microfiber merah yang dibasahi larutan disinfektan karbol.","Bersihkan wastafel dan kran stainless menggunakan sponge tapas dengan bubuk pembersih KIFA (jangan gunakan Bowl Cleaner pada kran stainless).","Sikat lantai kamar mandi dengan sikat lantai ID-51 dan cairan karbol disinfektan.","Tarik genangan air ke arah floor drain menggunakan floor squeegee hingga lantai tidak licin dan cepat kering.","Pasang kamper/urinal block, isi ulang sabun cuci tangan, dan pasang plastik sampah baru."],notes:"Kain lap warna MERAH mutlak hanya untuk toilet dan mangkuk spittoon. Dilarang mencampur cairan kimia sembarangan."},{id:5,title:"Pembersihan Area Parkir & Luar",category:"Pembersihan Rutin",icon:"\u{1F17F}\uFE0F",color:"#4F46E5",bg:"#EEF2FF",target:"Halaman parkir mobil & motor, saluran air luar, teras depan",chemical_ratio:"-",tools:["Sapu Lidi Gagang Panjang","Pengki Nagata","Plastik Sampah Hitam Besar (90x120)","Selang Air"],steps:["Sapu daun kering, puntung rokok, plastik, dan pasir dari batas trotoar/pagar menuju satu titik.","Kumpulkan dan angkat sampah ke dalam pengki, masukkan langsung ke plastik sampah hitam besar.","Periksa dan bersihkan saluran drainase air di area parkir dari sumbatan sampah padat.","Bila ada tumpahan oli atau kotoran tanah membandel, siram dengan air dan sikat.","Rapikan pembatas parkir dan pastikan area depan klinik terlihat bersih dan rapi menyambut pasien."],notes:"Lakukan sebelum jam operasional klinik dibuka (pukul 07.30 - 08.00)."},{id:6,title:"Pembersihan Kaca & Jendela",category:"Pembersihan Rutin",icon:"\u{1FA9F}",color:"#0891B2",bg:"#ECFEFF",target:"Pintu kaca utama, jendela kaca klinik, partisi kaca, cermin toilet",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Botol Spray Glass Cleaner Prokleen","Window Wiper / Window Squeegee","Kain Microfiber Kering / Kanebo","Tongkat Teleskopik (Kaca Tinggi)"],steps:["Semprotkan larutan glass cleaner merata pada permukaan kaca yang hendak dibersihkan.","Tarik wiper karet dari arah atas ke bawah secara lurus vertikal atau zigzag horizontal tanpa terputus.","Lap bilah karet wiper setiap selesai satu tarikan menggunakan kain lap bersih agar tidak meninggalkan garis bekas air.","Gunakan kain microfiber kering untuk membersihkan sisa tetesan air di sudut-sudut lis bingkai kaca.","Untuk kaca bagian luar yang tinggi, pasang wiper pada tongkat sambungan teleskopik."],notes:"Hindari membersihkan kaca saat terpapar terik sinar matahari langsung karena cairan akan cepat mengering dan membekas noda."},{id:7,title:"Mencuci Bola Mandi Bola (Playground)",category:"Pembersihan Khusus",icon:"\u{1F939}",color:"#7C3AED",bg:"#F5F3FF",target:"Area bermain anak (Kids Playground) dan bola-bola mandi bola",chemical_ratio:"30 ml Dishwashing Liquid per 1 Liter Air + Karbol Disinfektan",tools:["Baskom Plastik Besar 55cm","Dishwashing Liquid Prokleen","Karbol Disinfektan Pine","Lap Microfiber Bersih","Karung / Wadah Pengering Jaring"],steps:["Kumpulkan seluruh bola mandi bola ke dalam wadah atau karung bersih.","Siapkan baskom plastik besar berisi air hangat kuku dicampur Dishwashing Liquid dan sedikit karbol disinfektan.","Rendam bola dan cuci secara bertahap sambil digosok perlahan menggunakan sponge/lap bersih.","Bilas bola dengan air bersih mengalir hingga tidak ada sisa busa sabun yang menempel.","Tiriskan dan angin-anginkan bola di tempat bersih beralaskan terpal/kain kering hingga benar-benar kering.","Seka dan sterilkan area kolam mandi bola sebelum bola dimasukkan kembali."],notes:"Jadwal pencucian berkala dilakukan minimal 1 kali seminggu atau sesuai kebutuhan kebersihan arena anak."},{id:8,title:"Membersihkan Filter Air Purifier",category:"Elektronik & Utility",icon:"\u{1F4A8}",color:"#0D9488",bg:"#F0FDFA",target:"Unit Air Purifier di ruang tunggu, ruang dokter, dan koridor",chemical_ratio:"-",tools:["Vacuum Cleaner (Nozzle Sikat Halus)","Kain Microfiber Kering","Kuas Cat Halus"],steps:["Matikan saklar power unit air purifier dan cabut kabel daya dari stop kontak dinding.","Buka penutup bodi (casing) depan atau belakang air purifier secara hati-hati.","Keluarkan pre-filter jaring nilon dari unit.","Bersihkan debu tebal pada pre-filter menggunakan vacuum cleaner nozzle sikat halus atau cuci dengan air mengalir (jika tipe wash-able).","Untuk filter HEPA dan Karbon Aktif: HANYA sedot debu permukaannya secara perlahan dengan vacuum, JANGAN DICUCI DENGAN AIR.","Lap bodi luar dan kisi-kisi udara masuk menggunakan lap microfiber kering.","Pastikan pre-filter benar-benar kering sebelum dipasang kembali ke posisinya semula."],notes:"Jangan pernah mencuci filter HEPA dengan air karena akan merusak serat mikro penyaring bakteri/virus."},{id:9,title:"Membersihkan Filter & Bodi AC",category:"Elektronik & Utility",icon:"\u2744\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Filter udara dan penutup indoor unit AC split di seluruh ruangan",chemical_ratio:"-",tools:["Tangga Besi Lipat","Kain Lap Microfiber Lembap","Kuas Cat Halus","Kran Air Mengalir"],steps:["Matikan AC melalui remote dan matikan MCB/saklar AC bila diperlukan.","Gunakan tangga besi lipat yang stabil di lantai datar.","Buka tutup cover depan unit indoor AC ke atas hingga terkunci.","Lepaskan kedua kisi filter saringan udara plastik secara perlahan agar debu tidak berjatuhan.","Cuci saringan filter di bawah kran air mengalir dari arah bagian dalam ke bagian luar.","Keringkan filter di tempat teduh (jangan dijemur langsung di bawah terik matahari).","Lap bodi casing dan louver/swing AC dengan kain microfiber lembap.","Pasang kembali filter yang sudah kering, tutup cover AC, dan nyalakan AC untuk pengetesan."],notes:"Pembersihan rutin mandiri dilakukan 2 minggu sekali. Deep cleaning / cuci evaporator dilakukan oleh teknisi AC vendor."},{id:10,title:"Membersihkan Dental Unit",category:"Klinis & Medis",icon:"\u{1F9B7}",color:"#E11D48",bg:"#FFF1F2",target:"Kursi dental pasien, spittoon bowl, meja instrumen, lampu dental, cover bawah",chemical_ratio:"Glass Cleaner (netral) / Karbol Disinfektan (mangkuk) / Bowl Cleaner (kerak)",tools:["Sarung Tangan Karet (Wajib)","Kain Lap Microfiber Biru (Khusus Bodi Luar)","Kain Lap Microfiber Merah (Khusus Mangkuk Spittoon)","Botol Spray Glass Cleaner (25ml + Air)","Botol Spray Karbol (20ml + Air)","Botol Spray Bowl Cleaner (100ml + Air, bila berkerak)"],steps:["Pastikan saklar utama dental unit dalam posisi OFF dan kursi dinaikkan ke posisi tertinggi untuk memudahkan akses.","Lap seluruh jok kulit, sandaran kepala, sandaran tangan, dan tiang lampu menggunakan lap microfiber biru dan cairan pembersih netral/glass cleaner.","Bersihkan mangkuk kumur (spittoon bowl) menggunakan kain lap merah dan cairan karbol disinfektan.","Jika ditemukan kerak sisa darah/plak membandel di mangkuk spittoon, gunakan sedikit bowl cleaner, sikat perlahan, lalu bilas hingga bersih.","Lap permukaan meja instrumen meja dokter. Bila ada instrumen steril, geser dengan hati-hati dan jangan mengotori instrumen.","Buka penutup cover bawah unit dental, lap debu yang terjangkau dan pastikan kondisi area selang dalam keadaan kering."],notes:"Kain lap warna MERAH hanya untuk mangkuk spittoon. Dilarang keras menggosok jok kulit dental unit dengan bahan kimia asam keras."},{id:11,title:"Membersihkan Wastafel & Keran Air",category:"Sanitasi Toilet",icon:"\u{1F6B0}",color:"#0284C7",bg:"#F0F9FF",target:"Wastafel keramik ruang tindakan, wastafel lorong, kran stainless steel, afur sink",chemical_ratio:"Bubuk Pembersih KIFA / Dishwashing Liquid Prokleen",tools:["Sponge Scotch Brite Hijau/Kuning","Kain Microfiber Hijau (Area Basah)","Bubuk Pembersih KIFA","Sikat Gigi Bekas / Sikat Kecil"],steps:["Basahi seluruh mangkuk wastafel keramik dan kran air stainless steel.","Taburkan sedikit Bubuk KIFA pada tapas sponge yang basah.","Gosok perlahan noda sabun dan kerak air putih pada kran stainless steel searah guratan logam hingga mengkilap kembali.","Gunakan sikat kecil untuk membersihkan sela leher kran dan sudut sambungan afur pembuangan wastafel.","Bilas mangkuk wastafel dan kran dengan air bersih mengalir hingga seluruh sisa residu bubuk KIFA hilang.","Keringkan kran air stainless steel dan cermin wastafel menggunakan lap microfiber hijau kering."],notes:"DILARANG KERAS menggunakan Toilet Bowl Cleaner pada kran air stainless steel karena menyebabkan karat dan warna menghitam permanen."},{id:12,title:"Membersihkan Kulkas Klinik & Pantry",category:"Elektronik & Utility",icon:"\u{1F9CA}",color:"#475569",bg:"#F8FAFC",target:"Kulkas pantry staf dan kulkas penyimpan obat/material medis (eksterior & rak)",chemical_ratio:"30 ml Dishwashing Liquid : 1 Liter Air Hangat + Lap Kering",tools:["Kain Microfiber Biru & Hijau","Dishwashing Liquid","Ember Air Hangat","Sponge Scotch Brite Halus"],steps:["Keluarkan makanan/minuman yang telah kedaluwarsa atau berbau tidak sedap.","Lepas rak-rak kaca/plastik interior kulkas yang dapat dilepas, cuci dengan sabun cuci piring dan bilas air bersih.","Seka dinding interior kulkas menggunakan kain microfiber yang dicelupkan ke dalam air hangat bersabun lembut.","Lap kering seluruh permukaan dalam kulkas dengan microfiber kering bersih.","Pasang kembali rak interior dan susun kembali bahan/makanan secara tertib dan rapi.","Bersihkan karet pintu kulkas dari noda hitam/jamur menggunakan sikat halus dan lap kering.","Lap bodi luar kulkas dan pegangan pintu menggunakan cairan pembersih netral."],notes:"Untuk kulkas penyimpanan obat/material dental, koordinasikan terlebih dahulu dengan Perawat/PIC Obat sebelum membersihkan rak interior."},{id:13,title:"Membersihkan Sofa & Kursi Ruang Tunggu",category:"Pembersihan Khusus",icon:"\u{1F6CB}\uFE0F",color:"#9333EA",bg:"#FAF5FF",target:"Sofa ruang tunggu pasien, kursi reception, kursi kerja tim",chemical_ratio:"Baking Soda (penghilang bau) / Pembersih Upholstery secukupnya",tools:["Vacuum Cleaner (Nozzle Kecil & Sikat Lembut)","Kain Microfiber Bersih","Sikat Bulu Halus","Ember Air Hangat","Cairan Pembersih Khusus Sofa (Upholstery Cleaner)","Baking Soda (Penghilang Bau Tidak Sedap)","Sarung Tangan Karet"],steps:["Kenali bahan sofa (kain fabric, beludru, kulit asli, atau kulit sintetis/oscar).","Sedot debu di seluruh permukaan, lipatan jahitan, dan celah bantal sofa menggunakan vacuum cleaner nozzle kecil.","Gunakan sikat bulu halus untuk mengangkat rambut atau kotoran yang menempel kuat di serat kain.","Jika ada noda bercak (kopi, makanan, tinta), semprotkan pembersih upholstery secukupnya, gosok memutar perlahan dengan microfiber.","Celupkan kain microfiber ke air hangat suam-suam kuku, peras hingga lembap, lalu seka seluruh permukaan sofa.","Lap kembali dengan kain kering bersih untuk menyerap sisa kelembapan.","Jika sofa berbau apek: taburkan baking soda tipis di permukaan, diamkan 20 menit, lalu sedot kembali dengan vacuum cleaner.","Biarkan sofa mengering secara alami di dalam ruangan berventilasi atau ber-AC."],notes:"Jangan menjemur sofa langsung di bawah terik matahari karena dapat menyebabkan bahan kulit pecah-pecah dan warna kain memudar."}],Dn=[{no:1,name:"Sapu Lidi Gagang",qty:"1 unit",area:"Area parkir mobil & motor klinik"},{no:2,name:"Sapu & Pengki Nagata 742s",qty:"1 set",area:"Area pantry, balkon, dak lantai atas, atau teras"},{no:3,name:"Lobby Duster Set 60 cm",qty:"1 set",area:"Lobby utama dan koridor/lorong klinik"},{no:4,name:"Lobby Duster Set 40 cm / Refill",qty:"1 unit",area:"Area ruang tindakan dokter"},{no:5,name:"Single Bucket Mop Wringer 20L",qty:"1 paket",area:"Ember pel dengan pemeras + tongkat pel"},{no:6,name:"Double Bucket Wringer Trolley 46L",qty:"1 paket",area:"Khusus klinik besar / klinik dengan fasilitas lift"},{no:7,name:"Kemoceng Bulu Lembut",qty:"1 unit",area:"Membersihkan debu kasar meja dan perabot"},{no:8,name:"Lap Microfiber Warna Merah",qty:"2 pcs",area:"Khusus area toilet, sanitasi, dan mangkuk spittoon"},{no:9,name:"Lap Microfiber Warna Biru",qty:"2 pcs",area:"Area umum kering (meja, kursi, bodi peralatan)"},{no:10,name:"Lap Microfiber Warna Hijau",qty:"2 pcs",area:"Area umum basah (pantry, wastafel luar)"},{no:11,name:"Kanebo / Synthetic Chamois",qty:"1 pcs",area:"Area umum, cermin, dan pengering kaca"},{no:12,name:"Caddy Bag Cleaning",qty:"1 unit",area:"Tas perlengkapan pinggang petugas cleaning"},{no:13,name:"Sponge Scotch Brite 3M",qty:"2 pcs",area:"Pembersihan noda lantai dan kerak ringan (spotting)"},{no:14,name:"Scotch-Brite 3M Sikat Toilet",qty:"1 unit",area:"Sikat khusus kloset dan leher angsa"},{no:15,name:"Kapi Gagang Karet / Scrapper",qty:"1 unit",area:"Pembersihan kerak lantai, noda permen karet, stiker"},{no:16,name:"Kuas Cat 4 Inch",qty:"2 pcs",area:"Pembersih sela keyboard, ventilasi, dan sudut sempit"},{no:17,name:"Warning Sign (Lantai Licin)",qty:"1 unit",area:"Digunakan saat mopping dan general cleaning"},{no:18,name:"Teleskopik + Pembersih Kaca/Langit",qty:"1 set",area:"Membersihkan jendela tinggi, plafon, dan sarang laba-laba"},{no:19,name:"Sarung Tangan Karet 3M",qty:"1 pasang",area:"Perlindungan tangan area toilet dan bahan kimia"},{no:20,name:"Botol Spray 500ml",qty:"5 unit",area:"Wadah larutan chemical (diberi label masing-masing)"},{no:21,name:"Pengharum Ruangan Stella Aerosol",qty:"1 kaleng",area:"Digunakan darurat apabila timbul bau menyengat"},{no:22,name:"Kleaner Gantungan Sapu & Pel",qty:"1 set",area:"Gantungan penyimpanan alat cleaning di janitor"},{no:23,name:"Floor Squeegee 75 cm",qty:"1 unit",area:"Pendorong genangan air lantai toilet/teras"},{no:24,name:"Sikat Botol Ikea",qty:"1 unit",area:"Di setiap ruang tindakan dokter gigi"},{no:25,name:"Scotch-Brite 3M Sikat Lantai ID-51",qty:"1 unit",area:"Sikat lantai kamar mandi dan teras"},{no:26,name:"Kamper Toilet / Urinal Block",qty:"1 pack",area:"Pengharum di toilet dan urinoir"},{no:27,name:"Plastik Sampah Hitam Sedang (60x100)",qty:"1 pack",area:"Tempat sampah indoor ruangan"},{no:28,name:"Plastik Sampah Hitam Besar (90x120)",qty:"1 pack",area:"Tempat sampah luar & penampung sampah akhir"},{no:29,name:"Plastik Sampah Medis Kuning (90x120)",qty:"1 pack",area:"Limbah medis padat B3 klinik"},{no:30,name:"Plastik Sampah Medis Kuning (40x60)",qty:"1 pack",area:"Tempat sampah medis kecil di ruang tindakan"},{no:31,name:"Tangga Besi Lipat",qty:"1 unit",area:"Alat bantu membersihkan plafon, lampu, dan AC"},{no:32,name:"Baskom Plastik Besar 55cm",qty:"1 unit",area:"Mencuci bola kids playground"}],Pn=[{no:1,name:"Floor Cleaner Prokleen Lavender",size:"5 Liter",ratio:"25 ml : 1 Liter Air",function:"Pembersihan dan pengepelan rutin seluruh lantai ruangan"},{no:2,name:"Karbol Prokleen Pine Disinfektan",size:"5 Liter",ratio:"25 ml : 1 Liter Air (Mop) / 20 ml per botol spray",function:"Mengepel lantai toilet, disinfeksi area basah dan berbau tak sedap"},{no:3,name:"Glass Cleaner Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air / Siap pakai di botol spray",function:"Membersihkan pintu kaca, jendela, cermin, dan layar monitor"},{no:4,name:"Manual Dishwashing Liquid Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air",function:"Mencuci piring, gelas, sendok pantry, dan mencuci bola playground"},{no:5,name:"Hand Soap Prokleen Green Tea",size:"5 Liter",ratio:"Siap pakai (tanpa campur)",function:"Refill sabun pencuci tangan di wastafel pasien dan tim"},{no:6,name:"Liquid Detergent Prokleen",size:"5 Liter",ratio:"Sesuai kapasitas mesin",function:"Mencuci handuk, kain microfiber, dan seragam"},{no:7,name:"Toilet Bowl Cleaner Prokleen",size:"5 Liter",ratio:"100 ml per botol spray / Khusus kerak",function:"Pembersih kerak membandel pada kloset, wastafel keramik, dan nat keramik"},{no:8,name:"Bubuk Pembersih KIFA",size:"500 gram",ratio:"Taburkan tipis pada tapas basah",function:"Membersihkan kerak air pada kran air stainless steel, jet washer, dan sink"}],St="fcms_sop_fc_ik_v1",Ct="fcms_sop_fc_tools_v1",Ft="fcms_sop_fc_chem_v1";function ma(){try{let e=localStorage.getItem(St);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK",e)}return JSON.parse(JSON.stringify(Tn))}function kt(e){localStorage.setItem(St,JSON.stringify(e))}function ba(){try{let e=localStorage.getItem(Ct);if(e)return JSON.parse(e)}catch(e){console.error("Error load tools",e)}return JSON.parse(JSON.stringify(Dn))}function wt(e){localStorage.setItem(Ct,JSON.stringify(e))}function ha(){try{let e=localStorage.getItem(Ft);if(e)return JSON.parse(e)}catch(e){console.error("Error load chem",e)}return JSON.parse(JSON.stringify(Pn))}function Et(e){localStorage.setItem(Ft,JSON.stringify(e))}async function fa(e){let t="ik",i="",n="all",c="tools",x="",m=ma(),u=ba(),o=ha();function b(){e.innerHTML=`
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
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto; flex-wrap:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none;">
        <button class="sop-tab-btn ${t==="ik"?"active":""}" data-tab="ik" style="${B(t==="ik")}">
          \u{1F4CB} Instruksi Kerja (IK) <span class="badge" style="${D(t==="ik")}">${m.length}</span>
        </button>
        <button class="sop-tab-btn ${t==="tools"?"active":""}" data-tab="tools" style="${B(t==="tools")}">
          \u{1F6E0}\uFE0F Alat & Chemical <span class="badge" style="${D(t==="tools")}">${u.length+o.length}</span>
        </button>
        <button class="sop-tab-btn ${t==="policy"?"active":""}" data-tab="policy" style="${B(t==="policy")}">
          \u{1F4D6} Pedoman & Ketentuan
        </button>
        <button class="sop-tab-btn ${t==="links"?"active":""}" data-tab="links" style="${B(t==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sop-content-area"></div>
    `,S(),F()}function B(l){return l?"border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function D(l){return l?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function S(){e.querySelectorAll(".sop-tab-btn").forEach(L=>{L.addEventListener("click",()=>{t=L.dataset.tab,b()})});let l=e.querySelector("#btn-reset-default");l&&l.addEventListener("click",()=>{X("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja, Alat, dan Chemical ke data standar dokumen asli?",()=>{localStorage.removeItem(St),localStorage.removeItem(Ct),localStorage.removeItem(Ft),m=ma(),u=ba(),o=ha(),R("Data SOP Facility Care berhasil di-reset ke standar asli"),b()},"Konfirmasi Reset Standar")});let f=e.querySelector("#btn-print-sop");f&&f.addEventListener("click",()=>{window.print()});let v=e.querySelector("#btn-export-tools");v&&v.addEventListener("click",()=>{s()})}function F(){let l=e.querySelector("#sop-content-area");l&&(t==="ik"?$(l):t==="tools"?h(l):t==="policy"?d(l):t==="links"&&g(l))}function $(l){l.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="ik-search" class="form-control" placeholder="Cari pekerjaan, alat, atau chemical..." value="${i}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="cat-pills-wrap">
              <button class="cat-pill ${n==="all"?"active":""}" data-cat="all">Semua (${m.length})</button>
              <button class="cat-pill ${n==="Pembersihan Rutin"?"active":""}" data-cat="Pembersihan Rutin">Rutin (${m.filter(I=>I.category==="Pembersihan Rutin").length})</button>
              <button class="cat-pill ${n==="Sanitasi Toilet"?"active":""}" data-cat="Sanitasi Toilet">Toilet & Wastafel (${m.filter(I=>I.category==="Sanitasi Toilet").length})</button>
              <button class="cat-pill ${n==="Klinis & Medis"?"active":""}" data-cat="Klinis & Medis">Klinis Dental (${m.filter(I=>I.category==="Klinis & Medis").length})</button>
              <button class="cat-pill ${n==="Elektronik & Utility"?"active":""}" data-cat="Elektronik & Utility">Elektronik & AC (${m.filter(I=>I.category==="Elektronik & Utility").length})</button>
              <button class="cat-pill ${n==="Pembersihan Khusus"?"active":""}" data-cat="Pembersihan Khusus">Khusus (${m.filter(I=>I.category==="Pembersihan Khusus").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${m.length} dari ${m.length} Instruksi Kerja
            </div>
            <button id="btn-add-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
              \u2795 Tambah IK Baru
            </button>
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
    `,l.appendChild(f);let v=l.querySelector("#ik-cards-grid-container"),L=l.querySelector("#ik-count-label");function A(){let I=m.filter(z=>{let C=n==="all"||z.category===n,y=i.toLowerCase().trim(),_=!y||z.title.toLowerCase().includes(y)||z.target.toLowerCase().includes(y)||z.tools.some(T=>T.toLowerCase().includes(y))||z.steps.some(T=>T.toLowerCase().includes(y));return C&&_});if(L&&(L.textContent=`Menampilkan ${I.length} dari ${m.length} Instruksi Kerja`),I.length===0){v.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}v.innerHTML=`
        <div class="sop-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${I.map(z=>`
            <div class="card ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${z.bg||"#EFF6FF"}; border-radius:10px;">${z.icon||"\u{1F4CB}"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        ${z.id}. ${z.title}
                      </h3>
                      <span class="badge" style="background:${z.bg||"#EFF6FF"}; color:${z.color||"#2563EB"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${z.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    <button class="btn btn-sm btn-ghost btn-edit-ik" data-id="${z.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    <button class="btn btn-sm btn-ghost btn-delete-ik" data-id="${z.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${z.color||"#2563EB"};">
                  <strong>Area:</strong> ${z.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Bahan Utama:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${z.tools.slice(0,3).map(C=>`<li style="margin-bottom:2px;">${C}</li>`).join("")}
                    ${z.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${z.tools.length-3} alat lainnya...</li>`:""}
                  </ul>
                </div>

                ${z.chemical_ratio&&z.chemical_ratio!=="-"?`
                  <div style="font-size:0.8rem; background:#FEF3C7; color:#92400E; padding:6px 10px; border-radius:6px; margin-bottom:14px; font-weight:600;">
                    \u{1F9EA} Takaran: ${z.chemical_ratio}
                  </div>
                `:""}
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${z.steps.length} Langkah Kerja
                </span>
                <div style="display:flex; gap:6px;">
                  <button class="btn btn-sm btn-primary btn-view-ik" data-id="${z.id}" style="padding:6px 14px; font-weight:600; border-radius:6px;">
                    Lihat Panduan \u2794
                  </button>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,v.querySelectorAll(".btn-view-ik").forEach(z=>{z.addEventListener("click",()=>{let C=parseInt(z.dataset.id,10),y=m.find(_=>_.id===C);y&&E(y)})}),v.querySelectorAll(".btn-edit-ik").forEach(z=>{z.addEventListener("click",()=>{let C=parseInt(z.dataset.id,10),y=m.find(_=>_.id===C);y&&k(y)})}),v.querySelectorAll(".btn-delete-ik").forEach(z=>{z.addEventListener("click",()=>{let C=parseInt(z.dataset.id,10),y=m.find(_=>_.id===C);y&&X(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${y.title}"?`,()=>{m=m.filter(_=>_.id!==C),kt(m),R("Instruksi Kerja berhasil dihapus"),b()},"Hapus Instruksi Kerja")})})}A();let P=l.querySelector("#ik-search");P&&P.addEventListener("input",I=>{i=I.target.value,A()}),l.querySelectorAll(".cat-pill").forEach(I=>{I.addEventListener("click",()=>{l.querySelectorAll(".cat-pill").forEach(z=>z.classList.remove("active")),I.classList.add("active"),n=I.dataset.cat,A()})});let w=l.querySelector("#btn-add-ik");w&&w.addEventListener("click",()=>{k()})}function E(l){G({title:`${l.icon||"\u{1F4CB}"} IK-${l.id}: ${l.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${l.bg||"#EFF6FF"}; color:${l.color||"#2563EB"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${l.category}</span>
            <span>ID: IK-${String(l.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${l.color||"#2563EB"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${l.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Bahan yang Dibutuhkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${l.tools.map(f=>`<li>${f}</li>`).join("")}
              </ul>
            </div>
          </div>

          ${l.chemical_ratio&&l.chemical_ratio!=="-"?`
            <div style="background:#FEF3C7; border:1px solid #FCD34D; border-radius:8px; padding:10px 14px;">
              <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:2px;">\u{1F9EA} Rasio & Takaran Chemical:</strong>
              <span style="color:#78350F; font-size:0.9rem; font-weight:600;">${l.chemical_ratio}</span>
            </div>
          `:""}

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah-Langkah Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${l.steps.map(f=>`<li style="margin-bottom:6px;">${f}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${l.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Penting & Safety:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${l.notes}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(f,v)=>{v()}})}function k(l=null){let f=!!l,v=l?l.tools.join(`
`):"",L=l?l.steps.join(`
`):"",A=`
      <form id="ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-ik-title" class="form-control" required value="${l?r(l.title):""}" placeholder="Contoh: Membersihkan Dental Unit" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-ik-icon" class="form-control" value="${l?l.icon:"\u{1F4CB}"}" placeholder="\u{1F9F9}" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
            <select id="form-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
              <option value="Pembersihan Rutin" ${l?.category==="Pembersihan Rutin"?"selected":""}>Pembersihan Rutin</option>
              <option value="Sanitasi Toilet" ${l?.category==="Sanitasi Toilet"?"selected":""}>Sanitasi Toilet</option>
              <option value="Klinis & Medis" ${l?.category==="Klinis & Medis"?"selected":""}>Klinis & Medis</option>
              <option value="Elektronik & Utility" ${l?.category==="Elektronik & Utility"?"selected":""}>Elektronik & Utility</option>
              <option value="Pembersihan Khusus" ${l?.category==="Pembersihan Khusus"?"selected":""}>Pembersihan Khusus</option>
            </select>
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Takaran / Rasio Chemical</label>
            <input type="text" id="form-ik-ratio" class="form-control" value="${l?r(l.chemical_ratio||""):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-ik-target" class="form-control" required value="${l?r(l.target):""}" placeholder="Contoh: Kursi pasien, meja instrumen, mangkuk spittoon" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Bahan yang Digunakan (1 item per baris) *</label>
          <textarea id="form-ik-tools" class="form-control" rows="4" placeholder="Kain Microfiber Biru&#10;Sarung Tangan Karet&#10;Botol Spray Glass Cleaner" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${r(v)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah-Langkah Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-ik-steps" class="form-control" rows="6" placeholder="1. Pastikan posisi saklar OFF.&#10;2. Lap jok dental unit dengan kain microfiber biru.&#10;3. Sikat mangkuk spittoon..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${r(L)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Penting / Hal yang Perlu Diperhatikan (Safety)</label>
          <textarea id="form-ik-notes" class="form-control" rows="2" placeholder="Catatan larangan atau peringatan keselamatan..." style="width:100%; border-radius:6px; font-size:0.85rem;">${l?r(l.notes||""):""}</textarea>
        </div>
      </form>
    `;G({title:f?`\u270F\uFE0F Edit IK-${l.id}: ${l.title}`:"\u2795 Tambah Instruksi Kerja Baru",content:A,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(P,w)=>{let I=P.querySelector("#form-ik-title").value.trim(),z=P.querySelector("#form-ik-icon").value.trim()||"\u{1F4CB}",C=P.querySelector("#form-ik-cat").value,y=P.querySelector("#form-ik-ratio").value.trim(),_=P.querySelector("#form-ik-target").value.trim(),T=P.querySelector("#form-ik-tools").value.trim(),j=P.querySelector("#form-ik-steps").value.trim(),O=P.querySelector("#form-ik-notes").value.trim();if(!I||!_||!T||!j){N("Harap lengkapi seluruh field wajib (*)");return}let M=T.split(`
`).map(W=>W.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),q=j.split(`
`).map(W=>W.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),Y={"Pembersihan Rutin":{color:"#2563EB",bg:"#EFF6FF"},"Sanitasi Toilet":{color:"#DC2626",bg:"#FEF2F2"},"Klinis & Medis":{color:"#E11D48",bg:"#FFF1F2"},"Elektronik & Utility":{color:"#0D9488",bg:"#F0FDFA"},"Pembersihan Khusus":{color:"#7C3AED",bg:"#F5F3FF"}}[C]||{color:"#2563EB",bg:"#EFF6FF"};if(f)l.title=I,l.icon=z,l.category=C,l.chemical_ratio=y,l.target=_,l.tools=M,l.steps=q,l.notes=O,l.color=Y.color,l.bg=Y.bg,kt(m),R(`Instruksi Kerja "${I}" berhasil diperbarui`);else{let W=m.length>0?Math.max(...m.map(U=>U.id))+1:1,Z={id:W,title:I,category:C,icon:z,color:Y.color,bg:Y.bg,target:_,chemical_ratio:y||"-",tools:M,steps:q,notes:O};m.push(Z),kt(m),R(`Instruksi Kerja baru IK-${W} berhasil ditambahkan`)}w(),b()}})}function h(l){l.innerHTML=`
      <!-- Sub-tab Switcher & Actions -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-tools-btn" class="btn btn-sm ${c==="tools"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9F9} Alat Kebersihan (${u.length} Item)
          </button>
          <button id="view-chem-btn" class="btn btn-sm ${c==="chemical"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9EA} Bahan Chemical (${o.length} Item)
          </button>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="tool-search" class="form-control" placeholder="Cari nama / peruntukan / rasio..." value="${x}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
          <button id="btn-add-tool-chem" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u2795 ${c==="tools"?"Tambah Alat Baru":"Tambah Chemical Baru"}
          </button>
        </div>
      </div>

      <!-- Table Container -->
      <div id="tool-table-container"></div>
    `;let f=l.querySelector("#tool-table-container"),v=l.querySelector("#tool-search");function L(){let A=x.toLowerCase().trim();if(c==="tools"){let P=u.filter(w=>!A||w.name.toLowerCase().includes(A)||w.area.toLowerCase().includes(A));f.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Standar Peralatan Kebersihan (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-primary">${P.length} Item</span>
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
                  ${P.map((w,I)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${I%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${w.no||I+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${r(w.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#E0F2FE; color:#0369A1; font-weight:600;">${r(w.qty)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${r(w.area)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                        <button class="btn btn-sm btn-ghost btn-edit-tool" data-no="${w.no}" title="Edit Alat" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                        <button class="btn btn-sm btn-ghost btn-delete-tool" data-no="${w.no}" title="Hapus Alat" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,f.querySelectorAll(".btn-edit-tool").forEach(w=>{w.addEventListener("click",()=>{let I=parseInt(w.dataset.no,10),z=u.find(C=>C.no===I);z&&a(z)})}),f.querySelectorAll(".btn-delete-tool").forEach(w=>{w.addEventListener("click",()=>{let I=parseInt(w.dataset.no,10),z=u.find(C=>C.no===I);z&&X(`Apakah Anda yakin ingin menghapus alat: "${z.name}"?`,()=>{u=u.filter(C=>C.no!==I),u.forEach((C,y)=>{C.no=y+1}),wt(u),R("Alat kebersihan berhasil dihapus"),b()},"Hapus Alat Kebersihan")})})}else{let P=o.filter(w=>!A||w.name.toLowerCase().includes(A)||w.function.toLowerCase().includes(A));f.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Bahan Kimia & Standar Takaran (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-success">${P.length} Chemical</span>
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
                  ${P.map((w,I)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${I%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${w.no||I+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${r(w.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#FEF3C7; color:#92400E; font-weight:600;">${r(w.size)}</span></td>
                      <td style="padding:10px 14px;"><span style="background:#EFF6FF; color:#1E40AF; padding:3px 8px; border-radius:6px; font-weight:600; font-size:0.82rem;">${r(w.ratio)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${r(w.function)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                        <button class="btn btn-sm btn-ghost btn-edit-chem" data-no="${w.no}" title="Edit Chemical" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                        <button class="btn btn-sm btn-ghost btn-delete-chem" data-no="${w.no}" title="Hapus Chemical" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,f.querySelectorAll(".btn-edit-chem").forEach(w=>{w.addEventListener("click",()=>{let I=parseInt(w.dataset.no,10),z=o.find(C=>C.no===I);z&&p(z)})}),f.querySelectorAll(".btn-delete-chem").forEach(w=>{w.addEventListener("click",()=>{let I=parseInt(w.dataset.no,10),z=o.find(C=>C.no===I);z&&X(`Apakah Anda yakin ingin menghapus chemical: "${z.name}"?`,()=>{o=o.filter(C=>C.no!==I),o.forEach((C,y)=>{C.no=y+1}),Et(o),R("Bahan chemical berhasil dihapus"),b()},"Hapus Bahan Chemical")})})}}L(),l.querySelector("#view-tools-btn").addEventListener("click",()=>{c="tools",h(l)}),l.querySelector("#view-chem-btn").addEventListener("click",()=>{c="chemical",h(l)}),v&&v.addEventListener("input",A=>{x=A.target.value,L()}),l.querySelector("#btn-add-tool-chem").addEventListener("click",()=>{c==="tools"?a():p()})}function a(l=null){let f=!!l,v=`
      <form id="tool-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Alat / Item *</label>
          <input type="text" id="form-tool-name" class="form-control" required value="${l?r(l.name):""}" placeholder="Contoh: Lobby Duster Set 60 cm" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kuantitas Standar *</label>
          <input type="text" id="form-tool-qty" class="form-control" required value="${l?r(l.qty):"1 unit"}" placeholder="Contoh: 1 set, 2 pcs, 1 paket" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Peruntukan Area & Kegunaan *</label>
          <textarea id="form-tool-area" class="form-control" rows="3" required placeholder="Contoh: Lobby utama dan koridor/lorong klinik" style="width:100%; border-radius:6px; font-size:0.85rem;">${l?r(l.area):""}</textarea>
        </div>
      </form>
    `;G({title:f?`\u270F\uFE0F Edit Alat: ${l.name}`:"\u2795 Tambah Alat Kebersihan Baru",content:v,confirmText:"Simpan",onConfirm:(L,A)=>{let P=L.querySelector("#form-tool-name").value.trim(),w=L.querySelector("#form-tool-qty").value.trim(),I=L.querySelector("#form-tool-area").value.trim();if(!P||!w||!I){N("Harap lengkapi semua kolom");return}if(f)l.name=P,l.qty=w,l.area=I,wt(u),R(`Alat "${P}" berhasil diperbarui`);else{let z=u.length>0?Math.max(...u.map(C=>C.no))+1:1;u.push({no:z,name:P,qty:w,area:I}),wt(u),R(`Alat baru "${P}" berhasil ditambahkan`)}A(),b()}})}function p(l=null){let f=!!l,v=`
      <form id="chem-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Produk Chemical *</label>
          <input type="text" id="form-chem-name" class="form-control" required value="${l?r(l.name):""}" placeholder="Contoh: Floor Cleaner Prokleen Lavender" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kemasan Standar *</label>
            <input type="text" id="form-chem-size" class="form-control" required value="${l?r(l.size):"5 Liter"}" placeholder="Contoh: 5 Liter, 500 gram" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Rasio & Takaran Standar *</label>
            <input type="text" id="form-chem-ratio" class="form-control" required value="${l?r(l.ratio):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Fungsi & Cara Penggunaan *</label>
          <textarea id="form-chem-function" class="form-control" rows="3" required placeholder="Contoh: Pembersihan dan pengepelan rutin seluruh lantai ruangan" style="width:100%; border-radius:6px; font-size:0.85rem;">${l?r(l.function):""}</textarea>
        </div>
      </form>
    `;G({title:f?`\u270F\uFE0F Edit Chemical: ${l.name}`:"\u2795 Tambah Chemical Baru",content:v,confirmText:"Simpan",onConfirm:(L,A)=>{let P=L.querySelector("#form-chem-name").value.trim(),w=L.querySelector("#form-chem-size").value.trim(),I=L.querySelector("#form-chem-ratio").value.trim(),z=L.querySelector("#form-chem-function").value.trim();if(!P||!w||!I||!z){N("Harap lengkapi semua kolom");return}if(f)l.name=P,l.size=w,l.ratio=I,l.function=z,Et(o),R(`Chemical "${P}" berhasil diperbarui`);else{let C=o.length>0?Math.max(...o.map(y=>y.no))+1:1;o.push({no:C,name:P,size:w,ratio:I,function:z}),Et(o),R(`Chemical baru "${P}" berhasil ditambahkan`)}A(),b()}})}function d(l){l.innerHTML=`
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
    `}function g(l){l.innerHTML=`
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
    `}async function s(){try{let{downloadExcel:l}=await Promise.resolve().then(()=>(re(),he)),f=m.map(L=>({"No IK":`IK-${L.id}`,"Judul Pekerjaan":L.title,Kategori:L.category,"Sasaran Target Area":L.target,"Rasio Chemical":L.chemical_ratio||"-","Alat & Bahan":L.tools.join(", "),"Langkah Kerja":L.steps.map((A,P)=>`${P+1}. ${A}`).join(" | "),"Catatan Safety":L.notes||"-"})),v=new Date().toISOString().slice(0,10);l(f,`SOP_Facility_Care_Instruksi_Kerja_${v}`),R("Data Excel Instruksi Kerja berhasil diunduh")}catch(l){console.error(l),N("Gagal mendownload data Excel: "+l.message)}}function r(l){return l?String(l).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}b()}le();var Ln=[{id:1,title:"Pengamanan Area & Aset Perusahaan",category:"Pengamanan Fisik & Patroli",icon:"\u{1F6E1}\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Seluruh area dalam & luar Head Quarter dan Klinik Cabang Audy Dental",tools:["Google Form Checklist Monitoring Klinik: https://forms.gle/1EAvkqDJVdctXEGn7","Senter Patroli","Komputer dengan Akses Monitor CCTV (Live View)","Alat Komunikasi / Handphone / HT"],steps:["Lakukan Safety Patrol secara rutin mengelilingi seluruh area sesuai jadwal yang telah ditentukan (Patrol 1, 2, dan 3).","Catat hasil temuan dan kondisi keamanan secara berkala ke dalam Google Form Safety Patrol resmi.","Awasi kendaraan yang terparkir di area luar klinik; pastikan khusus sepeda motor dalam keadaan terkunci ganda.","Pantau situasi lingkungan dan area sensitif secara berkala melalui layar monitor CCTV saat berada di pos security.","Segera lakukan tindakan persuasif ramah apabila ditemukan potensi gangguan keamanan, dan segera hubungi pihak Kepolisian terdekat bila situasi darurat tidak dapat diatasi.","Bila ditemukan permasalahan fisik, kerusakan fasilitas, atau kendala di area klinik, segera dokumentasikan (foto) dan laporkan langsung ke SPV Klinik dan Leader Facility Care.","Disaat malam hari atau saat klinik telah tutup operasional, pastikan seluruh AC dan lampu ruangan dalam keadaan mati, kecuali lampu keamanan tertentu yang memang harus dibiarkan menyala terang."],notes:"Prioritaskan keselamatan diri dan tindakan persuasif. Dilarang meninggalkan area pengawasan tanpa personil pengganti yang jelas."},{id:2,title:"Penertiban & Pengaturan Parkir Kendaraan",category:"Operasional Parkir & Luar",icon:"\u{1F17F}\uFE0F",color:"#D97706",bg:"#FFFBEB",target:"Area parkir mobil, motor, akses keluar masuk, dan teras depan klinik",tools:["Lampu Lalu Lintas (Lampu Lalin)","Cone / Kerucut Pembatas Parkir","Peluit Parkir"],steps:["Berdiri di posisi strategis dan bantu mengarahkan setiap kendaraan pasien atau tamu yang datang secara tepat dan rapi ke slot parkir yang tersedia.","Berikan informasi dan solusi dengan sikap ramah serta santun apabila seluruh slot parkir klinik dalam kondisi penuh.","Dalam kondisi tertentu (bila ada mobil yang terhalang mobil lain dan harus dipindahkan, sementara pasien sedang dalam tindakan perawatan): bantu memindahkan kendaraan HANYA JIKA petugas security memiliki kemampuan mengemudi dan memiliki SIM aktif.","Apabila petugas security tidak memiliki kemampuan mengemudikan mobil: sampaikan permohonan maaf dengan tulus kepada pasien/tamu, minta untuk menunggu sejenak, dan segera minta bantuan kepada staf Resepsionis untuk menghubungi pemilik mobil yang menghalangi.","Bantu pasien atau tamu yang akan keluar parkir dengan memberikan aba-aba dan arahan lalu lintas jalan raya yang aman dan tepat."],notes:"Dilarang memaksakan memindahkan mobil pasien jika tidak memiliki lisensi mengemudi demi menghindari risiko gesekan atau kerusakan kendaraan."},{id:3,title:"Penerimaan Tamu & Pasien Klinik",category:"Layanan & Hospitalitas",icon:"\u{1F91D}",color:"#059669",bg:"#ECFDF5",target:"Pintu masuk utama klinik, lobby, area resepsionis, dan pos penerimaan tamu",tools:["Alat Tulis Kantor (Pulpen)","Buku Tamu / Log Book Keamanan","Alat Komunikasi / Handphone"],steps:['Berdiri dengan tegap dan sigap, bukakan pintu untuk setiap pasien/tamu yang datang dengan menerapkan budaya 5S (Senyum, Salam, Sapa, Sopan, Santun) dan menyapa: "Selamat pagi / siang / sore, ada yang bisa kami bantu?".',"Setelah pasien atau tamu memberitahukan tujuannya, persilakan dan arahkan pasien langsung menuju meja Resepsionis.","Apabila yang datang adalah tamu instansi, vendor, atau pihak Manajemen/BOD: persilakan mengisi Buku Tamu (Log Book) terlebih dahulu dan segera hubungi staf/karyawan yang ingin ditemui.",'Apabila pasien langsung melangkah ke arah tujuannya, berikan gestur tubuh hormat menyilangkan kedua tangan di depan, dan jika pasien/tamu mengucapkan terima kasih, jawab dengan ramah: "Dengan Senang Hati".',"Dalam kondisi khusus (pasien lansia, ibu hamil, atau pengguna kursi roda): petugas security wajib mendampingi, mengawal, dan membantu mendorong kursi roda sampai ke depan meja resepsionis.",'Buka kembali pintu untuk setiap pasien yang telah selesai melakukan perawatan saat hendak keluar klinik sambil mengucapkan ucapan hangat: "Terima Kasih, semoga lekas sehat kembali".'],notes:"Jaga senyum tulus dan keramahan sebagai garda terdepan representasi citra Audy Dental."}],In=[{no:1,type:"Aset Tetap",item:"Rumah & Bangunan",description:"Gedung klinik, pos satpam, pagar, kanopi, teras beton",pic:"Building Maintenance & Security"},{no:2,type:"Aset Tetap",item:"Genset & Lift",description:"Mesin generator listrik cadangan dan elevator penumpang tetap",pic:"Teknisi & Security"},{no:3,type:"Aset Tetap",item:"AC Central & Split Tetap",description:"Outdoor dan indoor AC yang terpasang permanen di dinding/plafon",pic:"Teknisi AC & Security"},{no:4,type:"Aset Tetap",item:"Dental Unit Permanen",description:"Kursi dental unit terpasang instalasi air dan kompresor",pic:"Tim Klinik & Security"},{no:5,type:"Aset Tetap",item:"Instalasi Listrik, Pipa & Jaringan IT Permanen",description:"Panel MCB, kabel fiber optik, pipa drainase, server rack",pic:"IT & Teknisi"},{no:6,type:"Aset Tetap",item:"Furniture Melekat (Built-in)",description:"Meja resepsionis cor/melekat, partisi kaca mati, kitchen set pantry",pic:"General Support"},{no:7,type:"Aset Bergerak",item:"Komputer, Laptop & Printer",description:"PC kasir/resepsionis, laptop staf, printer rontgen/dokumen",pic:"Staf Terkait & IT"},{no:8,type:"Aset Bergerak",item:"Kendaraan Operasional",description:"Mobil operasional dan sepeda motor inventaris perusahaan",pic:"Security & Driver"},{no:9,type:"Aset Bergerak",item:"Furniture Lepasan",description:"Meja kerja, kursi resepsionis, kursi tunggu pasien, lemari file",pic:"Facility Care & Staf"},{no:10,type:"Aset Bergerak",item:"Hand Tools, Alat Kerja & Alat Dental Portabel",description:"Handpiece dental, alat rontgen portable, toolbox teknisi",pic:"Dokter/Perawat & Teknisi"},{no:11,type:"Low Value Asset",item:"Peralatan Kantor (Telepon, ATK)",description:"Pesawat telepon, kalkulator, stapler, perlengkapan meja",pic:"Resepsionis"},{no:12,type:"Low Value Asset",item:"Alat-alat Kebersihan",description:"Lobby duster, ember wringer pel, sapu, tangga, botol spray",pic:"Facility Care"},{no:13,type:"Low Value Asset",item:"Perlengkapan Security",description:"Lampu lalin, cone parkir, peluit, senter patroli, jas hujan",pic:"Security"}],xa={cabang:[{zone:"Zona Hijau",level:"Akses Publik Bebas",color:"#059669",bg:"#ECFDF5",rooms:["Area Parkir Kendaraan","Lobby Lantai 1","Teras Luar Klinik"],desc:"Area dengan pengamanan umum yang bebas diakses pasien dan tamu setelah pengarahan security."},{zone:"Zona Kuning",level:"Akses Terbatas / Pendampingan",color:"#D97706",bg:"#FFFBEB",rooms:["Koridor Ruang Tindakan","Tangga & Lift Klinik","Ruang Tunggu Lantai 2 & 3","Kids Playground","Toilet Lantai 1 & 2","Mushola Staf & Pasien"],desc:"Area terbatas, tamu wajib didampingi petugas atau staf klinik terkait."},{zone:"Zona Merah",level:"Area Terlarang / Restricted",color:"#DC2626",bg:"#FEF2F2",rooms:["Area Meja Receptionist / Kasir","Seluruh Ruang Dental Unit (DU 101, 102, 201, 202, 301, 302)","Ruang Sterilisasi Alat","Ruang Rontgen Radiologi","Ruang Suction & Kompresor","Ruang Tim Staf & Gudang Material Medis"],desc:"Area sangat sensitif, hanya staf medis bertugas dan atasan pemilik area. Tamu wajib izin Manajemen/BOD dan didampingi penuh."}],hq:[{floor:"Lantai 1",hijau:["Area Parkir Mobil/Motor","Lobby Resepsionis Lantai 1"],kuning:["Seluruh Ruang Meeting Lantai 1","Ruang Training Lantai 1","Koridor Lift & Toilet"],merah:["Ruang Customer Services","Gudang Utama Lantai 1","Gudang Facility Care","Gudang Teknisi & Ruang Teknisi","Pos Satpam & Ruang Genset"]},{floor:"Lantai 2",hijau:["-"],kuning:["Ruang Meeting Lantai 2","Koridor Lift & Toilet","Pantry"],merah:["Ruang CEO","Ruang Sekretaris CEO","Ruang Human Capital","Ruang Dentist Relation","Ruang General Support","Ruang Eks. Patrel"]},{floor:"Lantai 3",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry","Mushola & Tempat Wudhu","Ruang Santai Karyawan"],merah:["Ruang Marketing","Studio Konten Marketing","Gudang Marketing","Ruang Patient Relation","Ruang OPS","Ruang Training Lantai 3","Gudang Lantai 3"]},{floor:"Lantai 4",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry"],merah:["Ruang Finance & Accounting","Ruang VP Finance & Accounting","Ruang Audit Internal","Ruang Training Lantai 4","Gudang Lantai 4"]},{floor:"Lantai 5",hijau:["-"],kuning:["Lobby Lantai 5","Pantry Lantai 5","Toilet Lantai 5","Ruang Training Besar Lantai 5"],merah:["Area Dak Toren Air & Rooftop"]}]},At="fcms_sop_sec_ik_v1",Bt="fcms_sop_sec_assets_v1";function ya(){try{let e=localStorage.getItem(At);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK Sec",e)}return JSON.parse(JSON.stringify(Ln))}function $t(e){localStorage.setItem(At,JSON.stringify(e))}function va(){try{let e=localStorage.getItem(Bt);if(e)return JSON.parse(e)}catch(e){console.error("Error load assets sec",e)}return JSON.parse(JSON.stringify(In))}function _t(e){localStorage.setItem(Bt,JSON.stringify(e))}async function ka(e){let t="ik",i="",n="all",c="",x="cabang",m=ya(),u=va();function o(){e.innerHTML=`
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
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto; flex-wrap:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none;">
        <button class="sec-tab-btn ${t==="ik"?"active":""}" data-tab="ik" style="${b(t==="ik")}">
          \u{1F6E1}\uFE0F Instruksi Kerja (IK) <span class="badge" style="${B(t==="ik")}">${m.length}</span>
        </button>
        <button class="sec-tab-btn ${t==="zonasi"?"active":""}" data-tab="zonasi" style="${b(t==="zonasi")}">
          \u{1F5FA}\uFE0F Zonasi Akses & Denah
        </button>
        <button class="sec-tab-btn ${t==="assets"?"active":""}" data-tab="assets" style="${b(t==="assets")}">
          \u{1F4E6} Klasifikasi Aset <span class="badge" style="${B(t==="assets")}">${u.length}</span>
        </button>
        <button class="sec-tab-btn ${t==="policy"?"active":""}" data-tab="policy" style="${b(t==="policy")}">
          \u{1F4CB} Shift, Patrol & Insiden
        </button>
        <button class="sec-tab-btn ${t==="links"?"active":""}" data-tab="links" style="${b(t==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sec-content-area"></div>
    `,D(),S()}function b(r){return r?"border:none; background:#1E40AF; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function B(r){return r?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function D(){e.querySelectorAll(".sec-tab-btn").forEach(v=>{v.addEventListener("click",()=>{t=v.dataset.tab,o()})});let r=e.querySelector("#btn-reset-sec-default");r&&r.addEventListener("click",()=>{X("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja dan Klasifikasi Aset Security ke standar dokumen asli?",()=>{localStorage.removeItem(At),localStorage.removeItem(Bt),m=ya(),u=va(),R("Data SOP Security berhasil di-reset ke standar dokumen"),o()},"Reset ke Standar Dokumen")});let l=e.querySelector("#btn-print-sec");l&&l.addEventListener("click",()=>{window.print()});let f=e.querySelector("#btn-export-sec");f&&f.addEventListener("click",()=>{g()})}function S(){let r=e.querySelector("#sec-content-area");r&&(t==="ik"?F(r):t==="zonasi"?k(r):t==="assets"?h(r):t==="policy"?p(r):t==="links"&&d(r))}function F(r){r.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="sec-ik-search" class="form-control" placeholder="Cari instruksi kerja security..." value="${i}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="sec-cat-pills-wrap">
              <button class="sec-cat-pill ${n==="all"?"active":""}" data-cat="all">Semua (${m.length})</button>
              <button class="sec-cat-pill ${n==="Pengamanan Fisik & Patroli"?"active":""}" data-cat="Pengamanan Fisik & Patroli">Patroli & Fisik (${m.filter(w=>w.category==="Pengamanan Fisik & Patroli").length})</button>
              <button class="sec-cat-pill ${n==="Operasional Parkir & Luar"?"active":""}" data-cat="Operasional Parkir & Luar">Parkir & Luar (${m.filter(w=>w.category==="Operasional Parkir & Luar").length})</button>
              <button class="sec-cat-pill ${n==="Layanan & Hospitalitas"?"active":""}" data-cat="Layanan & Hospitalitas">Tamu & Pasien (${m.filter(w=>w.category==="Layanan & Hospitalitas").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="sec-ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${m.length} Instruksi Kerja
            </div>
            <button id="btn-add-sec-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
              \u2795 Tambah IK Baru
            </button>
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="sec-ik-cards-grid-container"></div>
    `;let l=document.createElement("style");l.innerHTML=`
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
    `,r.appendChild(l);let f=r.querySelector("#sec-ik-cards-grid-container"),v=r.querySelector("#sec-ik-count-label");function L(){let w=m.filter(I=>{let z=n==="all"||I.category===n,C=i.toLowerCase().trim(),y=!C||I.title.toLowerCase().includes(C)||I.target.toLowerCase().includes(C)||I.tools.some(_=>_.toLowerCase().includes(C))||I.steps.some(_=>_.toLowerCase().includes(C));return z&&y});if(v&&(v.textContent=`Menampilkan ${w.length} dari ${m.length} Instruksi Kerja`),w.length===0){f.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}f.innerHTML=`
        <div class="sec-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${w.map(I=>`
            <div class="card sec-ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${I.bg||"#EFF6FF"}; border-radius:10px;">${I.icon||"\u{1F6E1}\uFE0F"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        IK-${String(I.id).padStart(2,"0")}. ${I.title}
                      </h3>
                      <span class="badge" style="background:${I.bg||"#EFF6FF"}; color:${I.color||"#1E40AF"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${I.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    <button class="btn btn-sm btn-ghost btn-edit-sec-ik" data-id="${I.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    <button class="btn btn-sm btn-ghost btn-delete-sec-ik" data-id="${I.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${I.color||"#1E40AF"};">
                  <strong>Area Sasaran:</strong> ${I.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Perlengkapan:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${I.tools.slice(0,3).map(z=>`<li style="margin-bottom:2px;">${s(z)}</li>`).join("")}
                    ${I.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${I.tools.length-3} perlengkapan lainnya...</li>`:""}
                  </ul>
                </div>
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${I.steps.length} Langkah Prosedur
                </span>
                <button class="btn btn-sm btn-primary btn-view-sec-ik" data-id="${I.id}" style="padding:6px 14px; font-weight:600; border-radius:6px; background:#1E40AF; border-color:#1E40AF;">
                  Lihat Panduan Lengkap \u2794
                </button>
              </div>
            </div>
          `).join("")}
        </div>
      `,f.querySelectorAll(".btn-view-sec-ik").forEach(I=>{I.addEventListener("click",()=>{let z=parseInt(I.dataset.id,10),C=m.find(y=>y.id===z);C&&$(C)})}),f.querySelectorAll(".btn-edit-sec-ik").forEach(I=>{I.addEventListener("click",()=>{let z=parseInt(I.dataset.id,10),C=m.find(y=>y.id===z);C&&E(C)})}),f.querySelectorAll(".btn-delete-sec-ik").forEach(I=>{I.addEventListener("click",()=>{let z=parseInt(I.dataset.id,10),C=m.find(y=>y.id===z);C&&X(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${C.title}"?`,()=>{m=m.filter(y=>y.id!==z),$t(m),R("Instruksi Kerja Security berhasil dihapus"),o()},"Hapus Instruksi Kerja")})})}L();let A=r.querySelector("#sec-ik-search");A&&A.addEventListener("input",w=>{i=w.target.value,L()}),r.querySelectorAll(".sec-cat-pill").forEach(w=>{w.addEventListener("click",()=>{r.querySelectorAll(".sec-cat-pill").forEach(I=>I.classList.remove("active")),w.classList.add("active"),n=w.dataset.cat,L()})});let P=r.querySelector("#btn-add-sec-ik");P&&P.addEventListener("click",()=>{E()})}function $(r){G({title:`${r.icon||"\u{1F6E1}\uFE0F"} IK-${String(r.id).padStart(2,"0")}: ${r.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${r.bg||"#EFF6FF"}; color:${r.color||"#1E40AF"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${r.category}</span>
            <span>Kode: IK-SEC-${String(r.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${r.color||"#1E40AF"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${r.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Perlengkapan yang Disiapkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${r.tools.map(l=>{if(l.includes("http")){let f=l.split(": ");return`<li>${f[0]}: <a href="${f.slice(1).join(": ")}" target="_blank" rel="noopener" style="color:#2563EB; font-weight:600; text-decoration:underline;">Buka Tautan Form \u2794</a></li>`}return`<li>${s(l)}</li>`}).join("")}
              </ul>
            </div>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah & Prosedur Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${r.steps.map(l=>`<li style="margin-bottom:6px;">${s(l)}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${r.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Khusus & Ketentuan Penting:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${s(r.notes)}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(l,f)=>{f()}})}function E(r=null){let l=!!r,f=r?r.tools.join(`
`):"",v=r?r.steps.join(`
`):"",L=`
      <form id="sec-ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-sec-ik-title" class="form-control" required value="${r?s(r.title):""}" placeholder="Contoh: Pengamanan Area & Aset Perusahaan" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-sec-ik-icon" class="form-control" value="${r?r.icon:"\u{1F6E1}\uFE0F"}" placeholder="\u{1F6E1}\uFE0F" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
          <select id="form-sec-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Pengamanan Fisik & Patroli" ${r?.category==="Pengamanan Fisik & Patroli"?"selected":""}>Pengamanan Fisik & Patroli</option>
            <option value="Operasional Parkir & Luar" ${r?.category==="Operasional Parkir & Luar"?"selected":""}>Operasional Parkir & Luar</option>
            <option value="Layanan & Hospitalitas" ${r?.category==="Layanan & Hospitalitas"?"selected":""}>Layanan & Hospitalitas</option>
          </select>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-sec-ik-target" class="form-control" required value="${r?s(r.target):""}" placeholder="Contoh: Seluruh area luar klinik dan akses parkir" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Perlengkapan (1 per baris) *</label>
          <textarea id="form-sec-ik-tools" class="form-control" rows="3" placeholder="Senter Patroli&#10;Lampu Lalin&#10;Peluit" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${s(f)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah & Prosedur Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-sec-ik-steps" class="form-control" rows="5" placeholder="1. Lakukan Safety Patrol sesuai jadwal.&#10;2. Arahkan kendaraan pasien dengan ramah..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${s(v)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Khusus / Ketentuan Penting</label>
          <textarea id="form-sec-ik-notes" class="form-control" rows="2" placeholder="Peringatan keselamatan atau pedoman khusus..." style="width:100%; border-radius:6px; font-size:0.85rem;">${r?s(r.notes||""):""}</textarea>
        </div>
      </form>
    `;G({title:l?`\u270F\uFE0F Edit IK-${r.id}: ${r.title}`:"\u2795 Tambah Instruksi Kerja Security Baru",content:L,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(A,P)=>{let w=A.querySelector("#form-sec-ik-title").value.trim(),I=A.querySelector("#form-sec-ik-icon").value.trim()||"\u{1F6E1}\uFE0F",z=A.querySelector("#form-sec-ik-cat").value,C=A.querySelector("#form-sec-ik-target").value.trim(),y=A.querySelector("#form-sec-ik-tools").value.trim(),_=A.querySelector("#form-sec-ik-steps").value.trim(),T=A.querySelector("#form-sec-ik-notes").value.trim();if(!w||!C||!y||!_){N("Harap lengkapi semua kolom bertanda bintang (*)");return}let j=y.split(`
`).map(H=>H.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),O=_.split(`
`).map(H=>H.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),q={"Pengamanan Fisik & Patroli":{color:"#2563EB",bg:"#EFF6FF"},"Operasional Parkir & Luar":{color:"#D97706",bg:"#FFFBEB"},"Layanan & Hospitalitas":{color:"#059669",bg:"#ECFDF5"}}[z]||{color:"#1E40AF",bg:"#EFF6FF"};if(l)r.title=w,r.icon=I,r.category=z,r.target=C,r.tools=j,r.steps=O,r.notes=T,r.color=q.color,r.bg=q.bg,$t(m),R(`Instruksi Kerja "${w}" berhasil diperbarui`);else{let H=m.length>0?Math.max(...m.map(W=>W.id))+1:1,Y={id:H,title:w,category:z,icon:I,color:q.color,bg:q.bg,target:C,tools:j,steps:O,notes:T};m.push(Y),$t(m),R(`Instruksi Kerja baru IK-${H} berhasil ditambahkan`)}P(),o()}})}function k(r){r.innerHTML=`
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:20px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-cabang-btn" class="btn btn-sm ${x==="cabang"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F3E5} Zonasi Klinik / Cabang
          </button>
          <button id="view-hq-btn" class="btn btn-sm ${x==="hq"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
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
    `;let l=r.querySelector("#zonasi-content-container");x==="cabang"?l.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
          ${xa.cabang.map(f=>`
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
                  ${f.rooms.map(v=>`<li>${v}</li>`).join("")}
                </ul>
              </div>
            </div>
          `).join("")}
        </div>
      `:l.innerHTML=`
        <div style="display:flex; flex-direction:column; gap:16px;">
          ${xa.hq.map(f=>`
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
                    ${f.hijau.map(v=>`<li>${v}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Kuning -->
                <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:8px; padding:12px;">
                  <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E1} ZONA KUNING (Terbatas):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#78350F; line-height:1.6;">
                    ${f.kuning.map(v=>`<li>${v}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Merah -->
                <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:8px; padding:12px;">
                  <strong style="color:#991B1B; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F534} ZONA MERAH (Restricted):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#7F1D1D; line-height:1.6;">
                    ${f.merah.map(v=>`<li>${v}</li>`).join("")}
                  </ul>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,r.querySelector("#view-cabang-btn").addEventListener("click",()=>{x="cabang",k(r)}),r.querySelector("#view-hq-btn").addEventListener("click",()=>{x="hq",k(r)})}function h(r){r.innerHTML=`
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
          <input type="text" id="sec-asset-search" class="form-control" placeholder="Cari nama aset / tipe / PIC..." value="${c}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
          <button id="btn-add-sec-asset" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
            \u2795 Tambah Aset Baru
          </button>
        </div>
      </div>

      <div id="sec-asset-table-container"></div>
    `;let l=r.querySelector("#sec-asset-table-container"),f=r.querySelector("#sec-asset-search");function v(){let A=c.toLowerCase().trim(),P=u.filter(w=>!A||w.item.toLowerCase().includes(A)||w.type.toLowerCase().includes(A)||w.description.toLowerCase().includes(A)||w.pic&&w.pic.toLowerCase().includes(A));l.innerHTML=`
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
            <strong style="color:var(--text-1); font-size:0.95rem;">Rekap Klasifikasi Aset Perusahaan</strong>
            <span class="badge badge-primary">${P.length} Aset</span>
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
                ${P.map((w,I)=>`
                  <tr style="border-bottom:1px solid #F1F5F9; background:${I%2===0?"#fff":"#FAFAFA"};">
                    <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${w.no||I+1}</td>
                    <td style="padding:10px 14px;">
                      <span class="badge" style="${L(w.type)}">${s(w.type)}</span>
                    </td>
                    <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${s(w.item)}</td>
                    <td style="padding:10px 14px; color:var(--text-2);">${s(w.description)}</td>
                    <td style="padding:10px 14px; font-weight:600; color:#1E40AF;">${s(w.pic||"-")}</td>
                    <td style="padding:10px 14px; text-align:center;">
                      <button class="btn btn-sm btn-ghost btn-edit-asset" data-no="${w.no}" title="Edit Aset" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-sm btn-ghost btn-delete-asset" data-no="${w.no}" title="Hapus Aset" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </div>
      `,l.querySelectorAll(".btn-edit-asset").forEach(w=>{w.addEventListener("click",()=>{let I=parseInt(w.dataset.no,10),z=u.find(C=>C.no===I);z&&a(z)})}),l.querySelectorAll(".btn-delete-asset").forEach(w=>{w.addEventListener("click",()=>{let I=parseInt(w.dataset.no,10),z=u.find(C=>C.no===I);z&&X(`Apakah Anda yakin ingin menghapus aset: "${z.item}"?`,()=>{u=u.filter(C=>C.no!==I),u.forEach((C,y)=>{C.no=y+1}),_t(u),R("Aset berhasil dihapus"),o()},"Hapus Aset")})})}function L(A){return A==="Aset Tetap"?"background:#EFF6FF; color:#1E40AF; font-weight:700; padding:3px 8px; border-radius:6px;":A==="Aset Bergerak"?"background:#ECFDF5; color:#047857; font-weight:700; padding:3px 8px; border-radius:6px;":"background:#FEF3C7; color:#92400E; font-weight:700; padding:3px 8px; border-radius:6px;"}v(),f&&f.addEventListener("input",A=>{c=A.target.value,v()}),r.querySelector("#btn-add-sec-asset").addEventListener("click",()=>{a()})}function a(r=null){let l=!!r,f=`
      <form id="asset-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Tipe / Klasifikasi Aset *</label>
          <select id="form-asset-type" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Aset Tetap" ${r?.type==="Aset Tetap"?"selected":""}>Aset Tetap</option>
            <option value="Aset Bergerak" ${r?.type==="Aset Bergerak"?"selected":""}>Aset Bergerak</option>
            <option value="Low Value Asset" ${r?.type==="Low Value Asset"?"selected":""}>Low Value Asset</option>
          </select>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Aset / Kelompok *</label>
          <input type="text" id="form-asset-item" class="form-control" required value="${r?s(r.item):""}" placeholder="Contoh: Genset & Elevator" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Deskripsi / Contoh Item *</label>
          <textarea id="form-asset-desc" class="form-control" rows="3" required placeholder="Contoh: Generator daya darurat dan panel listrik utama" style="width:100%; border-radius:6px; font-size:0.85rem;">${r?s(r.description):""}</textarea>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">PIC Pengawas / Penanggung Jawab *</label>
          <input type="text" id="form-asset-pic" class="form-control" required value="${r?s(r.pic):""}" placeholder="Contoh: Teknisi & Security" style="width:100%; border-radius:6px; height:36px;">
        </div>
      </form>
    `;G({title:l?`\u270F\uFE0F Edit Aset: ${r.item}`:"\u2795 Tambah Klasifikasi Aset Baru",content:f,confirmText:"Simpan",onConfirm:(v,L)=>{let A=v.querySelector("#form-asset-type").value,P=v.querySelector("#form-asset-item").value.trim(),w=v.querySelector("#form-asset-desc").value.trim(),I=v.querySelector("#form-asset-pic").value.trim();if(!P||!w||!I){N("Harap lengkapi semua kolom");return}if(l)r.type=A,r.item=P,r.description=w,r.pic=I,_t(u),R(`Aset "${P}" berhasil diperbarui`);else{let z=u.length>0?Math.max(...u.map(C=>C.no))+1:1;u.push({no:z,type:A,item:P,description:w,pic:I}),_t(u),R(`Aset baru "${P}" berhasil ditambahkan`)}L(),o()}})}function p(r){r.innerHTML=`
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
    `}function d(r){r.innerHTML=`
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
    `}async function g(){try{let{downloadExcel:r}=await Promise.resolve().then(()=>(re(),he)),l=m.map(v=>({"Kode IK":`IK-SEC-${String(v.id).padStart(2,"0")}`,"Judul Prosedur":v.title,Kategori:v.category,"Sasaran Target":v.target,"Alat & Perlengkapan":v.tools.join(", "),"Langkah Pelaksanaan":v.steps.map((L,A)=>`${A+1}. ${L}`).join(" | "),"Catatan Khusus":v.notes||"-"})),f=new Date().toISOString().slice(0,10);r(l,`SOP_Security_Instruksi_Kerja_${f}`),R("Data Excel SOP Security berhasil diunduh")}catch(r){console.error(r),N("Gagal mendownload data Excel: "+r.message)}}function s(r){return r?String(r).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}o()}le();var wa=[{id:1,name:"Berita Acara Kerusakan Alat",category:"Berita Acara",document_link:"https://docs.google.com/document/d/1example_ba_kerusakan",description:"Format berita acara resmi pencatatan dan eskalasi kerusakan peralatan klinik dan aset fasilitas."},{id:2,name:"Checklist Apar",category:"Apar & K3",document_link:"https://docs.google.com/spreadsheets/d/1example_apar",description:"Lembar inspeksi berkala tabung pemadam api ringan (APAR), tekanan manometer, dan masa kadaluarsa."},{id:3,name:"Checklist Deep Cleaning",category:"Kegiatan Deep Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_deep_cleaning",description:"Panduan dan verifikasi checklist pembersihan mendalam (deep cleaning) area dental unit dan klinik."},{id:4,name:"Checklist Fulldraine",category:"Drain Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_fulldraine",description:"Checklist monitoring pembersihan saluran pembuangan air kotor, floor drain, dan sedotan suction."},{id:5,name:"Checklist General Cleaning",category:"Kegiatan General Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_general_cleaning",description:"Standar checklist kebersihan harian/mingguan seluruh ruangan, lobby, toilet, dan koridor."},{id:6,name:"Form Aktivity Reliefer",category:"Form Operasional",document_link:"https://forms.gle/1example_reliefer",description:"Formulir pencatatan absensi dan penugasan harian reliefer Facility Care di cabang pengganti."},{id:7,name:"Form Aktivity Security HQ",category:"Form Operasional",document_link:"https://forms.gle/1example_security_hq",description:"Formulir monitoring shift dan pengawasan keamanan harian di Kantor Pusat (Head Quarter)."}];function Ea(e=""){let t=(e||"").toLowerCase();return t.includes("berita acara")?{label:e,bg:"#F3E8FF",color:"#7E22CE",border:"#D8B4FE",icon:"\u2696\uFE0F"}:t.includes("apar")||t.includes("k3")||t.includes("safety")?{label:e,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F9EF}"}:t.includes("deep cleaning")?{label:e,bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2728"}:t.includes("drain")?{label:e,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u{1F6BF}"}:t.includes("general cleaning")||t.includes("kebersihan")?{label:e,bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F9F9}"}:t.includes("form")?{label:e,bg:"#E0E7FF",color:"#4338CA",border:"#C7D2FE",icon:"\u{1F4DD}"}:{label:e||"Dokumen",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4C1}"}}function Sa(e=""){if(!e)return{type:"Tidak Ada Tautan",icon:"\u{1F4CE}",badgeClass:"bg-slate-100 text-slate-500"};let t=e.toLowerCase();return t.includes("forms.gle")||t.includes("google.com/forms")?{type:"Google Form",icon:"\u{1F7E3}",badgeClass:"badge-form",color:"#673AB7"}:t.includes("docs.google.com/spreadsheets")||t.includes("sheet")?{type:"Google Sheets",icon:"\u{1F7E2}",badgeClass:"badge-sheet",color:"#0F9D58"}:t.includes("docs.google.com/document")||t.includes("doc")?{type:"Google Docs",icon:"\u{1F535}",badgeClass:"badge-doc",color:"#4285F4"}:t.includes("drive.google.com")?{type:"Google Drive",icon:"\u{1F7E1}",badgeClass:"badge-drive",color:"#F4B400"}:t.endsWith(".pdf")?{type:"PDF Document",icon:"\u{1F534}",badgeClass:"badge-pdf",color:"#EA4335"}:{type:"Web Link",icon:"\u{1F310}",badgeClass:"badge-link",color:"#2563EB"}}async function Ca(e){let t=[],i="all",n="",c="fcms-master-doc-styles";if(!document.getElementById(c)){let $=document.createElement("style");$.id=c,$.textContent=`
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
    `,document.head.appendChild($)}async function x(){try{let $=await K("/api/checklist?limit=1000");if($.ok&&$.data?.data&&Array.isArray($.data.data)&&$.data.data.length>0)t=$.data.data,localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t));else{let E=localStorage.getItem("fcms_master_doc_cache");E?t=JSON.parse(E):(t=[...wa],localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)))}}catch($){console.warn("Fallback loading documents from cache/seed:",$);let E=localStorage.getItem("fcms_master_doc_cache");t=E?JSON.parse(E):[...wa]}S()}function m(){return t.filter($=>{let E=i==="all"||$.category===i,k=n.toLowerCase().trim(),h=!k||($.name||"").toLowerCase().includes(k)||($.category||"").toLowerCase().includes(k)||($.description||"").toLowerCase().includes(k);return E&&h})}function u(){let $={all:t.length};return t.forEach(E=>{let k=E.category||"Lainnya";$[k]=($[k]||0)+1}),$}function o($=null){let E=!!$,k=Array.from(new Set(t.map(a=>a.category).filter(Boolean))),h=document.createElement("div");h.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Nama Dokumen / Checklist <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Checklist Pembersihan Wastafel & Toilet" value="${$?.name?String($.name).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Kategori Dokumen <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-category" list="doc-categories-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Pilih atau ketik kategori baru..." value="${$?.category?String($.category).replace(/"/g,"&quot;"):""}" required />
          <datalist id="doc-categories-list">
            ${k.map(a=>`<option value="${a}">${a}</option>`).join("")}
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
            placeholder="https://docs.google.com/..." value="${$?.document_link?String($.document_link).replace(/"/g,"&quot;"):""}" />
          <span style="font-size:0.75rem; color:#64748B; margin-top:4px; display:block;">
            Bisa berupa tautan Google Forms untuk inspeksi harian, atau link Google Sheets / Docs sebagai acuan.
          </span>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Deskripsi / Petunjuk Penggunaan
          </label>
          <textarea id="doc-modal-desc" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan keterangan singkat mengenai tujuan dan instruksi pengisian dokumen ini...">${$?.description||""}</textarea>
        </div>
      </div>
    `,G({title:E?"\u270F\uFE0F Edit Master Dokumen":"\u2795 Tambah Master Dokumen Baru",content:h,size:"md",confirmText:E?"Simpan Perubahan":"Tambahkan Dokumen",confirmClass:"btn-primary",onConfirm:async(a,p)=>{let d=h.querySelector("#doc-modal-name"),g=h.querySelector("#doc-modal-category"),s=h.querySelector("#doc-modal-link"),r=h.querySelector("#doc-modal-desc"),l=(d.value||"").trim(),f=(g.value||"").trim(),v=(s.value||"").trim(),L=(r.value||"").trim();if(!l){N("Nama Dokumen wajib diisi"),d.focus();return}if(!f){N("Kategori Dokumen wajib diisi"),g.focus();return}let A={name:l,category:f,document_link:v,description:L};try{if(E)if((await K(`/api/checklist/${$.id}`,{method:"PUT",body:JSON.stringify(A)})).ok){let w=t.findIndex(I=>I.id===$.id);w!==-1&&(t[w]={...t[w],...A}),R("Dokumen berhasil diperbarui")}else{let w=t.findIndex(I=>I.id===$.id);w!==-1&&(t[w]={...t[w],...A}),R("Dokumen disimpan secara lokal")}else{let P=await K("/api/checklist",{method:"POST",body:JSON.stringify(A)});if(P.ok&&P.data?.id)t.unshift({id:P.data.id,...A}),R("Dokumen baru berhasil ditambahkan");else{let w=Date.now();t.unshift({id:w,...A}),R("Dokumen baru ditambahkan secara lokal")}}localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),p(),S()}catch(P){console.error("Error saving document:",P),N("Gagal menyimpan dokumen: "+P.message)}}})}function b($){X(`Apakah Anda yakin ingin menghapus dokumen <strong>"${$.name}"</strong>?`,async()=>{try{let E=await K(`/api/checklist/${$.id}`,{method:"DELETE"});E.ok||console.warn("API delete warning:",E),t=t.filter(k=>k.id!==$.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),R(`Dokumen "${$.name}" berhasil dihapus`),S()}catch{t=t.filter(k=>k.id!==$.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),R(`Dokumen "${$.name}" berhasil dihapus`),S()}},"Hapus Dokumen")}function B($,E){if(!$){ut("Dokumen belum memiliki tautan digital.");return}navigator.clipboard.writeText($).then(()=>{R(`Tautan "${E}" disalin ke clipboard!`)}).catch(()=>{ut(`Tautan: ${$}`)})}async function D(){try{let{downloadExcel:$}=await Promise.resolve().then(()=>(re(),he)),E=t.map((k,h)=>({No:h+1,"Nama Dokumen":k.name||"",Kategori:k.category||"","Platform / Format":Sa(k.document_link).type,"Link Dokumen":k.document_link||"","Deskripsi / Keterangan":k.description||""}));$(E,`Master_Dokumen_FCMS_${new Date().toISOString().slice(0,10)}`),R("Berhasil mengunduh data dokumen ke Excel.")}catch($){N("Gagal mengekspor data: "+$.message)}}function S(){let $=m(),E=u(),k=Object.keys(E).filter(a=>a!=="all"),h=t.filter(a=>!!a.document_link).length;e.innerHTML=`
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
              <button id="doc-btn-add" class="btn btn-primary" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap; border:none;">
                <span>\u2795</span>
                <span>Tambah Dokumen Baru</span>
              </button>
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
                <div class="doc-stat-value">${t.length}</div>
                <div class="doc-stat-label">Total Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon">\u{1F5C2}\uFE0F</div>
              <div>
                <div class="doc-stat-value">${k.length}</div>
                <div class="doc-stat-label">Kategori Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon" style="background:rgba(16,185,129,0.18); color:#34D399;">\u{1F310}</div>
              <div>
                <div class="doc-stat-value">${h} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">/ ${t.length}</span></div>
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
              <input type="text" id="doc-search-input" class="doc-search-input" placeholder="Cari nama dokumen, kategori, atau deskripsi..." value="${n?n.replace(/"/g,"&quot;"):""}" />
            </div>
          </div>

          <!-- Category Pills -->
          <div class="doc-filter-pills">
            <button class="doc-pill ${i==="all"?"active":""}" data-cat="all">
              <span>Semua Dokumen</span>
              <span class="doc-pill-count">${E.all||0}</span>
            </button>
            ${k.map(a=>{let p=Ea(a);return`
                <button class="doc-pill ${i===a?"active":""}" data-cat="${a}">
                  <span>${p.icon}</span>
                  <span>${a}</span>
                  <span class="doc-pill-count">${E[a]||0}</span>
                </button>
              `}).join("")}
          </div>
        </div>

        <!-- Documents View Container -->
        <div id="doc-list-view">
          ${$.length===0?`
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
              ${$.map(a=>{let p=Ea(a.category),d=Sa(a.document_link);return`
                  <div class="doc-card" data-id="${a.id}">
                    <div>
                      <div class="doc-card-top">
                        <span class="doc-category-badge" style="background:${p.bg}; color:${p.color}; border:1px solid ${p.border};">
                          <span>${p.icon}</span>
                          <span>${p.label}</span>
                        </span>
                        <span class="doc-platform-badge">
                          <span>${d.icon}</span>
                          <span>${d.type}</span>
                        </span>
                      </div>
                      
                      <h3 class="doc-card-title">${a.name}</h3>
                      <p class="doc-card-desc">${a.description||"Formulir acuan standar operasional dan administrasi fasilitas."}</p>
                    </div>

                    <div class="doc-card-footer">
                      ${a.document_link?`
                        <a href="${a.document_link}" target="_blank" rel="noopener noreferrer" class="doc-btn-open">
                          <span>Buka Dokumen</span>
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                        </a>
                      `:`
                        <button class="doc-btn-open disabled" title="Belum ada tautan yang dikaitkan">
                          <span>Tidak Ada Link</span>
                        </button>
                      `}

                      <div class="doc-card-actions">
                        ${a.document_link?`
                          <button class="doc-icon-btn copy-link-btn" title="Salin Tautan Dokumen" data-link="${a.document_link}" data-title="${a.name}">
                            <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                          </button>
                        `:""}
                        <button class="doc-icon-btn edit-doc-btn" title="Edit Dokumen" data-id="${a.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        <button class="doc-icon-btn delete delete-doc-btn" title="Hapus Dokumen" data-id="${a.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                        </button>
                      </div>
                    </div>
                  </div>
                `}).join("")}
            </div>
          `}
        </div>

      </div>
    `,F()}function F(){let $=e.querySelector("#doc-btn-add");$&&($.onclick=()=>o());let E=e.querySelector("#doc-btn-excel");E&&(E.onclick=()=>D());let k=e.querySelector("#doc-btn-refresh");k&&(k.onclick=()=>x());let h=e.querySelector("#doc-btn-reset-filter");h&&(h.onclick=()=>{n="",i="all",S()});let a=e.querySelector("#doc-search-input");a&&(a.oninput=p=>{n=p.target.value;let d=m();S();let g=e.querySelector("#doc-search-input");g&&(g.focus(),g.setSelectionRange(g.value.length,g.value.length))}),e.querySelectorAll(".doc-pill").forEach(p=>{p.onclick=()=>{i=p.getAttribute("data-cat")||"all",S()}}),e.querySelectorAll(".copy-link-btn").forEach(p=>{p.onclick=d=>{d.stopPropagation();let g=p.getAttribute("data-link"),s=p.getAttribute("data-title");B(g,s)}}),e.querySelectorAll(".edit-doc-btn").forEach(p=>{p.onclick=d=>{d.stopPropagation();let g=Number(p.getAttribute("data-id")),s=t.find(r=>r.id===g);s&&o(s)}}),e.querySelectorAll(".delete-doc-btn").forEach(p=>{p.onclick=d=>{d.stopPropagation();let g=Number(p.getAttribute("data-id")),s=t.find(r=>r.id===g);s&&b(s)}})}await x()}le();re();var ze={"Standar Umum":{icon:"\u{1F4CC}",bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE",rooms:["Lantai & Sudut Ruangan","Dinding, Partisi & Plint","Kaca, Cermin & Partisi Kaca","Sanitair & Area Basah Toilet","Plafon, Lampu & Kisi Ventilasi","Mebel, Meja & Kursi Kerja/Pasien","Peralatan Khusus & Medis Non-Steril","Wadah Sampah & Utilitas"]},"Area Depan & Pelayanan":{icon:"\u{1F6CB}\uFE0F",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",rooms:["Teras","Lobby 1","Receptionist","Snack Corner"]},"Area Klinis & Medis":{icon:"\u{1FA7A}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD",rooms:["Ruang Tindakan","Ronsen","Ruang Dokter","Ruang Steril"]},"Toilet & Sanitasi":{icon:"\u{1F6BF}",bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4",rooms:["Toilet 1","Janitor","Ruang Limbah"]},Sirkulasi:{icon:"\u{1F6B6}",bg:"#F1F5F9",color:"#334155",border:"#CBD5E1",rooms:["Area Tangga","Lift","Area Tangga Exit"]},"Area Anak & Fasilitas":{icon:"\u{1F9F8}",bg:"#FEF3C7",color:"#92400E",border:"#FDE68A",rooms:["Lobby 2","Playground","Musholla","Tempat Wudhu","Ruang Team"]},"Gudang & Teknis":{icon:"\u2699\uFE0F",bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA",rooms:["Gudang Farmasi & Bahan","RoofTop","Ruang Kompressor","Area Genset"]},Eksterior:{icon:"\u{1F333}",bg:"#DCFCE7",color:"#15803D",border:"#86EFAC",rooms:["Taman","Pos Security","Toilet Security","Halaman Parkir"]}};function Tt(e){for(let[t,i]of Object.entries(ze))if(i.rooms.includes(e))return{zoneName:t,...i};return{zoneName:"Lainnya",icon:"\u{1F4CD}",bg:"#F1F5F9",color:"#475569",border:"#E2E8F0"}}async function Fa(e){e.innerHTML=`
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
        ${Object.entries(ze).map(([p,d])=>`
          <button class="zone-tab-btn" data-zone="${p}" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
            <span>${d.icon}</span> ${p}
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
  `;let t="all",i="table",n=[],c=ee(),x=c&&["superadmin","admin","manager","spv"].includes(c.role);async function m(){try{let p=await K("/api/hygiene-standards?all=1");if(!p.ok)throw new Error("Gagal memuat standar hygiene");n=Array.isArray(p.data)?p.data:p.data?.data||[],o(),u(),B()}catch(p){console.error(p),N(p.message||"Terjadi kesalahan memuat data");let d=e.querySelector("#hygiene-view-container");d&&(d.innerHTML=`
          <div class="empty-state" style="text-align:center; padding:40px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:12px;">\u26A0\uFE0F</div>
            <h3 style="margin:0 0 8px 0; color:var(--text-1);">Gagal Memuat Data</h3>
            <p style="color:var(--text-3); margin-bottom:16px;">${p.message}</p>
            <button id="btn-retry-hygiene" class="btn btn-primary" style="padding:8px 16px; border-radius:8px;">Coba Lagi</button>
          </div>
        `,d.querySelector("#btn-retry-hygiene")?.addEventListener("click",m))}}function u(){let p=e.querySelector("#badge-count-all");p&&(p.textContent=n.length),Object.entries(ze).forEach(([s,r])=>{let l=s.replace(/[\s&]+/g,"-"),f=e.querySelector(`#badge-count-${l}`);if(f){let v=n.filter(L=>r.rooms.includes(L.room_name)).length;f.textContent=v}});let d=new Set(n.map(s=>s.room_name)),g=e.querySelector("#card-total-rooms");g&&(g.textContent=d.size)}function o(){let p=e.querySelector("#hygiene-room-filter");if(!p)return;let d=p.value,g=[];if(t==="all")g=Array.from(new Set(n.map(s=>s.room_name))).filter(Boolean).sort();else{let s=ze[t];s&&(g=s.rooms.filter(r=>n.some(l=>l.room_name===r)))}p.innerHTML=`<option value="">Semua Ruangan (${g.length})</option>`+g.map(s=>`<option value="${s}" ${s===d?"selected":""}>${s}</option>`).join("")}function b(){let p=(e.querySelector("#hygiene-search")?.value||"").toLowerCase().trim(),d=e.querySelector("#hygiene-room-filter")?.value||"";return n.filter(g=>{if(t!=="all"){let s=ze[t];if(!s||!s.rooms.includes(g.room_name))return!1}if(d&&g.room_name!==d)return!1;if(p){let s=(g.room_name||"").toLowerCase(),r=(g.item_name||"").toLowerCase(),l=(g.cleanliness_standard||"").toLowerCase();if(!s.includes(p)&&!r.includes(p)&&!l.includes(p))return!1}return!0})}function B(){let p=b(),d=e.querySelector("#hygiene-stats-text");d&&(d.innerHTML=`Menampilkan <strong>${p.length}</strong> dari <strong>${n.length}</strong> parameter`);let g=e.querySelector("#hygiene-view-container");if(g){if(p.length===0){g.innerHTML=`
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
      `,g.querySelector("#btn-clear-search-empty")?.addEventListener("click",()=>{let s=e.querySelector("#hygiene-search"),r=e.querySelector("#hygiene-room-filter");s&&(s.value=""),r&&(r.value=""),t="all",e.querySelectorAll(".zone-tab-btn").forEach(l=>l.classList.toggle("active",l.dataset.zone==="all")),e.querySelectorAll(".zone-tab-btn").forEach(l=>{l.dataset.zone==="all"?(l.style.background="#2563EB",l.style.color="#fff"):(l.style.background="#F1F5F9",l.style.color="var(--text-2)")}),o(),B()});return}i==="table"?D(g,p):S(g,p)}}function D(p,d){p.innerHTML=`
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse:collapse; text-align:left; font-size:0.9rem;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:2px solid #E2E8F0; color:#475569; font-size:0.78rem; text-transform:uppercase; letter-spacing:0.6px;">
                <th style="padding:14px 16px; width:50px; text-align:center;">No</th>
                <th style="padding:14px 16px; width:220px;">Ruangan & Zona</th>
                <th style="padding:14px 16px; width:280px;">Item / Objek Inspeksi</th>
                <th style="padding:14px 16px;">Standar Kebersihan Fisik</th>
                ${x?'<th style="padding:14px 16px; width:100px; text-align:center;">Aksi</th>':""}
              </tr>
            </thead>
            <tbody>
              ${d.map((g,s)=>{let r=Tt(g.room_name);return`
                  <tr style="border-bottom:1px solid #F1F5F9; transition:background-color 0.15s ease;" onmouseover="this.style.backgroundColor='#F8FAFC'" onmouseout="this.style.backgroundColor='transparent'">
                    <td style="padding:14px 16px; text-align:center; color:var(--text-3); font-weight:600; font-size:0.85rem;">
                      ${s+1}
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="margin-bottom:4px;">
                        <span style="display:inline-flex; align-items:center; gap:6px; background:${r.bg}; color:${r.color}; border:1px solid ${r.border}; padding:4px 10px; border-radius:6px; font-weight:700; font-size:0.83rem;">
                          <span>${r.icon}</span> ${g.room_name}
                        </span>
                      </div>
                      <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">
                        ${r.zoneName}
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
                    ${x?`
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
    `,F(p)}function S(p,d){let g={};d.forEach(s=>{let r=Tt(s.room_name);g[r.zoneName]||(g[r.zoneName]={info:r,rooms:{}}),g[r.zoneName].rooms[s.room_name]||(g[r.zoneName].rooms[s.room_name]=[]),g[r.zoneName].rooms[s.room_name].push(s)}),p.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:24px;">
        ${Object.entries(g).map(([s,r])=>{let l=Object.values(r.rooms).reduce((f,v)=>f+v.length,0);return`
            <div class="card" style="background:#fff; border:1px solid ${r.info.border}; border-radius:14px; overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.03);">
              <!-- Zone Card Header -->
              <div style="background:${r.info.bg}; border-bottom:1px solid ${r.info.border}; padding:14px 20px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="font-size:1.4rem;">${r.info.icon}</span>
                  <div>
                    <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:${r.info.color};">
                      ${s}
                    </h3>
                    <div style="font-size:0.78rem; color:${r.info.color}; opacity:0.85;">
                      ${Object.keys(r.rooms).length} Ruangan \xB7 ${l} Parameter Standar
                    </div>
                  </div>
                </div>
                <span class="badge" style="background:#fff; color:${r.info.color}; border:1px solid ${r.info.border}; font-weight:700; padding:4px 10px; border-radius:12px; font-size:0.8rem;">
                  ${l} Parameter
                </span>
              </div>

              <!-- Zone Rooms List -->
              <div style="padding:16px 20px;">
                ${Object.entries(r.rooms).map(([f,v])=>`
                  <div style="margin-bottom:20px; last-child:margin-bottom:0;">
                    <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                      <h4 style="margin:0; font-size:0.95rem; font-weight:700; color:var(--text-1);">
                        \u{1F6AA} ${f}
                      </h4>
                      <span class="badge badge-sm" style="background:#F1F5F9; color:var(--text-3); font-size:0.75rem;">
                        ${v.length} item
                      </span>
                    </div>

                    <div style="border:1px solid #E2E8F0; border-radius:8px; overflow:hidden;">
                      <table style="width:100%; border-collapse:collapse; font-size:0.875rem;">
                        <thead>
                          <tr style="background:#F8FAFC; border-bottom:1px solid #E2E8F0; color:#475569; font-size:0.75rem; text-transform:uppercase;">
                            <th style="padding:8px 12px; width:40px; text-align:center;">No</th>
                            <th style="padding:8px 12px; width:260px;">Item / Objek Inspeksi</th>
                            <th style="padding:8px 12px;">Standar Kebersihan Fisik</th>
                            ${x?'<th style="padding:8px 12px; width:80px; text-align:center;">Aksi</th>':""}
                          </tr>
                        </thead>
                        <tbody>
                          ${v.map((L,A)=>`
                            <tr style="border-bottom:1px solid #F1F5F9;">
                              <td style="padding:10px 12px; text-align:center; color:var(--text-3); font-size:0.8rem;">
                                ${A+1}
                              </td>
                              <td style="padding:10px 12px; font-weight:600; color:var(--text-1); vertical-align:top;">
                                ${L.item_name}
                              </td>
                              <td style="padding:10px 12px; vertical-align:top; color:#334155;">
                                <div style="display:flex; align-items:flex-start; gap:6px;">
                                  <span style="color:#10B981; font-weight:700;">\u2713</span>
                                  <span>${L.cleanliness_standard}</span>
                                </div>
                              </td>
                              ${x?`
                                <td style="padding:10px 12px; text-align:center; vertical-align:middle;">
                                  <div style="display:inline-flex; gap:4px;">
                                    <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${L.id}" title="Edit" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                                    <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${L.id}" data-name="${L.item_name}" title="Hapus" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
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
    `,F(p)}function F(p){p.querySelectorAll(".btn-edit-item").forEach(d=>{d.addEventListener("click",()=>{let g=d.dataset.id,s=n.find(r=>String(r.id)===String(g));s&&$(s)})}),p.querySelectorAll(".btn-delete-item").forEach(d=>{d.addEventListener("click",()=>{let g=d.dataset.id,s=d.dataset.name;X(`Yakin ingin menghapus parameter standar <strong>"${s}"</strong>?`,async()=>{try{let r=await K(`/api/hygiene-standards/${g}`,{method:"DELETE"});if(!r.ok)throw new Error(r.data?.error||"Gagal menghapus parameter");R("Parameter kebersihan berhasil dihapus"),m()}catch(r){N(r.message)}},"Hapus Parameter Standar")})})}function $(p=null){let d=!!p,g=Array.from(new Set(n.map(r=>r.room_name))).filter(Boolean).sort(),s=document.createElement("div");s.innerHTML=`
      <form id="form-hygiene-item" style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Ruangan / Lokasi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-room-name" list="room-suggestions" class="form-control" placeholder="Contoh: Ruang Tindakan, Lobby 1, Toilet 1..." value="${p?.room_name||""}" required style="width:100%; height:40px; border-radius:8px;">
          <datalist id="room-suggestions">
            ${g.map(r=>`<option value="${r}">`).join("")}
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
    `,G({title:d?"\u270F\uFE0F Edit Parameter Kebersihan":"\u2795 Tambah Parameter Kebersihan",content:s,confirmText:d?"Simpan Perubahan":"Tambah Parameter",size:"md",onConfirm:async(r,l)=>{let f=s.querySelector("#modal-room-name").value.trim(),v=s.querySelector("#modal-item-name").value.trim(),L=s.querySelector("#modal-standard").value.trim();if(!f||!v||!L){N("Semua kolom wajib diisi");return}try{let A={room_name:f,item_name:v,cleanliness_standard:L},P=d?`/api/hygiene-standards/${p.id}`:"/api/hygiene-standards",I=await K(P,{method:d?"PUT":"POST",body:JSON.stringify(A)});if(!I.ok)throw new Error(I.data?.error||"Gagal menyimpan parameter");R(d?"Parameter berhasil diperbarui":"Parameter berhasil ditambahkan"),l(),m()}catch(A){N(A.message)}}})}function E(){let p=window.open("","_blank");if(!p){N("Pop-up printer diblokir oleh browser. Izinkan pop-up untuk mencetak SOP.");return}let g=`
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

        ${Object.entries(ze).map(([s,r])=>{let l=n.filter(f=>r.rooms.includes(f.room_name));return l.length===0?"":`
            <div class="zone-title">${r.icon} ${s.toUpperCase()} (${l.length} PARAMETER)</div>
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
                ${l.map((f,v)=>`
                  <tr>
                    <td class="td-no">${v+1}</td>
                    <td class="td-room">${f.room_name}</td>
                    <td class="td-obj">${f.item_name}</td>
                    <td>${f.cleanliness_standard}</td>
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
    `;p.document.open(),p.document.write(g),p.document.close()}function k(){let p=n.map((g,s)=>{let r=Tt(g.room_name);return{No:s+1,"Kategori Zona":r.zoneName,"Ruangan / Lokasi":g.room_name,"Item / Objek Inspeksi":g.item_name,"Standar Kebersihan Fisik":g.cleanliness_standard}}),d=new Date().toISOString().slice(0,10);Q(p,`Master_Standar_Kebersihan_FCMS_${d}`),R("Data master kebersihan berhasil diekspor ke Excel")}e.querySelector("#hygiene-search")?.addEventListener("input",()=>B()),e.querySelector("#hygiene-room-filter")?.addEventListener("change",()=>B()),e.querySelector("#btn-reset-filters")?.addEventListener("click",()=>{let p=e.querySelector("#hygiene-search"),d=e.querySelector("#hygiene-room-filter");p&&(p.value=""),d&&(d.value=""),B()}),e.querySelectorAll(".zone-tab-btn").forEach(p=>{p.addEventListener("click",()=>{t=p.dataset.zone,e.querySelectorAll(".zone-tab-btn").forEach(d=>{let g=d.dataset.zone===t;if(d.classList.toggle("active",g),g){d.style.background="#2563EB",d.style.color="#fff";let s=d.querySelector(".badge");s&&(s.style.background="rgba(255,255,255,0.25)",s.style.color="#fff")}else{d.style.background="#F1F5F9",d.style.color="var(--text-2)";let s=d.querySelector(".badge");s&&(s.style.background="#CBD5E1",s.style.color="#334155")}}),o(),B()})});let h=e.querySelector("#view-table-btn"),a=e.querySelector("#view-zone-btn");h?.addEventListener("click",()=>{i="table",h.style.background="#fff",h.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",a.style.background="transparent",a.style.boxShadow="none",B()}),a?.addEventListener("click",()=>{i="zone",a.style.background="#fff",a.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",h.style.background="transparent",h.style.boxShadow="none",B()}),e.querySelector("#btn-print-hygiene")?.addEventListener("click",E),e.querySelector("#btn-export-hygiene")?.addEventListener("click",k),e.querySelector("#btn-add-hygiene")?.addEventListener("click",()=>$(null)),m()}le();re();async function $a(e,t="supply"){return zn(e)}async function zn(e){let i=((await K("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name}));de({container:e,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:n=>n?new Date(n).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(n,c)=>c.branch_name_ref||c.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:n=>{try{let c=JSON.parse(n);return Array.isArray(c)?c.join(", "):n}catch{return n||"-"}}},{key:"chemical_items",label:"Chemical",render:n=>{try{let c=JSON.parse(n);return Array.isArray(c)?c.join(", "):n}catch{return n||"-"}}},{key:"additional_notes",label:"Catatan",render:n=>n?.length>40?n.slice(0,40)+"\u2026":n||"-"},{key:"status",label:"Status",render:n=>ge(n)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:n=>{let c=n?.tools_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}let x=n?.chemical_items;try{x=Array.isArray(JSON.parse(x))?JSON.parse(x).join(", "):x}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:n?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:n?.branch_id&&!i.find(m=>m.value==n.branch_id)?[...i,{value:n.branch_id,label:n.branch_name||n.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:c},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:n?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:x},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:n?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:n?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:n?.status||""},{name:"processed_by",label:"Diproses Oleh",value:n?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async n=>{let c=new URLSearchParams(n||{}).toString(),x=await K(`/api/reports/supply?limit=10000&${c}`);if(x.ok){let m=x.data.data.map(u=>{let o=u.tools_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}let b=u.chemical_items;try{b=Array.isArray(JSON.parse(b))?JSON.parse(b).join(", "):b}catch{}return{Waktu:u.submitted_at||"",Pengirim:u.submitter_name||"",Cabang:u.branch_name_ref||u.branch_name||"","Alat/Barang":o||"",Chemical:b||"",Catatan:u.additional_notes||"",Status:u.status||"","Diproses Oleh":u.processed_by||""}});Q(m,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Q([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async n=>{let x=(await K("/api/branches?all=1")).data?.data||[],m=B=>{if(!B)return null;let D=String(B||"").toLowerCase(),S=x.find(F=>String(F.full_name||"").toLowerCase()===D||String(F.code||"").toLowerCase()===D||String(F.name||"").toLowerCase()===D);return S?S.id:null},u=B=>{if(B==null||B==="")return"";if(B instanceof Date&&!isNaN(B.getTime()))return B.toISOString().slice(0,10);let D=String(B).trim();if(D===""||D==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(D))return D.slice(0,10);if(/^\d{4,5}$/.test(D)){let F=Number(D);if(F>2e4&&F<99999){let $=new Date(Date.UTC(1899,11,30)+F*864e5);return isNaN($.getTime())?"":$.toISOString().slice(0,10)}}let S=D.split(/[\/\-\.]/);if(S.length===3){let[F,$,E]=S.map(k=>k.trim());if(F.length===4&&$.length<=2&&E.length<=2)return`${F}-${$.padStart(2,"0")}-${E.padStart(2,"0")}`;if(E.length===4&&$.length<=2&&F.length<=2)return`${E}-${$.padStart(2,"0")}-${F.padStart(2,"0")}`}return D},o=n.map(B=>({submitted_at:u(B.Waktu||B.Tanggal),submitter_name:String(B.Pengirim||"").trim(),branch_id:m(String(B.Cabang||"").trim()),tools_items:String(B["Alat/Barang"]||B.Alat||"").trim(),chemical_items:String(B.Chemical||"").trim(),additional_notes:String(B.Catatan||B.Keterangan||"").trim(),status:String(B.Status||"").trim(),processed_by:String(B["Diproses Oleh"]||B.PIC||"").trim()})).filter(B=>B.submitted_at&&B.submitter_name&&B.branch_id),b=await K("/api/reports/supply/import",{method:"POST",body:JSON.stringify(o)});if(!b.ok)throw new Error(b.data?.error||"Import gagal");return b.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(n,c)=>{let x=G({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(m,u)=>{let o=m.querySelector("#supply-status").value,b=m.querySelector("#supply-processed-by").value;(await K(`/api/reports/supply/${n.id}`,{method:"PUT",body:JSON.stringify({status:o,processed_by:b})})).ok?(R("Status diperbarui."),u(),c()):N("Gagal update status.")}})}}]})}re();async function _a(e){let t=ee();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}de({container:e,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",editor_khusus:"badge-warning",input_lembur:"badge-info",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let n=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!n,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!n,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"editor_khusus",label:"Editor (Hanya Lembur & Masalah)"},{value:"input_lembur",label:"Input Lembur (Hanya Isi Lembur)"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:n?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!n,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:n?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await K(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let n=i.data.data.map(c=>({"Nama Lengkap":c.full_name||"",Username:c.username||"",Email:c.email||"",Role:c.role||"",Status:c.is_active?"Aktif":"Nonaktif"}));Q(n,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{Q([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let n=i.map(x=>({full_name:String(x["Nama Lengkap"]||"").trim(),username:String(x.Username||"").trim(),email:String(x.Email||"").trim(),role:String(x.Role||"").trim()||"viewer",password:String(x.Password||"").trim()})).filter(x=>x.username&&x.password&&x.email&&x.full_name),c=await K("/api/users/import",{method:"POST",body:JSON.stringify(n)});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}}})}re();async function Aa(e){de({container:e,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:t=>t?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:t=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:t?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:t?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:t?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:t?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:t?.is_active!==void 0?t.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let t=await K(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(t.ok)Q(t.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{Q([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async t=>{let i=t.map(c=>({code:String(c["Kode Cabang"]||"").trim(),name:String(c["Nama Pendek"]||"").trim(),full_name:String(c["Nama Lengkap"]||"").trim(),city:String(c.Kota||"").trim()})).filter(c=>c.code&&c.name),n=await K("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}async function Ba(e){let t=new Date,i=[];e.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{t.setMonth(t.getMonth()-1),c()}),document.getElementById("cal-next").addEventListener("click",()=>{t.setMonth(t.getMonth()+1),c()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(m=>m.addEventListener("change",c));async function n(){try{let m=`${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,"0")}`;i=(await K(`/api/dashboard/calendar?month=${m}`)).data?.data||[]}catch(m){console.warn("[Calendar] Failed to load events, rendering empty grid:",m),i=[]}}async function c(){let m=document.getElementById("calendar-grid");if(m){m.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await n();try{let u=t.getFullYear(),o=t.getMonth(),b=t.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),B=document.getElementById("cal-month-label");B&&(B.textContent=b);let D=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(g=>g.value)),S=i.filter(g=>D.has(g.type)),F={};S.forEach(g=>{let s=Mn(g.event_date);s&&(F[s]||(F[s]=[]),F[s].push(g))});let $=new Date(u,o,1).getDay(),E=new Date(u,o+1,0).getDate(),k=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],h=new Date().toISOString().slice(0,10),a='<div class="calendar-grid">';k.forEach(g=>{a+=`<div class="cal-day-header">${g}</div>`});for(let g=0;g<$;g++)a+='<div class="cal-cell cal-cell-empty"></div>';for(let g=1;g<=E;g++){let s=`${u}-${String(o+1).padStart(2,"0")}-${String(g).padStart(2,"0")}`,r=F[s]||[],l=s===h;a+=`
          <div class="cal-cell ${l?"cal-today":""} ${r.length?"cal-has-events":""}"
               data-date="${s}" tabindex="0" role="button" aria-label="${s}">
            <div class="cal-day-num ${l?"today-num":""}">${g}</div>
            <div class="cal-events-preview">
              ${r.slice(0,3).map(f=>`
                <div class="cal-event-dot cal-color-${f.color||"gray"}" title="${we(f.title||f.type)}">
                  <span class="cal-event-dot-label">${jn(f.title||f.branch_name||f.type,18)}</span>
                </div>
              `).join("")}
              ${r.length>3?`<div class="cal-more">+${r.length-3} lagi</div>`:""}
            </div>
          </div>`}let d=($+E)%7;if(d!==0)for(let g=0;g<7-d;g++)a+='<div class="cal-cell cal-cell-empty"></div>';a+="</div>",m.innerHTML=a,m.querySelectorAll(".cal-cell[data-date]").forEach(g=>{g.addEventListener("click",()=>{let s=g.dataset.date,r=F[s]||[];if(!r.length)return;let l=document.getElementById("cal-event-list"),f=new Date(s+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=f,document.getElementById("cal-event-items").innerHTML=r.map(v=>`
            <div class="cal-event-item cal-color-border-${v.color||"gray"}">
              <div class="cal-event-type">
                ${Kn(v.type)}
                ${v.date_type?` <span class="badge" style="font-size:0.75rem;padding:2px 6px;background:rgba(0,0,0,0.06);margin-left:4px;">${we(v.date_type)}</span>`:""}
              </div>
              <div class="cal-event-title" style="font-weight:600;font-size:0.95rem;margin-top:2px;">${we(v.title||"-")}</div>
              ${v.branch_name?`<div class="cal-event-branch" style="color:var(--text-2);font-size:0.85rem;margin-top:2px;">\u{1F4CD} ${we(v.branch_name)}</div>`:""}
              ${v.pic||v.period||v.opening_date||v.target_date||v.completion_date||v.notes?`
                <div class="cal-event-meta" style="margin-top:6px;padding-top:6px;border-top:1px dashed var(--border);font-size:0.8rem;color:var(--text-2);display:flex;flex-direction:column;gap:3px;">
                  ${v.pic?`<div>\u{1F464} <b>PIC:</b> ${we(v.pic)}</div>`:""}
                  ${v.period?`<div>\u{1F4C5} <b>Periode:</b> ${we(v.period)}</div>`:""}
                  ${v.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${we(Dt(v.opening_date))}</div>`:""}
                  ${v.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${we(Dt(v.target_date))}</div>`:""}
                  ${v.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${we(Dt(v.completion_date))}</div>`:""}
                  ${v.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${we(v.notes)}</div>`:""}
                </div>
              `:""}
              ${v.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${v.status==="Done"?"badge-success":"badge-warning"}">${we(v.status)}</span></div>`:""}
              ${v.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${v.days_remaining} hari</div>`:""}
            </div>
          `).join(""),l.style.display="block"})})}catch(u){console.error("[Calendar] Render error:",u),m&&(m.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let x=()=>{n().then(()=>c())};at.on("data:changed",x),c()}function Mn(e){if(!e||e==="-"||String(e).trim()==="")return"";let t=String(e).trim();if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);let i=t.split(/[\/\-\.]/);if(i.length===3){let[n,c,x]=i.map(m=>m.trim());if(x.length===4)return`${x}-${c.padStart(2,"0")}-${n.padStart(2,"0")}`;if(n.length===4)return`${n}-${c.padStart(2,"0")}-${x.padStart(2,"0")}`}return t.slice(0,10)}function Dt(e){if(!e||e==="-"||String(e).trim()==="")return"";let i=String(e).trim().split(/[\/\-\.]/);if(i.length===3){let[n,c,x]=i.map(m=>m.trim());if(n.length===4)return`${x.padStart(2,"0")}-${c.padStart(2,"0")}-${n}`;if(x.length===4)return`${n.padStart(2,"0")}-${c.padStart(2,"0")}-${x}`}return e}function jn(e,t){return e?e.length>t?e.slice(0,t)+"\u2026":e:""}function we(e){return e?String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Kn(e){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging"}[e]||e}async function Ta(e){let t=ee(),i=(t?.full_name||t?.username||"U")[0].toUpperCase(),c={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[t?.role]||"#64748B";e.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${c},${c}99)">
            ${i}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${t?.full_name||"\u2014"}</div>
            <div class="profile-username">@${t?.username||"\u2014"}</div>
            <span class="badge badge-info" style="background:${c}18;color:${c};margin-top:6px">
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
            <span class="info-value" style="color:${c};font-weight:700">${t?.role||"\u2014"}</span>
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
  `;let x=localStorage.getItem("fm_token"),m=document.getElementById("session-info");if(x&&m)try{let u=JSON.parse(atob(x.split(".")[1])),o=new Date(u.exp*1e3);m.textContent=`Berakhir: ${o.toLocaleString("id-ID")}`}catch{m.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async u=>{u.preventDefault();let o=document.getElementById("pwd-error"),b=document.getElementById("pwd-success"),B=document.getElementById("btn-save-pwd");o.style.display="none",b.style.display="none";let D=u.target,S=D.current_password.value,F=D.new_password.value,$=D.confirm_password.value;if(F!==$){o.textContent="\u274C Konfirmasi password tidak cocok.",o.style.display="block";return}if(F.length<6){o.textContent="\u274C Password baru minimal 6 karakter.",o.style.display="block";return}B.disabled=!0,B.textContent="\u23F3 Menyimpan...";let E=await K("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:S,new_password:F})});B.disabled=!1,B.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',E.ok?(b.textContent="\u2705 Password berhasil diubah.",b.style.display="block",D.reset(),R("Password berhasil diubah.")):(o.textContent=E.data?.error||"Gagal mengubah password.",o.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}var ot={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function pe(e){if(e==null||e==="")return null;if(e instanceof Date)return isNaN(e.getTime())?null:e.toISOString().slice(0,10);let t=String(e).trim();if(t===""||t==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);if(/^\d{4,5}$/.test(t)){let c=Number(t);if(c>2e4&&c<99999){let x=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(x.getTime())?null:x.toISOString().slice(0,10)}}let i=t.split(/[\/\-\.]/);if(i.length===3){let[c,x,m]=i.map(B=>B.trim()),u=Number(c),o=Number(x),b=Number(m);if(c.length===4&&u>1900)return`${c}-${x.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&b>1900)return u>12?`${m}-${x.padStart(2,"0")}-${c.padStart(2,"0")}`:o>12?`${m}-${c.padStart(2,"0")}-${x.padStart(2,"0")}`:`${m}-${x.padStart(2,"0")}-${c.padStart(2,"0")}`;if(m.length===2&&!isNaN(b)){let B=b>=50?`19${m}`:`20${m}`;return u>12?`${B}-${x.padStart(2,"0")}-${c.padStart(2,"0")}`:`${B}-${x.padStart(2,"0")}-${c.padStart(2,"0")}`}}let n=new Date(t);return isNaN(n.getTime())?null:n.toISOString().slice(0,10)}function Da(e){return Object.values(e).every(t=>t==null||String(t).trim()==="")}var Nn={validation:{required:[],map:e=>({cabang:e.CABANG,pic:e.PIC,kegiatan:e.KEGIATAN,quartal:e.QUARTAL,masa_pkwt:e["MASA PKWT"],pic_pelapor:e["PIC PELAPOR"],kontrak:e.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:e=>({name:e["Nama SOP"],category:e.Kategori||"Umum",document_link:e["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({full_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||e.Divisi||"FACILITY CARE",phone:e["No. Hp"]||e["No. HP"],join_date:pe(e["Tanggal Masuk"]||e["Tgl Masuk"]),status:e.Status||"",target_pindah_os:pe(e["Target Pindah OS"]),target_selesai_os:pe(e["Target Selesai OS"]),notes:e.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({employee_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||"FACILITY CARE",start_date:pe(e["Tanggal Mulai"]),end_date:pe(e["Tanggal Selesai"]),contract_type:e["Tipe Kontrak"]||"",pkwt_number:e.PKWT||"",status:e.Status||"",notes:e.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:e=>({report_date:pe(e["Tanggal Info"]),branch_name:e.Cabang,category:e.Kategori,source:e["Sumber Laporan"],complaint:e.Keluhan,employee_name:e["Nama FC"],fc_specialist:e["FC Spesialis"],solution:e.Solusi,status:e.Status||"",completion_date:pe(e["Tanggal Selesai"])})},one_on_one:{required:[],map:e=>({meeting_date:pe(e.Tanggal),branch_name:e.Cabang,employee_name:e["Nama Karyawan"],pic:e.Pic,problem:e.Masalah,solution:e.Solusi,status:e.Status||"",completion_date:pe(e["Tanggal Selesai"]),document_link:e["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:e=>({branch_name:e.Cabang,activity_type:e.Kegiatan,period:e.Periode,pic:e.Pic||e.PIC,opening_date:pe(e["Tanggal Opening"]||e["Tgl Opening"]),target_date:pe(e["Tanggal Target"]||e["Tgl Target"]),completion_date:pe(e["Tanggal Selesai"]||e["Tgl Selesai"]),status:e.Status||"",notes:e.Keterangan||e.Catatan})},inspection:{required:[],map:e=>({inspection_date:pe(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",fc_score:e["Point FC SP"]!==void 0&&e["Point FC SP"]!==null?parseFloat(String(e["Point FC SP"]).replace(",",".")):null,spv_score:e["Point SPV"]!==void 0&&e["Point SPV"]!==null?parseFloat(String(e["Point SPV"]).replace(",",".")):null,document_link:e.Link,notes:""})},cleaning:{required:[],map:e=>({activity_date:pe(e.Tanggal),branch_name:e.Cabang,activity_type:e["Jenis Kegiatan"]||"General Cleaning",period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},fogging:{required:[],map:e=>({activity_date:pe(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:e=>({info_date:pe(e["Tgl Info"]),branch_name:e.Cabang,problem:e.Permasalahan,pic:e.PIC,done_date:pe(e["Tgl Done"]),status:e.Status||"",notes:e.Ket})},relievers:{required:[],map:e=>({branch_name:e.Cabang,original_fc_name:e["Nama Facility care"],period:e.Periode,reliever_name:e.Relifer,backup_date:pe(e["Tanggal Back Up"]),completion_date:pe(e["Tanggal Selesai"]),reason:e.Keterangan,shift:e.Shift,status:e.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:e=>({training_date:pe(e.Tanggal),batch:e.Batch,subject:e.Materi,participants:e.Peserta,branch_name:e.Cabang,trainer:e.Trainer,score:e.Nilai!==void 0&&e.Nilai!==null?parseFloat(String(e.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:e=>({name:e["Master Checklist"],category:"Umum",document_link:e["Link Document"],description:""})},supply:{required:[],map:e=>({submitted_at:pe(e.Timestamp),submitter_name:e["Nama Lengkap"],branch_name:e["Kebutuhan Untuk Cabang"],tools_items:e["Alat - Alat / Barang"],tools_quantity:e["Jumlah Permintaan Alat / Barang"],chemical_items:e.Chemical,chemical_quantity:e["Jumlah Permintaan Chemical"],additional_notes:e["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:e.Status||""})}};function qn(e,t){let i=ot[e];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let n=Nn[i.module];if(!n)return{valid:[],errors:[],mapped:[],skipped:!0};let c=[],x=[],m=[];return t.filter(o=>!Da(o)).forEach((o,b)=>{let B=t.indexOf(o)+2,D=[];n.required.forEach(({key:F,label:$})=>{let E=o[F];if(E==null||String(E).trim()===""){let k=Object.keys(o).filter(h=>h.trim()).join(", ");D.push({column:$,originalValue:E||"",reason:`Kolom "${$}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${k.slice(0,120)}`})}});let S=n.map(o);D.length>0?x.push({row:B,data:S,raw:o,errors:D}):(c.push(o),m.push(S))}),{valid:c,errors:x,mapped:m}}function Pa(e){let t=[];return e.SheetNames.forEach(i=>{let n=ot[i];if(!n)return;let c=e.Sheets[i],x=window.XLSX.utils.sheet_to_json(c,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),m=qn(i,x),u=x.filter(o=>!Da(o));t.push({sheetName:i,module:n.module,label:n.label,total:u.length,valid:m.mapped.length,errorCount:m.errors.length,errors:m.errors,mapped:m.mapped,skipped:!1})}),t}function La(){let e=window.XLSX,t=e.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([n,c])=>{e.utils.book_append_sheet(t,e.utils.json_to_sheet(c),n)}),e.writeFile(t,"Template_Import_Data_Awal_FCMS.xlsx")}function Ia(e){let t=window.XLSX,i=t.utils.book_new(),n=!1;return e.forEach(c=>{if(!c.errors||c.errors.length===0)return;n=!0;let x=c.errors.map(u=>({"No. Baris":u.row,"Kolom Gagal":(u.errors||[]).map(o=>o.column||o).join("; "),"Alasan Error":(u.errors||[]).map(o=>o.reason||o).join("; "),...Object.fromEntries(Object.entries(u.data||{}).map(([o,b])=>[o,b??""]))})),m=t.utils.json_to_sheet(x);t.utils.book_append_sheet(i,m,c.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),n?(t.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Rn=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","sop","inspection","cleaning","fogging","basecamp","supply"];function za(e){e.innerHTML=`
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
              ${Object.entries(ot).map(([E,{label:k}])=>`<span class="import-sheet-tag">\u{1F4C4} ${E} \u2192 ${k}</span>`).join("")}
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
  `;let t=null,i=null,n=0,c={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function x(E){Object.entries(c).forEach(([k,h])=>{h.style.display=k===E?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let E=document.getElementById("btn-backup-db");E.disabled=!0,E.textContent="\u23F3 Memproses Backup...";try{let k=await K("/api/import/backup");if(k.ok){if(!window.XLSX){N("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let h=window.XLSX,a=h.utils.book_new();Object.entries(k.data.database).forEach(([p,d])=>{let g=d.length>0?d:[{}],s=h.utils.json_to_sheet(g);h.utils.book_append_sheet(a,s,p.substring(0,31))}),h.writeFile(a,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),R("Backup berhasil diunduh!")}else N("Gagal memproses backup: "+(k.data?.error||"Unknown error"))}catch(k){N("Gagal memproses backup: "+k.message)}finally{E.disabled=!1,E.textContent="\u{1F4E6} Backup Database"}});let m=document.getElementById("btn-sync-google");m&&m.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let E=m.innerHTML;m.innerHTML='<span class="spinner"></span> Menyinkronkan...',m.disabled=!0;try{let k=await K("/api/sync/google-sheets",{method:"POST"});k.ok?alert("Sinkronisasi Berhasil: "+(k.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(k.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{m.innerHTML=E,m.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{La(),R("Template Excel berhasil didownload!")});let u=document.getElementById("file-input"),o=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",E=>{E.stopPropagation(),u.click()}),u.addEventListener("change",E=>{E.target.files[0]&&b(E.target.files[0])}),o.addEventListener("dragover",E=>{E.preventDefault(),o.classList.add("drag-over")}),o.addEventListener("dragleave",()=>o.classList.remove("drag-over")),o.addEventListener("drop",E=>{E.preventDefault(),o.classList.remove("drag-over");let k=E.dataTransfer.files[0];k&&k.name.match(/\.xlsx?$/i)?b(k):N("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{t=null,u.value="",document.getElementById("file-info").style.display="none",o.style.display="",x("upload")});async function b(E){t=E,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${E.name} (${(E.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",o.style.display="none",await B(E)}async function B(E){x("validating");let k=document.getElementById("validation-status"),h=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");k.textContent="Membaca file Excel...",h.style.width="20%",await He(200);let a=await E.arrayBuffer(),p=window.XLSX.read(a,{type:"array",cellDates:!0});k.textContent=`Memvalidasi ${p.SheetNames.length} sheet...`,h.style.width="50%",await He(100),i=Pa(p),h.style.width="100%",k.textContent="Validasi selesai!",await He(300),D()}catch(a){x("upload"),N("Gagal memproses file: "+a.message),document.getElementById("file-info").style.display="flex",o.style.display="none"}}function D(){x("preview");let E=i.filter(l=>!l.skipped).length,k=i.reduce((l,f)=>l+f.total,0),h=i.reduce((l,f)=>l+f.valid,0),a=i.reduce((l,f)=>l+f.errorCount,0),p=k>0?Math.round(h/k*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${E} sheet</span>
      <span class="badge badge-secondary">${k} baris</span>
      <span class="badge badge-success">${h} valid (${p}%)</span>
      ${a>0?`<span class="badge badge-danger">${a} error</span>`:""}
    `;let d=document.getElementById("preview-table-container");d.innerHTML=`
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
          ${i.map((l,f)=>`
            <tr class="${l.errorCount>0?"row-error":l.skipped?"row-skipped":"row-ok"}">
              <td><strong>${l.sheetName}</strong></td>
              <td>${l.label}</td>
              <td style="text-align:center">${l.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${l.valid}</span></td>
              <td style="text-align:center">${l.errorCount>0?`<span class="badge badge-danger">${l.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${l.skipped?'<span class="badge badge-neutral">Dilewati</span>':l.errorCount>0&&l.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':l.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':l.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${l.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${f}">\u{1F50D} ${l.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,d.querySelectorAll(".btn-detail-error").forEach(l=>{l.addEventListener("click",()=>{let f=i[Number(l.dataset.idx)];S(f)})});let g=document.getElementById("error-detail-section"),s=document.getElementById("error-detail-container");s.innerHTML="",g.style.display="none";let r=document.getElementById("btn-start-import");h===0?(r.disabled=!0,r.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(r.disabled=!1,a>0?(r.innerHTML=`\u{1F680} Import ${h} Data Valid (${a} dilewati)`,r.title="Baris error akan dilewati, baris valid tetap diimport"):r.innerHTML=`\u{1F680} Mulai Import ${h} Data`)}function S(E){let k=document.getElementById("error-detail-section"),h=document.getElementById("error-detail-container");k.style.display="";let a=E.errors.slice(0,100).map(p=>(Array.isArray(p.errors)?p.errors:[]).map(g=>{let s=typeof g=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${p.row}</span></td>
            <td><strong>${s?g.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${s&&g.originalValue!==void 0?g.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${s?g.reason:g}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${s&&g.aliases?`Gunakan salah satu nama kolom:<br><em>${g.aliases}</em>`:s&&g.hint?g.hint:""}
            </td>
          </tr>
        `}).join("")).join("");h.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${E.sheetName} \u2014 ${E.errorCount} baris error dari ${E.total} total
          ${E.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
            <tbody>${a||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${E.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,k.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{x("upload"),document.getElementById("file-info").style.display="none",o.style.display="",t=null,u.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;Ia(i)?R("Log error berhasil didownload."):R("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let E=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";F(E)});async function F(E){x("importing"),n=Date.now();let k=[];Rn.forEach(g=>{let s=i?.find(r=>r.module===g&&r.mapped?.length>0);s&&k.push(s)});let h=document.getElementById("import-steps-list");h.innerHTML=k.map(g=>`
      <div class="import-step-item" id="step-item-${g.module}">
        <span class="step-item-icon" id="step-icon-${g.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${g.label} <span class="step-item-count">(${g.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${g.module}"></span>
      </div>
    `).join("");let a=document.getElementById("import-bar"),p=document.getElementById("import-current-status"),d={totalSheets:k.length,totalRows:k.reduce((g,s)=>g+s.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let g=0;g<k.length;g++){let s=k[g],r=document.getElementById(`step-icon-${s.module}`),l=document.getElementById(`step-status-${s.module}`);r.textContent="\u{1F504}",l.textContent="Mengimport...",p.textContent=`Mengimport ${s.label}...`,a.style.width=`${Math.round(g/k.length*100)}%`;try{let f=await K(`/api/import/${s.module}`,{method:"POST",body:JSON.stringify({rows:s.mapped,onDuplicate:E})});if(f.ok){let v=f.data;d.inserted+=v.inserted||0,d.skipped+=v.skipped||0,d.moduleResults.push({label:s.label,inserted:v.inserted||0,skipped:v.skipped||0,status:"ok"}),r.textContent="\u2705",l.innerHTML=`<span class="badge badge-success">${v.inserted||0} berhasil</span>${v.skipped>0?` <span class="badge badge-neutral">${v.skipped} skip</span>`:""}`}else d.failed++,d.moduleResults.push({label:s.label,inserted:0,skipped:0,status:"error",error:f.data?.error}),r.textContent="\u274C",l.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(f){d.failed++,d.moduleResults.push({label:s.label,inserted:0,skipped:0,status:"error",error:f.message}),r.textContent="\u274C",l.innerHTML='<span class="badge badge-danger">Gagal</span>'}await He(150)}a.style.width="100%",p.textContent="Selesai!",await He(400),$(d)}function $(E){x("summary");let k=((Date.now()-n)/1e3).toFixed(1),h=E.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${h?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${h?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${E.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${E.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${E.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${E.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${E.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${E.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
      <div class="summary-stat-card info">
        <div class="stat-value">${k}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${E.moduleResults.map(a=>`
            <tr>
              <td>${a.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${a.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${a.skipped}</span></td>
              <td style="text-align:center">
                ${a.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${a.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{t=null,i=null,u.value="",document.getElementById("file-info").style.display="none",o.style.display="",x("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function He(e){return new Promise(t=>setTimeout(t,e))}le();var Ma=[{id:1,employee_name:"Teguh Ariyandi",division:"FACILITY CARE",branch_id:18,branch_name:"018. Harapan Indah",tanggal:"2025-08-19",akhir_sp:"2026-02-18",sp_type:"SP 1",status:"Aktif",document_link:"",notes:"Keterlambatan berulang dan ketidaklengkapan checklist harian lantai dasar."},{id:2,employee_name:"Andri",division:"FACILITY CARE",branch_id:1,branch_name:"001. Pondok Bambu",tanggal:"2025-04-10",akhir_sp:"2025-10-09",sp_type:"SP 1",status:"Selesai",document_link:"",notes:"Pelanggaran grooming dan standar pembersihan area tindakan tanpa APD lengkap."},{id:3,employee_name:"M. FIRMANSYAH",division:"FACILITY CARE",branch_id:25,branch_name:"025. Bogor Baru",tanggal:"2024-03-21",akhir_sp:"2024-09-20",sp_type:"SP 2",status:"Selesai",document_link:"",notes:"Meninggalkan area tugas tanpa konfirmasi supervisor selama jam operasional klinik."}];function Me(e){if(!e)return{status:"Tidak Ditentukan",label:"Tidak Ada Tanggal",bg:"#F1F5F9",color:"#64748B",border:"#CBD5E1",icon:"\u26AA",daysLeft:0,isExpired:!1,isExpiringSoon:!1};let t=new Date;t.setHours(0,0,0,0);let i=new Date(e);i.setHours(0,0,0,0);let n=i.getTime()-t.getTime(),c=Math.ceil(n/(1e3*60*60*24));return c<0?{status:"Selesai",label:`Masa Sanksi Selesai (${Math.abs(c)} hari lalu)`,shortLabel:"Selesai",bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2705",daysLeft:c,isExpired:!0,isExpiringSoon:!1}:c<=30?{status:"Mendekati Berakhir",label:`Segera Berakhir (Sisa ${c} hari)`,shortLabel:`Sisa ${c} hari`,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F",daysLeft:c,isExpired:!1,isExpiringSoon:!0}:{status:"Aktif",label:`Aktif Berlaku (Sisa ${c} hari)`,shortLabel:`Aktif (${c} hari)`,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F534}",daysLeft:c,isExpired:!1,isExpiringSoon:!1}}function ja(e=""){let t=(e||"").toUpperCase().trim();return t==="SP 3"?{label:"SP 3 (Terakhir)",bg:"#7F1D1D",color:"#FFFFFF",border:"#991B1B",icon:"\u{1F6A8}"}:t==="SP 2"?{label:"SP 2",bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F6D1}"}:t==="SP 1"?{label:"SP 1",bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F"}:t.includes("TEGURAN")?{label:"Teguran Lisan",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F4AC}"}:{label:e||"SP",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4DC}"}}function st(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Ka(e){let t=[],i=[],n=[],c="all",x="all",m="all",u="",o="fcms-sp-page-styles";if(!document.getElementById(o)){let a=document.createElement("style");a.id=o,a.textContent=`
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
        .sp-table-card {
          border-radius: 14px !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
        }
        .sp-table {
          min-width: 650px !important;
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
    `,document.head.appendChild(a)}async function b(){try{let[a,p]=await Promise.all([te(),ue()]);i=a||[],n=p||[];let d=await K("/api/sp?limit=1000");if(d.ok&&d.data?.data&&Array.isArray(d.data.data)&&d.data.data.length>0)t=d.data.data,localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t));else{let g=localStorage.getItem("fcms_sp_data_cache");g?t=JSON.parse(g):(t=[...Ma],localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t)))}}catch(a){console.warn("Fallback loading SP data from cache/seed:",a);let p=localStorage.getItem("fcms_sp_data_cache");t=p?JSON.parse(p):[...Ma]}k()}function B(){return t.filter(a=>{let p=u.toLowerCase().trim(),d=!p||(a.employee_name||"").toLowerCase().includes(p)||(a.branch_name||"").toLowerCase().includes(p)||(a.division||"").toLowerCase().includes(p)||(a.sp_type||"").toLowerCase().includes(p),g=x==="all"||String(a.branch_id)===String(x),s=m==="all"||a.division===m,r=Me(a.akhir_sp),l=!0;return c==="active"?l=!r.isExpired:c==="expiring"?l=r.isExpiringSoon:c==="expired"?l=r.isExpired:c==="sp1"?l=(a.sp_type||"").toUpperCase()==="SP 1":c==="sp2"?l=(a.sp_type||"").toUpperCase()==="SP 2":c==="sp3"?l=(a.sp_type||"").toUpperCase()==="SP 3":c==="teguran"&&(l=(a.sp_type||"").toUpperCase().includes("TEGURAN")),d&&g&&s&&l})}function D(){let a=0,p=0,d=0,g=0,s=0,r=0,l=0,f=new Set;return t.forEach(v=>{v.employee_name&&f.add(v.employee_name.toLowerCase().trim());let L=Me(v.akhir_sp);L.isExpired?d++:a++,L.isExpiringSoon&&p++;let A=(v.sp_type||"").toUpperCase();A==="SP 1"?g++:A==="SP 2"?s++:A==="SP 3"?r++:A.includes("TEGURAN")&&l++}),{total:t.length,active:a,expiring:p,expired:d,sp1:g,sp2:s,sp3:r,teguran:l,uniqueEmployees:f.size}}function S(a=null){let p=!!a,d=a?.tanggal?a.tanggal.slice(0,10):new Date().toISOString().slice(0,10),g=a?.akhir_sp?a.akhir_sp.slice(0,10):"";if(!g){let f=new Date(d);f.setMonth(f.getMonth()+6),g=f.toISOString().slice(0,10)}let s=document.createElement("div");s.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <input type="text" id="sp-modal-emp" list="sp-emp-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih atau ketik nama..." value="${a?.employee_name?String(a.employee_name).replace(/"/g,"&quot;"):""}" required />
            <datalist id="sp-emp-list">
              ${n.map(f=>`<option value="${f.label}">${f.label}</option>`).join("")}
            </datalist>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-div" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="FACILITY CARE" ${a?.division==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${a?.division==="SECURITY"?"selected":""}>SECURITY</option>
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
              ${i.map(f=>`<option value="${f.value}" ${String(a?.branch_id)===String(f.value)?"selected":""}>${f.label}</option>`).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tingkat / Jenis SP <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="SP 1" ${a?.sp_type==="SP 1"?"selected":""}>SP 1 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 2" ${a?.sp_type==="SP 2"?"selected":""}>SP 2 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 3" ${a?.sp_type==="SP 3"?"selected":""}>SP 3 (Masa Berlaku 6 Bulan / Terakhir)</option>
              <option value="Teguran Lisan" ${a?.sp_type==="Teguran Lisan"?"selected":""}>Teguran Lisan (Pembinaan)</option>
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
              value="${g}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Link Berkas / Foto BAP / Scan Surat SP
          </label>
          <input type="url" id="sp-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://drive.google.com/..." value="${a?.document_link?String(a.document_link).replace(/"/g,"&quot;"):""}" />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Kronologi Pelanggaran / Catatan Pembinaan
          </label>
          <textarea id="sp-modal-notes" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan ringkasan pelanggaran disiplin kerja yang terjadi...">${a?.notes||a?.description||""}</textarea>
        </div>
      </div>
    `;let r=s.querySelector("#sp-modal-tanggal"),l=s.querySelector("#sp-modal-akhir");r.addEventListener("change",()=>{if(r.value){let f=new Date(r.value);f.setMonth(f.getMonth()+6),l.value=f.toISOString().slice(0,10)}}),G({title:p?"\u270F\uFE0F Edit Data Surat Peringatan (SP)":"\u2795 Terbitkan Surat Peringatan (SP) Baru",content:s,size:"md",confirmText:p?"Simpan Perubahan":"Terbitkan SP",confirmClass:"btn-primary",onConfirm:async(f,v)=>{let L=(s.querySelector("#sp-modal-emp").value||"").trim(),A=s.querySelector("#sp-modal-div").value,P=s.querySelector("#sp-modal-branch").value,w=s.querySelector("#sp-modal-type").value,I=s.querySelector("#sp-modal-tanggal").value,z=s.querySelector("#sp-modal-akhir").value,C=(s.querySelector("#sp-modal-link").value||"").trim(),y=(s.querySelector("#sp-modal-notes").value||"").trim();if(!L){N("Nama Karyawan wajib diisi");return}if(!P){N("Cabang penempatan wajib dipilih");return}if(!I||!z){N("Tanggal SP dan Akhir SP wajib diisi");return}let _=i.find(O=>String(O.value)===String(P)),T=_?_.label:"Cabang",j={employee_name:L,division:A,branch_id:Number(P),branch_name:T,sp_type:w,tanggal:I,akhir_sp:z,document_link:C,notes:y,status:Me(z).status};try{if(p){let O=await K(`/api/sp/${a.id}`,{method:"PUT",body:JSON.stringify(j)}),M=t.findIndex(q=>q.id===a.id);M!==-1&&(t[M]={...t[M],...j}),R("Data SP berhasil diperbarui")}else{let O=await K("/api/sp",{method:"POST",body:JSON.stringify(j)}),M=O.ok&&O.data?.id?O.data.id:Date.now();t.unshift({id:M,...j}),R("Surat Peringatan baru berhasil diterbitkan")}localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t)),v(),k()}catch(O){console.error("Error saving SP:",O),N("Gagal menyimpan SP: "+O.message)}}})}function F(a){X(`Hapus data Surat Peringatan untuk karyawan <strong>${a.employee_name}</strong> (${a.sp_type})?`,async()=>{try{await K(`/api/sp/${a.id}`,{method:"DELETE"})}catch(p){console.warn("API delete error fallback to local:",p)}t=t.filter(p=>p.id!==a.id),localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t)),R(`Data SP ${a.employee_name} berhasil dihapus`),k()},"Hapus Data SP")}function $(a){let p=Me(a.akhir_sp),d=ja(a.sp_type),g=document.createElement("div");g.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div style="display:flex; align-items:center; gap:14px; padding-bottom:14px; border-bottom:1px solid #E2E8F0;">
          <div class="sp-avatar-box" style="width:54px; height:54px; font-size:1.2rem;">
            ${a.employee_name.slice(0,2).toUpperCase()}
          </div>
          <div>
            <h3 style="margin:0 0 4px 0; font-size:1.15rem; font-weight:700; color:#0F172A;">
              ${a.employee_name}
            </h3>
            <div style="display:flex; align-items:center; gap:8px;">
              <span class="badge" style="background:#EFF6FF; color:#1D4ED8; font-weight:700;">${a.division||"FACILITY CARE"}</span>
              <span style="color:#64748B; font-size:0.85rem;">\u{1F4CD} ${a.branch_name||"Cabang"}</span>
            </div>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Tingkat Sanksi</div>
            <span class="sp-type-badge" style="background:${d.bg}; color:${d.color}; border:1px solid ${d.border}; font-size:0.85rem;">
              <span>${d.icon}</span>
              <span>${d.label}</span>
            </span>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Status Masa Berlaku</div>
            <span class="sp-status-badge" style="background:${p.bg}; color:${p.color}; border:1px solid ${p.border}; font-size:0.85rem;">
              <span>${p.icon}</span>
              <span>${p.label}</span>
            </span>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Tanggal Terbit SP</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${st(a.tanggal)}</div>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Akhir Masa Berlaku</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${st(a.akhir_sp)}</div>
          </div>
        </div>

        ${a.notes||a.description?`
          <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:10px; padding:14px;">
            <div style="font-size:0.78rem; font-weight:700; color:#92400E; margin-bottom:4px;">\u{1F4DD} Catatan Pelanggaran / Pembinaan:</div>
            <div style="font-size:0.88rem; color:#78350F; line-height:1.5;">${a.notes||a.description}</div>
          </div>
        `:""}

        ${a.document_link?`
          <div>
            <a href="${a.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-primary" style="width:100%; justify-content:center; padding:10px; font-weight:600; display:flex; align-items:center; gap:8px;">
              <span>\u{1F4C4}</span> Buka Dokumen BAP / Scan SP Fisik
            </a>
          </div>
        `:`
          <div style="text-align:center; padding:8px; font-size:0.82rem; color:#94A3B8;">
            (Belum ada scan dokumen / BAP yang diunggah)
          </div>
        `}
      </div>
    `,G({title:"\u{1F4CB} Lembar Riwayat Surat Peringatan (SP)",content:g,size:"md",cancelText:"Tutup"})}async function E(){try{let{downloadExcel:a}=await Promise.resolve().then(()=>(re(),he)),p=t.map((d,g)=>{let s=Me(d.akhir_sp);return{No:g+1,"Nama Karyawan":d.employee_name||"",Divisi:d.division||"",Cabang:d.branch_name||"","Jenis SP":d.sp_type||"","Tanggal SP":d.tanggal||"","Akhir SP":d.akhir_sp||"","Status Masa Berlaku":s.status,"Sisa Hari":s.daysLeft>0?s.daysLeft:0,"Link BAP / Foto":d.document_link||"","Catatan / Kronologi":d.notes||d.description||""}});a(p,`Data_Surat_Peringatan_FCMS_${new Date().toISOString().slice(0,10)}`),R("Berhasil mengunduh data SP ke format Excel.")}catch(a){N("Gagal ekspor: "+a.message)}}function k(){let a=B(),p=D();e.innerHTML=`
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
              <button id="sp-btn-add" class="btn btn-primary sp-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                <span>Terbitkan SP</span>
              </button>
              
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
                <div class="sp-stat-value">${p.active}</div>
                <div class="sp-stat-label">SP Aktif Berlaku</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">\u26A0\uFE0F</div>
              <div>
                <div class="sp-stat-value">${p.expiring}</div>
                <div class="sp-stat-label">Segera Berakhir (&lt;30 Hari)</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">\u2705</div>
              <div>
                <div class="sp-stat-value">${p.expired}</div>
                <div class="sp-stat-label">Masa Sanksi Selesai</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">\u{1F465}</div>
              <div>
                <div class="sp-stat-value">${p.uniqueEmployees} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">karyawan</span></div>
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
              <input type="text" id="sp-search-input" class="sp-search-input" placeholder="Cari nama karyawan, cabang, atau divisi..." value="${u?u.replace(/"/g,"&quot;"):""}" />
            </div>

            <!-- Branch Dropdown -->
            <select id="sp-filter-branch" class="sp-filter-select">
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(d=>`<option value="${d.value}" ${String(x)===String(d.value)?"selected":""}>${d.label}</option>`).join("")}
            </select>

            <!-- Division Dropdown -->
            <select id="sp-filter-division" class="sp-filter-select">
              <option value="all" ${m==="all"?"selected":""}>Semua Divisi</option>
              <option value="FACILITY CARE" ${m==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${m==="SECURITY"?"selected":""}>SECURITY</option>
            </select>
          </div>

          <!-- Filter Pills -->
          <div class="sp-filter-pills">
            <button class="sp-pill ${c==="all"?"active":""}" data-filter="all">
              <span>Semua SP</span>
              <span class="sp-pill-count">${p.total}</span>
            </button>
            <button class="sp-pill ${c==="active"?"active":""}" data-filter="active">
              <span>\u{1F534} SP Aktif</span>
              <span class="sp-pill-count">${p.active}</span>
            </button>
            <button class="sp-pill ${c==="expiring"?"active":""}" data-filter="expiring">
              <span>\u26A0\uFE0F &lt;30 Hari</span>
              <span class="sp-pill-count">${p.expiring}</span>
            </button>
            <button class="sp-pill ${c==="expired"?"active":""}" data-filter="expired">
              <span>\u2705 Selesai</span>
              <span class="sp-pill-count">${p.expired}</span>
            </button>
            <button class="sp-pill ${c==="sp1"?"active":""}" data-filter="sp1">
              <span>SP 1</span>
              <span class="sp-pill-count">${p.sp1}</span>
            </button>
            <button class="sp-pill ${c==="sp2"?"active":""}" data-filter="sp2">
              <span>SP 2</span>
              <span class="sp-pill-count">${p.sp2}</span>
            </button>
            <button class="sp-pill ${c==="sp3"?"active":""}" data-filter="sp3">
              <span>SP 3</span>
              <span class="sp-pill-count">${p.sp3}</span>
            </button>
          </div>
        </div>

        <!-- Content Area -->
        <div id="sp-list-container">
          ${a.length===0?`
            <div class="sp-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data SP Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="sp-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:`
            <!-- Grid Card View -->
            <div class="sp-cards-grid">
              ${a.map(d=>{let g=Me(d.akhir_sp),s=ja(d.sp_type);return`
                  <div class="sp-card" data-id="${d.id}">
                    <div>
                      <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px;">
                        <span class="sp-type-badge" style="background:${s.bg}; color:${s.color}; border:1px solid ${s.border};">
                          <span>${s.icon}</span>
                          <span>${s.label}</span>
                        </span>
                        <span class="sp-status-badge" style="background:${g.bg}; color:${g.color}; border:1px solid ${g.border};">
                          <span>${g.icon}</span>
                          <span>${g.shortLabel}</span>
                        </span>
                      </div>

                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="sp-avatar-box" style="width:48px; height:48px; font-size:1.05rem;">
                          ${d.employee_name.slice(0,2).toUpperCase()}
                        </div>
                        <div>
                          <h3 style="margin:0 0 3px 0; font-size:1.1rem; font-weight:700; color:#0F172A;">
                            ${d.employee_name}
                          </h3>
                          <div style="font-size:0.8rem; color:#64748B;">
                            \u{1F4CD} ${d.branch_name||"Cabang"} \u2022 <span style="color:#2563EB; font-weight:600;">${d.division||"FC"}</span>
                          </div>
                        </div>
                      </div>

                      <div style="background:#F8FAFC; border-radius:10px; padding:10px 12px; margin-bottom:14px; font-size:0.82rem; display:flex; justify-content:space-between;">
                        <div>
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Tanggal Terbit</div>
                          <div style="font-weight:700; color:#1E293B;">${st(d.tanggal)}</div>
                        </div>
                        <div style="text-align:right;">
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Akhir SP</div>
                          <div style="font-weight:700; color:#0F172A;">${st(d.akhir_sp)}</div>
                        </div>
                      </div>

                      <p style="font-size:0.83rem; color:#64748B; line-height:1.45; margin:0 0 16px 0; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">
                        ${d.notes||d.description||"Pencatatan sanksi peringatan disiplin operasional."}
                      </p>
                    </div>

                    <div style="padding-top:12px; border-top:1px solid #F1F5F9; display:flex; align-items:center; justify-content:space-between;">
                      <button class="btn btn-sm btn-ghost sp-detail-btn" style="border:1px solid #CBD5E1; border-radius:8px; font-weight:600; font-size:0.82rem; padding:6px 12px;" data-id="${d.id}">
                        \u{1F441}\uFE0F Rincian BAP
                      </button>

                      <div style="display:flex; align-items:center; gap:6px;">
                        ${d.document_link?`
                          <a href="${d.document_link}" target="_blank" rel="noopener noreferrer" class="sp-icon-btn" title="Buka Dokumen">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                          </a>
                        `:""}
                        <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${d.id}">
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${d.id}">
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
    `,h()}function h(){let a=e.querySelector("#sp-btn-add");a&&(a.onclick=()=>S());let p=e.querySelector("#sp-btn-excel");p&&(p.onclick=()=>E());let d=e.querySelector("#sp-btn-refresh");d&&(d.onclick=()=>b());let g=e.querySelector("#sp-btn-reset-filter");g&&(g.onclick=()=>{u="",x="all",m="all",c="all",k()});let s=e.querySelector("#sp-search-input");s&&(s.oninput=f=>{u=f.target.value,k();let v=e.querySelector("#sp-search-input");v&&(v.focus(),v.setSelectionRange(v.value.length,v.value.length))});let r=e.querySelector("#sp-filter-branch");r&&(r.onchange=f=>{x=f.target.value,k()});let l=e.querySelector("#sp-filter-division");l&&(l.onchange=f=>{m=f.target.value,k()}),e.querySelectorAll(".sp-pill").forEach(f=>{f.onclick=()=>{c=f.getAttribute("data-filter")||"all",k()}}),e.querySelectorAll(".sp-detail-btn, .sp-view-detail-link").forEach(f=>{f.onclick=v=>{v.stopPropagation();let L=Number(f.getAttribute("data-id")),A=t.find(P=>P.id===L);A&&$(A)}}),e.querySelectorAll(".sp-edit-btn").forEach(f=>{f.onclick=v=>{v.stopPropagation();let L=Number(f.getAttribute("data-id")),A=t.find(P=>P.id===L);A&&S(A)}}),e.querySelectorAll(".sp-delete-btn").forEach(f=>{f.onclick=v=>{v.stopPropagation();let L=Number(f.getAttribute("data-id")),A=t.find(P=>P.id===L);A&&F(A)}})}await b()}le();var Pt=[{id:1,tanggal:"2026-10-02",employee_name:"TAUFIK HIDAYAT",from_branch_id:12,from_branch_name:"012. BSD",to_branch_id:0,to_branch_name:"000. Management",status:"Selesai",document_link:""},{id:2,tanggal:"2026-09-30",employee_name:"ELVAN JANGGA PRAMANA",from_branch_id:55,from_branch_name:"055. Gajah Mada",to_branch_id:0,to_branch_name:"000. Management",status:"Proses",document_link:""},{id:3,tanggal:"2026-06-23",employee_name:"Widya Astuti",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:7,to_branch_name:"007. Bekasi",status:"Selesai",document_link:""},{id:4,tanggal:"2026-06-22",employee_name:"Gema Azan Saputra",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:21,to_branch_name:"021. Kelapa Gading Boulevard",status:"Selesai",document_link:""},{id:5,tanggal:"2026-06-22",employee_name:"Ahmad Farhan",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:62,to_branch_name:"062. Summarecon Bekasi",status:"Selesai",document_link:""}];function On(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Na(e){let t=[],i=[],n=[],c="all",x="all",m="all",u="",o=!1,B=ee()?.role!=="viewer";try{[i,n]=await Promise.all([te(),ue()])}catch(a){console.error("Error fetching cache:",a)}async function D(){o=!0,k();try{let a=await K("/api/mutasi?limit=10000");if(a.ok){let p=a.data?.data||a.data||[];t=Array.isArray(p)&&p.length>0?p:Pt}else t=Pt}catch(a){console.warn("Gagal load mutasi, menggunakan data lokal:",a),t=Pt}finally{o=!1,k()}}async function S(){try{let{downloadExcel:a}=await Promise.resolve().then(()=>(re(),he)),p=t.map((d,g)=>({No:g+1,Tanggal:d.tanggal||"","Nama Karyawan":d.employee_name||"","Cabang Asal":d.from_branch_name||"","Cabang Tujuan":d.to_branch_name||"",Status:d.status||"","Link Dokumen":d.document_link||""}));a(p,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`),R("Data mutasi berhasil diekspor ke Excel!")}catch(a){N("Gagal ekspor Excel: "+a.message)}}async function F(a){if(!B)return;let p=t.find(d=>d.id===a);X({title:"Hapus Data Mutasi",message:`Apakah Anda yakin ingin menghapus data mutasi untuk "${p?.employee_name||"karyawan ini"}"?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let d=await K(`/api/mutasi/${a}`,{method:"DELETE"});d.ok?(R("Data mutasi berhasil dihapus"),t=t.filter(g=>g.id!==a),k()):N(d.data?.error||"Gagal menghapus mutasi")}catch(d){N("Gagal menghapus data mutasi: "+d.message)}}})}function $(a=null){if(!B)return;let p=!!a,d=a?.tanggal?a.tanggal.slice(0,10):new Date().toISOString().slice(0,10),g=document.createElement("div");g.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="mutasi-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${d}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Proses" ${a?.status==="Proses"?"selected":""}>\u23F3 Proses</option>
              <option value="Selesai" ${a?.status==="Selesai"||!a?"selected":""}>\u2705 Selesai</option>
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Karyawan <span style="color:#DC2626;">*</span>
          </label>
          <input list="emp-list-options" id="mutasi-modal-emp" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Ketik atau pilih nama karyawan..." value="${a?.employee_name?String(a.employee_name).replace(/"/g,"&quot;"):""}" required />
          <datalist id="emp-list-options">
            ${n.map(s=>`<option value="${typeof s=="string"?s:s.label||s.name||""}">`).join("")}
          </datalist>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Asal <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-from" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang Asal --</option>
              ${i.map(s=>`
                <option value="${s.value}" ${String(a?.from_branch_id)===String(s.value)?"selected":""}>
                  ${s.label}
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
              ${i.map(s=>`
                <option value="${s.value}" ${String(a?.to_branch_id)===String(s.value)?"selected":""}>
                  ${s.label}
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
            placeholder="https://drive.google.com/file/d/..." value="${a?.document_link?String(a.document_link).replace(/"/g,"&quot;"):""}" />
        </div>
      </div>
    `,G({title:p?"\u270F\uFE0F Edit Data Mutasi Karyawan":"\u2795 Tambah Data Mutasi Karyawan",content:g,size:"md",confirmText:p?"Simpan Perubahan":"Simpan Mutasi",confirmClass:"btn-primary",onConfirm:async(s,r)=>{let l=g.querySelector("#mutasi-modal-tanggal").value,f=(g.querySelector("#mutasi-modal-emp").value||"").trim(),v=g.querySelector("#mutasi-modal-from").value,L=g.querySelector("#mutasi-modal-to").value,A=g.querySelector("#mutasi-modal-status").value,P=(g.querySelector("#mutasi-modal-link").value||"").trim();if(!l){N("Tanggal mutasi wajib diisi");return}if(!f){N("Nama karyawan wajib diisi");return}if(!v){N("Cabang asal wajib dipilih");return}if(!L){N("Cabang tujuan wajib dipilih");return}let w=i.find(C=>String(C.value)===String(v)),I=i.find(C=>String(C.value)===String(L)),z={tanggal:l,employee_name:f,from_branch_id:Number(v),from_branch_name:w?w.label:"Cabang Asal",to_branch_id:Number(L),to_branch_name:I?I.label:"Cabang Tujuan",status:A,document_link:P};try{if(p){let C=await K(`/api/mutasi/${a.id}`,{method:"PUT",body:JSON.stringify(z)});if(C.ok){R("Data mutasi berhasil diperbarui");let y=t.findIndex(_=>_.id===a.id);y!==-1&&(t[y]={...t[y],...z}),r(),k()}else N(C.data?.error||"Gagal memperbarui data mutasi")}else{let C=await K("/api/mutasi",{method:"POST",body:JSON.stringify(z)});if(C.ok){R("Data mutasi baru berhasil ditambahkan");let y=C.data?.id||Date.now();t.unshift({id:y,...z}),r(),k()}else N(C.data?.error||"Gagal menambahkan mutasi")}}catch(C){N("Terjadi kesalahan: "+C.message)}}})}function E(){return t.filter(a=>{if(u){let p=u.toLowerCase(),d=(a.employee_name||"").toLowerCase().includes(p),g=(a.from_branch_name||"").toLowerCase().includes(p),s=(a.to_branch_name||"").toLowerCase().includes(p);if(!d&&!g&&!s)return!1}return!(x!=="all"&&String(a.from_branch_id)!==String(x)||m!=="all"&&String(a.to_branch_id)!==String(m)||c!=="all"&&a.status!==c)})}function k(){let a=E(),p=t.length,d=t.filter(v=>v.status==="Proses").length,g=t.filter(v=>v.status==="Selesai").length,s=new Date,r=s.getFullYear(),l=s.getMonth(),f=t.filter(v=>{if(!v.tanggal)return!1;let L=new Date(v.tanggal);return!isNaN(L.getTime())&&L.getFullYear()===r&&L.getMonth()===l}).length;e.innerHTML=`
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
              ${B?`
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
                <div class="mutasi-stat-val">${p}</div>
                <div class="mutasi-stat-lbl">Total Mutasi</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u23F3
              </div>
              <div>
                <div class="mutasi-stat-val">${d}</div>
                <div class="mutasi-stat-lbl">Dalam Proses</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(34,197,94,0.2); color:#4ADE80;">
                \u2705
              </div>
              <div>
                <div class="mutasi-stat-val">${g}</div>
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
              <input type="text" id="mutasi-search-input" class="mutasi-search-input" placeholder="Cari nama karyawan, cabang asal, atau cabang tujuan..." value="${u}" />
            </div>

            <select id="mutasi-filter-from" class="mutasi-select">
              <option value="all">Semua Cabang Asal (${i.length})</option>
              ${i.map(v=>`<option value="${v.value}" ${String(x)===String(v.value)?"selected":""}>${v.label}</option>`).join("")}
            </select>

            <select id="mutasi-filter-to" class="mutasi-select">
              <option value="all">Semua Cabang Tujuan (${i.length})</option>
              ${i.map(v=>`<option value="${v.value}" ${String(m)===String(v.value)?"selected":""}>${v.label}</option>`).join("")}
            </select>
          </div>

          <!-- Status Filter Pills (Swipeable on Mobile) -->
          <div class="mutasi-filter-pills">
            <div class="mutasi-pill ${c==="all"?"active":""}" data-status="all">
              <span>Semua Mutasi</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${p}</span>
            </div>
            <div class="mutasi-pill ${c==="Proses"?"active":""}" data-status="Proses">
              <span style="color:#D97706;">\u23F3</span>
              <span>Proses</span>
              <span class="badge" style="background:rgba(217,119,6,0.15); color:#B45309; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${d}</span>
            </div>
            <div class="mutasi-pill ${c==="Selesai"?"active":""}" data-status="Selesai">
              <span style="color:#16A34A;">\u2705</span>
              <span>Selesai</span>
              <span class="badge" style="background:rgba(22,163,74,0.15); color:#15803D; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${g}</span>
            </div>
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="mutasi-list-container">
          ${a.length===0?`
            <div class="mutasi-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">\u{1F50D}</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data Mutasi Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="mutasi-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          `:`
            <div class="mutasi-cards-grid">
              ${a.map(v=>{let L=v.status==="Selesai",A=(v.employee_name||"Karyawan").split(" ").map(P=>P[0]).filter(Boolean).slice(0,2).join("").toUpperCase();return`
                  <div class="mutasi-card" data-id="${v.id}">
                    <div>
                      <!-- Card Header: Status & Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="mutasi-status-badge ${L?"selesai":"proses"}">
                          <span>${L?"\u2705":"\u23F3"}</span>
                          <span>${v.status||"Proses"}</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B; display:inline-flex; align-items:center; gap:4px;">
                          <span>\u{1F5D3}\uFE0F</span>
                          <span>${On(v.tanggal)}</span>
                        </span>
                      </div>

                      <!-- Employee Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="mutasi-avatar-box">
                          ${A}
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

                      ${B?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${v.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                          <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${v.id}">
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
    `,h()}function h(){let a=e.querySelector("#mutasi-btn-add");a&&(a.onclick=()=>$());let p=e.querySelector("#mutasi-btn-excel");p&&(p.onclick=()=>S());let d=e.querySelector("#mutasi-btn-refresh");d&&(d.onclick=()=>D());let g=e.querySelector("#mutasi-btn-reset-filter");g&&(g.onclick=()=>{u="",x="all",m="all",c="all",k()});let s=e.querySelector("#mutasi-search-input");s&&(s.oninput=f=>{u=f.target.value,k();let v=e.querySelector("#mutasi-search-input");v&&(v.focus(),v.setSelectionRange(v.value.length,v.value.length))});let r=e.querySelector("#mutasi-filter-from");r&&(r.onchange=f=>{x=f.target.value,k()});let l=e.querySelector("#mutasi-filter-to");l&&(l.onchange=f=>{m=f.target.value,k()}),e.querySelectorAll(".mutasi-pill").forEach(f=>{f.onclick=()=>{c=f.getAttribute("data-status")||"all",k()}}),e.querySelectorAll(".mutasi-edit-btn").forEach(f=>{f.onclick=v=>{v.stopPropagation();let L=Number(f.getAttribute("data-id")),A=t.find(P=>P.id===L);A&&$(A)}}),e.querySelectorAll(".mutasi-delete-btn").forEach(f=>{f.onclick=v=>{v.stopPropagation();let L=Number(f.getAttribute("data-id"));F(L)}})}await D()}le();var Lt=[{id:1,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:1,employee_name:"ADE SURAHMAN",start_time:"23:00",end_time:"08:00",break_hours:1,total_hours:8,reason:"General Cleaning"},{id:2,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:2,employee_name:"TAUFIK HIDAYAT",start_time:"00:00",end_time:"08:00",break_hours:1,total_hours:7,reason:"General Cleaning"},{id:3,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:3,employee_name:"YUDA WILDAN WIGUNA",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"},{id:4,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:4,employee_name:"REZA RIZKI",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"},{id:5,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:5,employee_name:"ADITYA PRATAMA",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"}],qa=["General Cleaning","Deep Cleaning","Cover manpower","Longshift","Pekerjaan urgent"];function Ra(e,t,i){if(!e||!t)return 0;try{let[n,c]=e.split(":").map(Number),[x,m]=t.split(":").map(Number),u=x*60+m-(n*60+c);u<0&&(u+=24*60);let o=u/60-(parseFloat(i)||0);return o<0&&(o=0),Number.isInteger(o)?o:parseFloat(o.toFixed(2))}catch{return 0}}function It(e=""){let t=(e||"").toLowerCase();return t.includes("general")?{bg:"#EFF6FF",color:"#1D4ED8",border:"#BFDBFE",icon:"\u{1F9F9}"}:t.includes("deep")?{bg:"#F5F3FF",color:"#6D28D9",border:"#DDD6FE",icon:"\u2728"}:t.includes("cover")?{bg:"#ECFDF5",color:"#047857",border:"#A7F3D0",icon:"\u{1F465}"}:t.includes("longshift")?{bg:"#FFFBEB",color:"#B45309",border:"#FDE68A",icon:"\u23F3"}:t.includes("urgent")?{bg:"#FFF1F2",color:"#BE123C",border:"#FECDD3",icon:"\u{1F6A8}"}:{bg:"#F8FAFC",color:"#475569",border:"#CBD5E1",icon:"\u23F1\uFE0F"}}function zt(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Oa(e){let t=[],i=[],n=[],c="all",x="all",m="",u=!1,o=localStorage.getItem("fcms_overtime_view_mode")||"card",b=ee(),B=b?.role==="viewer",D=b?.role==="input_lembur",S=!B,F=!B,$=!B&&!D;try{[i,n]=await Promise.all([te(),_e()])}catch(s){console.error("Error fetching cache:",s)}async function E(){u=!0,d();try{let s=await K("/api/overtime?limit=10000");if(s.ok){let r=s.data?.data||s.data||[];t=Array.isArray(r)&&r.length>0?r:Lt}else t=Lt}catch(s){console.warn("Gagal load lembur, menggunakan data lokal:",s),t=Lt}finally{u=!1,d()}}async function k(){try{let{downloadExcel:s}=await Promise.resolve().then(()=>(re(),he)),r=t.map((l,f)=>({No:f+1,Tanggal:l.date||"","Cabang Lembur":l.branch_name||"","Nama Karyawan":l.employee_name||"","Jam Mulai":l.start_time||"","Jam Selesai":l.end_time||"","Istirahat (Jam)":l.break_hours!=null?l.break_hours:0,"Total Jam":l.total_hours!=null?l.total_hours:0,"Alasan / Keterangan":l.reason||""}));s(r,`Data_Lembur_${new Date().toISOString().slice(0,10)}`),R("Data lembur berhasil diekspor ke Excel!")}catch(s){N("Gagal ekspor Excel: "+s.message)}}async function h(s){if(!$)return;let r=t.find(l=>l.id===s);X({title:"Hapus Catatan Lembur",message:`Apakah Anda yakin ingin menghapus catatan lembur untuk "${r?.employee_name||"karyawan ini"}" pada tanggal ${zt(r?.date)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let l=await K(`/api/overtime/${s}`,{method:"DELETE"});l.ok?(R("Catatan lembur berhasil dihapus"),t=t.filter(f=>f.id!==s),d()):N(l.data?.error||"Gagal menghapus lembur")}catch(l){N("Gagal menghapus catatan lembur: "+l.message)}}})}function a(s=null){if(!F)return;let r=!!s,l=s?.date?s.date.slice(0,10):new Date().toISOString().slice(0,10),f=s?.start_time||"22:00",v=s?.end_time||"08:00",L=s?.break_hours!=null?s.break_hours:1,A=s?.total_hours!=null?s.total_hours:Ra(f,v,L),P=document.createElement("div");P.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="ot-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${l}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penugasan <span style="color:#DC2626;">*</span>
            </label>
            <select id="ot-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${i.map(_=>`
                <option value="${_.value}" ${String(s?.branch_id)===String(_.value)?"selected":""}>
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
            ${n.map(_=>{let T=typeof _=="string"?_:_.label||_.full_name||_.name||"",j=typeof _=="object"&&_!==null?_.value!=null?_.value:_.id!=null?_.id:T:_,O=String(s?.employee_id)===String(j)||s?.employee_name===T;return`<option value="${j}" ${O?"selected":""}>${T}</option>`}).join("")}
          </select>
        </div>

        <!-- Shift & Waktu Lembur -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Mulai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${f}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Selesai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${v}" required />
          </div>
        </div>

        <!-- Durasi & Istirahat -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Istirahat (Jam) <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.5" min="0" max="10" id="ot-modal-break" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${L}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Total Jam Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.1" min="0" id="ot-modal-total" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; background:#F8FAFC; font-weight:700; color:#2563EB;" 
              value="${A}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Alasan / Keterangan Penugasan Lembur <span style="color:#DC2626;">*</span>
          </label>
          <select id="ot-modal-reason" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
            ${qa.map(_=>`
              <option value="${_}" ${s?.reason===_?"selected":""}>${_}</option>
            `).join("")}
          </select>
        </div>
      </div>
    `;let w=P.querySelector("#ot-modal-start"),I=P.querySelector("#ot-modal-end"),z=P.querySelector("#ot-modal-break"),C=P.querySelector("#ot-modal-total");function y(){let _=w.value,T=I.value,j=z.value;C.value=Ra(_,T,j)}w.addEventListener("input",y),I.addEventListener("input",y),z.addEventListener("input",y),G({title:r?"\u270F\uFE0F Edit Data Lembur Karyawan":"\u2795 Catat Penugasan Lembur Baru",content:P,size:"md",confirmText:r?"Simpan Perubahan":"Catat Lembur",confirmClass:"btn-primary",onConfirm:async(_,T)=>{let j=P.querySelector("#ot-modal-date").value,O=P.querySelector("#ot-modal-branch").value,M=P.querySelector("#ot-modal-emp").value,q=P.querySelector("#ot-modal-start").value,H=P.querySelector("#ot-modal-end").value,Y=parseFloat(P.querySelector("#ot-modal-break").value)||0,W=parseFloat(P.querySelector("#ot-modal-total").value)||0,Z=P.querySelector("#ot-modal-reason").value;if(!j){N("Tanggal lembur wajib diisi");return}if(!O){N("Cabang penugasan wajib dipilih");return}if(!M){N("Nama karyawan wajib dipilih");return}if(!q||!H){N("Jam mulai dan jam selesai wajib diisi");return}let U=i.find(ie=>String(ie.value)===String(O)),V=n.find(ie=>String(ie.value)===String(M)||String(ie.id)===String(M)),ae=V?V.label||V.full_name||V.name||"Karyawan":typeof M=="string"&&isNaN(Number(M))?M:"Karyawan",J={date:j,branch_id:Number(O),branch_name:U?U.label:"Cabang Lembur",employee_id:Number(M),employee_name:V?V.label:"Karyawan",start_time:q,end_time:H,break_hours:Y,total_hours:W,reason:Z};try{if(r){let ie=await K(`/api/overtime/${s.id}`,{method:"PUT",body:JSON.stringify(J)});if(ie.ok){R("Data lembur berhasil diperbarui");let me=t.findIndex(se=>se.id===s.id);me!==-1&&(t[me]={...t[me],...J}),T(),d()}else N(ie.data?.error||"Gagal memperbarui data lembur")}else{let ie=await K("/api/overtime",{method:"POST",body:JSON.stringify(J)});if(ie.ok){R("Data lembur berhasil ditambahkan");let me=ie.data?.id||Date.now();t.unshift({id:me,...J}),T(),d()}else N(ie.data?.error||"Gagal menambahkan lembur")}}catch(ie){N("Terjadi kesalahan: "+ie.message)}}})}function p(){return t.filter(s=>{if(m){let r=m.toLowerCase(),l=(s.employee_name||"").toLowerCase().includes(r),f=(s.branch_name||"").toLowerCase().includes(r),v=(s.reason||"").toLowerCase().includes(r);if(!l&&!f&&!v)return!1}return!(x!=="all"&&String(s.branch_id)!==String(x)||c!=="all"&&s.reason!==c)})}function d(){let s=p(),r=t.length,l=t.reduce((w,I)=>w+(parseFloat(I.total_hours)||0),0),f=r>0?(l/r).toFixed(1):0,v=new Date,L=v.getFullYear(),A=v.getMonth(),P=t.filter(w=>{if(!w.date)return!1;let I=new Date(w.date);return!isNaN(I.getTime())&&I.getFullYear()===L&&I.getMonth()===A}).length;e.innerHTML=`
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
              ${S?`
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
                <div class="lembur-stat-val">${l} Jam</div>
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
                <div class="lembur-stat-val">${f} Jam</div>
                <div class="lembur-stat-lbl">Rata-rata Durasi</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="lembur-stat-val">${P}</div>
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
              <input type="text" id="lembur-search-input" class="lembur-search-input" placeholder="Cari nama karyawan, cabang, atau alasan lembur..." value="${m}" />
            </div>

            <select id="lembur-filter-branch" class="lembur-select">
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(w=>`<option value="${w.value}" ${String(x)===String(w.value)?"selected":""}>${w.label}</option>`).join("")}
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
            <div class="lembur-pill ${c==="all"?"active":""}" data-reason="all">
              <span>Semua Lembur</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${r}</span>
            </div>
            ${qa.map(w=>{let I=t.filter(C=>C.reason===w).length,z=It(w);return`
                <div class="lembur-pill ${c===w?"active":""}" data-reason="${w}">
                  <span>${z.icon}</span>
                  <span>${w}</span>
                  ${I>0?`<span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${I}</span>`:""}
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
                      ${F?'<th style="width:90px; text-align:center;">Aksi</th>':""}
                    </tr>
                  </thead>
                  <tbody>
                    ${s.map((w,I)=>{let z=(w.employee_name||"Karyawan").split(" ").map(y=>y[0]).filter(Boolean).slice(0,2).join("").toUpperCase(),C=It(w.reason);return`
                        <tr data-id="${w.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${I+1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="lembur-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${z}
                              </div>
                              <div style="font-weight:700; color:#0F172A; white-space:nowrap;">
                                ${w.employee_name}
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#334155; font-size:0.85rem; display:inline-flex; align-items:center; gap:4px;">
                              <span>\u{1F4CD}</span>
                              <span>${w.branch_name||"Cabang"}</span>
                            </div>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#1E293B;">
                            \u{1F4C5} ${zt(w.date)}
                          </td>
                          <td style="white-space:nowrap;">
                            <span style="font-weight:700; color:#2563EB;">\u23F0 ${w.start_time||"00:00"} - ${w.end_time||"00:00"}</span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="lembur-hours-badge" style="padding:3px 8px; font-size:0.8rem;">
                              <span>\u23F1\uFE0F</span>
                              <span>${w.total_hours!=null?w.total_hours:0} Jam</span>
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap; color:#64748B; font-weight:600; font-size:0.82rem;">
                            \u2615 ${w.break_hours!=null?w.break_hours:0} Jam
                          </td>
                          <td style="white-space:nowrap;">
                            <span class="lembur-reason-badge" style="background:${C.bg}; color:${C.color}; border:1px solid ${C.border};">
                              <span>${C.icon}</span>
                              <span>${w.reason||"Lembur Operasional"}</span>
                            </span>
                          </td>
                          ${F?`
                            <td style="text-align:center; white-space:nowrap;">
                              <div style="display:inline-flex; align-items:center; gap:6px;">
                                <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${w.id}">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                                ${$?`
                                  <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${w.id}">
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
              ${s.map(w=>{let I=(w.employee_name||"Karyawan").split(" ").map(C=>C[0]).filter(Boolean).slice(0,2).join("").toUpperCase(),z=It(w.reason);return`
                  <div class="lembur-card" data-id="${w.id}">
                    <div>
                      <!-- Card Header: Jam & Istirahat -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="lembur-hours-badge">
                          <span>\u23F1\uFE0F</span>
                          <span>${w.total_hours!=null?w.total_hours:0} Jam Lembur</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                          \u2615 Istirahat: ${w.break_hours!=null?w.break_hours:0} Jam
                        </span>
                      </div>

                      <!-- Employee & Branch -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="lembur-avatar-box">
                          ${I}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${w.employee_name}
                          </h4>
                          <div style="font-size:0.8rem; color:#64748B; display:flex; align-items:center; gap:4px;">
                            <span>\u{1F4CD}</span>
                            <span style="font-weight:600; color:#334155;">${w.branch_name||"Cabang"}</span>
                          </div>
                        </div>
                      </div>

                      <!-- Shift Times & Date Box -->
                      <div class="lembur-time-box">
                        <div>
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Tanggal</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#1E293B;">
                            \u{1F5D3}\uFE0F ${zt(w.date)}
                          </div>
                        </div>

                        <div style="text-align:right;">
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Waktu Shift</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#2563EB;">
                            \u23F0 ${w.start_time||"00:00"} - ${w.end_time||"00:00"}
                          </div>
                        </div>
                      </div>

                      <!-- Reason Badge -->
                      <div>
                        <span class="lembur-reason-badge" style="background:${z.bg}; color:${z.color}; border:1px solid ${z.border};">
                          <span>${z.icon}</span>
                          <span>${w.reason||"Lembur Operasional"}</span>
                        </span>
                      </div>
                    </div>

                    <!-- Card Footer: Aksi -->
                    ${F?`
                      <div style="display:flex; justify-content:flex-end; align-items:center; gap:6px; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:12px;">
                        <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${w.id}">
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        ${$?`
                          <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${w.id}">
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
    `,g()}function g(){let s=e.querySelector("#lembur-btn-add");s&&(s.onclick=()=>a());let r=e.querySelector("#lembur-btn-excel");r&&(r.onclick=()=>k());let l=e.querySelector("#lembur-btn-refresh");l&&(l.onclick=()=>E());let f=e.querySelector("#lembur-btn-reset-filter");f&&(f.onclick=()=>{m="",x="all",c="all",d()});let v=e.querySelector("#lembur-search-input");v&&(v.oninput=A=>{m=A.target.value,d();let P=e.querySelector("#lembur-search-input");P&&(P.focus(),P.setSelectionRange(P.value.length,P.value.length))});let L=e.querySelector("#lembur-filter-branch");L&&(L.onchange=A=>{x=A.target.value,d()}),e.querySelectorAll(".lembur-pill").forEach(A=>{A.onclick=()=>{c=A.getAttribute("data-reason")||"all",d()}}),e.querySelectorAll(".lembur-view-btn").forEach(A=>{A.onclick=()=>{let P=A.getAttribute("data-view");P&&P!==o&&(o=P,localStorage.setItem("fcms_overtime_view_mode",P),d())}}),e.querySelectorAll(".lembur-edit-btn").forEach(A=>{A.onclick=P=>{P.stopPropagation();let w=Number(A.getAttribute("data-id")),I=t.find(z=>z.id===w);I&&a(I)}}),e.querySelectorAll(".lembur-delete-btn").forEach(A=>{A.onclick=P=>{P.stopPropagation();let w=Number(A.getAttribute("data-id"));h(w)}})}await E()}async function Ha(e){let t=ee();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}de({container:e,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:i=>new Date(i).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(i,n)=>`<strong>${i||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${n.user_role||""}</small>`},{key:"action",label:"Aksi",render:i=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[i]||"badge-neutral"}">${i}</span>`},{key:"module",label:"Modul",render:i=>`<span style="text-transform:capitalize">${(i||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(i,n)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${i}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async i=>{try{let x=((await(await fetch(`/api/audit-logs?search=${i}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(b=>String(b.id)===String(i));if(!x)return alert("Data tidak ditemukan");let m=b=>{if(!b)return"Tidak ada data";try{return JSON.stringify(JSON.parse(b),null,2)}catch{return b}},u=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${m(x.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${m(x.new_data)}</pre>
           </div>
         </div>
       `,{createModal:o}=await Promise.resolve().then(()=>(le(),Rt));o({title:`Detail Audit Log #${i}`,content:u,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}le();re();var Mt={"2026-09":"September 2026","2026-10":"Oktober 2026","2026-11":"November 2026","2026-12":"Desember 2026"},Ue={scheduled:{label:"Terjadwal",badgeClass:"badge-warning",icon:"\u{1F7E1}",bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},in_progress:{label:"Sedang Proses",badgeClass:"badge-info",icon:"\u{1F535}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},completed:{label:"Selesai",badgeClass:"badge-success",icon:"\u{1F7E2}",bg:"#DCFCE7",color:"#166534",border:"#86EFAC"}};async function Ua(e){e.innerHTML=`
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
  `;let t="board",i=null,n=[],c=ee(),x=c&&["superadmin","admin","manager","spv"].includes(c.role);K("/api/branches?all=1").then(h=>{h.ok&&h.data?.data&&(n=h.data.data)}).catch(()=>{});async function m(){try{let h=await K("/api/os-migrations/summary");if(!h.ok)throw new Error("Gagal memuat data migrasi");i=h.data?.data||h.data,u(i),b()}catch(h){document.getElementById("migration-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${h.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function u(h){let a=document.getElementById("migration-summary-cards");if(!a||!h)return;let p=h.total_branches||0,d=h.vendors?.AEON?.total||0,g=h.vendors?.["Triple S"]?.total||0,s=h.by_status?.completed||0,r=h.by_status?.in_progress||0,l=h.by_status?.scheduled||0,f=p>0?Math.round(s/p*100):0;a.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${p} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.8rem; font-weight:800; color:#1E40AF; line-height:1.2;">${d} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
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
          <span style="font-size:0.85rem; font-weight:700; color:#15803D;">${f}%</span>
        </div>
        <div style="width:100%; height:8px; background:#E2E8F0; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${f}%; height:100%; background:#16A34A; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.75rem; color:var(--text-2); font-weight:500;">
          <span>\u{1F7E2} Selesai: <strong>${s}</strong></span>
          <span>\u{1F535} Proses: <strong>${r}</strong></span>
          <span>\u{1F7E1} Terjadwal: <strong>${l}</strong></span>
        </div>
      </div>
    `}function o(){if(!i||!i.matrix)return[];let h=(document.getElementById("migration-search")?.value||"").toLowerCase().trim(),a=document.getElementById("filter-vendor")?.value||"",p=document.getElementById("filter-month")?.value||"",d=document.getElementById("filter-status")?.value||"";return i.matrix.filter(g=>!(a&&g.vendor!==a||p&&g.target_month!==p||d&&g.status!==d||h&&!`${g.branch_code} ${g.branch_name} ${g.branch_full_name} ${g.vendor} ${g.notes||""} ${g.pic_vendor||""}`.toLowerCase().includes(h)))}function b(){let h=o(),a=document.getElementById("filter-stats-text");a&&(a.innerHTML=`Menampilkan <strong>${h.length}</strong> dari <strong>${i?.total_branches||0}</strong> cabang`);let p=document.getElementById("migration-view-container");t==="board"?B(p,h):D(p,h)}function B(h,a){if(!a.length){h.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;return}let p=["AEON","Triple S"],d=["2026-09","2026-10","2026-11","2026-12"],g="";p.forEach(s=>{let r=a.filter(P=>P.vendor===s);if(document.getElementById("filter-vendor")?.value&&document.getElementById("filter-vendor").value!==s||r.length===0)return;let l=r.length,f=s==="AEON"?"#1E40AF":"#0F766E",v=s==="AEON"?"#EFF6FF":"#F0FDFA",L=s==="AEON"?"#BFDBFE":"#99F6E4",A=d.filter(P=>r.some(w=>w.target_month===P));g+=`
        <div class="vendor-board-section" style="margin-bottom:32px; background:#fff; border:1px solid var(--border); border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <!-- Vendor Section Header -->
          <div style="padding:14px 20px; background:${v}; border-bottom:1px solid ${L}; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.1rem; font-weight:800; color:${f};">${s==="AEON"?"\u{1F3E2} AEON":"\u{1F6E1}\uFE0F Triple S"}</span>
              <span class="badge" style="background:#fff; color:${f}; border:1px solid ${L}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px;">
                ${l} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
              Format Timeline Berdasarkan Dokumen Migrasi
            </div>
          </div>

          <!-- Months Grid (Hanya menampilkan bulan yang memiliki jadwal) -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px; padding:18px;">
            ${A.map(P=>{let w=r.filter(C=>C.target_month===P);return`
                <div class="month-column" style="background:#F8FAFC; border:1px solid var(--border); border-radius:10px; display:flex; flex-direction:column; min-height:220px;">
                  <!-- Month Header -->
                  <div style="padding:12px 14px; background:#F1F5F9; border-bottom:1px solid var(--border); border-radius:9px 9px 0 0; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">
                      \u{1F4C5} ${(Mt[P]||P).split(" ")[0]}
                    </div>
                    <span class="badge" style="background:#E2E8F0; color:var(--text-2); font-weight:700; font-size:0.75rem; padding:2px 7px; border-radius:4px;">
                      ${w.length}
                    </span>
                  </div>

                  <!-- Branch Cards Container -->
                  <div style="padding:10px; display:flex; flex-direction:column; gap:8px; flex:1; max-height:480px; overflow-y:auto;">
                    ${w.map(C=>{let y=Ue[C.status]||Ue.scheduled;return`
                        <div class="branch-card" style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 12px; box-shadow:0 1px 2px rgba(0,0,0,0.03); transition:transform 0.15s, box-shadow 0.15s;">
                          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px; margin-bottom:6px;">
                            <div style="font-weight:700; font-size:0.9rem; color:var(--text-1); line-height:1.3;">
                              <span style="color:#2563EB; font-family:monospace; margin-right:4px;">${C.branch_code}.</span>${C.branch_name}
                            </div>
                            ${x?`
                              <div style="display:flex; gap:2px;">
                                <button class="btn-icon-mini edit-card-btn" data-id="${C.id}" title="Edit" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--text-3);">\u270F\uFE0F</button>
                                <button class="btn-icon-mini delete-card-btn" data-id="${C.id}" data-name="${C.branch_name}" title="Hapus" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--danger);">\u{1F5D1}\uFE0F</button>
                              </div>
                            `:""}
                          </div>

                          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px; margin-top:8px;">
                            <!-- Status Selector Dropdown -->
                            <div style="position:relative;">
                              <button class="status-toggle-btn" data-id="${C.id}" data-status="${C.status}" style="background:${y.bg}; color:${y.color}; border:1px solid ${y.border}; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; display:inline-flex; align-items:center; gap:4px; cursor:pointer;" title="Klik untuk ubah status">
                                ${y.icon} ${y.label} \u25BE
                              </button>
                            </div>

                            ${C.employee_count>0?`
                              <span style="font-size:0.75rem; color:var(--text-3); font-weight:600; display:inline-flex; align-items:center; gap:3px;" title="${C.employee_count} personil Facility Care aktif di cabang ini">
                                \u{1F464} ${C.employee_count} FC
                              </span>
                            `:""}
                          </div>

                          ${C.actual_date?`
                            <div style="font-size:0.7rem; color:var(--text-3); margin-top:6px;">
                              Realisasi: <strong>${C.actual_date}</strong>
                            </div>
                          `:""}
                        </div>
                      `}).join("")}
                  </div>
                </div>
              `}).join("")}
          </div>
        </div>
      `}),h.innerHTML=g,S(h)}function D(h,a){if(!a.length){h.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
        </div>
      `;return}h.innerHTML=`
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
                ${x?'<th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${a.map((p,d)=>{let g=Ue[p.status]||Ue.scheduled;return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3);">${d+1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">${p.branch_full_name}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${p.vendor==="AEON"?"#DBEAFE":"#CCFBF1"}; color:${p.vendor==="AEON"?"#1E40AF":"#0F766E"}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${p.vendor}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:var(--text-2);">
                      ${Mt[p.target_month]||p.target_month}
                    </td>
                    <td style="padding:12px 16px;">
                      <button class="status-toggle-btn" data-id="${p.id}" data-status="${p.status}" style="background:${g.bg}; color:${g.color}; border:1px solid ${g.border}; padding:4px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; cursor:pointer;" title="Klik untuk ubah status">
                        ${g.icon} ${g.label} \u25BE
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
                    ${x?`
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
    `,S(h)}function S(h){h.querySelectorAll(".status-toggle-btn").forEach(a=>{a.addEventListener("click",p=>{p.stopPropagation();let d=a.dataset.id,g=a.dataset.status;F(a,d,g)})}),h.querySelectorAll(".edit-card-btn").forEach(a=>{a.addEventListener("click",()=>{let p=a.dataset.id,d=i?.matrix?.find(g=>String(g.id)===String(p));d&&$(d)})}),h.querySelectorAll(".delete-card-btn").forEach(a=>{a.addEventListener("click",()=>{let p=a.dataset.id,d=a.dataset.name||"Cabang";X(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${d}</strong>?`,async()=>{try{if(!(await K(`/api/os-migrations/${p}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus");R(`Jadwal migrasi ${d} berhasil dihapus`),m()}catch(g){N(g.message)}})})})}function F(h,a,p){document.querySelectorAll(".status-popup-menu").forEach(l=>l.remove());let d=h.getBoundingClientRect(),g=document.createElement("div");g.className="status-popup-menu",g.style.cssText=`
      position: fixed;
      top: ${d.bottom+window.scrollY+4}px;
      left: ${d.left+window.scrollX}px;
      background: #fff;
      border: 1px solid var(--border);
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 9999;
      padding: 6px 0;
      min-width: 160px;
    `;let s=[{key:"scheduled",label:"\u{1F7E1} Terjadwal"},{key:"in_progress",label:"\u{1F535} Sedang Proses"},{key:"completed",label:"\u{1F7E2} Selesai"}];g.innerHTML=s.map(l=>`
      <button class="status-option-item" data-key="${l.key}" style="
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
        color: ${l.key===p?"#2563EB":"var(--text-1)"};
        font-weight: ${l.key===p?"700":"500"};
      ">
        ${l.label} ${l.key===p?"\u2713":""}
      </button>
    `).join("");let r=l=>{!g.contains(l.target)&&l.target!==h&&(g.remove(),document.removeEventListener("click",r))};setTimeout(()=>document.addEventListener("click",r),50),g.querySelectorAll(".status-option-item").forEach(l=>{l.addEventListener("click",async()=>{let f=l.dataset.key;if(g.remove(),document.removeEventListener("click",r),f!==p)try{if(!(await K(`/api/os-migrations/${a}/status`,{method:"PATCH",body:JSON.stringify({status:f})})).ok)throw new Error("Gagal memperbarui status");R("Status berhasil diubah"),m()}catch(v){N(v.message)}})}),document.body.appendChild(g)}function $(h=null){let a=!!h,p=a?"Edit Jadwal Migrasi OS":"Tambah Jadwal Migrasi OS",g=`
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${n.map(s=>`
      <option value="${s.id}" ${h&&h.branch_id===s.id?"selected":""}>
        ${s.full_name}
      </option>
    `).join("")}
          </select>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Vendor Outsourcing *</label>
            <select name="vendor" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="AEON" ${h?.vendor==="AEON"?"selected":""}>AEON</option>
              <option value="Triple S" ${h?.vendor==="Triple S"?"selected":""}>Triple S</option>
              <option value="Lainnya" ${h&&!["AEON","Triple S"].includes(h.vendor)?"selected":""}>Lainnya</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Target Bulan *</label>
            <select name="target_month" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="2026-09" ${h?.target_month==="2026-09"?"selected":""}>September 2026</option>
              <option value="2026-10" ${h?.target_month==="2026-10"?"selected":""}>Oktober 2026</option>
              <option value="2026-11" ${h?.target_month==="2026-11"?"selected":""}>November 2026</option>
              <option value="2026-12" ${h?.target_month==="2026-12"?"selected":""}>Desember 2026</option>
            </select>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Status Pelaksanaan</label>
            <select name="status" class="form-control" style="width:100%; height:38px; border-radius:8px;">
              <option value="scheduled" ${h?.status==="scheduled"?"selected":""}>\u{1F7E1} Terjadwal</option>
              <option value="in_progress" ${h?.status==="in_progress"?"selected":""}>\u{1F535} Sedang Proses</option>
              <option value="completed" ${h?.status==="completed"?"selected":""}>\u{1F7E2} Selesai</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Tanggal Realisasi (Opsional)</label>
            <input type="date" name="actual_date" class="form-control" value="${h?.actual_date||""}" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">PIC / Koordinator Vendor (Opsional)</label>
          <input type="text" name="pic_vendor" class="form-control" placeholder="Nama PIC vendor / no HP" value="${h?.pic_vendor||""}" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Catatan / Keterangan</label>
          <textarea name="notes" class="form-control" rows="2" placeholder="Catatan serah terima / info tambahan..." style="width:100%; border-radius:8px;">${h?.notes||""}</textarea>
        </div>
      </form>
    `;G({title:p,content:g,confirmText:a?"Simpan Perubahan":"Tambah Cabang",onConfirm:async(s,r)=>{let l=s.querySelector("#migration-form");if(!l.checkValidity()){l.reportValidity();return}let f=new FormData(l),v={branch_id:parseInt(f.get("branch_id"),10),vendor:f.get("vendor"),target_month:f.get("target_month"),status:f.get("status"),actual_date:f.get("actual_date")||null,pic_vendor:f.get("pic_vendor")||null,notes:f.get("notes")||null};try{let L=a?`/api/os-migrations/${h.id}`:"/api/os-migrations",P=await K(L,{method:a?"PUT":"POST",body:JSON.stringify(v)});if(!P.ok){let w=await P.json().catch(()=>({}));throw new Error(w.error||"Gagal menyimpan data")}R(a?"Jadwal migrasi berhasil diupdate":"Jadwal migrasi berhasil ditambahkan"),r(),m()}catch(L){N(L.message)}}})}let E=document.getElementById("view-board-btn"),k=document.getElementById("view-table-btn");E.addEventListener("click",()=>{t="board",E.style.background="#fff",E.classList.remove("btn-ghost"),k.style.background="transparent",k.classList.add("btn-ghost"),b()}),k.addEventListener("click",()=>{t="table",k.style.background="#fff",k.classList.remove("btn-ghost"),E.style.background="transparent",E.classList.add("btn-ghost"),b()}),["migration-search","filter-vendor","filter-month","filter-status"].forEach(h=>{let a=document.getElementById(h);a&&(a.addEventListener("input",b),a.addEventListener("change",b))}),document.getElementById("btn-add-migration").addEventListener("click",()=>{$()}),document.getElementById("btn-export-migrations").addEventListener("click",()=>{let h=o();if(!h.length){N("Tidak ada data untuk diexport");return}let a=h.map((p,d)=>({No:d+1,"Kode Cabang":p.branch_code,"Nama Cabang":p.branch_name,"Cabang Lengkap":p.branch_full_name,"Vendor OS":p.vendor,"Target Bulan":Mt[p.target_month]||p.target_month,Status:Ue[p.status]?.label||p.status,"Tanggal Realisasi":p.actual_date||"","Jml Personil FC":p.employee_count||0,"PIC Vendor":p.pic_vendor||"",Catatan:p.notes||""}));Q(a,`Timeline_Migrasi_OS_${new Date().toISOString().slice(0,10)}`),R("File Excel berhasil diunduh")}),m()}le();re();var Ga={Persiapan:{bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE"},"Area Umum":{bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD"},"Ruang Tindakan":{bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},"Ruang Penunjang":{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},Toilet:{bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4"},"Area Luar":{bg:"#FEF3C7",color:"#92400E",border:"#FDE68A"},"Kontrol & Check":{bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA"},Istirahat:{bg:"#E2E8F0",color:"#475569",border:"#CBD5E1"},"Layanan Pasien":{bg:"#DCFCE7",color:"#15803D",border:"#86EFAC"},"Sampah Umum":{bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},"Sampah Medis":{bg:"#FEE2E2",color:"#991B1B",border:"#FCA5A5"},Penutup:{bg:"#E0E7FF",color:"#3730A3",border:"#C7D2FE"}};async function Ja(e){e.innerHTML=`
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
  `;let t="pagi",i="timeline",n=[],c=ee(),x=c&&["superadmin","admin","manager","spv"].includes(c.role);async function m(){try{let h=await K("/api/dac?all=1");if(!h.ok)throw new Error("Gagal memuat standar DAC");n=h.data?.data||h.data||[];let a=n.filter(s=>s.shift==="pagi").length,p=n.filter(s=>s.shift==="siang").length;document.getElementById("badge-count-pagi").textContent=a,document.getElementById("badge-count-siang").textContent=p,document.getElementById("badge-count-all").textContent=n.length;let d=document.getElementById("dac-category-filter"),g=Array.from(new Set(n.map(s=>s.category))).sort();d.innerHTML='<option value="">Semua Kategori</option>'+g.map(s=>`<option value="${s}">${s}</option>`).join(""),o(),b()}catch(h){document.getElementById("dac-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${h.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function u(){let h=(document.getElementById("dac-search")?.value||"").toLowerCase().trim(),a=document.getElementById("dac-category-filter")?.value||"",p=document.getElementById("dac-period-filter")?.value||"";return n.filter(d=>{if(t!=="all"&&d.shift!==t||a&&d.category!==a)return!1;if(p==="morning"){let g=parseInt(d.start_time.split(":")[0],10);if(d.shift==="pagi"&&g>=12||d.shift==="siang"&&g>=17)return!1}else if(p==="afternoon"){let g=parseInt(d.start_time.split(":")[0],10);if(d.shift==="pagi"&&g<12||d.shift==="siang"&&g<17)return!1}return!(h&&!`${d.start_time} ${d.end_time} ${d.category} ${d.location} ${d.activity}`.toLowerCase().includes(h))})}function o(){let h=document.getElementById("dac-summary-cards");if(!h)return;let a=n.filter(r=>r.shift==="pagi"),p=n.filter(r=>r.shift==="siang"),d=t==="pagi"?a:t==="siang"?p:n,g=t==="pagi"?"08:00 \u2013 17:00":t==="siang"?"13:00 \u2013 22:00":"08:00 \u2013 22:00",s=t==="pagi"?"Shift Pagi":t==="siang"?"Shift Siang":"Semua Shift";h.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Aktivitas (${s})</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${d.length} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Aktivitas</span></div>
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
    `}function b(){let h=u(),a=document.getElementById("dac-stats-text");a&&(a.innerHTML=`Menampilkan <strong>${h.length}</strong> aktivitas`);let p=document.getElementById("dac-view-container");i==="timeline"?B(p,h):D(p,h)}function B(h,a){if(!a.length){h.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;return}h.innerHTML=`
      <div class="timeline-container" style="display:flex; flex-direction:column; gap:12px; max-width:960px; margin:0 auto;">
        ${a.map(p=>{let d=Ga[p.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},g=p.shift==="pagi"?"\u{1F305} Pagi":"\u{1F307} Siang",s=p.shift==="pagi"?"#FEF3C7":"#E0E7FF",r=p.shift==="pagi"?"#92400E":"#3730A3";return`
            <div class="dac-timeline-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:16px 20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; gap:16px; align-items:flex-start; position:relative;">
              <!-- Time Column -->
              <div style="min-width:130px; text-align:center; background:#F8FAFC; border:1px solid var(--border); border-radius:8px; padding:10px 8px;">
                <div style="font-size:1.1rem; font-weight:800; color:var(--text-1); font-family:monospace;">
                  ${p.start_time} - ${p.end_time}
                </div>
                <div style="font-size:0.75rem; color:var(--text-3); font-weight:600; margin-top:4px;">
                  Item #${p.item_no}
                </div>
                ${t==="all"?`
                  <div style="margin-top:6px;">
                    <span class="badge" style="background:${s}; color:${r}; font-size:0.7rem; font-weight:700; padding:2px 6px; border-radius:4px;">
                      ${g}
                    </span>
                  </div>
                `:""}
              </div>

              <!-- Content Column -->
              <div style="flex:1;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                  <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span class="badge" style="background:${d.bg}; color:${d.color}; border:1px solid ${d.border}; font-weight:700; font-size:0.75rem; padding:3px 10px; border-radius:6px;">
                      ${p.category}
                    </span>
                    <span style="font-weight:700; font-size:0.95rem; color:var(--text-1);">
                      \u{1F4CD} ${p.location}
                    </span>
                  </div>
                  ${x?`
                    <div style="display:flex; gap:4px;">
                      <button class="btn btn-ghost btn-sm edit-dac-btn" data-id="${p.id}" title="Edit" style="padding:4px 8px; font-size:0.8rem; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-ghost btn-sm delete-dac-btn" data-id="${p.id}" data-name="${p.location}" title="Hapus" style="padding:4px 8px; font-size:0.8rem; color:var(--danger); cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </div>
                  `:""}
                </div>

                <div style="font-size:0.9rem; color:var(--text-2); line-height:1.5; background:#FAFBFD; padding:10px 14px; border-radius:8px; border-left:3px solid ${d.color}; margin-top:8px;">
                  ${p.activity}
                </div>
              </div>
            </div>
          `}).join("")}
      </div>
    `,S(h)}function D(h,a){if(!a.length){h.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
        </div>
      `;return}h.innerHTML=`
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
                ${x?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right; width:100px;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${a.map(p=>{let d=Ga[p.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"};return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 14px; font-size:0.85rem; color:var(--text-3); font-weight:600;">${p.item_no}</td>
                    ${t==="all"?`
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
                      <span class="badge" style="background:${d.bg}; color:${d.color}; border:1px solid ${d.border}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px; white-space:nowrap;">
                        ${p.category}
                      </span>
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; font-weight:600; color:var(--text-1);">
                      ${p.location}
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; color:var(--text-2); line-height:1.4;">
                      ${p.activity}
                    </td>
                    ${x?`
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
    `,S(h)}function S(h){h.querySelectorAll(".edit-dac-btn").forEach(a=>{a.addEventListener("click",()=>{let p=a.dataset.id,d=n.find(g=>String(g.id)===String(p));d&&F(d)})}),h.querySelectorAll(".delete-dac-btn").forEach(a=>{a.addEventListener("click",()=>{let p=a.dataset.id,d=a.dataset.name||"item";X(`Apakah Anda yakin ingin menghapus aktivitas pada lokasi <strong>${d}</strong>?`,async()=>{try{if(!(await K(`/api/dac/${p}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus item");R("Aktivitas DAC berhasil dihapus"),m()}catch(g){N(g.message)}})})})}function F(h=null){let a=!!h,p=a?"Edit Aktivitas DAC":"Tambah Aktivitas DAC",d=["Persiapan","Area Umum","Ruang Tindakan","Ruang Penunjang","Toilet","Area Luar","Kontrol & Check","Istirahat","Layanan Pasien","Sampah Umum","Sampah Medis","Penutup"],g=`
      <form id="dac-form" style="display:flex; flex-direction:column; gap:14px;">
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Shift *</label>
            <select name="shift" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="pagi" ${h?.shift==="pagi"||!h&&t==="pagi"?"selected":""}>\u{1F305} Shift Pagi (08:00 - 17:00)</option>
              <option value="siang" ${h?.shift==="siang"||!h&&t==="siang"?"selected":""}>\u{1F307} Shift Siang (13:00 - 22:00)</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Nomor Urut *</label>
            <input type="number" name="item_no" class="form-control" value="${h?.item_no||""}" placeholder="Otomatis jika kosong" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Mulai (HH:MM) *</label>
            <input type="time" name="start_time" class="form-control" value="${h?.start_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Selesai (HH:MM) *</label>
            <input type="time" name="end_time" class="form-control" value="${h?.end_time||""}" required style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Kategori *</label>
          <input list="cat-suggestions" name="category" class="form-control" value="${h?.category||""}" required placeholder="Pilih atau ketik kategori..." style="width:100%; height:38px; border-radius:8px;">
          <datalist id="cat-suggestions">
            ${d.map(s=>`<option value="${s}">`).join("")}
          </datalist>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Lokasi / Area *</label>
          <input type="text" name="location" class="form-control" value="${h?.location||""}" required placeholder="Contoh: Ruang tindakan Du 101, Lobby, Toilet GF" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Aktivitas Detail *</label>
          <textarea name="activity" class="form-control" rows="3" required placeholder="Deskripsi aktivitas pembersihan..." style="width:100%; border-radius:8px;">${h?.activity||""}</textarea>
        </div>
      </form>
    `;G({title:p,content:g,confirmText:a?"Simpan Perubahan":"Tambah Aktivitas",onConfirm:async(s,r)=>{let l=s.querySelector("#dac-form");if(!l.checkValidity()){l.reportValidity();return}let f=new FormData(l),v={shift:f.get("shift"),item_no:f.get("item_no")?parseInt(f.get("item_no"),10):void 0,start_time:f.get("start_time"),end_time:f.get("end_time"),category:f.get("category"),location:f.get("location"),activity:f.get("activity")};try{let L=a?`/api/dac/${h.id}`:"/api/dac",P=await K(L,{method:a?"PUT":"POST",body:JSON.stringify(v)});if(!P.ok){let w=await P.json().catch(()=>({}));throw new Error(w.error||"Gagal menyimpan aktivitas")}R(a?"Aktivitas berhasil diperbarui":"Aktivitas berhasil ditambahkan"),r(),m()}catch(L){N(L.message)}}})}function $(){let h=u();if(!h.length){N("Tidak ada data untuk dicetak");return}let a=t==="pagi"?"SHIFT PAGI (08:00 - 17:00)":t==="siang"?"SHIFT SIANG (13:00 - 22:00)":"SELURUH SHIFT (PAGI & SIANG)",p=window.open("","_blank");if(!p){N("Popup diblokir browser, mohon izinkan popup untuk mencetak");return}p.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>STANDAR ACUAN DAC - ${a}</title>
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
          <div class="subtitle">STANDAR RUTINITAS DAILY ACTIVITY CLEANING (DAC) \u2014 ${a}</div>
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
            ${h.map(d=>`
              <tr>
                <td style="text-align:center; font-weight:bold;">${d.item_no}</td>
                ${t==="all"?`<td>${d.shift==="pagi"?"Pagi":"Siang"}</td>`:""}
                <td class="jam">${d.start_time} - ${d.end_time}</td>
                <td><strong>${d.category}</strong></td>
                <td><strong>${d.location}</strong></td>
                <td>${d.activity}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
        <div style="margin-top:20px; text-align:right; font-size:9pt; color:#64748b;">
          Dicetak pada: ${new Date().toLocaleString("id-ID")} | Standar Resmi Operasional Facility Care Audy Dental
        </div>
      </body>
      </html>
    `),p.document.close(),p.focus(),setTimeout(()=>{p.print()},400)}document.querySelectorAll(".shift-tab-btn").forEach(h=>{h.addEventListener("click",()=>{document.querySelectorAll(".shift-tab-btn").forEach(a=>{a.style.background="#F1F5F9",a.style.color="var(--text-2)"}),h.style.background="#2563EB",h.style.color="#fff",t=h.dataset.shift,o(),b()})});let E=document.getElementById("view-timeline-btn"),k=document.getElementById("view-table-btn");E.addEventListener("click",()=>{i="timeline",E.style.background="#fff",E.classList.remove("btn-ghost"),k.style.background="transparent",k.classList.add("btn-ghost"),b()}),k.addEventListener("click",()=>{i="table",k.style.background="#fff",k.classList.remove("btn-ghost"),E.style.background="transparent",E.classList.add("btn-ghost"),b()}),["dac-search","dac-category-filter","dac-period-filter"].forEach(h=>{let a=document.getElementById(h);a&&(a.addEventListener("input",b),a.addEventListener("change",b))}),document.getElementById("btn-add-dac").addEventListener("click",()=>{F()}),document.getElementById("btn-print-dac").addEventListener("click",()=>{$()}),document.getElementById("btn-export-dac").addEventListener("click",()=>{let h=u();if(!h.length){N("Tidak ada data untuk diexport");return}let a=h.map(p=>({No:p.item_no,Shift:p.shift==="pagi"?"Shift Pagi":"Shift Siang","Jam Mulai":p.start_time,"Jam Selesai":p.end_time,Kategori:p.category,"Lokasi / Area":p.location,"Rincian Aktivitas":p.activity}));Q(a,`DAC_Daily_Activity_Cleaning_${t}_${new Date().toISOString().slice(0,10)}`),R("File Excel berhasil diunduh")}),m()}window.parseFlexibleDate=e=>{if(!e||e==="-")return"";if(e=String(e).trim(),/^\d{5}$/.test(e)){let i=Math.floor(Number(e)-25569);return new Date(i*86400*1e3).toISOString().split("T")[0]}if(e.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let i=e.split(/[\/\-]/);return`${i[2]}-${i[1]}-${i[0]}`}let t=e.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(t){let i=t[1],n=parseInt(t[2],10),c=parseInt(t[3],10);if(n>12&&c<=12)return`${i}-${t[3]}-${t[2]}`}return e.split("T")[0]};window.formatDate=e=>{let t=window.parseFlexibleDate(e);if(!t)return"";let i=t.split("-");if(i.length===3&&i[0].length===4){let n=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],c=parseInt(i[2],10),x=n[parseInt(i[1],10)-1];return`${c} ${x} ${i[0]}`}return t};function oe(e){return async t=>{if(!Ye()){$e("/login");return}return e(t)}}var Ge=null;function Hn(){Ge&&clearInterval(Ge);let e=()=>{let t=new Date,i=t.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),n=t.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),c=document.getElementById("header-clock-time"),x=document.getElementById("header-clock-date");c&&(c.textContent=i),x&&(x.textContent=n)};e(),Ge=setInterval(e,1e3)}async function Un(){try{let e=await K("/api/dashboard/kpi");if(!e.ok)return;let t=e.data?.data||e.data||{},i=(n,c)=>{let x=document.getElementById(n);x&&(x.textContent=c>0?c:"",x.style.display=c>0?"inline-flex":"none")};i("badge-issues",t.issues?.current||0),i("badge-contracts",t.expiring30?.current||0),i("badge-oo1",t.one_on_one?.current||0),i("badge-schedule",t.schedule?.current||0),i("badge-supply",t.supply?.current||0)}catch{}}var je=[];async function Gn(){try{let e=await K("/api/dashboard/notifications");if(!e.ok)return;je=e.data?.data||e.data||[];let t=document.getElementById("notif-dot");t&&(t.style.display=je.length>0?"block":"none",t.textContent=je.length)}catch{}}function Jn(){if(!je.length){G({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(t,i)=>i()});return}let e=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${je.map(t=>`
        <div class="notif-item notif-severity-${t.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${t.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${t.date}</span>
            <span class="badge badge-${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"info"}">${t.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;G({title:`Notifikasi (${je.length})`,content:e,confirmText:"Tutup",onConfirm:(t,i)=>i()})}function Ya(){let e=ee(),t=(e?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
                <span class="topbar-greeting-time">${(()=>{let b=new Date().getHours();return b>=4&&b<11?"Selamat Pagi":b>=11&&b<15?"Selamat Siang":b>=15&&b<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${e?.full_name||e?.username||"Guest"}</span> \u{1F44B}
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
  `;let i=document.getElementById("sidebar"),n=document.getElementById("sidebar-overlay"),c=document.getElementById("topbar-menu-btn"),x=document.getElementById("sidebar-close"),m=()=>{i.classList.add("open"),n.classList.add("show")},u=()=>{i.classList.remove("open"),n.classList.remove("show")};c?.addEventListener("click",m),x?.addEventListener("click",u),n?.addEventListener("click",u),document.querySelectorAll(".nav-item").forEach(b=>b.addEventListener("click",u)),e&&e.role==="input_lembur"&&(document.querySelectorAll(".nav-item").forEach(b=>{b.dataset.route!=="/overtime"&&(b.style.display="none")}),document.querySelectorAll(".nav-section").forEach(b=>{Array.from(b.querySelectorAll(".nav-item")).some(D=>D.style.display!=="none")||(b.style.display="none")}));function o(){let b=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(S=>{let F=S.dataset.route;S.classList.toggle("active",b===F||F!=="/dashboard"&&b.startsWith(F))});let B=document.getElementById("topbar-title"),D=document.querySelector(".nav-item.active .nav-label");B&&D&&(B.textContent=D.textContent)}window.addEventListener("hashchange",o),o(),Hn(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await K("/api/auth/logout",{method:"POST"}),Ve(),Ge&&clearInterval(Ge),$e("/login")}),Un(),Gn(),document.getElementById("btn-notif")?.addEventListener("click",b=>{b.preventDefault(),Jn()})}async function Yn(){ne("/login",({main:t})=>Wt(t)),ne("/dashboard",oe(({main:t})=>Ut(t))),ne("/calendar",oe(({main:t})=>Ba(t))),ne("/employees",oe(({main:t,params:i})=>Xt(t,i))),ne("/contracts",oe(({main:t,params:i})=>Zt(t,i))),ne("/sp",oe(({main:t})=>Ka(t))),ne("/mutasi",oe(({main:t})=>Na(t))),ne("/overtime",oe(({main:t})=>Oa(t))),ne("/timeline",oe(({main:t,params:i})=>ea(t,i))),ne("/os-migrations",oe(({main:t})=>Ua(t))),ne("/issues",oe(({main:t,params:i})=>na(t,i))),ne("/one-on-one",oe(({main:t,params:i})=>ia(t,i))),ne("/training",oe(({main:t})=>sa(t))),ne("/relievers",oe(({main:t,params:i})=>da(t,i))),ne("/reports/inspection",oe(({main:t})=>pa(t))),ne("/reports/cleaning",oe(({main:t})=>ca(t))),ne("/reports/fogging",oe(({main:t})=>ua(t))),ne("/reports/basecamp",oe(({main:t})=>ga(t))),ne("/reports/supply",oe(({main:t})=>$a(t,"supply"))),ne("/sop",()=>{window.location.hash="#/sop-fc"}),ne("/sop-fc",oe(({main:t})=>fa(t))),ne("/sop-security",oe(({main:t})=>ka(t))),ne("/checklist",oe(({main:t})=>Ca(t))),ne("/hygiene-standards",oe(({main:t})=>Fa(t))),ne("/dac",oe(({main:t})=>Ja(t))),ne("/users",oe(({main:t})=>_a(t))),ne("/branches",oe(({main:t})=>Aa(t))),ne("/profile",oe(({main:t})=>Ta(t))),ne("/settings/import",oe(({main:t})=>za(t))),ne("/audit-logs",oe(({main:t})=>Ha(t)));let e=Ye();if(!e&&window.location.hash!=="#/login"&&$e("/login"),e){let t=await K("/api/auth/me");if(t.ok){We(t.data.data),Ya();let i=ee(),n=window.location.hash;i&&i.role==="input_lembur"&&(n==="#/dashboard"||n===""||n==="#/")&&$e("/overtime")}else Ve(),$e("/login")}window.addEventListener("fm:login",()=>{Ya();let t=ee();t&&t.role==="input_lembur"?$e("/overtime"):$e("/dashboard")}),qt()}Yn();
