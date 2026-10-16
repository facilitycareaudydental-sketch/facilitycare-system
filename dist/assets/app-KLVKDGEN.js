var Ba=Object.defineProperty;var bt=(e,t)=>()=>(e&&(t=e(e=0)),t);var ht=(e,t)=>{for(var i in t)Ba(e,i,{get:t[i],enumerable:!0})};var De={};ht(De,{API:()=>Nt,CLIENT_SIDE_MAX_ROWS:()=>Fe,IS_DEVELOPMENT:()=>ze,apiFetch:()=>F,clearToken:()=>Re,getToken:()=>je,getUser:()=>ie,setToken:()=>ft,setUser:()=>qe});function je(){return localStorage.getItem("fm_token")}function ft(e){localStorage.setItem("fm_token",e)}function Re(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ie(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function qe(e){localStorage.setItem("fm_user",JSON.stringify(e))}async function F(e,t={}){let i=je(),n={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...t.headers||{}};try{let r=`cb=${Date.now()}`,d=e.includes("?")?"&":"?",o=`${Nt}${e}${d}${r}`,s=await fetch(o,{...t,headers:n}),a;try{let l=await s.text();try{a=JSON.parse(l)}catch{a={error:`Server Error (${s.status}): ${l.substring(0,80)}...`}}}catch{a={error:"Gagal membaca respon dari server"}}return s.status===401&&(Re(),window.location.hash="#/login"),{ok:s.ok,status:s.status,data:a}}catch(r){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${r.message})`}}}}var ze,Pa,Nt,Fe,H=bt(()=>{ze=!1,Pa="https://fm-operations-api.facilitycare-audydental.workers.dev",Nt=Pa,Fe=1e4});var vt={};ht(vt,{confirmDialog:()=>pe,createModal:()=>ae});function ae({title:e,content:t,onConfirm:i,onCancel:n,confirmText:r="Simpan",cancelText:d="Batal",size:o="md",confirmClass:s="btn-primary"}){let a={sm:"400px",md:"560px",lg:"720px",xl:"900px"},l=document.createElement("div");l.className="modal-overlay",l.innerHTML=`
    <div class="modal" style="max-width:${a[o]||a.md}">
      <div class="modal-header">
        <h3 class="modal-title">${e}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof t=="string"?t:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${d}</button>
        ${i?`<button class="btn ${s} modal-confirm">${r}</button>`:""}
      </div>
    </div>
  `,t instanceof HTMLElement&&l.querySelector(".modal-body").appendChild(t);let v=()=>{l.classList.remove("show"),setTimeout(()=>l.remove(),250)};return l.querySelector(".modal-close").addEventListener("click",()=>{n&&n(),v()}),l.querySelector(".modal-cancel").addEventListener("click",()=>{n&&n(),v()}),i&&l.querySelector(".modal-confirm").addEventListener("click",()=>i(l,v)),l.addEventListener("click",x=>{x.target===l&&(n&&n(),v())}),document.body.appendChild(l),requestAnimationFrame(()=>l.classList.add("show")),{overlay:l,close:v}}function pe(e,t,i="Konfirmasi"){return ae({title:i,content:`<p>${e}</p>`,onConfirm:(n,r)=>{t(),r()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var he=bt(()=>{});var me={};ht(me,{downloadExcel:()=>K,parseExcel:()=>Ue,renderExcelButtons:()=>Na});function Ue(e){return new Promise((t,i)=>{let n=new FileReader;n.onload=r=>{try{let d=new Uint8Array(r.target.result),o=XLSX.read(d,{type:"array"}),s=o.SheetNames[0],a=o.Sheets[s];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${e.name}`),console.log(`File Size: ${(e.size/1024).toFixed(2)} KB`),console.log(`File Type: ${e.type||"unknown"}`),console.log(`Sheets Found: ${o.SheetNames.join(", ")}`),console.log(`Sheet Used: ${s}`);let l=XLSX.utils.decode_range(a["!ref"]||"A1:A1"),v=l.e.r-l.s.r+1,x=l.e.c-l.s.c+1;console.log(`Total Rows (including empty): ${v}`),console.log(`Total Columns: ${x}`);let h=[];for(let S=l.s.c;S<=l.e.c;++S){let k=a[XLSX.utils.encode_cell({c:S,r:l.s.r})];k&&k.v&&h.push(k.v)}console.log(`Headers Found: ${h.join(", ")}`),console.log("---------------------------");let m=XLSX.utils.sheet_to_json(a,{defval:""});Object.defineProperty(m,"__worksheet",{value:a,enumerable:!1}),Object.defineProperty(m,"__headers",{value:h,enumerable:!1}),t(m)}catch(d){i(d)}},n.onerror=r=>i(r),n.readAsArrayBuffer(e)})}function K(e,t){try{let i=XLSX.utils.json_to_sheet(e),n=XLSX.utils.book_new();XLSX.utils.book_append_sheet(n,i,"Data"),XLSX.writeFile(n,`${t}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function Na(e){return`
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
  `}var G=bt(()=>{});H();var yt={},Xe=null;function X(e,t){yt[e]=t}function $e(e){window.location.hash=e}function Ot(){async function e(){let t=window.location.hash.replace("#","")||"/dashboard",[i,...n]=t.split("?"),r=yt[i];if(!r){for(let[o,s]of Object.entries(yt))if(o.endsWith("/*")&&i.startsWith(o.slice(0,-2))){r=s;break}}Xe&&(Xe(),Xe=null);let d=document.getElementById("main-content");if(d&&(d.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),r){let o=new URLSearchParams(n.join("?")),s=i.split("/").filter(Boolean),a=await r({path:i,params:o,segments:s,main:d});a&&(Xe=a)}else{let o=d||document.getElementById("app");o&&(o.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",e),e()}var He;function Ma(){return He||(He=document.createElement("div"),He.id="toast-container",document.body.appendChild(He)),He}function Kt(e,t="info",i=3500){let n=Ma(),r=document.createElement("div");r.className=`toast toast-${t}`;let d={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};r.innerHTML=`<span class="toast-icon">${d[t]||"\u2139"}</span><span class="toast-msg">${e}</span>`,n.appendChild(r),requestAnimationFrame(()=>r.classList.add("show")),setTimeout(()=>{r.classList.remove("show"),setTimeout(()=>r.remove(),350)},i)}var j=e=>Kt(e,"success"),R=e=>Kt(e,"error");he();H();H();he();function Ze({columns:e,data:t,onEdit:i,onDelete:n,onView:r,actions:d=[],emptyText:o="Tidak ada data",bulkSelect:s=null}){let a=document.createElement("div");if(a.className="table-wrapper",!t||t.length===0)return a.innerHTML=`<div class="empty-state"><p>${o}</p></div>`,a;let l=document.createElement("table");l.className="data-table";let v=document.createElement("thead"),x=document.createElement("tr");if(s){let m=document.createElement("th");m.style.width="40px",m.style.textAlign="center";let S=document.createElement("input");S.type="checkbox",S.id="select-all-checkbox",S.title="Pilih semua",S.addEventListener("change",()=>{t.forEach(k=>{S.checked?s.selectedIds.add(k.id):s.selectedIds.delete(k.id)}),a.querySelectorAll(".row-checkbox").forEach(k=>k.checked=S.checked),s.onToggle()}),m.appendChild(S),x.appendChild(m)}if(e.forEach(m=>{let S=document.createElement("th");S.textContent=m.label,m.width&&(S.style.width=m.width),x.appendChild(S)}),i||n||r||d.length>0){let m=document.createElement("th");m.textContent="Aksi",m.style.width="120px",x.appendChild(m)}v.appendChild(x),l.appendChild(v);let h=document.createElement("tbody");return t.forEach(m=>{let S=document.createElement("tr");if(s){let k=document.createElement("td");k.style.textAlign="center",k.style.width="40px";let g=document.createElement("input");g.type="checkbox",g.className="row-checkbox",g.checked=s.selectedIds.has(m.id),g.addEventListener("change",()=>{if(g.checked)s.selectedIds.add(m.id);else{s.selectedIds.delete(m.id);let c=document.getElementById("select-all-checkbox");c&&(c.checked=!1)}s.onToggle()}),k.appendChild(g),S.appendChild(k)}if(e.forEach(k=>{let g=document.createElement("td");if(k.render){let c=k.render(m[k.key],m);c instanceof HTMLElement?g.appendChild(c):g.innerHTML=c||""}else g.textContent=m[k.key]!==null&&m[k.key]!==void 0&&m[k.key]!==""?m[k.key]:"";k.nowrap&&(g.style.whiteSpace="nowrap"),S.appendChild(g)}),i||n||r||d.length>0){let k=document.createElement("td");k.className="actions-cell";let g=document.createElement("div");if(g.className="btn-group",r){let c=document.createElement("button");c.className="btn btn-xs btn-ghost",c.innerHTML="\u{1F441}",c.title="Lihat",c.addEventListener("click",()=>r(m)),g.appendChild(c)}if(i){let c=document.createElement("button");c.className="btn btn-xs btn-secondary",c.innerHTML="\u270F\uFE0F",c.title="Edit",c.addEventListener("click",()=>i(m)),g.appendChild(c)}d.forEach(c=>{let f=document.createElement("button");f.className=`btn btn-xs ${c.class||"btn-ghost"}`,f.innerHTML=c.icon||c.label,f.title=c.label,f.addEventListener("click",()=>c.handler(m)),g.appendChild(f)}),k.appendChild(g),S.appendChild(k)}h.appendChild(S)}),l.appendChild(h),a.appendChild(l),a}function et({page:e,pages:t,total:i,limit:n,onPage:r}){if(t<=1)return null;let d=document.createElement("div");d.className="pagination";let o=document.createElement("span");o.className="pagination-info",o.textContent=`Total: ${i} data`,d.appendChild(o);let s=document.createElement("div");s.className="pagination-btns";let a=(x,h,m=!1,S=!1)=>{let k=document.createElement("button");k.className=`btn btn-sm ${S?"btn-primary":"btn-ghost"} pagination-btn`,k.textContent=x,k.disabled=m,k.addEventListener("click",()=>r(h)),s.appendChild(k)};a("\xAB",1,e===1),a("\u2039",e-1,e===1);let l=Math.max(1,e-2),v=Math.min(t,e+2);for(let x=l;x<=v;x++)a(x,x,!1,x===e);return a("\u203A",e+1,e===t),a("\xBB",t,e===t),d.appendChild(s),d}he();function Je(e){return e.map(t=>{if(t.type==="hidden")return`<input type="hidden" name="${t.name}" value="${t.value||""}">`;if(t.type==="html")return t.html||"";if(t.type==="row")return`<div class="form-row">${Je(t.fields)}</div>`;let i=t.required?"required":"",n=t.label?`<label class="form-label">${t.label}${t.required?' <span class="required">*</span>':""}</label>`:"",r="";switch(t.type){case"textarea":r=`<textarea name="${t.name}" class="form-control" placeholder="${t.placeholder||""}" ${i} rows="${t.rows||3}">${t.value||""}</textarea>`;break;case"select":let o=(t.options||[]).map(m=>{let S=typeof m=="object"?m.value:m,k=typeof m=="object"?m.label:m,g=t.value==S?"selected":"";return`<option value="${S}" ${g}>${k}</option>`}).join("");r=`<select name="${t.name}" class="form-control" ${i}><option value="">-- Pilih ${t.label||""} --</option>${o}</select>`;break;case"combobox":let s=`dl-${t.name}-${Math.random().toString(36).substring(7)}`,a=(t.options||[]).map(m=>{let S=typeof m=="object"?m.value:m,k=typeof m=="object"?m.label||m.value||"":m||"";return(k==="undefined"||k==="[object Object]"||k==="null")&&(k=""),k?`<option value="${k}"></option>`:""}).join(""),l=t.value||"";if(t.value){let m=(t.options||[]).find(S=>(typeof S=="object"?S.value:S)==t.value);if(m){let S=typeof m=="object"?m.label||m.value||"":m||"";S&&S!=="undefined"&&S!=="[object Object]"&&S!=="null"&&(l=S)}}r=`
          <input type="text" name="${t.name}" list="${s}" class="form-control" value="${l}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${s}">${a}</datalist>
        `;break;case"checkbox":r=`<label class="checkbox-label"><input type="checkbox" name="${t.name}" value="1" ${t.value?"checked":""}> ${t.checkLabel||t.label}</label>`;break;case"date":let v=window.parseFlexibleDate&&t.value?window.parseFlexibleDate(t.value):t.value||"";r=`<input type="date" name="${t.name}" class="form-control" value="${v}" ${i}>`;break;case"month":let x=t.value?String(t.value).slice(0,7):"";r=`<input type="month" name="${t.name}" class="form-control" value="${x}" ${i}>`;break;case"number":let h=t.readonly?"readonly":"";r=`<input type="number" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" min="${t.min||""}" max="${t.max||""}" step="${t.step||"1"}" ${i} ${h}>`;break;case"email":r=`<input type="email" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${i}>`;break;case"url":r=`<input type="url" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||"https://..."}" ${i}>`;break;default:r=`<input type="${t.type||"text"}" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${i} autocomplete="off">`}let d=t.hint?`<div class="form-hint">${t.hint}</div>`:"";return`<div class="form-group ${t.class||""}">${n}${r}${d}</div>`}).join("")}function tt(e){let t={},i=new FormData(e);for(let[n,r]of i.entries())t[n]=r===""?null:r;return e.querySelectorAll("input[type=checkbox]").forEach(n=>{n.checked||(t[n.name]=null)}),t}function at(e,t){t&&Object.entries(t).forEach(([i,n])=>{let r=e.querySelector(`[name="${i}"]`);r&&(r.hasAttribute("list")||(r.type==="checkbox"?r.checked=!!n:r.type==="date"&&n&&window.parseFlexibleDate?r.value=window.parseFlexibleDate(n):r.type==="month"&&n?r.value=String(n).slice(0,7):r.value=n??""))})}G();var Ce={},nt={on(e,t){Ce[e]||(Ce[e]=new Set),Ce[e].add(t)},off(e,t){Ce[e]&&Ce[e].delete(t)},emit(e,t){Ce[e]&&Ce[e].forEach(i=>{try{i(t)}catch(n){console.warn("[calendarBus] Handler error:",n)}})},clear(){Object.keys(Ce).forEach(e=>delete Ce[e])}},Oa=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function be(e){if(!e){nt.emit("data:changed",{module:"unknown"});return}let t=String(e).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");nt.emit("data:changed",{module:t,relevant:Oa.has(t)})}function q({container:e,title:t,icon:i,apiPath:n,columns:r,formFields:d,filterFields:o,defaultFilters:s={},itemLabel:a="Data",canCreate:l=!0,canEdit:v=!0,canDelete:x=!0,onBeforeSubmit:h,onAfterLoad:m,onDataLoaded:S,extraActions:k=[],initialSearch:g="",exportOptions:c=null,bulkDelete:f=!1,paginationMode:u="server"}){let w=ie();w&&w.role==="viewer"&&(l=!1,v=!1,x=!1,f=!1,c=null),w&&w.role==="editor_khusus"&&n!=="/api/issues"&&n!=="/api/overtime"&&(l=!1,v=!1,x=!1,f=!1,c=null),w&&w.role==="input_lembur"&&(n==="/api/overtime"?(x=!1,f=!1,c=null):(l=!1,v=!1,x=!1,f=!1,c=null));let b=1,C={...s};g&&(C.search=g);let _=new Set;e.innerHTML=`
    ${f?`
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
        ${l?`<button class="btn btn-primary" id="btn-create">+ Tambah ${a}</button>`:""}
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
        ${o.filter(y=>y.type==="search").map(y=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${y.placeholder||"Cari..."}" id="filter-search" value="${C.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${o.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let E=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select" name="${y.name}" id="filter-${y.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${E}</option>${(y.options||[]).map($=>`<option value="${typeof $=="object"?$.value:$}" ${C[y.name]===(typeof $=="object"?$.value:$)?"selected":""}>${typeof $=="object"?$.label:$}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let E=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select filter-select-sheet" name="${y.name}-sheet" id="filter-sheet-${y.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${E}</option>${(y.options||[]).map($=>`<option value="${typeof $=="object"?$.value:$}" ${C[y.name]===(typeof $=="object"?$.value:$)?"selected":""}>${typeof $=="object"?$.label:$}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function p(){let y=document.getElementById("bulk-toolbar");if(!y)return;let E=document.getElementById("bulk-count"),$=document.getElementById("btn-bulk-delete"),U=document.getElementById("btn-bulk-cancel");E.textContent=`${_.size} item dipilih`,_.size>0?(y.style.display="flex",$.disabled=!1,U.disabled=!1):(y.style.display="none",$.disabled=!0,U.disabled=!0);let J=document.getElementById("select-all-checkbox");if(J){let Y=document.querySelectorAll(".row-checkbox");if(Y.length>0){let M=[...Y].every(oe=>oe.checked),z=[...Y].some(oe=>oe.checked);J.checked=M,J.indeterminate=z&&!M}else J.checked=!1,J.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{_.clear(),document.querySelectorAll(".row-checkbox").forEach(E=>E.checked=!1);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1),p()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(_.size===0)return;let y=[..._],E=document.createElement("div");E.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",E.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${y.length} ${a}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${y.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(E),E.querySelector("#bulk-cancel-btn").addEventListener("click",()=>E.remove()),E.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let $=E.querySelector("#bulk-confirm-btn");$.disabled=!0,$.textContent="Menghapus...";let U=await F(`${n}/bulk`,{method:"DELETE",body:JSON.stringify({ids:y})});E.remove(),U.ok?(j(`${y.length} ${a} berhasil dihapus.`),_.clear(),p(),be(n),I()):R(U.data?.error||"Gagal menghapus data.")})});let T=document.getElementById("filter-search"),D;T?.addEventListener("input",y=>{clearTimeout(D),D=setTimeout(()=>{C.search=y.target.value,b=1,_.clear(),p(),I()},400)}),o?.forEach(y=>{(y.type==="select"||y.type==="combobox")&&(document.getElementById(`filter-${y.name}`)?.addEventListener("change",E=>{C[y.name]=E.target.value;let $=document.getElementById(`filter-sheet-${y.name}`);$&&($.value=E.target.value),b=1,_.clear(),p(),I()}),document.getElementById(`filter-sheet-${y.name}`)?.addEventListener("change",E=>{C[y.name]=E.target.value;let $=document.getElementById(`filter-${y.name}`);$&&($.value=E.target.value),b=1,_.clear(),p(),I(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{C={},T&&(T.value=""),o?.forEach(y=>{let E=document.getElementById(`filter-${y.name}`);E&&(E.value="");let $=document.getElementById(`filter-sheet-${y.name}`);$&&($.value="")}),b=1,_.clear(),p(),I()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{C={},T&&(T.value=""),o?.forEach(y=>{let E=document.getElementById(`filter-${y.name}`);E&&(E.value="");let $=document.getElementById(`filter-sheet-${y.name}`);$&&($.value="")}),b=1,_.clear(),p(),I(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>L(null)),c&&document.addEventListener("click",function(y){let E=document.getElementById("aksi-menu-main"),$=document.getElementById("btn-aksi-main");E&&$&&!$.contains(y.target)&&!E.contains(y.target)&&E.classList.remove("show-aksi-menu")});let A=document.getElementById("btn-mobile-filter"),O=document.getElementById("filter-options-wrapper"),P=document.getElementById("btn-close-filter-sheet");if(A&&O&&(A.addEventListener("click",y=>{y.preventDefault(),O.classList.add("sheet-open")}),P&&P.addEventListener("click",y=>{y.preventDefault(),O.classList.remove("sheet-open")})),c){document.getElementById(`btn-export-${c.moduleName}`)?.addEventListener("click",async E=>{let $=E.target,U=$.innerHTML;$.innerHTML="\u23F3 Loading...",$.disabled=!0;try{await c.onExport()}catch{R("Gagal export data")}finally{$.innerHTML=U,$.disabled=!1}}),document.getElementById(`btn-template-${c.moduleName}`)?.addEventListener("click",()=>{c.onTemplate()});let y=document.getElementById(`input-import-${c.moduleName}`);y?.addEventListener("change",async E=>{let $=E.target.files[0];if(!$)return;let U=document.getElementById(`label-import-${c.moduleName}`),J=U?U.querySelector(".import-text"):null,Y=J?J.innerText:"";J&&(J.innerText="\u231B Memproses..."),U&&(U.style.pointerEvents="none"),y.disabled=!0;try{let M=await Ue($);if(M.length===0)throw new Error("File kosong atau format salah");await c.onImport(M),j("Import berhasil!"),be(n),I()}catch(M){R(M.message||"Gagal import data")}finally{J&&(J.innerText=Y),U&&(U.style.pointerEvents="auto"),y.disabled=!1,y.value=""}})}async function I(){p();let y=document.getElementById("table-container");if(!y)return;y.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let E=u==="client",$=E?1:b,U=E?Fe:20,J=new URLSearchParams({page:$,limit:U,...Object.fromEntries(Object.entries(C).filter(([,Q])=>Q))}),Y=await F(`${n}?${J}`);if(!Y.ok){y.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${Y.data?.error||"Error"}</p></div>`;return}let M=Y.data?.data||Y.data||[],z=Y.data?.pagination,oe=M.length;if(E){M=S(M);let Q=M.length,te=20,le=Math.ceil(Q/te);b>le&&le>0&&(b=le);let V=(b-1)*te,ge=b*te;M=M.slice(V,ge),z={page:b,limit:te,total:Q,pages:le}}!1,m&&m(M);let ce=Ze({columns:r,data:M,onEdit:v?Q=>L(Q):null,actions:k.map(Q=>({...Q,handler:te=>Q.handler(te,I)})),emptyText:`Tidak ada ${String(a||"").toLowerCase()}`,bulkSelect:f?{selectedIds:_,onToggle:p}:null});y.innerHTML="",y.appendChild(ce);let ue=document.getElementById("pagination-container");if(ue&&(ue.innerHTML="",z&&z.pages>1)){let Q=et({page:z.page,pages:z.pages,total:z.total,limit:z.limit,onPage:te=>{b=te,I()}});Q&&ue.appendChild(Q)}}function B(y){let E=typeof d=="function"?d(y):d;return Je(E)}function L(y){let E=!!y,$=document.createElement("form");if($.noValidate=!0,$.innerHTML=B(y),E){let J=typeof d=="function"?d(y):d;at($,y)}let{close:U}=ae({title:E?`Edit ${a}`:`Tambah ${a}`,content:$,size:"lg",confirmText:E?"Simpan Perubahan":`Tambah ${a}`,onConfirm:async(J,Y)=>{if(!$.reportValidity())return;let M=J.querySelector(".modal-confirm");M.disabled=!0,M.textContent="Menyimpan...";let z=tt($),oe=typeof d=="function"?d(y):d,ce=async le=>{for(let V of le)if(V.type==="row")await ce(V.fields);else if(V.type==="combobox"&&z[V.name]){let ge=z[V.name],_e=(V.options||[]).find(ne=>{let de=String(typeof ne=="object"?ne.value:ne),gt=String(typeof ne=="object"?ne.label:ne);return de===ge||gt===ge});if(_e)z[V.name]=typeof _e=="object"?_e.value:_e;else if(V.createApi){let ne={};ne[V.createApi.field]=ge,V.createApi.extra&&Object.assign(ne,V.createApi.extra);let de=await F(V.createApi.path,{method:"POST",body:JSON.stringify(ne)});if(de.ok&&de.data?.id)z[V.name]=de.data.id;else if(de.ok&&!de.data?.id)z[V.name]=ge;else throw new Error(`Gagal membuat master data: ${de.data?.error||"Unknown error"}`)}}};try{await ce(oe)}catch(le){R(le.message),M.disabled=!1,M.textContent=E?"Simpan Perubahan":`Tambah ${a}`;return}h&&(z=await h(z,y));let ue=E?"PUT":"POST",Q=E?`${n}/${y.id}`:n,te=await F(Q,{method:ue,body:JSON.stringify(z)});te.ok?(j(E?`${a} berhasil diperbarui.`:`${a} berhasil ditambahkan.`),Y(),be(n),I()):(R(te.data?.error||"Gagal menyimpan data."),M.disabled=!1,M.textContent=E?"Simpan Perubahan":`Tambah ${a}`)}})}function N(y){pe(`Hapus ${a} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let E=await F(`${n}/${y.id}`,{method:"DELETE"});E.ok?(j(`${a} berhasil dihapus.`),be(n),I()):R(E.data?.error||"Gagal menghapus.")},`Hapus ${a}`)}return I(),I}H();H();var Ie=null,it=null;async function Te(e=!1){if(Ie&&!e)return console.log("Employees Raw (Cache Hit)",Ie.slice(0,5)),Ie;let t=await F(`/api/employees?limit=${Fe}&status=Aktif`);return Ie=(t.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",t.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",Ie.slice(0,5)),Ie}async function se(e=!1){let i=(await Te(e)).map(n=>({value:n.label,label:n.label}));return console.log("Employee Options",i.slice(0,5)),i}async function W(e=!1){return it&&!e||(it=((await F("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),it}function ee(e){let t={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!e||e==="-"||String(e).trim()===""?"":`<span class="badge ${t[e]||"badge-neutral"}">${e}</span>`}function kt(e){return e==null?'<span class="badge badge-neutral">-</span>':e<0?`<span class="badge badge-danger">Expired (${Math.abs(e)}h)</span>`:e<=14?`<span class="badge badge-danger">${e} hari</span>`:e<=30?`<span class="badge badge-warning">${e} hari</span>`:`<span class="badge badge-success">${e} hari</span>`}function Ae(e){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[e]||"badge-neutral"}">${e||"-"}</span>`}function xt(e){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[e]||"badge-neutral"}">${e||"-"}</span>`}function ye(e){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[e]||"badge-neutral"}">${e||"-"}</span>`}G();H();G();H();G();H();G();function rt(e){if(e.target.name==="report_date"||e.target.name==="completion_date"){let t=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),n=document.querySelector('input[name="day_count"]');if(t&&i&&n)if(t.value&&i.value){let r=new Date(t.value),d=new Date(i.value),o=Math.floor((d-r)/864e5);n.value=isNaN(o)?"":o}else n.value=""}}document.body.removeEventListener("input",rt);document.body.addEventListener("input",rt);document.body.removeEventListener("change",rt);document.body.addEventListener("change",rt);H();var Ee={};function Pe(e){if(Ee[e]){try{Ee[e].destroy()}catch{}delete Ee[e]}}function Ka(){Object.keys(Ee).forEach(Pe)}var fe=(e,t=0)=>{let i=Number(e);return isNaN(i)||e===null||e===void 0?t:i},Le=(e,t="\u2014")=>{if(e==null||e==="")return t;let i=String(e).trim();return i===""||i==="[object Object]"?t:i};function ot(e,t,i=900){if(!e)return;e._animFrame&&cancelAnimationFrame(e._animFrame);let n=Math.max(0,Math.round(fe(t))),r=e.textContent?e.textContent.replace(/[^\d]/g,""):"",d=r!==""&&parseInt(r)||0;if(n===d){e.textContent=n.toLocaleString("id-ID"),e._animFrame=null;return}let o=Date.now(),s=()=>{let a=Math.min((Date.now()-o)/i,1),l=1-Math.pow(1-a,3);e.textContent=Math.round(d+l*(n-d)).toLocaleString("id-ID"),a<1?e._animFrame=requestAnimationFrame(s):(e.textContent=n.toLocaleString("id-ID"),e._animFrame=null)};e._animFrame=requestAnimationFrame(s)}var za={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},ja=e=>{let t=Le(e,"\u2014");return`<span class="status-pill ${za[t]||"pill-neutral"}">${t}</span>`};var ke={family:"Inter",size:11},Se="#94A3B8",Be="#F1F5F9",wt=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],Ra=()=>window.innerWidth<768;function st(e={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:Ra()?"bottom":"top",labels:{font:ke,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ke,titleFont:{...ke,weight:"700"}}},scales:{x:{grid:{color:Be},ticks:{font:ke,color:Se,maxRotation:0}},y:{grid:{color:Be},ticks:{font:ke,color:Se},beginAtZero:!0}},...e}}var qa=()=>Array(6).fill(0).map(()=>`
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
  </div>`).join(""),Ha=()=>Array(6).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function zt(e=3){return Array(e).fill(0).map((t,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<e-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function ve(e,t,i=8e3){try{let n=new AbortController,r=setTimeout(()=>n.abort(),i),d=await F(e,{signal:n.signal}).catch(()=>null);if(clearTimeout(r),!d||!d.ok)return t;let o=d.data;return o?o.data!==void 0?o.data??t:o:t}catch{return t}}function Ja(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(n=>{let r=document.getElementById(n);r&&(r.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(n=>{let r=document.getElementById(n);if(r&&r.style.display==="none"){r.style.display="block";let d=r.parentElement;if(d&&!d.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent="Belum ada data",r.style.display="none",d.appendChild(o)}}});let e=document.getElementById("kpi-row");e&&e.querySelector(".skeleton")&&qt({});let t=document.getElementById("mini-stats-row");t&&t.querySelector(".skeleton")&&Ht({}),["table-contracts","table-issues"].forEach(n=>{let r=document.getElementById(n);r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Rt(e){Ka(),e._dashRefresh&&clearInterval(e._dashRefresh),e._skelTimeout&&clearTimeout(e._skelTimeout);let t=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});e.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">
      
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h2 style="margin:0; font-size:1.2rem; color:var(--text-1);">Dashboard Overview</h2>
        <button id="btn-dash-refresh" class="btn-primary" style="padding:6px 12px; font-size:0.8rem; display:flex; align-items:center; gap:6px; cursor:pointer; background:var(--primary); color:#fff; border:none; border-radius:6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
          Refresh Data
        </button>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${qa()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${Ha()}</div>

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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${zt(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${zt(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>jt(e)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async i=>{let n=i.target.value,r=document.getElementById("jadwal-year-label");r&&(r.textContent=n);let d=document.getElementById("skel-jadwal"),o=document.getElementById("chart-jadwal");d&&(d.style.display="block",d.style.position="absolute"),o&&(o.style.display="none");let s=await ve(`/api/dashboard/schedule-chart?year=${n}`,{},8e3);try{Jt(s)}catch(a){console.warn("ScheduleChart render:",a),xe("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let n=i.target.value,r=n?`/api/dashboard/inspection-bar?month=${n}`:"/api/dashboard/inspection-bar",d=document.getElementById("skel-insp"),o=document.getElementById("chart-insp");d&&(d.style.display="block",d.style.position="absolute"),o&&(o.style.display="none");let s=await ve(r,{},8e3);try{Ut(s)}catch(a){console.warn("InspBar render:",a),xe("skel-insp","chart-insp")}}),e._skelTimeout=setTimeout(()=>Ja(),5e3),await jt(e)}async function jt(e){e._skelTimeout&&(clearTimeout(e._skelTimeout),e._skelTimeout=null);let[t,i,n,r,d,o,s,a]=await Promise.all([ve("/api/dashboard/kpi",{},8e3),ve("/api/dashboard/issues-trend",{},8e3),ve("/api/dashboard/issues-summary",{},8e3),ve("/api/dashboard/stats",{},8e3),ve("/api/dashboard/calendar",[],8e3),ve("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),ve(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),ve("/api/dashboard/dropdown-aggregates",{},8e3)]),l=document.getElementById("filter-insp-month"),v=l?l.value:"",x=v?`/api/dashboard/inspection-bar?month=${v}`:"/api/dashboard/inspection-bar",h=await ve(x,{},8e3);window.dashboardAggregates=a||{};try{qt(t)}catch(m){console.warn("KPI render:",m)}try{Ht(t)}catch(m){console.warn("MiniStats render:",m)}try{Jt(s)}catch(m){console.warn("ScheduleChart render:",m),xe("skel-jadwal","chart-jadwal")}try{Ua(Array.isArray(n?.by_category)?n.by_category:[])}catch(m){console.warn("Donut render:",m),xe("skel-donut","chart-donut")}try{Ga(i)}catch(m){console.warn("Trend render:",m),xe("skel-trend","chart-trend")}try{Ut(h)}catch(m){console.warn("InspBar render:",m),xe("skel-insp","chart-insp")}try{let m=Array.isArray(r)?r:Array.isArray(r?.recent_issues)?r.recent_issues:[];Wa(m)}catch(m){console.warn("IssuesTable render:",m)}try{let m=Array.isArray(r?.expiring_contracts)?r.expiring_contracts:[];Va(o)}catch(m){console.warn("ContractsTable render:",m)}try{Qa(Array.isArray(d)?d:[])}catch(m){console.warn("Agenda render:",m)}try{Ya()}catch(m){console.warn("Quick Actions render:",m)}}function qt(e){let t=document.getElementById("kpi-row");if(!t)return;e=e||{};let i=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,n=window.dashboardAggregates||{},r=a=>{let l=window.dashboardAggregates?.schedule_quarter_details?.[a];return l?`Insp: ${l.insp} \u2022 GC: ${l.gc} \u2022 DC: ${l.dc} \u2022 Fog: ${l.fog}`:""},d=a=>{let l=window.dashboardAggregates?.schedule_quarter_details?.[a];return l?`Rincian ${a}: Inspeksi: ${l.insp} | General Cleaning: ${l.gc} | Deep Cleaning: ${l.dc} | Fogging: ${l.fog}`:`Total Kegiatan ${a}`},o=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
          <option value="Q1" ${i==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${i==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${i==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${i==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:n.schedule_by_quarter?.[i]??e.schedule?.current??0,href:`#/timeline?dash_filter=period_${i.toLowerCase()}`,color:"kpi-blue",sub:r(i),subId:"mini-jadwal-detail",tooltip:d(i)}];t.innerHTML=o.map(a=>{let l=a.val!==void 0?fe(a.val,0):fe(e[a.key]?.current,0);return a.dropdown?`
        <div class="kpi-card ${a.color}" ${a.id?`id="${a.id}"`:""} ${a.tooltip?`title="${a.tooltip}"`:""} data-href="${a.href}" tabindex="0" role="link" style="cursor:pointer;padding:10px 12px">
          <div style="display:flex; gap:10px; align-items:center;">
            <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
            <div style="flex:1;min-width:0;">
              <div style="display:flex; align-items:baseline; gap:3px;">
                <div class="kpi-value" data-target="${l}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${l}</div>
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
              <div class="kpi-value" data-target="${l}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${l}</div>
              ${a.dropdown?a.dropdown:""}
            </div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" ${a.subId?`id="${a.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${a.tooltip||""}">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),t.querySelectorAll(".kpi-card[data-href]").forEach(a=>{a.addEventListener("click",l=>{l.target.closest("select")||a.dataset.href&&(window.location.hash=a.dataset.href)}),a.addEventListener("keydown",l=>{l.key==="Enter"&&!l.target.closest("select")&&a.dataset.href&&(window.location.hash=a.dataset.href)})}),t.querySelectorAll(".kpi-value").forEach(a=>{ot(a,parseInt(a.dataset.target)||0)});let s=document.getElementById("dash-jadwal-period");if(s){let a=l=>{let x=(window.dashboardAggregates?.schedule_by_quarter||{})[l]||0,h=document.querySelector("#mini-jadwal .kpi-value");h&&(h.dataset.target=x,ot(h,x,400));let m=document.getElementById("mini-jadwal-detail");m&&(m.textContent=r(l),m.title=d(l));let S=document.getElementById("mini-jadwal");if(S){let k=`#/timeline?dash_filter=period_${l.toLowerCase()}`;S.tagName==="A"?S.href=k:S.dataset.href=k,S.title=d(l)}};s.addEventListener("change",l=>a(l.target.value))}}function Ht(e){let t=document.getElementById("mini-stats-row");if(!t)return;e=e||{};let i=new Date().getFullYear(),n=String(new Date().getMonth()+1).padStart(2,"0"),r=`${i}-${n}`,d=l=>`
    <select id="${l}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
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
  `,o=window.dashboardAggregates||{},s=[{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:d("dash-inspeksi-month"),val:o.inspeksi_by_month?.[r]??e.inspection_month?.current??0,href:`#/timeline?dash_filter=inspeksi&month=${r}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:d("dash-gcdc-month"),val:o.gcdc_by_month?.[r]??e.cleaning_month?.current??0,href:`#/timeline?dash_filter=gcdc&month=${r}`,color:"mini-green"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:d("dash-reliefer-month"),val:o.relievers_by_month?.[r]??e.reliever_completed?.current??0,href:`#/relievers?dash_filter=reliever&month=${r}`,color:"mini-teal"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:d("dash-fogging-month"),val:o.fogging_by_month?.[r]??e.fogging_month?.current??0,href:`#/reports/fogging?dash_filter=fogging&month=${r}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:e.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:e.branches?.current,href:"#/branches",color:"mini-teal"}];t.innerHTML=s.map(l=>l.dropdown?`
        <div class="mini-stat ${l.color}" style="cursor:pointer;" id="${l.id||""}" data-href="${l.href}" tabindex="0" role="link">
          <div class="mini-stat-icon">${l.icon}</div>
          <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="mini-stat-value" data-target="${fe(l.val)}">0</div>
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
            <div class="mini-stat-value" data-target="${fe(l.val)}">0</div>
            ${l.dropdown?l.dropdown:""}
          </div>
          <div class="mini-stat-text">${l.label}</div>
          ${l.extra||""}
        </div>
      </a>`).join(""),t.querySelectorAll(".mini-stat[data-href]").forEach(l=>{l.addEventListener("click",v=>{v.target.closest("select")||l.dataset.href&&(window.location.hash=l.dataset.href)}),l.addEventListener("keydown",v=>{v.key==="Enter"&&!v.target.closest("select")&&l.dataset.href&&(window.location.hash=l.dataset.href)})}),t.querySelectorAll(".mini-stat-value").forEach(l=>ot(l,parseInt(l.dataset.target)||0,700));let a=(l,v,x,h)=>{let m=document.getElementById(l);if(m){let S=k=>{let c=(window.dashboardAggregates?.[x]||{})[k]||0,f=document.querySelector(`#${v} .mini-stat-value`);f&&(f.dataset.target=c,ot(f,c,400));let u=document.getElementById(v);if(u){let w=`${h}&month=${k}`;u.tagName==="A"?u.href=w:u.dataset.href=w}};m.addEventListener("change",k=>S(k.target.value))}};a("dash-reliefer-month","mini-reliefer","relievers_by_month","#/relievers?dash_filter=reliever"),a("dash-inspeksi-month","mini-inspeksi","inspeksi_by_month","#/timeline?dash_filter=inspeksi"),a("dash-gcdc-month","mini-gcdc","gcdc_by_month","#/timeline?dash_filter=gcdc"),a("dash-fogging-month","mini-fogging","fogging_by_month","#/reports/fogging?dash_filter=fogging")}function Ua(e){xe("skel-donut","chart-donut");let t=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!t||!i)return;Pe("donut");let n=(e||[]).filter(a=>fe(a.count)>0);if(!n.length){Ge(t,"Belum ada data permasalahan");return}let r=n.map(a=>`${Le(a.category,"Lainnya")}`),d=n.map(a=>fe(a.count)),o=d.reduce((a,l)=>a+l,0);i.innerHTML=n.map((a,l)=>{let v=wt[l%wt.length],x=o>0?Math.round(a.count/o*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${v}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${a.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${x}%)</span></div>
          <div class="donut-legend-label">${r[l]}</div>
        </div>
      </div>
    `}).join("");let s={id:"centerText",beforeDraw:function(a){let l=a.width,v=a.height,x=a.ctx;x.restore();let h=(v/80).toFixed(2);x.font="bold "+h+"em Inter",x.textBaseline="middle",x.fillStyle="#1E293B";let m=o.toString(),S=Math.round((l-x.measureText(m).width)/2),k=v/2;x.fillText(m,S,k-4),x.font="600 "+(h*.35).toFixed(2)+"em Inter",x.fillStyle="#64748B";let g="Total",c=Math.round((l-x.measureText(g).width)/2);x.fillText(g,c,k+10),x.save()}};Ee.donut=new Chart(t,{type:"doughnut",data:{labels:r,datasets:[{data:d,backgroundColor:wt,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ke,titleFont:{...ke,weight:"700"},callbacks:{label:a=>` ${a.label}: ${a.parsed} kasus`}}},cutout:"75%"},plugins:[s]})}function Ga(e){xe("skel-trend","chart-trend");let t=document.getElementById("chart-trend");if(!t)return;Pe("trend"),e=e||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],n=(e.labels||[]).map(o=>{if(!o||typeof o!="string")return"";try{let[s,a]=o.split("-");return(i[Number(a)-1]||a)+" "+String(s).slice(-2)}catch{return o}}),r=(e.open||[]).map(o=>fe(o)),d=(e.closed||[]).map(o=>fe(o));if(!n.length){Ge(t,"Belum ada data trend");return}Ee.trend=new Chart(t,{type:"line",data:{labels:n,datasets:[{label:"Open",data:r,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:d,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:st({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Se,maxRotation:45,autoSkip:!0}},y:{grid:{color:Be},ticks:{font:{family:"Inter",size:9},color:Se},beginAtZero:!0}}})})}function Jt(e){xe("skel-jadwal","chart-jadwal");let t=document.getElementById("chart-jadwal");if(!t)return;Pe("jadwal"),e=e||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(e).some(a=>Array.isArray(a)&&a.some(l=>l>0))){Ge(t,"Belum ada data jadwal");return}let r=e["Inspeksi Hygiene"]||Array(12).fill(0),d=e["General Cleaning"]||Array(12).fill(0),o=e["Deep Cleaning"]||Array(12).fill(0),s=e.Fogging||Array(12).fill(0);Ee.jadwal=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Inspeksi",data:r,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:d,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:o,backgroundColor:"#F59E0B"},{label:"Fogging",data:s,backgroundColor:"#EF4444"}]},options:st({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Se,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Be},ticks:{font:{family:"Inter",size:9},color:Se},min:0}}})})}function Ut(e){xe("skel-insp","chart-insp");let t=document.getElementById("chart-insp");if(!t)return;Pe("inspBar"),e=e||{};let i=e.labels||[],n=(e.fc||[]).map(d=>fe(d)),r=(e.spv||[]).map(d=>fe(d));if(!i.length){Ge(t,"Belum ada data inspeksi");return}Ee.inspBar=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:n,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:r,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:st({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ke,color:Se,maxRotation:45,minRotation:45}},y:{grid:{color:Be},ticks:{font:ke,color:Se},min:0,max:100}}})})}function Va(e){xe("skel-contract-mini","chart-contract-mini");let t=document.getElementById("chart-contract-mini");if(!t)return;Pe("contractMiniBar"),e=e||{};let i={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},n=(e.labels||[]).map(o=>{let s=o.split("-")[1];return i[s]||o}),r=(e.data||[]).map(o=>fe(o));if(!n.length){Ge(t,"Belum ada data");return}let d=t.getContext("2d");Ee.contractMiniBar=new Chart(t,{type:"bar",data:{labels:n,datasets:[{label:"Kontrak Habis",data:r,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:st({onClick:(o,s)=>{if(s&&s.length>0){let a=s[0].index,l=(e.labels||[])[a];l&&(window.location.hash="#/contracts?month_expiry="+l)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ke,color:Se,maxRotation:0,autoSkip:!1}},y:{grid:{color:Be,borderDash:[4,4],drawBorder:!1},ticks:{font:ke,color:Se,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function Wa(e){let t=document.getElementById("table-issues");if(!t)return;let i=(e||[]).slice(0,8);if(!i.length){t.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}t.innerHTML=`
    <div class="dash-list">
      ${i.map(n=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${ja(n.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Le(n.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Le(n.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function Qa(e){let t=document.getElementById("widget-agenda");if(!t)return;let i=new Date,n=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,d=(e||[]).filter(o=>(o.event_date||"").startsWith(n)).slice(0,10);if(!d.length){t.innerHTML="";return}t.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${d.map(o=>{let s="#3B82F6",a="#EFF6FF",l="Agenda",v=(o.title||"").toLowerCase();return v.includes("inspeksi")?(s="#10B981",a="#ECFDF5",l="Inspeksi"):v.includes("cleaning")||v.includes("gcdc")?(s="#3B82F6",a="#EFF6FF",l="Cleaning"):v.includes("reliefer")?(s="#F59E0B",a="#FFFBEB",l="Reliefer"):v.includes("fogging")&&(s="#8B5CF6",a="#F5F3FF",l="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(o.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${s};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Le(o.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Le(o.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${a};color:${s}">${l}</div>
        </div>
      `}).join("")}
    </div>
  `}function Ya(){let e=document.getElementById("quick-actions");if(!e)return;let t=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];e.innerHTML=t.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function xe(e,t){let i=document.getElementById(e),n=document.getElementById(t);if(i&&(i.style.display="none",i.style.position=""),n){n.style.display="block";let r=n.parentElement;if(r){let d=r.querySelector(".chart-empty");d&&d.remove()}}}function Ge(e,t="Belum ada data"){if(!e)return;e.style.display="none";let i=e.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent=t,i.appendChild(r)}}H();async function Gt(e){document.getElementById("app").innerHTML=`
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
  `;let t=document.getElementById("login-form"),i=document.getElementById("login-error"),n=document.getElementById("login-btn"),r=document.getElementById("toggle-password"),d=document.getElementById("login-password");r?.addEventListener("click",()=>{let o=d.type==="text";d.type=o?"password":"text",r.style.color=o?"":"var(--primary)"}),t?.addEventListener("submit",async o=>{o.preventDefault(),i.style.display="none";let s=t.username.value.trim(),a=t.password.value;if(!s||!a){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}n.querySelector(".btn-text").style.display="none",n.querySelector(".btn-spinner").style.display="",n.disabled=!0;try{let l=await F("/api/auth/login",{method:"POST",body:JSON.stringify({username:s,password:a})});l.ok&&l.data.success?(ft(l.data.data.token),qe(l.data.data.user),j("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=l.data.error||"Username atau password salah.",i.style.display="block",n.classList.add("shake"),setTimeout(()=>n.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{n.querySelector(".btn-text").style.display="",n.querySelector(".btn-spinner").style.display="none",n.disabled=!1}})}H();G();async function Xa(){return await W()}function lt(e){if(!e||e==="-"||String(e).trim()==="")return"-";let t=String(e).trim(),i=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],n=t.match(/^(\d{4})-(\d{1,2})/);if(n){let d=n[1],o=parseInt(n[2],10)-1;if(o>=0&&o<12)return`${i[o]} ${d}`}let r=t.split(/[\/\-\.]/);if(r.length===2&&r[1].length===4){let d=parseInt(r[0],10)-1;if(d>=0&&d<12)return`${i[d]} ${r[1]}`}if(r.length===3&&r[2].length===4){let d=parseInt(r[1],10)-1;if(d>=0&&d<12)return`${i[d]} ${r[2]}`}for(let d=0;d<12;d++)if(t.toLowerCase().includes(i[d].toLowerCase()))return t;return t}function Za(e,t){let i=String(e.status||"").toLowerCase();return t==="active"?i==="aktif":t==="reliefer"?e.division==="FC - RELIEFER"&&i==="aktif":!1}async function Vt(e,t){let i=await Xa(),n=t?t.get("dash_filter"):null;q({container:e,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",enableMobileFilterSheet:!0,itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:r=>n?r.filter(d=>Za(d,n)):r,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:r=>Ae(r)},{key:"phone",label:"No. HP",render:r=>r?`<a href="tel:${r}">${r}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:r=>window.formatDate(r)},{key:"status",label:"Status",render:r=>ee(r)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:r=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:r?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:r?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:i,value:r?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:r?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:r?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:r?.status||""}]},{type:"row",fields:[{name:"target_pindah_os",label:"Target Pindah OS (Bulan)",type:"month",value:r?.target_pindah_os?String(r.target_pindah_os).slice(0,7):""},{name:"target_selesai_os",label:"Target Selesai OS (Bulan)",type:"month",value:r?.target_selesai_os?String(r.target_selesai_os).slice(0,7):""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let r=await F(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let d=r.data.data.map(o=>({"Nama Lengkap":o.full_name,Cabang:o.branch_name||"",Divisi:o.division||"","No. HP":o.phone||"","Tgl Masuk":o.join_date||"",Status:o.status||"","Target Pindah OS":lt(o.target_pindah_os)==="-"?"":lt(o.target_pindah_os),"Target Selesai OS":lt(o.target_selesai_os)==="-"?"":lt(o.target_selesai_os)}));K(d,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif","Target Pindah OS":"September 2026","Target Selesai OS":"Oktober 2026"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif","Target Pindah OS":"November 2026","Target Selesai OS":"Desember 2026"}],"Template_Import_Karyawan")},onImport:async r=>{let d=a=>{if(!a)return null;let l=String(a||"").toLowerCase(),v=i.find(x=>String(x.label||"").toLowerCase()===l);return v?v.value:null},o=r.map(a=>({full_name:String(a["Nama Lengkap"]||"").trim(),branch_id:d(String(a.Cabang||"").trim()),division:String(a.Divisi||a["Div / Bagian"]||"").trim()||"FACILITY CARE",phone:String(a["No. HP"]||a["No. Hp"]||"").trim(),join_date:String(a["Tgl Masuk"]||a["Tanggal Masuk"]||"").trim(),status:String(a.Status||"").trim(),target_pindah_os:String(a["Target Pindah OS"]||"").trim(),target_selesai_os:String(a["Target Selesai OS"]||"").trim(),notes:String(a.Catatan||"").trim()})).filter(a=>a.full_name),s=await F("/api/import/employees",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}H();G();var _t=[],Wt=[];async function en(){_t=await W(),Wt=await Te()}var St=async e=>{let t=[],i=1;for(;;){let r=await(await Promise.resolve().then(()=>(H(),De))).apiFetch(`${e}${e.includes("?")?"&":"?"}limit=100&page=${i}`);if(!r.ok)break;let d=r.data?.data||r.data||[],o=Array.isArray(d)?d:[];if(t=t.concat(o),o.length<100||r.data?.pagination&&i>=r.data.pagination.pages)break;i++}return t};async function dt(e,t){await en(),q({container:e,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",defaultFilters:{},onDataLoaded:n=>n,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:n=>Ae(n)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:n=>window.formatDate(n)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:n=>!n||String(n).startsWith("2099")?"Tetap / PKWTT":window.formatDate(n)},{key:"days_remaining",label:"Sisa Kontrak",render:(n,r)=>r.end_date&&String(r.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':kt(n)},{key:"status",label:"Status",render:n=>ee(n)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:_t},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:n=>(n.end_date||(n.end_date="2099-12-31"),n),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let n=document.createElement("button");n.id="btn-find-missing",n.className="btn btn-ghost",n.innerHTML="\u{1F50D} Cek Selisih Karyawan",n.style.marginLeft="8px",n.style.color="#EF4444",n.style.border="1px solid currentColor",n.onclick=async()=>{n.innerHTML="\u231B Mencari...",n.disabled=!0;try{let[d,o]=await Promise.all([St("/api/employees?status=Aktif"),St("/api/contracts")]);if(d.length>0){let s=o.filter(x=>x.status==="Aktif"),a=new Set(s.map(x=>x.employee_id)),l=d.filter(x=>!a.has(x.id)),v=`<p style="margin-bottom:12px">Data yang terbaca: <b>${d.length}</b> Karyawan Aktif, dan <b>${s.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${l.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;l.forEach(x=>{let h=o.filter(S=>S.employee_id===x.id),m='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(h.length>0){let S=h[0];m=`Pernah ada kontrak (Status: <b style="color:#EF4444">${S.status}</b>, Selesai: ${window.formatDate(S.end_date)})`}v+=`<li style="margin-bottom:8px"><b>${x.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${x.branch_name||"-"} | ${m}</span></li>`}),v+="</ul>",Promise.resolve().then(()=>(he(),vt)).then(x=>x.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:v,cancelText:"Tutup"}))}}catch(d){console.error(d)}n.innerHTML="\u{1F50D} Cek Selisih Karyawan",n.disabled=!1};let r=document.querySelector(".page-actions");r&&r.appendChild(n)}},formFields:n=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:Wt,value:n?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:_t,value:n?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:n?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:n?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:n?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:n?.end_date&&!String(n.end_date).startsWith("2099")?n.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:n?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:n?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let n=await F(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let r=n.data.data.map(d=>({"Nama Lengkap":d.employee_name,Cabang:d.branch_name||"","Div / Bagian":d.division||"","Tanggal Mulai":d.start_date||"","Tanggal Selesai":d.end_date&&String(d.end_date).startsWith("2099")?"":d.end_date||"","Sisa Kontrak":d.end_date&&String(d.end_date).startsWith("2099")?"Tetap":d.days_remaining!==null&&d.days_remaining!==void 0?`${d.days_remaining} Hari`:"",Status:d.status||""}));K(r,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async n=>{let[r,d]=await Promise.all([F("/api/branches?limit=10000"),St("/api/employees")]),o=r.data?.data||[],s=d||[];console.log(`Total employee yang berhasil dimuat dari database : ${s.length}`),s.length>0&&(console.log("Contoh 5 employee pertama:"),s.slice(0,5).forEach((g,c)=>{console.log(`${c+1}. ID: ${g.id}, Name: ${g.full_name}, Status: ${g.status}`)}));let a=g=>{if(!g)return null;let c=String(g||"").replace(/\s+/g," ").toLowerCase().trim(),f=o.find(u=>String(u.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===c||String(u.code||"").replace(/\s+/g," ").toLowerCase().trim()===c||String(u.name||"").replace(/\s+/g," ").toLowerCase().trim()===c);return f?f.id:null},l=(g,c)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${c}`),console.log(`Nama dari Excel : "${g}"`),!g)return console.log("Alasan gagal mapping : Nama kosong"),null;let f=String(g||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${f}"`),console.log(`Jumlah employee di database : ${s.length}`);let u=s.find(w=>String(w.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===f);return u?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${u.id}`),u.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},v=g=>{if(!g)return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let c=String(g).trim();if(/^\d{4,5}(\.\d+)?$/.test(c)){let u=Math.floor(Number(c));if(u>2e4&&u<99999){let w=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);let f=c.split(/[\/\-\.]/);if(f.length===3){let[u,w,b]=f.map(C=>C.trim());if(u.length===4&&w.length<=2&&b.length<=2)return`${u}-${w.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&w.length<=2&&u.length<=2)return`${b}-${w.padStart(2,"0")}-${u.padStart(2,"0")}`}return c},x=n.map((g,c)=>{let f=c+2,u=String(g["Nama Lengkap"]||"").trim(),w=g["Tanggal Mulai"],b=v(w);if(!b){let p=n.__worksheet,T=n.__headers||[],D=T.indexOf("Tanggal Mulai"),A="N/A",O="N/A",P="N/A";if(D!==-1&&p&&window.XLSX){let B=window.XLSX.utils.encode_cell({c:D,r:f-1});P=B;let L=p[B];L?(A=L.t||"undefined",O=L.w||"undefined"):A="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let I="Unknown";w==null||w===""?I="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":w instanceof Date&&isNaN(w.getTime())?I="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":I="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${f}`),console.log(`Employee Name : ${u}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${D})`),console.log(`Raw Cell Value : "${w}"`),console.log(`JavaScript Type : ${typeof w}`),console.log(`SheetJS Cell Type : ${A}`),console.log(`SheetJS Formatted Value : "${O}"`),console.log(`Value After Trim : "${String(w||"").trim()}"`),console.log(`Value After Date Parser : "${b}"`),console.log(`Is Empty : ${!w}`),console.log(`Is Invalid Date : ${w instanceof Date?isNaN(w.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${I}`),console.log(`Workbook Sheet : ${p?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${P}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(g,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(T)),console.log(`==========================
`)}let C=l(u,f),_=null;return C?b||(_="Tanggal Mulai kosong atau tidak berformat tanggal"):_="Karyawan tidak ditemukan di Database",{isValid:!!(C&&b),invalidReason:_,rowNum:f,data:{employee_id:C,branch_id:a(String(g.Cabang||"").trim()),division:String(g["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:b,end_date:v(g["Tanggal Selesai"])||"2099-12-31",status:String(g.Status||"").trim(),_rawName:u}}}),h=[],m=[];if(x.forEach(g=>{g.isValid?h.push(g.data):m.push({rowNum:g.rowNum,name:g.data._rawName,reason:g.invalidReason})}),console.log(`Split Validation - Valid: ${h.length}, Invalid: ${m.length}`),h.length===0){let g=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${n.length}
Valid: 0
Invalid: ${m.length}

Daftar Kegagalan (Contoh):
`;m.slice(0,10).forEach(c=>{g+=`- Row ${c.rowNum} | Nama: ${c.name} | Alasan: ${c.reason}
`}),m.length>10&&(g+=`- ... dan ${m.length-10} lainnya.
`),alert(g);return}let S=await F("/api/contracts/import",{method:"POST",body:JSON.stringify(h)}),k=`IMPORT SUMMARY
======================
`;k+=`Total Baris Excel : ${n.length}
`,k+=`Baris Valid       : ${h.length}
`,k+=`Baris Invalid     : ${m.length}

`,S&&S.data&&S.data.metrics?(k+=`Berhasil INSERT   : ${S.data.metrics.inserted}
`,k+=`Berhasil UPDATE   : ${S.data.metrics.updated}
`):k+=`Berhasil diproses : ${h.length}
`,m.length>0&&(k+=`
DAFTAR DATA DILEWATI:
`,m.forEach(g=>{k+=`- Row ${g.rowNum} | ${g.name} | ${g.reason}
`})),alert(k),typeof dt=="function"&&dt()}}})}H();G();var Ct=[],Ve=[];function tn(e){if(!Array.isArray(e))return"Q3";let t=["Q4","Q3","Q2","Q1"];for(let i of t)if(e.some(n=>n.period===i))return i;return"Q3"}async function Qt(e,t){Ct=await W();let i=await se();Ve=["Berlin Ariansyah","Ade Surahman"];let n=g=>g&&!Ve.find(c=>String(typeof c=="object"?c.value:c).toLowerCase()===String(g).toLowerCase())?[...Ve,g]:Ve,r=await F(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),d=g=>{if(!g||g==="-"||String(g).trim()==="")return"";let c=String(g).split("-");return c.length===3&&c[0].length===4?`${c[2]}-${c[1]}-${c[0]}`:g},o=r.data?.data||[],s=tn(o),a=t?t.get("dash_filter"):null,l=new Date,v=`${l.getFullYear()}-${String(l.getMonth()+1).padStart(2,"0")}`,x={},h=t&&t.get("month")?t.get("month"):null;a==="inspeksi"?x={status:"Done",activity_type:"Inspeksi Hygiene",month:h}:a==="gcdc"?x={status:"Done",activity_type:"GCDC",month:h}:a&&a.startsWith("period_")&&(x={period:a.replace("period_","").toUpperCase()});let m=new Date().getFullYear(),k=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((g,c)=>{let f=String(c+1).padStart(2,"0");return{value:`${m}-${f}`,label:`${g} ${m}`}});q({container:e,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:x,onDataLoaded:g=>g.sort((c,f)=>{let u=c.opening_date?new Date(c.opening_date).getTime():0;return(f.opening_date?new Date(f.opening_date).getTime():0)-u}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:g=>xt(g)},{key:"period",label:"Periode",render:g=>ye(g)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:g=>d(g)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:g=>d(g)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:g=>d(g)},{key:"status",label:"Status",render:g=>ee(g)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:Ct},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:k},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Ve}],formFields:g=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Ct,value:g?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:g?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:g?.period},{name:"pic",label:"PIC",type:"combobox",options:n(g?.pic),value:g?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:g?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:g?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:g?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:g?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:g?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let g=await F(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(g.ok){let c=g.data.data.map(f=>({Cabang:f.branch_name||"",Kegiatan:f.activity_type||"",Periode:f.period||"",PIC:f.pic||"","Tgl Opening":f.opening_date||"","Tgl Target":f.target_date||"","Tgl Selesai":f.completion_date||"",Status:f.status||""}));K(c,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async g=>{let f=(await F("/api/branches?all=1")).data?.data||[],u=_=>{if(!_)return null;let p=String(_||"").toLowerCase().trim(),T=p.match(/^(\d{3})/);if(T){let A=f.find(O=>String(O.code||"")===T[1]);if(A)return A.id}let D=f.find(A=>String(A.full_name||"").toLowerCase()===p||String(A.code||"").toLowerCase()===p||String(A.name||"").toLowerCase()===p);return D?D.id:null},w=_=>{if(_==null||_==="")return"";if(_ instanceof Date&&!isNaN(_.getTime()))return _.toISOString().slice(0,10);let p=String(_).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let D=Number(p);if(D>2e4&&D<99999){let A=new Date(Date.UTC(1899,11,30)+D*864e5);return isNaN(A.getTime())?"":A.toISOString().slice(0,10)}}let T=p.split(/[\/\-\.]/);if(T.length===3){let[D,A,O]=T.map(P=>P.trim());if(D.length===4&&A.length<=2&&O.length<=2)return`${D}-${A.padStart(2,"0")}-${O.padStart(2,"0")}`;if(O.length===4&&A.length<=2&&D.length<=2)return`${O}-${A.padStart(2,"0")}-${D.padStart(2,"0")}`}return p},b=g.map(_=>({branch_id:u(String(_.Cabang||"").trim()),branch_name:String(_.Cabang||"").trim(),activity_type:String(_.Kegiatan||"").trim(),period:String(_.Periode||"").trim(),pic:String(_.PIC||_.Pic||"").trim(),opening_date:w(_["Tgl Opening"]||_["Tanggal Opening"]||_["Tgl Openir"]),target_date:w(_["Tgl Target"]||_["Tanggal Target"]),completion_date:w(_["Tgl Selesai"]||_["Tanggal Selesai"]),status:String(_.Status||"").trim(),notes:String(_.Catatan||_.Keterangan||"").trim()})).filter(_=>_.activity_type&&_.period),C=await F("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:b,onDuplicate:"update"})});if(!C.ok)throw new Error(C.data?.error||"Import gagal");return be("schedule"),C.data}}})}H();G();var Et=[],ct=[];function pt(e){if(e.target.name==="report_date"||e.target.name==="completion_date"){let t=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),n=document.querySelector('input[name="day_count"]');if(t&&i&&n)if(t.value&&i.value){let r=new Date(t.value),d=new Date(i.value),o=Math.floor((d-r)/864e5);n.value=isNaN(o)?"":o}else n.value=""}}document.body.removeEventListener("input",pt);document.body.addEventListener("input",pt);document.body.removeEventListener("change",pt);document.body.addEventListener("change",pt);function an(e,t){let i=String(e.status||"").toLowerCase();return t==="open"?i==="open":!1}async function Yt(e,t){let i=t?t.get("dash_filter"):null;Et=await W(),ct=await se();let n=a=>a&&!ct.find(l=>l.value===a)?[...ct,{value:a,label:a}]:ct,r=new Date().getFullYear(),d=["2025","2026","2027","2028","2029","2030"],s=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((a,l)=>{let v=String(l+1).padStart(2,"0");return{value:`${r}-${v}`,label:`${a} ${r}`}});q({container:e,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:a=>i?a.filter(l=>an(l,i)):a,columns:[{key:"report_date",label:"Tanggal Info",nowrap:!0,render:a=>window.formatDate(a)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:a=>`<span class="badge badge-secondary">${a}</span>`},{key:"source",label:"Sumber Laporan"},{key:"complaint",label:"Keluhan",render:a=>`<span title="${a}">${a?.length>50?a.slice(0,50)+"...":a}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:a=>`<span title="${a||""}">${a?.length>40?a.slice(0,40)+"...":a||"-"}</span>`},{key:"status",label:"Status",render:a=>ee(a)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:a=>window.formatDate(a)},{key:"day_count",label:"Day",render:a=>a??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:Et},{type:"select",name:"month",label:"Bulan",options:s},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:d}],formFields:a=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:a?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Et,value:a?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:a?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:a?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:a?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:n(a?.employee_name),value:a?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:n(a?.fc_specialist),value:a?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:a?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:a?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:a?.completion_date},{name:"day_count",label:"Day",type:"number",value:a?.day_count,readonly:!0}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let a=await F(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let l=a.data.data.map(v=>({"Tanggal Info":v.report_date||"",Cabang:v.branch_name||"",Kategori:v.category||"","Sumber Laporan":v.source||"",Keluhan:v.complaint||"","Nama FC":v.employee_name||"","FC Spesialis":v.fc_specialist||"",Solusi:v.solution||"",Status:v.status||"","Tanggal Selesai":v.completion_date||"",Day:v.day_count!==null?v.day_count:""}));K(l,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan",Status:"Done","Tanggal Selesai":"2024-03-02",Day:1}],"Template_Import_Permasalahan")},onImport:async a=>{let v=(await F("/api/branches?all=1")).data?.data||[],x=S=>{if(!S)return null;let k=String(S||"").toLowerCase(),g=v.find(c=>String(c.full_name||"").toLowerCase()===k||String(c.code||"").toLowerCase()===k||String(c.name||"").toLowerCase()===k);return g?g.id:null},h=a.map(S=>({branch_id:x(String(S.Cabang||"").trim()),report_date:String(S["Tanggal Info"]||S.Tanggal||"").trim(),category:String(S.Kategori||"").trim(),source:String(S["Sumber Laporan"]||S.Sumber||"").trim(),complaint:String(S.Keluhan||"").trim(),employee_name:String(S["Nama FC"]||"").trim(),fc_specialist:String(S["FC Spesialis"]||"").trim(),solution:String(S.Solusi||"").trim(),completion_date:String(S["Tanggal Selesai"]||S["Tgl Selesai"]||"").trim(),status:String(S.Status||"").trim(),day_count:S.Day||S.Hari||null})).filter(S=>S.report_date&&S.complaint&&S.category),m=await F("/api/import/issues",{method:"POST",body:JSON.stringify({rows:h,onDuplicate:"update"})});if(!m.ok)throw new Error(m.data?.error||"Import gagal");return m.data}}})}H();var Me=[];function nn(e,t){let i=String(e.status||"").toLowerCase();return t==="pending"?i==="pending":!1}async function Xt(e,t){let i=t?t.get("dash_filter"):null;Me=await W();let n=await se(),r=["Ade","Berlin"],d=s=>s&&!n.find(a=>a.value===s)?[...n,{value:s,label:s}]:n,o=s=>s&&!r.find(a=>(typeof a=="object"?a.value:a)===s)?[...r,s]:r;q({container:e,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:s=>i?s.filter(a=>nn(a,i)):s,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:s=>`<span title="${s||""}">${s?.length>50?s.slice(0,50)+"\u2026":s||"-"}</span>`},{key:"solution",label:"Solusi",render:s=>`<span title="${s||""}">${s?.length>40?s.slice(0,40)+"\u2026":s||"-"}</span>`},{key:"status",label:"Status",render:s=>ee(s)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:s=>window.formatDate(s)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:s=>s?`<a href="${s}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:Me},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async s=>{let a=new URLSearchParams(s||{}).toString(),l=await F(`/api/one-on-one?limit=10000&${a}`);if(l.ok){let v=l.data.data.map(h=>({Tanggal:h.meeting_date||"",Cabang:h.branch_name||"","Nama Karyawan":h.employee_name||"",PIC:h.pic||"",Masalah:h.problem||"",Solusi:h.solution||"",Status:h.status||"","Tgl Selesai":h.completion_date||"",Dokumen:h.document_link||""})),{downloadExcel:x}=await Promise.resolve().then(()=>(G(),me));x(v,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let s=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:a}=await Promise.resolve().then(()=>(G(),me));a(s,"Template_Import_OneOnOne")},onImport:async s=>{let a=h=>{if(!h)return null;let m=String(h||"").toLowerCase(),S=Me.find(k=>String(k.label||"").toLowerCase()===m);return S?S.value:null},l=h=>{if(!h)return"";if(h instanceof Date&&!isNaN(h.getTime()))return h.toISOString().slice(0,10);let m=String(h).trim();if(/^\d{4,5}$/.test(m)){let k=Number(m);if(k>2e4&&k<99999){let g=new Date(Date.UTC(1899,11,30)+k*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);let S=m.split(/[\/\-\.]/);if(S.length===3){let[k,g,c]=S.map(f=>f.trim());if(k.length===4&&g.length<=2&&c.length<=2)return`${k}-${g.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&g.length<=2&&k.length<=2)return`${c}-${g.padStart(2,"0")}-${k.padStart(2,"0")}`}return m},v=s.map(h=>({meeting_date:l(h.Tanggal),employee_name:String(h["Nama Karyawan"]||"").trim(),branch_id:a(String(h.Cabang||"").trim()),pic:String(h.PIC||"").trim(),problem:String(h.Masalah||"").trim(),solution:String(h.Solusi||"").trim(),status:String(h.Status||"").trim(),completion_date:l(h["Tgl Selesai"]),document_link:String(h.Dokumen||"").trim()})).filter(h=>h.meeting_date&&h.employee_name&&h.branch_id),x=await F("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:v,onDuplicate:"update"})});if(!x.ok)throw new Error(x.data?.error||"Import gagal");return x.data}},formFields:s=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:s?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:s?.branch_id&&!Me.find(a=>a.value==s.branch_id)?[...Me,{value:s.branch_id,label:s.branch_name||s.branch_id}]:Me,createApi:{path:"/api/branches",field:"full_name"},value:s?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:d(s?.employee_name),value:s?.employee_name},{name:"pic",label:"PIC",type:"select",options:o(s?.pic),createApi:{path:"/api/pic",field:"name"},value:s?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:s?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:s?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:s?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:s?.document_link}]})}H();async function Zt(e){let t=await W(),i=await se(),n=["Ade Surahman","Berlin Ariansyah","Mizwar","Fajar","Ade","Berlin"],r=a=>a&&!i.find(l=>l.value===a)?[...i,{value:a,label:a}]:i,d=a=>a&&!n.find(l=>(typeof l=="object"?l.value:l)===a)?[...n,a]:n,o=Array.from({length:5},(a,l)=>String(new Date().getFullYear()-l));q({container:e,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a||"-"}catch{return a||"-"}}},{key:"score",label:"Nilai",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer / peserta..."},{type:"select",name:"batch",label:"Batch",options:["Batch 1","Batch 2","Batch 3","Batch 4","Batch 5"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"trainer",label:"Trainer",options:["Berlin Ariansyah"]},{type:"select",name:"year",label:"Tahun",options:o}],exportOptions:{moduleName:"training",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),v=await F(`/api/training?limit=10000&${l}`);if(v.ok){let x=v.data.data.map(m=>{let S=m.participants||"";try{let k=JSON.parse(S);S=Array.isArray(k)?k.join(", "):S}catch{}return{Tanggal:m.training_date||"",Batch:m.batch||"",Materi:m.subject||"",Cabang:m.branch_name||"",Trainer:m.trainer||"",Peserta:S,Nilai:m.score!==null&&m.score!==void 0?m.score:"",Dokumen:m.document_link||""}}),{downloadExcel:h}=await Promise.resolve().then(()=>(G(),me));h(x,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let a=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:l}=await Promise.resolve().then(()=>(G(),me));l(a,"Template_Import_Training")},onImport:async a=>{let l=m=>{if(!m)return null;let S=String(m||"").toLowerCase(),k=t.find(g=>String(g.label||"").toLowerCase()===S);return k?k.value:null},v=m=>{if(!m)return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let S=String(m).trim();if(/^\d{4,5}$/.test(S)){let g=Number(S);if(g>2e4&&g<99999){let c=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(c.getTime())?"":c.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(S))return S.slice(0,10);let k=S.split(/[\/\-\.]/);if(k.length===3){let[g,c,f]=k.map(u=>u.trim());if(g.length===4&&c.length<=2&&f.length<=2)return`${g}-${c.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&c.length<=2&&g.length<=2)return`${f}-${c.padStart(2,"0")}-${g.padStart(2,"0")}`}return S},x=a.map(m=>({training_date:v(m.Tanggal),batch:String(m.Batch||"").trim(),subject:String(m.Materi||"").trim(),branch_id:l(String(m.Cabang||"").trim()),trainer:String(m.Trainer||"").trim(),participants:String(m.Peserta||"").trim(),score:m.Nilai?Number(m.Nilai):null,document_link:String(m.Dokumen||"").trim()})).filter(m=>m.training_date&&m.subject&&m.branch_id),h=await F("/api/training/import",{method:"POST",body:JSON.stringify(x)});if(!h.ok)throw new Error(h.data?.error||"Import gagal")}},formFields:a=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:a?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:a?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:a?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:t,value:a?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:d(a?.trainer),value:a?.trainer}]},{name:"participants",label:"Peserta",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let l=JSON.parse(a?.participants);return Array.isArray(l)?l.join(", "):a?.participants||""}catch{return a?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:a?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:a?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],onBeforeSubmit:async a=>a})}H();he();G();function ea({container:e,title:t,icon:i,apiPath:n,columns:r,formFields:d,filterFields:o,defaultFilters:s={},itemLabel:a="Data",canCreate:l=!0,canEdit:v=!0,canDelete:x=!0,onBeforeSubmit:h,onAfterLoad:m,onDataLoaded:S,extraActions:k=[],initialSearch:g="",exportOptions:c=null,bulkDelete:f=!1,paginationMode:u="server"}){let w=ie();w&&w.role==="viewer"&&(l=!1,v=!1,x=!1,f=!1,c=null),w&&w.role==="editor_khusus"&&n!=="/api/issues"&&n!=="/api/overtime"&&(l=!1,v=!1,x=!1,f=!1,c=null),w&&w.role==="input_lembur"&&(n==="/api/overtime"?(x=!1,f=!1,c=null):(l=!1,v=!1,x=!1,f=!1,c=null));let b=1,C={...s};g&&(C.search=g);let _=new Set;e.innerHTML=`
    ${f?`
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
        ${l?`<button class="btn btn-primary" id="btn-create">+ Tambah ${a}</button>`:""}
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
        ${o.filter(y=>y.type==="search").map(y=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${y.placeholder||"Cari..."}" id="filter-search" value="${C.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${o.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let E=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select" name="${y.name}" id="filter-${y.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${E}</option>${(y.options||[]).map($=>`<option value="${typeof $=="object"?$.value:$}" ${C[y.name]===(typeof $=="object"?$.value:$)?"selected":""}>${typeof $=="object"?$.label:$}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(y=>y.type!=="search").map(y=>{if(y.type==="select"||y.type==="combobox"){let E=(y.label||"").startsWith("Pilih")?y.label:`Pilih ${y.label||""}`;return`<select class="filter-select filter-select-sheet" name="${y.name}-sheet" id="filter-sheet-${y.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${E}</option>${(y.options||[]).map($=>`<option value="${typeof $=="object"?$.value:$}" ${C[y.name]===(typeof $=="object"?$.value:$)?"selected":""}>${typeof $=="object"?$.label:$}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function p(){let y=document.getElementById("bulk-toolbar");if(!y)return;let E=document.getElementById("bulk-count"),$=document.getElementById("btn-bulk-delete"),U=document.getElementById("btn-bulk-cancel");E.textContent=`${_.size} item dipilih`,_.size>0?(y.style.display="flex",$.disabled=!1,U.disabled=!1):(y.style.display="none",$.disabled=!0,U.disabled=!0);let J=document.getElementById("select-all-checkbox");if(J){let Y=document.querySelectorAll(".row-checkbox");if(Y.length>0){let M=[...Y].every(oe=>oe.checked),z=[...Y].some(oe=>oe.checked);J.checked=M,J.indeterminate=z&&!M}else J.checked=!1,J.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{_.clear(),document.querySelectorAll(".row-checkbox").forEach(E=>E.checked=!1);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1),p()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(_.size===0)return;let y=[..._],E=document.createElement("div");E.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",E.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${y.length} ${a}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${y.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(E),E.querySelector("#bulk-cancel-btn").addEventListener("click",()=>E.remove()),E.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let $=E.querySelector("#bulk-confirm-btn");$.disabled=!0,$.textContent="Menghapus...";let U=await F(`${n}/bulk`,{method:"DELETE",body:JSON.stringify({ids:y})});E.remove(),U.ok?(j(`${y.length} ${a} berhasil dihapus.`),_.clear(),p(),be(n),I()):R(U.data?.error||"Gagal menghapus data.")})});let T=document.getElementById("filter-search"),D;T?.addEventListener("input",y=>{clearTimeout(D),D=setTimeout(()=>{C.search=y.target.value,b=1,_.clear(),p(),I()},400)}),o?.forEach(y=>{(y.type==="select"||y.type==="combobox")&&(document.getElementById(`filter-${y.name}`)?.addEventListener("change",E=>{C[y.name]=E.target.value;let $=document.getElementById(`filter-sheet-${y.name}`);$&&($.value=E.target.value),b=1,_.clear(),p(),I()}),document.getElementById(`filter-sheet-${y.name}`)?.addEventListener("change",E=>{C[y.name]=E.target.value;let $=document.getElementById(`filter-${y.name}`);$&&($.value=E.target.value),b=1,_.clear(),p(),I(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{C={},T&&(T.value=""),o?.forEach(y=>{let E=document.getElementById(`filter-${y.name}`);E&&(E.value="");let $=document.getElementById(`filter-sheet-${y.name}`);$&&($.value="")}),b=1,_.clear(),p(),I()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{C={},T&&(T.value=""),o?.forEach(y=>{let E=document.getElementById(`filter-${y.name}`);E&&(E.value="");let $=document.getElementById(`filter-sheet-${y.name}`);$&&($.value="")}),b=1,_.clear(),p(),I(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>L(null)),c&&document.addEventListener("click",function(y){let E=document.getElementById("aksi-menu-main"),$=document.getElementById("btn-aksi-main");E&&$&&!$.contains(y.target)&&!E.contains(y.target)&&E.classList.remove("show-aksi-menu")});let A=document.getElementById("btn-mobile-filter"),O=document.getElementById("filter-options-wrapper"),P=document.getElementById("btn-close-filter-sheet");if(A&&O&&(A.addEventListener("click",y=>{y.preventDefault(),O.classList.add("sheet-open")}),P&&P.addEventListener("click",y=>{y.preventDefault(),O.classList.remove("sheet-open")})),c){document.getElementById(`btn-export-${c.moduleName}`)?.addEventListener("click",async E=>{let $=E.target,U=$.innerHTML;$.innerHTML="\u23F3 Loading...",$.disabled=!0;try{await c.onExport()}catch{R("Gagal export data")}finally{$.innerHTML=U,$.disabled=!1}}),document.getElementById(`btn-template-${c.moduleName}`)?.addEventListener("click",()=>{c.onTemplate()});let y=document.getElementById(`input-import-${c.moduleName}`);y?.addEventListener("change",async E=>{let $=E.target.files[0];if(!$)return;let U=document.getElementById(`label-import-${c.moduleName}`),J=U?U.querySelector(".import-text"):null,Y=J?J.innerText:"";J&&(J.innerText="\u231B Memproses..."),U&&(U.style.pointerEvents="none"),y.disabled=!0;try{let M=await Ue($);if(M.length===0)throw new Error("File kosong atau format salah");await c.onImport(M),j("Import berhasil!"),be(n),I()}catch(M){R(M.message||"Gagal import data")}finally{J&&(J.innerText=Y),U&&(U.style.pointerEvents="auto"),y.disabled=!1,y.value=""}})}async function I(){p();let y=document.getElementById("table-container");if(!y)return;y.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let E=u==="client",$=E?1:b,U=E?Fe:20,J=new URLSearchParams({page:$,limit:U,...Object.fromEntries(Object.entries(C).filter(([,Q])=>Q))}),Y=await F(`${n}?${J}`);if(!Y.ok){y.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${Y.data?.error||"Error"}</p></div>`;return}let M=Y.data?.data||Y.data||[],z=Y.data?.pagination,oe=M.length;if(E){M=S(M);let Q=M.length,te=20,le=Math.ceil(Q/te);b>le&&le>0&&(b=le);let V=(b-1)*te,ge=b*te;M=M.slice(V,ge),z={page:b,limit:te,total:Q,pages:le}}!1,m&&m(M);let ce=Ze({columns:r,data:M,onEdit:v?Q=>L(Q):null,actions:k.map(Q=>({...Q,handler:te=>Q.handler(te,I)})),emptyText:`Tidak ada ${String(a||"").toLowerCase()}`,bulkSelect:f?{selectedIds:_,onToggle:p}:null});y.innerHTML="",y.appendChild(ce);let ue=document.getElementById("pagination-container");if(ue&&(ue.innerHTML="",z&&z.pages>1)){let Q=et({page:z.page,pages:z.pages,total:z.total,limit:z.limit,onPage:te=>{b=te,I()}});Q&&ue.appendChild(Q)}}function B(y){let E=typeof d=="function"?d(y):d;return Je(E)}function L(y){let E=!!y,$=document.createElement("form");if($.noValidate=!0,$.innerHTML=B(y),E){let J=typeof d=="function"?d(y):d;at($,y)}let{close:U}=ae({title:E?`Edit ${a}`:`Tambah ${a}`,content:$,size:"lg",confirmText:E?"Simpan Perubahan":`Tambah ${a}`,onConfirm:async(J,Y)=>{if(!$.reportValidity())return;let M=J.querySelector(".modal-confirm");M.disabled=!0,M.textContent="Menyimpan...";let z=tt($),oe=typeof d=="function"?d(y):d,ce=async le=>{for(let V of le)if(V.type==="row")await ce(V.fields);else if(V.type==="combobox"&&z[V.name]){let ge=z[V.name],_e=(V.options||[]).find(ne=>{let de=String(typeof ne=="object"?ne.value:ne),gt=String(typeof ne=="object"?ne.label:ne);return de===ge||gt===ge});if(_e)z[V.name]=typeof _e=="object"?_e.value:_e;else if(V.createApi){let ne={};ne[V.createApi.field]=ge,V.createApi.extra&&Object.assign(ne,V.createApi.extra);let de=await F(V.createApi.path,{method:"POST",body:JSON.stringify(ne)});if(de.ok&&de.data?.id)z[V.name]=de.data.id;else if(de.ok&&!de.data?.id)z[V.name]=ge;else throw new Error(`Gagal membuat master data: ${de.data?.error||"Unknown error"}`)}}};try{await ce(oe)}catch(le){R(le.message),M.disabled=!1,M.textContent=E?"Simpan Perubahan":`Tambah ${a}`;return}h&&(z=await h(z,y));let ue=E?"PUT":"POST",Q=E?`${n}/${y.id}`:n,te=await F(Q,{method:ue,body:JSON.stringify(z)});te.ok?(j(E?`${a} berhasil diperbarui.`:`${a} berhasil ditambahkan.`),Y(),be(n),I()):(R(te.data?.error||"Gagal menyimpan data."),M.disabled=!1,M.textContent=E?"Simpan Perubahan":`Tambah ${a}`)}})}function N(y){pe(`Hapus ${a} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let E=await F(`${n}/${y.id}`,{method:"DELETE"});E.ok?(j(`${a} berhasil dihapus.`),be(n),I()):R(E.data?.error||"Gagal menghapus.")},`Hapus ${a}`)}return I(),I}H();G();async function ta(e,t){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await W(),n=await se(),r=t?t.get("dash_filter"):null,d={};if(r==="reliever"){let h=new Date,m=`${h.getFullYear()}-${String(h.getMonth()+1).padStart(2,"0")}`;d={status:"Done",month:t&&t.get("month")?t.get("month"):m}}console.log("RAW",await Te()),console.log("OPTIONS",n);let o=h=>h&&!n.find(m=>m.value===h)?[...n,{value:h,label:h}]:n,s=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],a=h=>h&&!s.includes(h)?[...s,h]:s,l=new Date().getFullYear(),x=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((h,m)=>{let S=String(m+1).padStart(2,"0");return{value:`${l}-${S}`,label:`${h} ${l}`}});ea({container:e,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:d,onDataLoaded:h=>h.sort((m,S)=>{let k=m.backup_date?new Date(m.backup_date).getTime():0;return(S.backup_date?new Date(S.backup_date).getTime():0)-k}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:h=>ye(h)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:h=>window.formatDate(h)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:h=>window.formatDate(h)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:h=>h?`<span class="badge badge-info">${h}</span>`:"-"},{key:"status",label:"Status",render:h=>ee(h)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:s},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"month",label:"Bulan",options:x},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:h=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:h?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:h?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:o(h?.original_fc_name),value:h?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:a(h?.reliever_name),value:h?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:h?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:h?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:h?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:h?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:h?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let h=await F(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(h.ok){let m=h.data.data.map(S=>({Cabang:S.branch_name||"","Nama Facility care":S.original_fc_name||"",Periode:S.period||"",Relifer:S.reliever_name||"","Tanggal Back Up":S.backup_date||"","Tanggal Selesai":S.completion_date||"",Keterangan:S.reason||"",Shift:S.shift||"",Status:S.status||""}));m.length===0&&m.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),K(m,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async h=>{let S=(await F("/api/branches?all=1")).data?.data||[],k=f=>{if(!f)return null;let u=String(f||"").toLowerCase(),w=S.find(b=>String(b.full_name||"").toLowerCase()===u||String(b.code||"").toLowerCase()===u||String(b.name||"").toLowerCase()===u);return w?w.id:null},g=h.map(f=>({branch_name:String(f.Cabang||"").trim(),backup_date:String(f["Tanggal Back Up"]||f["Tanggal Backup"]||"").trim(),original_fc_name:String(f["Nama Facility care"]||f["FC Digantikan"]||"").trim(),reliever_name:String(f.Relifer||f.Reliefer||"").trim(),period:String(f.Periode||"").trim(),reason:String(f.Keterangan||"").trim(),shift:String(f.Shift||"").trim(),completion_date:String(f["Tanggal Selesai"]||"").trim(),status:String(f.Status||"").trim()})).filter(f=>f.reliever_name&&f.backup_date),c=await F("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:g,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}}})}H();G();async function aa(e){let t=await W(),i=Array.from({length:4},(n,r)=>String(new Date().getFullYear()-r));q({container:e,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:n=>ye(n)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"fc_score",label:"Point FC",render:n=>n!=null?`<strong class="${n>=80?"text-success":n>=60?"text-warning":"text-danger"}">${n}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"status",label:"Status",render:n=>ee(n)},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:n?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:n?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:n?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:n?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:n?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async n=>{let r=new URLSearchParams(n||{}).toString(),d=await F(`/api/reports/inspection?limit=10000&${r}`);if(d.ok){let o=d.data.data.map(s=>({Cabang:s.branch_name||"",Periode:s.period||"",Tanggal:s.inspection_date||"","Point FC":s.fc_score!==null&&s.fc_score!==void 0?s.fc_score:"","Point SPV":s.spv_score!==null&&s.spv_score!==void 0?s.spv_score:"",Status:s.status||"","Link Dokumen":s.document_link||""}));K(o,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async n=>{let r=a=>{if(!a)return null;let l=String(a||"").toLowerCase(),v=t.find(x=>String(x.label||"").toLowerCase()===l);return v?v.value:null},d=a=>{if(a==null||a==="")return"";if(a instanceof Date&&!isNaN(a.getTime()))return a.toISOString().slice(0,10);let l=String(a).trim();if(l===""||l==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);if(/^\d{4,5}$/.test(l)){let x=Number(l);if(x>2e4&&x<99999){let h=new Date(Date.UTC(1899,11,30)+x*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}let v=l.split(/[\/\-\.]/);if(v.length===3){let[x,h,m]=v.map(S=>S.trim());if(x.length===4&&h.length<=2&&m.length<=2)return`${x}-${h.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&h.length<=2&&x.length<=2)return`${m}-${h.padStart(2,"0")}-${x.padStart(2,"0")}`}return l},o=n.map(a=>({branch_id:r(String(a.Cabang||"").trim()),period:String(a.Periode||"").trim(),inspection_date:d(a.Tanggal),fc_score:a["Point FC"]!==void 0&&a["Point FC"]!==""?Number(a["Point FC"]):null,spv_score:a["Point SPV"]!==void 0&&a["Point SPV"]!==""?Number(a["Point SPV"]):null,status:String(a.Status||"").trim(),document_link:String(a["Link Dokumen"]||"").trim(),notes:String(a.Catatan||a.Keterangan||"").trim()})).filter(a=>a.branch_id&&a.period&&a.inspection_date),s=await F("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}H();G();async function na(e){let t=await W(),i=Array.from({length:4},(n,r)=>String(new Date().getFullYear()-r));q({container:e,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:n=>`<span class="badge ${n==="Deep Cleaning"?"badge-purple":"badge-success"}">${n}</span>`},{key:"period",label:"Periode",render:n=>ye(n)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"status",label:"Status",render:n=>ee(n)},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t,value:n?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:n?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:n?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:n?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:n?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async n=>{let r=new URLSearchParams(n||{}).toString(),d=await F(`/api/reports/cleaning?limit=10000&${r}`);if(d.ok){let o=d.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||""}));K(o,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async n=>{let r=a=>{if(!a)return null;let l=String(a||"").toLowerCase(),v=t.find(x=>String(x.label||"").toLowerCase()===l);return v?v.value:null},d=a=>{if(a==null||a==="")return"";if(a instanceof Date&&!isNaN(a.getTime()))return a.toISOString().slice(0,10);let l=String(a).trim();if(l===""||l==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);if(/^\d{4,5}$/.test(l)){let x=Number(l);if(x>2e4&&x<99999){let h=new Date(Date.UTC(1899,11,30)+x*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}let v=l.split(/[\/\-\.]/);if(v.length===3){let[x,h,m]=v.map(S=>S.trim());if(x.length===4&&h.length<=2&&m.length<=2)return`${x}-${h.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&h.length<=2&&x.length<=2)return`${m}-${h.padStart(2,"0")}-${x.padStart(2,"0")}`}return l},o=n.map(a=>({branch_id:r(String(a.Cabang||"").trim()),activity_type:String(a.Jenis||a.Kegiatan||"").trim(),period:String(a.Periode||"").trim(),activity_date:d(a.Tanggal),status:String(a.Status||"").trim(),document_link:String(a["Link Dokumen"]||"").trim(),notes:String(a.Catatan||a.Keterangan||"").trim()})).filter(a=>a.branch_id&&a.activity_type&&a.period&&a.activity_date),s=await F("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}H();G();async function ia(e,t){let i=await W(),n=Array.from({length:4},(o,s)=>String(new Date().getFullYear()-s)),r=t?t.get("dash_filter"):null,d={};if(r==="fogging"){let o=new Date,s=String(o.getMonth()+1).padStart(2,"0"),a=String(o.getFullYear()),l=t?t.get("month"):null;l&&l.length===7&&(a=l.split("-")[0],s=l.split("-")[1]),d={status:"Done",month:s,year:a}}q({container:e,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:d,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:o=>`<span class="badge badge-warning">${o}</span>`},{key:"period",label:"Periode",render:o=>ye(o)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>ee(o)},{key:"document_link",label:"Dokumen",render:o=>o?`<a href="${o}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:o=>o||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:i,value:o?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:o?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:o?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:o?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:o?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async o=>{let s=new URLSearchParams(o||{}).toString(),a=await F(`/api/reports/fogging?limit=10000&${s}`);if(a.ok){let l=a.data.data.map(v=>({Cabang:v.branch_name||"",Jenis:v.activity_type||"Fogging",Periode:v.period||"",Tanggal:v.activity_date||"",Status:v.status||"","Link Dokumen":v.document_link||""}));K(l,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async o=>{let s=x=>{if(!x)return null;let h=String(x||"").toLowerCase(),m=i.find(S=>String(S.label||"").toLowerCase()===h);return m?m.value:null},a=x=>{if(x==null||x==="")return"";if(x instanceof Date&&!isNaN(x.getTime()))return x.toISOString().slice(0,10);let h=String(x).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let S=Number(h);if(S>2e4&&S<99999){let k=new Date(Date.UTC(1899,11,30)+S*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}let m=h.split(/[\/\-\.]/);if(m.length===3){let[S,k,g]=m.map(c=>c.trim());if(S.length===4&&k.length<=2&&g.length<=2)return`${S}-${k.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&k.length<=2&&S.length<=2)return`${g}-${k.padStart(2,"0")}-${S.padStart(2,"0")}`}return h},l=o.map(x=>({branch_id:s(String(x.Cabang||"").trim()),activity_type:String(x.Jenis||x.Kegiatan||"Fogging").trim(),period:String(x.Periode||"").trim(),activity_date:a(x.Tanggal),status:String(x.Status||"").trim(),document_link:String(x["Link Dokumen"]||"").trim(),notes:String(x.Catatan||x.Keterangan||"").trim()})).filter(x=>x.branch_id&&x.period&&x.activity_date),v=await F("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(l)});if(!v.ok)throw new Error(v.data?.error||"Import gagal");return v.data}}})}H();G();async function ra(e){let t=await W(),i=await se(),n=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],r=Array.from({length:4},(s,a)=>String(new Date().getFullYear()-a)),d=s=>s&&!i.find(a=>a.value===s)?[...i,{value:s,label:s}]:i,o=s=>s&&!n.find(a=>a.value===s)?[...n,{value:s,label:s}]:n;q({container:e,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:s=>`<span title="${s||""}">${s?.length>60?s.slice(0,60)+"\u2026":s||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:s=>window.formatDate(s)},{key:"status",label:"Status",render:s=>ee(s)},{key:"notes",label:"Keterangan",render:s=>s?.length>40?s.slice(0,40)+"\u2026":s||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:s=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:s?.branch_id},{name:"pic",label:"PIC",type:"select",options:o(s?.pic),value:s?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:s?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:s?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:s?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:s?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:s?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async s=>{let a=new URLSearchParams(s||{}).toString(),l=await F(`/api/reports/basecamp?limit=10000&${a}`);if(l.ok){let v=l.data.data.map(x=>({"Tgl Info":x.info_date||"",Cabang:x.branch_name||"",Permasalahan:x.problem||"",PIC:x.pic||"","Tgl Done":x.done_date||"",Status:x.status||"",Keterangan:x.notes||""}));K(v,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async s=>{let a=h=>{if(!h)return null;let m=String(h||"").toLowerCase(),S=t.find(k=>String(k.label||"").toLowerCase()===m);return S?S.value:null},l=h=>{if(h==null||h==="")return"";if(h instanceof Date&&!isNaN(h.getTime()))return h.toISOString().slice(0,10);let m=String(h).trim();if(m===""||m==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);if(/^\d{4,5}$/.test(m)){let k=Number(m);if(k>2e4&&k<99999){let g=new Date(Date.UTC(1899,11,30)+k*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}let S=m.split(/[\/\-\.]/);if(S.length===3){let[k,g,c]=S.map(f=>f.trim());if(k.length===4&&g.length<=2&&c.length<=2)return`${k}-${g.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&g.length<=2&&k.length<=2)return`${c}-${g.padStart(2,"0")}-${k.padStart(2,"0")}`}return m},v=s.map(h=>({info_date:l(h["Tgl Info"]||h["Tanggal Info"]),branch_id:a(String(h.Cabang||"").trim()),problem:String(h.Permasalahan||"").trim(),pic:String(h.PIC||"").trim(),done_date:l(h["Tgl Done"]||h["Tanggal Done"]),status:String(h.Status||"").trim(),notes:String(h.Keterangan||h.Catatan||"").trim()})).filter(h=>h.info_date&&h.branch_id&&h.problem),x=await F("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(v)});if(!x.ok)throw new Error(x.data?.error||"Import gagal");return x.data}}})}async function oa(e){q({container:e,title:"SOP",icon:"\u{1F4DA}",apiPath:"/api/sop",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async t=>{let i=new URLSearchParams(t||{}).toString(),{apiFetch:n}=await Promise.resolve().then(()=>(H(),De)),r=await n(`/api/sop?limit=10000&${i}`);if(r.ok){let d=r.data.data.map(s=>({"Nama SOP":s.name||"",Kategori:s.category||"",Dokumen:s.document_link||"",Catatan:s.notes||s.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(G(),me));o(d,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let t=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(G(),me));i(t,"Template_Import_SOP")},onImport:async t=>{let i=t.map(d=>({name:String(d["Nama SOP"]||"").trim(),category:String(d.Kategori||"").trim(),document_link:String(d.Dokumen||"").trim(),description:String(d.Catatan||"").trim()})).filter(d=>d.name),{apiFetch:n}=await Promise.resolve().then(()=>(H(),De)),r=await n("/api/sop/import",{method:"POST",body:JSON.stringify(i)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}},formFields:t=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:t?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:t?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:t?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:t?.description}]})}he();var rn=[{id:1,title:"Dusting (Pembersihan Debu)",category:"Pembersihan Rutin",icon:"\u{1F9F9}",color:"#2563EB",bg:"#EFF6FF",target:"Meja resepsionis, meja pasien, meja dokter, furniture, kaca, monitor & keyboard",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Kain Lap Microfiber Biru (Area Umum)","Kanebo (Kaca & Layar Monitor)","Kemoceng","Botol Spray Glass Cleaner","Kuas 4 Inch (Keyboard & Sela Sempit)"],steps:["Siapkan larutan glass cleaner di botol spray khusus kaca dan perabot.","Lipat kain microfiber menjadi 8 bidang kerja agar efisien dan bersih merata.","Semprotkan cairan pembersih ke permukaan kain lap, bukan langsung ke monitor atau bodi elektronik.","Usap permukaan perabot dari bagian paling atas ke bagian bawah secara searah atau pola S.","Gunakan kuas cat 4 inch untuk membersihkan sela-sela keyboard dan lubang ventilasi.","Keringkan kaca atau monitor menggunakan lap microfiber kering atau kanebo bersih."],notes:"Khusus layar monitor medis & dental unit, dilarang menyemprotkan cairan langsung ke layar."},{id:2,title:"Sweeping (Lobby Duster & Sapu Nylon)",category:"Pembersihan Rutin",icon:"\u{1F33E}",color:"#D97706",bg:"#FFFBEB",target:"Lobby utama, lorong, ruang tindakan, balkon, pantry, teras",chemical_ratio:"-",tools:["Lobby Duster Set 60cm (Area Lobby & Lorong)","Lobby Duster Set 40cm (Area Ruang Tindakan)","Sapu & Pengki Nagata 742s","Kapi Scrapper (Kerak/Permen Karet)"],steps:["Pastikan kain lobby duster dalam keadaan kering, bersih, dan bebas pasir kasar.","Dorong lobby duster menyusuri tepi dinding ruangan terlebih dahulu membentuk perimeter.","Lanjutkan mendorong dari sudut terjauh menuju arah pintu keluar tanpa mengangkat kain pel lobi dari lantai.","Kumpulkan debu dan kotoran di dekat pintu keluar, lalu angkat menggunakan sapu dan pengki Nagata.","Jika ada noda kerak atau bekas stiker/permen karet menempel di lantai, kikis perlahan menggunakan kapi scrapper.","Kibaskan kain duster di luar ruangan atau sedot sisa debu pada kain duster."],notes:"Jangan pernah menggunakan lobby duster di area lantai yang masih basah atau tergenang air."},{id:3,title:"Mopping (Pengepelan Lantai)",category:"Pembersihan Rutin",icon:"\u{1FAA3}",color:"#059669",bg:"#ECFDF5",target:"Seluruh lantai ruangan, koridor, dan area tunggu klinik",chemical_ratio:"25 ml Floor Cleaner per 1 Liter air bersih",tools:["Single Bucket Mop Wringer 20L / Double Bucket 46L","Mopping Stick / Tongkat & Kain Pel","Chemical Floor Cleaner Prokleen Lavender","Warning Sign Lantai Licin"],steps:['Pasang papan peringatan "Caution Wet Floor" di titik strategis yang mudah terlihat.',"Campurkan cairan Floor Cleaner ke dalam ember air bersih dengan takaran 25 ml per 1 Liter air.","Celupkan kain pel, masukkan ke dalam alat pemeras (wringer), dan peras hingga kondisi kain pel lembap (tidak menetes).","Mulai mengepel dari sudut terjauh ruangan mundur ke arah pintu keluar dengan gerakan membentuk pola angka delapan (8).","Bilas kain pel secara berkala apabila air pada ember bilas sudah terlihat keruh.","Biarkan lantai mengering sempurna secara alami sebelum papan peringatan wet floor dilepas."],notes:'Wajib pasang papan peringatan "Caution Wet Floor" sebelum mulai mengepel untuk mencegah risiko terpeleset.'},{id:4,title:"Membersihkan Toilet & Wastafel",category:"Sanitasi Toilet",icon:"\u{1F6BD}",color:"#DC2626",bg:"#FEF2F2",target:"Toilet pasien, toilet tim, closet, wastafel, dinding keramik, floor drain",chemical_ratio:"Karbol 25 ml:1 L Air / Bowl Cleaner 100 ml spray (kerak) / Bubuk KIFA",tools:["Sarung Tangan Karet (Wajib)","Sikat Kloset Scotch-Brite 3M","Kain Lap Microfiber Merah (Khusus Toilet)","Sponge Scotch Brite & Sikat Lantai ID-51","Floor Squeegee 75cm","Chemical Karbol Pine, Bowl Cleaner, Bubuk KIFA"],steps:["Kenakan sarung tangan karet dan siapkan perlengkapan sanitasi toilet.","Flush kloset, tuangkan Toilet Bowl Cleaner ke bagian dalam dinding kloset dan leher angsa, diamkan 5\u201310 menit.","Sikat bagian dalam kloset secara menyeluruh dengan sikat kloset, lalu bilas (flush) kembali hingga bersih.","Seka dudukan, penutup, dan tabung luar kloset menggunakan kain lap microfiber merah yang dibasahi larutan disinfektan karbol.","Bersihkan wastafel dan kran stainless menggunakan sponge tapas dengan bubuk pembersih KIFA (jangan gunakan Bowl Cleaner pada kran stainless).","Sikat lantai kamar mandi dengan sikat lantai ID-51 dan cairan karbol disinfektan.","Tarik genangan air ke arah floor drain menggunakan floor squeegee hingga lantai tidak licin dan cepat kering.","Pasang kamper/urinal block, isi ulang sabun cuci tangan, dan pasang plastik sampah baru."],notes:"Kain lap warna MERAH mutlak hanya untuk toilet dan mangkuk spittoon. Dilarang mencampur cairan kimia sembarangan."},{id:5,title:"Pembersihan Area Parkir & Luar",category:"Pembersihan Rutin",icon:"\u{1F17F}\uFE0F",color:"#4F46E5",bg:"#EEF2FF",target:"Halaman parkir mobil & motor, saluran air luar, teras depan",chemical_ratio:"-",tools:["Sapu Lidi Gagang Panjang","Pengki Nagata","Plastik Sampah Hitam Besar (90x120)","Selang Air"],steps:["Sapu daun kering, puntung rokok, plastik, dan pasir dari batas trotoar/pagar menuju satu titik.","Kumpulkan dan angkat sampah ke dalam pengki, masukkan langsung ke plastik sampah hitam besar.","Periksa dan bersihkan saluran drainase air di area parkir dari sumbatan sampah padat.","Bila ada tumpahan oli atau kotoran tanah membandel, siram dengan air dan sikat.","Rapikan pembatas parkir dan pastikan area depan klinik terlihat bersih dan rapi menyambut pasien."],notes:"Lakukan sebelum jam operasional klinik dibuka (pukul 07.30 - 08.00)."},{id:6,title:"Pembersihan Kaca & Jendela",category:"Pembersihan Rutin",icon:"\u{1FA9F}",color:"#0891B2",bg:"#ECFEFF",target:"Pintu kaca utama, jendela kaca klinik, partisi kaca, cermin toilet",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Botol Spray Glass Cleaner Prokleen","Window Wiper / Window Squeegee","Kain Microfiber Kering / Kanebo","Tongkat Teleskopik (Kaca Tinggi)"],steps:["Semprotkan larutan glass cleaner merata pada permukaan kaca yang hendak dibersihkan.","Tarik wiper karet dari arah atas ke bawah secara lurus vertikal atau zigzag horizontal tanpa terputus.","Lap bilah karet wiper setiap selesai satu tarikan menggunakan kain lap bersih agar tidak meninggalkan garis bekas air.","Gunakan kain microfiber kering untuk membersihkan sisa tetesan air di sudut-sudut lis bingkai kaca.","Untuk kaca bagian luar yang tinggi, pasang wiper pada tongkat sambungan teleskopik."],notes:"Hindari membersihkan kaca saat terpapar terik sinar matahari langsung karena cairan akan cepat mengering dan membekas noda."},{id:7,title:"Mencuci Bola Mandi Bola (Playground)",category:"Pembersihan Khusus",icon:"\u{1F939}",color:"#7C3AED",bg:"#F5F3FF",target:"Area bermain anak (Kids Playground) dan bola-bola mandi bola",chemical_ratio:"30 ml Dishwashing Liquid per 1 Liter Air + Karbol Disinfektan",tools:["Baskom Plastik Besar 55cm","Dishwashing Liquid Prokleen","Karbol Disinfektan Pine","Lap Microfiber Bersih","Karung / Wadah Pengering Jaring"],steps:["Kumpulkan seluruh bola mandi bola ke dalam wadah atau karung bersih.","Siapkan baskom plastik besar berisi air hangat kuku dicampur Dishwashing Liquid dan sedikit karbol disinfektan.","Rendam bola dan cuci secara bertahap sambil digosok perlahan menggunakan sponge/lap bersih.","Bilas bola dengan air bersih mengalir hingga tidak ada sisa busa sabun yang menempel.","Tiriskan dan angin-anginkan bola di tempat bersih beralaskan terpal/kain kering hingga benar-benar kering.","Seka dan sterilkan area kolam mandi bola sebelum bola dimasukkan kembali."],notes:"Jadwal pencucian berkala dilakukan minimal 1 kali seminggu atau sesuai kebutuhan kebersihan arena anak."},{id:8,title:"Membersihkan Filter Air Purifier",category:"Elektronik & Utility",icon:"\u{1F4A8}",color:"#0D9488",bg:"#F0FDFA",target:"Unit Air Purifier di ruang tunggu, ruang dokter, dan koridor",chemical_ratio:"-",tools:["Vacuum Cleaner (Nozzle Sikat Halus)","Kain Microfiber Kering","Kuas Cat Halus"],steps:["Matikan saklar power unit air purifier dan cabut kabel daya dari stop kontak dinding.","Buka penutup bodi (casing) depan atau belakang air purifier secara hati-hati.","Keluarkan pre-filter jaring nilon dari unit.","Bersihkan debu tebal pada pre-filter menggunakan vacuum cleaner nozzle sikat halus atau cuci dengan air mengalir (jika tipe wash-able).","Untuk filter HEPA dan Karbon Aktif: HANYA sedot debu permukaannya secara perlahan dengan vacuum, JANGAN DICUCI DENGAN AIR.","Lap bodi luar dan kisi-kisi udara masuk menggunakan lap microfiber kering.","Pastikan pre-filter benar-benar kering sebelum dipasang kembali ke posisinya semula."],notes:"Jangan pernah mencuci filter HEPA dengan air karena akan merusak serat mikro penyaring bakteri/virus."},{id:9,title:"Membersihkan Filter & Bodi AC",category:"Elektronik & Utility",icon:"\u2744\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Filter udara dan penutup indoor unit AC split di seluruh ruangan",chemical_ratio:"-",tools:["Tangga Besi Lipat","Kain Lap Microfiber Lembap","Kuas Cat Halus","Kran Air Mengalir"],steps:["Matikan AC melalui remote dan matikan MCB/saklar AC bila diperlukan.","Gunakan tangga besi lipat yang stabil di lantai datar.","Buka tutup cover depan unit indoor AC ke atas hingga terkunci.","Lepaskan kedua kisi filter saringan udara plastik secara perlahan agar debu tidak berjatuhan.","Cuci saringan filter di bawah kran air mengalir dari arah bagian dalam ke bagian luar.","Keringkan filter di tempat teduh (jangan dijemur langsung di bawah terik matahari).","Lap bodi casing dan louver/swing AC dengan kain microfiber lembap.","Pasang kembali filter yang sudah kering, tutup cover AC, dan nyalakan AC untuk pengetesan."],notes:"Pembersihan rutin mandiri dilakukan 2 minggu sekali. Deep cleaning / cuci evaporator dilakukan oleh teknisi AC vendor."},{id:10,title:"Membersihkan Dental Unit",category:"Klinis & Medis",icon:"\u{1F9B7}",color:"#E11D48",bg:"#FFF1F2",target:"Kursi dental pasien, spittoon bowl, meja instrumen, lampu dental, cover bawah",chemical_ratio:"Glass Cleaner (netral) / Karbol Disinfektan (mangkuk) / Bowl Cleaner (kerak)",tools:["Sarung Tangan Karet (Wajib)","Kain Lap Microfiber Biru (Khusus Bodi Luar)","Kain Lap Microfiber Merah (Khusus Mangkuk Spittoon)","Botol Spray Glass Cleaner (25ml + Air)","Botol Spray Karbol (20ml + Air)","Botol Spray Bowl Cleaner (100ml + Air, bila berkerak)"],steps:["Pastikan saklar utama dental unit dalam posisi OFF dan kursi dinaikkan ke posisi tertinggi untuk memudahkan akses.","Lap seluruh jok kulit, sandaran kepala, sandaran tangan, dan tiang lampu menggunakan lap microfiber biru dan cairan pembersih netral/glass cleaner.","Bersihkan mangkuk kumur (spittoon bowl) menggunakan kain lap merah dan cairan karbol disinfektan.","Jika ditemukan kerak sisa darah/plak membandel di mangkuk spittoon, gunakan sedikit bowl cleaner, sikat perlahan, lalu bilas hingga bersih.","Lap permukaan meja instrumen meja dokter. Bila ada instrumen steril, geser dengan hati-hati dan jangan mengotori instrumen.","Buka penutup cover bawah unit dental, lap debu yang terjangkau dan pastikan kondisi area selang dalam keadaan kering."],notes:"Kain lap warna MERAH hanya untuk mangkuk spittoon. Dilarang keras menggosok jok kulit dental unit dengan bahan kimia asam keras."},{id:11,title:"Membersihkan Wastafel & Keran Air",category:"Sanitasi Toilet",icon:"\u{1F6B0}",color:"#0284C7",bg:"#F0F9FF",target:"Wastafel keramik ruang tindakan, wastafel lorong, kran stainless steel, afur sink",chemical_ratio:"Bubuk Pembersih KIFA / Dishwashing Liquid Prokleen",tools:["Sponge Scotch Brite Hijau/Kuning","Kain Microfiber Hijau (Area Basah)","Bubuk Pembersih KIFA","Sikat Gigi Bekas / Sikat Kecil"],steps:["Basahi seluruh mangkuk wastafel keramik dan kran air stainless steel.","Taburkan sedikit Bubuk KIFA pada tapas sponge yang basah.","Gosok perlahan noda sabun dan kerak air putih pada kran stainless steel searah guratan logam hingga mengkilap kembali.","Gunakan sikat kecil untuk membersihkan sela leher kran dan sudut sambungan afur pembuangan wastafel.","Bilas mangkuk wastafel dan kran dengan air bersih mengalir hingga seluruh sisa residu bubuk KIFA hilang.","Keringkan kran air stainless steel dan cermin wastafel menggunakan lap microfiber hijau kering."],notes:"DILARANG KERAS menggunakan Toilet Bowl Cleaner pada kran air stainless steel karena menyebabkan karat dan warna menghitam permanen."},{id:12,title:"Membersihkan Kulkas Klinik & Pantry",category:"Elektronik & Utility",icon:"\u{1F9CA}",color:"#475569",bg:"#F8FAFC",target:"Kulkas pantry staf dan kulkas penyimpan obat/material medis (eksterior & rak)",chemical_ratio:"30 ml Dishwashing Liquid : 1 Liter Air Hangat + Lap Kering",tools:["Kain Microfiber Biru & Hijau","Dishwashing Liquid","Ember Air Hangat","Sponge Scotch Brite Halus"],steps:["Keluarkan makanan/minuman yang telah kedaluwarsa atau berbau tidak sedap.","Lepas rak-rak kaca/plastik interior kulkas yang dapat dilepas, cuci dengan sabun cuci piring dan bilas air bersih.","Seka dinding interior kulkas menggunakan kain microfiber yang dicelupkan ke dalam air hangat bersabun lembut.","Lap kering seluruh permukaan dalam kulkas dengan microfiber kering bersih.","Pasang kembali rak interior dan susun kembali bahan/makanan secara tertib dan rapi.","Bersihkan karet pintu kulkas dari noda hitam/jamur menggunakan sikat halus dan lap kering.","Lap bodi luar kulkas dan pegangan pintu menggunakan cairan pembersih netral."],notes:"Untuk kulkas penyimpanan obat/material dental, koordinasikan terlebih dahulu dengan Perawat/PIC Obat sebelum membersihkan rak interior."},{id:13,title:"Membersihkan Sofa & Kursi Ruang Tunggu",category:"Pembersihan Khusus",icon:"\u{1F6CB}\uFE0F",color:"#9333EA",bg:"#FAF5FF",target:"Sofa ruang tunggu pasien, kursi reception, kursi kerja tim",chemical_ratio:"Baking Soda (penghilang bau) / Pembersih Upholstery secukupnya",tools:["Vacuum Cleaner (Nozzle Kecil & Sikat Lembut)","Kain Microfiber Bersih","Sikat Bulu Halus","Ember Air Hangat","Cairan Pembersih Khusus Sofa (Upholstery Cleaner)","Baking Soda (Penghilang Bau Tidak Sedap)","Sarung Tangan Karet"],steps:["Kenali bahan sofa (kain fabric, beludru, kulit asli, atau kulit sintetis/oscar).","Sedot debu di seluruh permukaan, lipatan jahitan, dan celah bantal sofa menggunakan vacuum cleaner nozzle kecil.","Gunakan sikat bulu halus untuk mengangkat rambut atau kotoran yang menempel kuat di serat kain.","Jika ada noda bercak (kopi, makanan, tinta), semprotkan pembersih upholstery secukupnya, gosok memutar perlahan dengan microfiber.","Celupkan kain microfiber ke air hangat suam-suam kuku, peras hingga lembap, lalu seka seluruh permukaan sofa.","Lap kembali dengan kain kering bersih untuk menyerap sisa kelembapan.","Jika sofa berbau apek: taburkan baking soda tipis di permukaan, diamkan 20 menit, lalu sedot kembali dengan vacuum cleaner.","Biarkan sofa mengering secara alami di dalam ruangan berventilasi atau ber-AC."],notes:"Jangan menjemur sofa langsung di bawah terik matahari karena dapat menyebabkan bahan kulit pecah-pecah dan warna kain memudar."}],on=[{no:1,name:"Sapu Lidi Gagang",qty:"1 unit",area:"Area parkir mobil & motor klinik"},{no:2,name:"Sapu & Pengki Nagata 742s",qty:"1 set",area:"Area pantry, balkon, dak lantai atas, atau teras"},{no:3,name:"Lobby Duster Set 60 cm",qty:"1 set",area:"Lobby utama dan koridor/lorong klinik"},{no:4,name:"Lobby Duster Set 40 cm / Refill",qty:"1 unit",area:"Area ruang tindakan dokter"},{no:5,name:"Single Bucket Mop Wringer 20L",qty:"1 paket",area:"Ember pel dengan pemeras + tongkat pel"},{no:6,name:"Double Bucket Wringer Trolley 46L",qty:"1 paket",area:"Khusus klinik besar / klinik dengan fasilitas lift"},{no:7,name:"Kemoceng Bulu Lembut",qty:"1 unit",area:"Membersihkan debu kasar meja dan perabot"},{no:8,name:"Lap Microfiber Warna Merah",qty:"2 pcs",area:"Khusus area toilet, sanitasi, dan mangkuk spittoon"},{no:9,name:"Lap Microfiber Warna Biru",qty:"2 pcs",area:"Area umum kering (meja, kursi, bodi peralatan)"},{no:10,name:"Lap Microfiber Warna Hijau",qty:"2 pcs",area:"Area umum basah (pantry, wastafel luar)"},{no:11,name:"Kanebo / Synthetic Chamois",qty:"1 pcs",area:"Area umum, cermin, dan pengering kaca"},{no:12,name:"Caddy Bag Cleaning",qty:"1 unit",area:"Tas perlengkapan pinggang petugas cleaning"},{no:13,name:"Sponge Scotch Brite 3M",qty:"2 pcs",area:"Pembersihan noda lantai dan kerak ringan (spotting)"},{no:14,name:"Scotch-Brite 3M Sikat Toilet",qty:"1 unit",area:"Sikat khusus kloset dan leher angsa"},{no:15,name:"Kapi Gagang Karet / Scrapper",qty:"1 unit",area:"Pembersihan kerak lantai, noda permen karet, stiker"},{no:16,name:"Kuas Cat 4 Inch",qty:"2 pcs",area:"Pembersih sela keyboard, ventilasi, dan sudut sempit"},{no:17,name:"Warning Sign (Lantai Licin)",qty:"1 unit",area:"Digunakan saat mopping dan general cleaning"},{no:18,name:"Teleskopik + Pembersih Kaca/Langit",qty:"1 set",area:"Membersihkan jendela tinggi, plafon, dan sarang laba-laba"},{no:19,name:"Sarung Tangan Karet 3M",qty:"1 pasang",area:"Perlindungan tangan area toilet dan bahan kimia"},{no:20,name:"Botol Spray 500ml",qty:"5 unit",area:"Wadah larutan chemical (diberi label masing-masing)"},{no:21,name:"Pengharum Ruangan Stella Aerosol",qty:"1 kaleng",area:"Digunakan darurat apabila timbul bau menyengat"},{no:22,name:"Kleaner Gantungan Sapu & Pel",qty:"1 set",area:"Gantungan penyimpanan alat cleaning di janitor"},{no:23,name:"Floor Squeegee 75 cm",qty:"1 unit",area:"Pendorong genangan air lantai toilet/teras"},{no:24,name:"Sikat Botol Ikea",qty:"1 unit",area:"Di setiap ruang tindakan dokter gigi"},{no:25,name:"Scotch-Brite 3M Sikat Lantai ID-51",qty:"1 unit",area:"Sikat lantai kamar mandi dan teras"},{no:26,name:"Kamper Toilet / Urinal Block",qty:"1 pack",area:"Pengharum di toilet dan urinoir"},{no:27,name:"Plastik Sampah Hitam Sedang (60x100)",qty:"1 pack",area:"Tempat sampah indoor ruangan"},{no:28,name:"Plastik Sampah Hitam Besar (90x120)",qty:"1 pack",area:"Tempat sampah luar & penampung sampah akhir"},{no:29,name:"Plastik Sampah Medis Kuning (90x120)",qty:"1 pack",area:"Limbah medis padat B3 klinik"},{no:30,name:"Plastik Sampah Medis Kuning (40x60)",qty:"1 pack",area:"Tempat sampah medis kecil di ruang tindakan"},{no:31,name:"Tangga Besi Lipat",qty:"1 unit",area:"Alat bantu membersihkan plafon, lampu, dan AC"},{no:32,name:"Baskom Plastik Besar 55cm",qty:"1 unit",area:"Mencuci bola kids playground"}],sn=[{no:1,name:"Floor Cleaner Prokleen Lavender",size:"5 Liter",ratio:"25 ml : 1 Liter Air",function:"Pembersihan dan pengepelan rutin seluruh lantai ruangan"},{no:2,name:"Karbol Prokleen Pine Disinfektan",size:"5 Liter",ratio:"25 ml : 1 Liter Air (Mop) / 20 ml per botol spray",function:"Mengepel lantai toilet, disinfeksi area basah dan berbau tak sedap"},{no:3,name:"Glass Cleaner Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air / Siap pakai di botol spray",function:"Membersihkan pintu kaca, jendela, cermin, dan layar monitor"},{no:4,name:"Manual Dishwashing Liquid Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air",function:"Mencuci piring, gelas, sendok pantry, dan mencuci bola playground"},{no:5,name:"Hand Soap Prokleen Green Tea",size:"5 Liter",ratio:"Siap pakai (tanpa campur)",function:"Refill sabun pencuci tangan di wastafel pasien dan tim"},{no:6,name:"Liquid Detergent Prokleen",size:"5 Liter",ratio:"Sesuai kapasitas mesin",function:"Mencuci handuk, kain microfiber, dan seragam"},{no:7,name:"Toilet Bowl Cleaner Prokleen",size:"5 Liter",ratio:"100 ml per botol spray / Khusus kerak",function:"Pembersih kerak membandel pada kloset, wastafel keramik, dan nat keramik"},{no:8,name:"Bubuk Pembersih KIFA",size:"500 gram",ratio:"Taburkan tipis pada tapas basah",function:"Membersihkan kerak air pada kran air stainless steel, jet washer, dan sink"}],Dt="fcms_sop_fc_ik_v1",It="fcms_sop_fc_tools_v1",At="fcms_sop_fc_chem_v1";function sa(){try{let e=localStorage.getItem(Dt);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK",e)}return JSON.parse(JSON.stringify(rn))}function $t(e){localStorage.setItem(Dt,JSON.stringify(e))}function la(){try{let e=localStorage.getItem(It);if(e)return JSON.parse(e)}catch(e){console.error("Error load tools",e)}return JSON.parse(JSON.stringify(on))}function Tt(e){localStorage.setItem(It,JSON.stringify(e))}function da(){try{let e=localStorage.getItem(At);if(e)return JSON.parse(e)}catch(e){console.error("Error load chem",e)}return JSON.parse(JSON.stringify(sn))}function Ft(e){localStorage.setItem(At,JSON.stringify(e))}async function ca(e){let t="ik",i="",n="all",r="tools",d="",o=sa(),s=la(),a=da();function l(){e.innerHTML=`
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
        <button class="sop-tab-btn ${t==="ik"?"active":""}" data-tab="ik" style="${v(t==="ik")}">
          \u{1F4CB} Instruksi Kerja (IK) <span class="badge" style="${x(t==="ik")}">${o.length}</span>
        </button>
        <button class="sop-tab-btn ${t==="tools"?"active":""}" data-tab="tools" style="${v(t==="tools")}">
          \u{1F6E0}\uFE0F Alat & Chemical <span class="badge" style="${x(t==="tools")}">${s.length+a.length}</span>
        </button>
        <button class="sop-tab-btn ${t==="policy"?"active":""}" data-tab="policy" style="${v(t==="policy")}">
          \u{1F4D6} Pedoman & Ketentuan
        </button>
        <button class="sop-tab-btn ${t==="links"?"active":""}" data-tab="links" style="${v(t==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sop-content-area"></div>
    `,h(),m()}function v(p){return p?"border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function x(p){return p?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function h(){e.querySelectorAll(".sop-tab-btn").forEach(A=>{A.addEventListener("click",()=>{t=A.dataset.tab,l()})});let p=e.querySelector("#btn-reset-default");p&&p.addEventListener("click",()=>{pe("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja, Alat, dan Chemical ke data standar dokumen asli?",()=>{localStorage.removeItem(Dt),localStorage.removeItem(It),localStorage.removeItem(At),o=sa(),s=la(),a=da(),j("Data SOP Facility Care berhasil di-reset ke standar asli"),l()},"Konfirmasi Reset Standar")});let T=e.querySelector("#btn-print-sop");T&&T.addEventListener("click",()=>{window.print()});let D=e.querySelector("#btn-export-tools");D&&D.addEventListener("click",()=>{C()})}function m(){let p=e.querySelector("#sop-content-area");p&&(t==="ik"?S(p):t==="tools"?c(p):t==="policy"?w(p):t==="links"&&b(p))}function S(p){p.innerHTML=`
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
              <button class="cat-pill ${n==="Pembersihan Rutin"?"active":""}" data-cat="Pembersihan Rutin">Rutin (${o.filter(B=>B.category==="Pembersihan Rutin").length})</button>
              <button class="cat-pill ${n==="Sanitasi Toilet"?"active":""}" data-cat="Sanitasi Toilet">Toilet & Wastafel (${o.filter(B=>B.category==="Sanitasi Toilet").length})</button>
              <button class="cat-pill ${n==="Klinis & Medis"?"active":""}" data-cat="Klinis & Medis">Klinis Dental (${o.filter(B=>B.category==="Klinis & Medis").length})</button>
              <button class="cat-pill ${n==="Elektronik & Utility"?"active":""}" data-cat="Elektronik & Utility">Elektronik & AC (${o.filter(B=>B.category==="Elektronik & Utility").length})</button>
              <button class="cat-pill ${n==="Pembersihan Khusus"?"active":""}" data-cat="Pembersihan Khusus">Khusus (${o.filter(B=>B.category==="Pembersihan Khusus").length})</button>
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
    `;let T=document.createElement("style");T.innerHTML=`
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
    `,p.appendChild(T);let D=p.querySelector("#ik-cards-grid-container"),A=p.querySelector("#ik-count-label");function O(){let B=o.filter(L=>{let N=n==="all"||L.category===n,y=i.toLowerCase().trim(),E=!y||L.title.toLowerCase().includes(y)||L.target.toLowerCase().includes(y)||L.tools.some($=>$.toLowerCase().includes(y))||L.steps.some($=>$.toLowerCase().includes(y));return N&&E});if(A&&(A.textContent=`Menampilkan ${B.length} dari ${o.length} Instruksi Kerja`),B.length===0){D.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}D.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fill, minmax(360px, 1fr)); gap:18px;">
          ${B.map(L=>`
            <div class="card ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${L.bg||"#EFF6FF"}; border-radius:10px;">${L.icon||"\u{1F4CB}"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        ${L.id}. ${L.title}
                      </h3>
                      <span class="badge" style="background:${L.bg||"#EFF6FF"}; color:${L.color||"#2563EB"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${L.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    <button class="btn btn-sm btn-ghost btn-edit-ik" data-id="${L.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    <button class="btn btn-sm btn-ghost btn-delete-ik" data-id="${L.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${L.color||"#2563EB"};">
                  <strong>Area:</strong> ${L.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Bahan Utama:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${L.tools.slice(0,3).map(N=>`<li style="margin-bottom:2px;">${N}</li>`).join("")}
                    ${L.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${L.tools.length-3} alat lainnya...</li>`:""}
                  </ul>
                </div>

                ${L.chemical_ratio&&L.chemical_ratio!=="-"?`
                  <div style="font-size:0.8rem; background:#FEF3C7; color:#92400E; padding:6px 10px; border-radius:6px; margin-bottom:14px; font-weight:600;">
                    \u{1F9EA} Takaran: ${L.chemical_ratio}
                  </div>
                `:""}
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${L.steps.length} Langkah Kerja
                </span>
                <div style="display:flex; gap:6px;">
                  <button class="btn btn-sm btn-primary btn-view-ik" data-id="${L.id}" style="padding:6px 14px; font-weight:600; border-radius:6px;">
                    Lihat Panduan \u2794
                  </button>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,D.querySelectorAll(".btn-view-ik").forEach(L=>{L.addEventListener("click",()=>{let N=parseInt(L.dataset.id,10),y=o.find(E=>E.id===N);y&&k(y)})}),D.querySelectorAll(".btn-edit-ik").forEach(L=>{L.addEventListener("click",()=>{let N=parseInt(L.dataset.id,10),y=o.find(E=>E.id===N);y&&g(y)})}),D.querySelectorAll(".btn-delete-ik").forEach(L=>{L.addEventListener("click",()=>{let N=parseInt(L.dataset.id,10),y=o.find(E=>E.id===N);y&&pe(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${y.title}"?`,()=>{o=o.filter(E=>E.id!==N),$t(o),j("Instruksi Kerja berhasil dihapus"),l()},"Hapus Instruksi Kerja")})})}O();let P=p.querySelector("#ik-search");P&&P.addEventListener("input",B=>{i=B.target.value,O()}),p.querySelectorAll(".cat-pill").forEach(B=>{B.addEventListener("click",()=>{p.querySelectorAll(".cat-pill").forEach(L=>L.classList.remove("active")),B.classList.add("active"),n=B.dataset.cat,O()})});let I=p.querySelector("#btn-add-ik");I&&I.addEventListener("click",()=>{g()})}function k(p){ae({title:`${p.icon||"\u{1F4CB}"} IK-${p.id}: ${p.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${p.bg||"#EFF6FF"}; color:${p.color||"#2563EB"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${p.category}</span>
            <span>ID: IK-${String(p.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${p.color||"#2563EB"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${p.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Bahan yang Dibutuhkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${p.tools.map(T=>`<li>${T}</li>`).join("")}
              </ul>
            </div>
          </div>

          ${p.chemical_ratio&&p.chemical_ratio!=="-"?`
            <div style="background:#FEF3C7; border:1px solid #FCD34D; border-radius:8px; padding:10px 14px;">
              <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:2px;">\u{1F9EA} Rasio & Takaran Chemical:</strong>
              <span style="color:#78350F; font-size:0.9rem; font-weight:600;">${p.chemical_ratio}</span>
            </div>
          `:""}

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah-Langkah Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${p.steps.map(T=>`<li style="margin-bottom:6px;">${T}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${p.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Penting & Safety:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${p.notes}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(T,D)=>{D()}})}function g(p=null){let T=!!p,D=p?p.tools.join(`
`):"",A=p?p.steps.join(`
`):"",O=`
      <form id="ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-ik-title" class="form-control" required value="${p?_(p.title):""}" placeholder="Contoh: Membersihkan Dental Unit" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-ik-icon" class="form-control" value="${p?p.icon:"\u{1F4CB}"}" placeholder="\u{1F9F9}" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
            <select id="form-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
              <option value="Pembersihan Rutin" ${p?.category==="Pembersihan Rutin"?"selected":""}>Pembersihan Rutin</option>
              <option value="Sanitasi Toilet" ${p?.category==="Sanitasi Toilet"?"selected":""}>Sanitasi Toilet</option>
              <option value="Klinis & Medis" ${p?.category==="Klinis & Medis"?"selected":""}>Klinis & Medis</option>
              <option value="Elektronik & Utility" ${p?.category==="Elektronik & Utility"?"selected":""}>Elektronik & Utility</option>
              <option value="Pembersihan Khusus" ${p?.category==="Pembersihan Khusus"?"selected":""}>Pembersihan Khusus</option>
            </select>
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Takaran / Rasio Chemical</label>
            <input type="text" id="form-ik-ratio" class="form-control" value="${p?_(p.chemical_ratio||""):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-ik-target" class="form-control" required value="${p?_(p.target):""}" placeholder="Contoh: Kursi pasien, meja instrumen, mangkuk spittoon" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Bahan yang Digunakan (1 item per baris) *</label>
          <textarea id="form-ik-tools" class="form-control" rows="4" placeholder="Kain Microfiber Biru&#10;Sarung Tangan Karet&#10;Botol Spray Glass Cleaner" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${_(D)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah-Langkah Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-ik-steps" class="form-control" rows="6" placeholder="1. Pastikan posisi saklar OFF.&#10;2. Lap jok dental unit dengan kain microfiber biru.&#10;3. Sikat mangkuk spittoon..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${_(A)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Penting / Hal yang Perlu Diperhatikan (Safety)</label>
          <textarea id="form-ik-notes" class="form-control" rows="2" placeholder="Catatan larangan atau peringatan keselamatan..." style="width:100%; border-radius:6px; font-size:0.85rem;">${p?_(p.notes||""):""}</textarea>
        </div>
      </form>
    `;ae({title:T?`\u270F\uFE0F Edit IK-${p.id}: ${p.title}`:"\u2795 Tambah Instruksi Kerja Baru",content:O,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(P,I)=>{let B=P.querySelector("#form-ik-title").value.trim(),L=P.querySelector("#form-ik-icon").value.trim()||"\u{1F4CB}",N=P.querySelector("#form-ik-cat").value,y=P.querySelector("#form-ik-ratio").value.trim(),E=P.querySelector("#form-ik-target").value.trim(),$=P.querySelector("#form-ik-tools").value.trim(),U=P.querySelector("#form-ik-steps").value.trim(),J=P.querySelector("#form-ik-notes").value.trim();if(!B||!E||!$||!U){R("Harap lengkapi seluruh field wajib (*)");return}let Y=$.split(`
`).map(ce=>ce.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),M=U.split(`
`).map(ce=>ce.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),oe={"Pembersihan Rutin":{color:"#2563EB",bg:"#EFF6FF"},"Sanitasi Toilet":{color:"#DC2626",bg:"#FEF2F2"},"Klinis & Medis":{color:"#E11D48",bg:"#FFF1F2"},"Elektronik & Utility":{color:"#0D9488",bg:"#F0FDFA"},"Pembersihan Khusus":{color:"#7C3AED",bg:"#F5F3FF"}}[N]||{color:"#2563EB",bg:"#EFF6FF"};if(T)p.title=B,p.icon=L,p.category=N,p.chemical_ratio=y,p.target=E,p.tools=Y,p.steps=M,p.notes=J,p.color=oe.color,p.bg=oe.bg,$t(o),j(`Instruksi Kerja "${B}" berhasil diperbarui`);else{let ce=o.length>0?Math.max(...o.map(Q=>Q.id))+1:1,ue={id:ce,title:B,category:N,icon:L,color:oe.color,bg:oe.bg,target:E,chemical_ratio:y||"-",tools:Y,steps:M,notes:J};o.push(ue),$t(o),j(`Instruksi Kerja baru IK-${ce} berhasil ditambahkan`)}I(),l()}})}function c(p){p.innerHTML=`
      <!-- Sub-tab Switcher & Actions -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-tools-btn" class="btn btn-sm ${r==="tools"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9F9} Alat Kebersihan (${s.length} Item)
          </button>
          <button id="view-chem-btn" class="btn btn-sm ${r==="chemical"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9EA} Bahan Chemical (${a.length} Item)
          </button>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="tool-search" class="form-control" placeholder="Cari nama / peruntukan / rasio..." value="${d}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
          <button id="btn-add-tool-chem" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u2795 ${r==="tools"?"Tambah Alat Baru":"Tambah Chemical Baru"}
          </button>
        </div>
      </div>

      <!-- Table Container -->
      <div id="tool-table-container"></div>
    `;let T=p.querySelector("#tool-table-container"),D=p.querySelector("#tool-search");function A(){let O=d.toLowerCase().trim();if(r==="tools"){let P=s.filter(I=>!O||I.name.toLowerCase().includes(O)||I.area.toLowerCase().includes(O));T.innerHTML=`
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
                  ${P.map((I,B)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${B%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${I.no||B+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${_(I.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#E0F2FE; color:#0369A1; font-weight:600;">${_(I.qty)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${_(I.area)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                        <button class="btn btn-sm btn-ghost btn-edit-tool" data-no="${I.no}" title="Edit Alat" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                        <button class="btn btn-sm btn-ghost btn-delete-tool" data-no="${I.no}" title="Hapus Alat" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,T.querySelectorAll(".btn-edit-tool").forEach(I=>{I.addEventListener("click",()=>{let B=parseInt(I.dataset.no,10),L=s.find(N=>N.no===B);L&&f(L)})}),T.querySelectorAll(".btn-delete-tool").forEach(I=>{I.addEventListener("click",()=>{let B=parseInt(I.dataset.no,10),L=s.find(N=>N.no===B);L&&pe(`Apakah Anda yakin ingin menghapus alat: "${L.name}"?`,()=>{s=s.filter(N=>N.no!==B),s.forEach((N,y)=>{N.no=y+1}),Tt(s),j("Alat kebersihan berhasil dihapus"),l()},"Hapus Alat Kebersihan")})})}else{let P=a.filter(I=>!O||I.name.toLowerCase().includes(O)||I.function.toLowerCase().includes(O));T.innerHTML=`
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
                  ${P.map((I,B)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${B%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${I.no||B+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${_(I.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#FEF3C7; color:#92400E; font-weight:600;">${_(I.size)}</span></td>
                      <td style="padding:10px 14px;"><span style="background:#EFF6FF; color:#1E40AF; padding:3px 8px; border-radius:6px; font-weight:600; font-size:0.82rem;">${_(I.ratio)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${_(I.function)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                        <button class="btn btn-sm btn-ghost btn-edit-chem" data-no="${I.no}" title="Edit Chemical" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                        <button class="btn btn-sm btn-ghost btn-delete-chem" data-no="${I.no}" title="Hapus Chemical" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,T.querySelectorAll(".btn-edit-chem").forEach(I=>{I.addEventListener("click",()=>{let B=parseInt(I.dataset.no,10),L=a.find(N=>N.no===B);L&&u(L)})}),T.querySelectorAll(".btn-delete-chem").forEach(I=>{I.addEventListener("click",()=>{let B=parseInt(I.dataset.no,10),L=a.find(N=>N.no===B);L&&pe(`Apakah Anda yakin ingin menghapus chemical: "${L.name}"?`,()=>{a=a.filter(N=>N.no!==B),a.forEach((N,y)=>{N.no=y+1}),Ft(a),j("Bahan chemical berhasil dihapus"),l()},"Hapus Bahan Chemical")})})}}A(),p.querySelector("#view-tools-btn").addEventListener("click",()=>{r="tools",c(p)}),p.querySelector("#view-chem-btn").addEventListener("click",()=>{r="chemical",c(p)}),D&&D.addEventListener("input",O=>{d=O.target.value,A()}),p.querySelector("#btn-add-tool-chem").addEventListener("click",()=>{r==="tools"?f():u()})}function f(p=null){let T=!!p,D=`
      <form id="tool-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Alat / Item *</label>
          <input type="text" id="form-tool-name" class="form-control" required value="${p?_(p.name):""}" placeholder="Contoh: Lobby Duster Set 60 cm" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kuantitas Standar *</label>
          <input type="text" id="form-tool-qty" class="form-control" required value="${p?_(p.qty):"1 unit"}" placeholder="Contoh: 1 set, 2 pcs, 1 paket" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Peruntukan Area & Kegunaan *</label>
          <textarea id="form-tool-area" class="form-control" rows="3" required placeholder="Contoh: Lobby utama dan koridor/lorong klinik" style="width:100%; border-radius:6px; font-size:0.85rem;">${p?_(p.area):""}</textarea>
        </div>
      </form>
    `;ae({title:T?`\u270F\uFE0F Edit Alat: ${p.name}`:"\u2795 Tambah Alat Kebersihan Baru",content:D,confirmText:"Simpan",onConfirm:(A,O)=>{let P=A.querySelector("#form-tool-name").value.trim(),I=A.querySelector("#form-tool-qty").value.trim(),B=A.querySelector("#form-tool-area").value.trim();if(!P||!I||!B){R("Harap lengkapi semua kolom");return}if(T)p.name=P,p.qty=I,p.area=B,Tt(s),j(`Alat "${P}" berhasil diperbarui`);else{let L=s.length>0?Math.max(...s.map(N=>N.no))+1:1;s.push({no:L,name:P,qty:I,area:B}),Tt(s),j(`Alat baru "${P}" berhasil ditambahkan`)}O(),l()}})}function u(p=null){let T=!!p,D=`
      <form id="chem-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Produk Chemical *</label>
          <input type="text" id="form-chem-name" class="form-control" required value="${p?_(p.name):""}" placeholder="Contoh: Floor Cleaner Prokleen Lavender" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kemasan Standar *</label>
            <input type="text" id="form-chem-size" class="form-control" required value="${p?_(p.size):"5 Liter"}" placeholder="Contoh: 5 Liter, 500 gram" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Rasio & Takaran Standar *</label>
            <input type="text" id="form-chem-ratio" class="form-control" required value="${p?_(p.ratio):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Fungsi & Cara Penggunaan *</label>
          <textarea id="form-chem-function" class="form-control" rows="3" required placeholder="Contoh: Pembersihan dan pengepelan rutin seluruh lantai ruangan" style="width:100%; border-radius:6px; font-size:0.85rem;">${p?_(p.function):""}</textarea>
        </div>
      </form>
    `;ae({title:T?`\u270F\uFE0F Edit Chemical: ${p.name}`:"\u2795 Tambah Chemical Baru",content:D,confirmText:"Simpan",onConfirm:(A,O)=>{let P=A.querySelector("#form-chem-name").value.trim(),I=A.querySelector("#form-chem-size").value.trim(),B=A.querySelector("#form-chem-ratio").value.trim(),L=A.querySelector("#form-chem-function").value.trim();if(!P||!I||!B||!L){R("Harap lengkapi semua kolom");return}if(T)p.name=P,p.size=I,p.ratio=B,p.function=L,Ft(a),j(`Chemical "${P}" berhasil diperbarui`);else{let N=a.length>0?Math.max(...a.map(y=>y.no))+1:1;a.push({no:N,name:P,size:I,ratio:B,function:L}),Ft(a),j(`Chemical baru "${P}" berhasil ditambahkan`)}O(),l()}})}function w(p){p.innerHTML=`
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
    `}function b(p){p.innerHTML=`
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
    `}async function C(){try{let{downloadExcel:p}=await Promise.resolve().then(()=>(G(),me)),T=o.map(A=>({"No IK":`IK-${A.id}`,"Judul Pekerjaan":A.title,Kategori:A.category,"Sasaran Target Area":A.target,"Rasio Chemical":A.chemical_ratio||"-","Alat & Bahan":A.tools.join(", "),"Langkah Kerja":A.steps.map((O,P)=>`${P+1}. ${O}`).join(" | "),"Catatan Safety":A.notes||"-"})),D=new Date().toISOString().slice(0,10);p(T,`SOP_Facility_Care_Instruksi_Kerja_${D}`),j("Data Excel Instruksi Kerja berhasil diunduh")}catch(p){console.error(p),R("Gagal mendownload data Excel: "+p.message)}}function _(p){return p?String(p).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}l()}async function pa(e){q({container:e,title:"Master Document",icon:"\u{1F4CB}",apiPath:"/api/checklist",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Document",columns:[{key:"name",label:"Nama Document"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari document..."}],exportOptions:{moduleName:"checklist",onExport:async t=>{let i=new URLSearchParams(t||{}).toString(),{apiFetch:n}=await Promise.resolve().then(()=>(H(),De)),r=await n(`/api/checklist?limit=10000&${i}`);if(r.ok){let d=r.data.data.map(s=>({"Nama Document":s.name||"",Kategori:s.category||"",Dokumen:s.document_link||"",Deskripsi:s.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(G(),me));o(d,`Master_Document_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let t=[{"Nama Document":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(G(),me));i(t,"Template_Import_Document")},onImport:async t=>{let i=t.map(d=>({name:String(d["Nama Document"]||d["Nama Checklist"]||"").trim(),category:String(d.Kategori||"").trim(),document_link:String(d.Dokumen||"").trim(),description:String(d.Deskripsi||"").trim()})).filter(d=>d.name),{apiFetch:n}=await Promise.resolve().then(()=>(H(),De)),r=await n("/api/checklist/import",{method:"POST",body:JSON.stringify(i)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}},formFields:t=>[{name:"name",label:"Nama Document",required:!0,placeholder:"Nama document",value:t?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:t?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:t?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:t?.description}]})}H();he();G();var Ne={"Standar Umum":{icon:"\u{1F4CC}",bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE",rooms:["Lantai & Sudut Ruangan","Dinding, Partisi & Plint","Kaca, Cermin & Partisi Kaca","Sanitair & Area Basah Toilet","Plafon, Lampu & Kisi Ventilasi","Mebel, Meja & Kursi Kerja/Pasien","Peralatan Khusus & Medis Non-Steril","Wadah Sampah & Utilitas"]},"Area Depan & Pelayanan":{icon:"\u{1F6CB}\uFE0F",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",rooms:["Teras","Lobby 1","Receptionist","Snack Corner"]},"Area Klinis & Medis":{icon:"\u{1FA7A}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD",rooms:["Ruang Tindakan","Ronsen","Ruang Dokter","Ruang Steril"]},"Toilet & Sanitasi":{icon:"\u{1F6BF}",bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4",rooms:["Toilet 1","Janitor","Ruang Limbah"]},Sirkulasi:{icon:"\u{1F6B6}",bg:"#F1F5F9",color:"#334155",border:"#CBD5E1",rooms:["Area Tangga","Lift","Area Tangga Exit"]},"Area Anak & Fasilitas":{icon:"\u{1F9F8}",bg:"#FEF3C7",color:"#92400E",border:"#FDE68A",rooms:["Lobby 2","Playground","Musholla","Tempat Wudhu","Ruang Team"]},"Gudang & Teknis":{icon:"\u2699\uFE0F",bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA",rooms:["Gudang Farmasi & Bahan","RoofTop","Ruang Kompressor","Area Genset"]},Eksterior:{icon:"\u{1F333}",bg:"#DCFCE7",color:"#15803D",border:"#86EFAC",rooms:["Taman","Pos Security","Toilet Security","Halaman Parkir"]}};function Lt(e){for(let[t,i]of Object.entries(Ne))if(i.rooms.includes(e))return{zoneName:t,...i};return{zoneName:"Lainnya",icon:"\u{1F4CD}",bg:"#F1F5F9",color:"#475569",border:"#E2E8F0"}}async function ma(e){e.innerHTML=`
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
        ${Object.entries(Ne).map(([u,w])=>`
          <button class="zone-tab-btn" data-zone="${u}" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
            <span>${w.icon}</span> ${u}
            <span class="badge" style="background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-${u.replace(/[\s&]+/g,"-")}">-</span>
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
  `;let t="all",i="table",n=[],r=ie(),d=r&&["superadmin","admin","manager","spv"].includes(r.role);async function o(){try{let u=await F("/api/hygiene-standards?all=1");if(!u.ok)throw new Error("Gagal memuat standar hygiene");n=Array.isArray(u.data)?u.data:u.data?.data||[],a(),s(),v()}catch(u){console.error(u),R(u.message||"Terjadi kesalahan memuat data");let w=e.querySelector("#hygiene-view-container");w&&(w.innerHTML=`
          <div class="empty-state" style="text-align:center; padding:40px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:12px;">\u26A0\uFE0F</div>
            <h3 style="margin:0 0 8px 0; color:var(--text-1);">Gagal Memuat Data</h3>
            <p style="color:var(--text-3); margin-bottom:16px;">${u.message}</p>
            <button id="btn-retry-hygiene" class="btn btn-primary" style="padding:8px 16px; border-radius:8px;">Coba Lagi</button>
          </div>
        `,w.querySelector("#btn-retry-hygiene")?.addEventListener("click",o))}}function s(){let u=e.querySelector("#badge-count-all");u&&(u.textContent=n.length),Object.entries(Ne).forEach(([C,_])=>{let p=C.replace(/[\s&]+/g,"-"),T=e.querySelector(`#badge-count-${p}`);if(T){let D=n.filter(A=>_.rooms.includes(A.room_name)).length;T.textContent=D}});let w=new Set(n.map(C=>C.room_name)),b=e.querySelector("#card-total-rooms");b&&(b.textContent=w.size)}function a(){let u=e.querySelector("#hygiene-room-filter");if(!u)return;let w=u.value,b=[];if(t==="all")b=Array.from(new Set(n.map(C=>C.room_name))).filter(Boolean).sort();else{let C=Ne[t];C&&(b=C.rooms.filter(_=>n.some(p=>p.room_name===_)))}u.innerHTML=`<option value="">Semua Ruangan (${b.length})</option>`+b.map(C=>`<option value="${C}" ${C===w?"selected":""}>${C}</option>`).join("")}function l(){let u=(e.querySelector("#hygiene-search")?.value||"").toLowerCase().trim(),w=e.querySelector("#hygiene-room-filter")?.value||"";return n.filter(b=>{if(t!=="all"){let C=Ne[t];if(!C||!C.rooms.includes(b.room_name))return!1}if(w&&b.room_name!==w)return!1;if(u){let C=(b.room_name||"").toLowerCase(),_=(b.item_name||"").toLowerCase(),p=(b.cleanliness_standard||"").toLowerCase();if(!C.includes(u)&&!_.includes(u)&&!p.includes(u))return!1}return!0})}function v(){let u=l(),w=e.querySelector("#hygiene-stats-text");w&&(w.innerHTML=`Menampilkan <strong>${u.length}</strong> dari <strong>${n.length}</strong> parameter`);let b=e.querySelector("#hygiene-view-container");if(b){if(u.length===0){b.innerHTML=`
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
      `,b.querySelector("#btn-clear-search-empty")?.addEventListener("click",()=>{let C=e.querySelector("#hygiene-search"),_=e.querySelector("#hygiene-room-filter");C&&(C.value=""),_&&(_.value=""),t="all",e.querySelectorAll(".zone-tab-btn").forEach(p=>p.classList.toggle("active",p.dataset.zone==="all")),e.querySelectorAll(".zone-tab-btn").forEach(p=>{p.dataset.zone==="all"?(p.style.background="#2563EB",p.style.color="#fff"):(p.style.background="#F1F5F9",p.style.color="var(--text-2)")}),a(),v()});return}i==="table"?x(b,u):h(b,u)}}function x(u,w){u.innerHTML=`
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse:collapse; text-align:left; font-size:0.9rem;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:2px solid #E2E8F0; color:#475569; font-size:0.78rem; text-transform:uppercase; letter-spacing:0.6px;">
                <th style="padding:14px 16px; width:50px; text-align:center;">No</th>
                <th style="padding:14px 16px; width:220px;">Ruangan & Zona</th>
                <th style="padding:14px 16px; width:280px;">Item / Objek Inspeksi</th>
                <th style="padding:14px 16px;">Standar Kebersihan Fisik</th>
                ${d?'<th style="padding:14px 16px; width:100px; text-align:center;">Aksi</th>':""}
              </tr>
            </thead>
            <tbody>
              ${w.map((b,C)=>{let _=Lt(b.room_name);return`
                  <tr style="border-bottom:1px solid #F1F5F9; transition:background-color 0.15s ease;" onmouseover="this.style.backgroundColor='#F8FAFC'" onmouseout="this.style.backgroundColor='transparent'">
                    <td style="padding:14px 16px; text-align:center; color:var(--text-3); font-weight:600; font-size:0.85rem;">
                      ${C+1}
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="margin-bottom:4px;">
                        <span style="display:inline-flex; align-items:center; gap:6px; background:${_.bg}; color:${_.color}; border:1px solid ${_.border}; padding:4px 10px; border-radius:6px; font-weight:700; font-size:0.83rem;">
                          <span>${_.icon}</span> ${b.room_name}
                        </span>
                      </div>
                      <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">
                        ${_.zoneName}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="font-weight:700; color:var(--text-1); line-height:1.45; font-size:0.9rem;">
                        ${b.item_name}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-left:3px solid #10B981; border-radius:6px; padding:10px 12px; line-height:1.55; color:#334155; font-size:0.875rem;">
                        <span style="color:#059669; font-weight:700; margin-right:4px;">\u2713</span> ${b.cleanliness_standard}
                      </div>
                    </td>
                    ${d?`
                      <td style="padding:14px 16px; text-align:center; vertical-align:middle;">
                        <div style="display:inline-flex; gap:6px;">
                          <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${b.id}" title="Edit Parameter" style="padding:6px 8px; border-radius:6px; color:#2563EB; cursor:pointer;">
                            \u270F\uFE0F
                          </button>
                          <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${b.id}" data-name="${b.item_name}" title="Hapus Parameter" style="padding:6px 8px; border-radius:6px; color:#DC2626; cursor:pointer;">
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
    `,m(u)}function h(u,w){let b={};w.forEach(C=>{let _=Lt(C.room_name);b[_.zoneName]||(b[_.zoneName]={info:_,rooms:{}}),b[_.zoneName].rooms[C.room_name]||(b[_.zoneName].rooms[C.room_name]=[]),b[_.zoneName].rooms[C.room_name].push(C)}),u.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:24px;">
        ${Object.entries(b).map(([C,_])=>{let p=Object.values(_.rooms).reduce((T,D)=>T+D.length,0);return`
            <div class="card" style="background:#fff; border:1px solid ${_.info.border}; border-radius:14px; overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.03);">
              <!-- Zone Card Header -->
              <div style="background:${_.info.bg}; border-bottom:1px solid ${_.info.border}; padding:14px 20px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="font-size:1.4rem;">${_.info.icon}</span>
                  <div>
                    <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:${_.info.color};">
                      ${C}
                    </h3>
                    <div style="font-size:0.78rem; color:${_.info.color}; opacity:0.85;">
                      ${Object.keys(_.rooms).length} Ruangan \xB7 ${p} Parameter Standar
                    </div>
                  </div>
                </div>
                <span class="badge" style="background:#fff; color:${_.info.color}; border:1px solid ${_.info.border}; font-weight:700; padding:4px 10px; border-radius:12px; font-size:0.8rem;">
                  ${p} Parameter
                </span>
              </div>

              <!-- Zone Rooms List -->
              <div style="padding:16px 20px;">
                ${Object.entries(_.rooms).map(([T,D])=>`
                  <div style="margin-bottom:20px; last-child:margin-bottom:0;">
                    <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                      <h4 style="margin:0; font-size:0.95rem; font-weight:700; color:var(--text-1);">
                        \u{1F6AA} ${T}
                      </h4>
                      <span class="badge badge-sm" style="background:#F1F5F9; color:var(--text-3); font-size:0.75rem;">
                        ${D.length} item
                      </span>
                    </div>

                    <div style="border:1px solid #E2E8F0; border-radius:8px; overflow:hidden;">
                      <table style="width:100%; border-collapse:collapse; font-size:0.875rem;">
                        <thead>
                          <tr style="background:#F8FAFC; border-bottom:1px solid #E2E8F0; color:#475569; font-size:0.75rem; text-transform:uppercase;">
                            <th style="padding:8px 12px; width:40px; text-align:center;">No</th>
                            <th style="padding:8px 12px; width:260px;">Item / Objek Inspeksi</th>
                            <th style="padding:8px 12px;">Standar Kebersihan Fisik</th>
                            ${d?'<th style="padding:8px 12px; width:80px; text-align:center;">Aksi</th>':""}
                          </tr>
                        </thead>
                        <tbody>
                          ${D.map((A,O)=>`
                            <tr style="border-bottom:1px solid #F1F5F9;">
                              <td style="padding:10px 12px; text-align:center; color:var(--text-3); font-size:0.8rem;">
                                ${O+1}
                              </td>
                              <td style="padding:10px 12px; font-weight:600; color:var(--text-1); vertical-align:top;">
                                ${A.item_name}
                              </td>
                              <td style="padding:10px 12px; vertical-align:top; color:#334155;">
                                <div style="display:flex; align-items:flex-start; gap:6px;">
                                  <span style="color:#10B981; font-weight:700;">\u2713</span>
                                  <span>${A.cleanliness_standard}</span>
                                </div>
                              </td>
                              ${d?`
                                <td style="padding:10px 12px; text-align:center; vertical-align:middle;">
                                  <div style="display:inline-flex; gap:4px;">
                                    <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${A.id}" title="Edit" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                                    <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${A.id}" data-name="${A.item_name}" title="Hapus" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
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
    `,m(u)}function m(u){u.querySelectorAll(".btn-edit-item").forEach(w=>{w.addEventListener("click",()=>{let b=w.dataset.id,C=n.find(_=>String(_.id)===String(b));C&&S(C)})}),u.querySelectorAll(".btn-delete-item").forEach(w=>{w.addEventListener("click",()=>{let b=w.dataset.id,C=w.dataset.name;pe(`Yakin ingin menghapus parameter standar <strong>"${C}"</strong>?`,async()=>{try{let _=await F(`/api/hygiene-standards/${b}`,{method:"DELETE"});if(!_.ok)throw new Error(_.data?.error||"Gagal menghapus parameter");j("Parameter kebersihan berhasil dihapus"),o()}catch(_){R(_.message)}},"Hapus Parameter Standar")})})}function S(u=null){let w=!!u,b=Array.from(new Set(n.map(_=>_.room_name))).filter(Boolean).sort(),C=document.createElement("div");C.innerHTML=`
      <form id="form-hygiene-item" style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Ruangan / Lokasi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-room-name" list="room-suggestions" class="form-control" placeholder="Contoh: Ruang Tindakan, Lobby 1, Toilet 1..." value="${u?.room_name||""}" required style="width:100%; height:40px; border-radius:8px;">
          <datalist id="room-suggestions">
            ${b.map(_=>`<option value="${_}">`).join("")}
          </datalist>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Pilih dari daftar atau ketik nama ruangan baru</span>
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Item / Objek Inspeksi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-item-name" class="form-control" placeholder="Contoh: Eksterior Dental Unit & Kursi, Lantai & Sudut..." value="${u?.item_name||""}" required style="width:100%; height:40px; border-radius:8px;">
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Standar Kebersihan Fisik <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="modal-standard" class="form-control" rows="4" placeholder="Deskripsi kondisi fisik kebersihan yang dipersyaratkan..." required style="width:100%; border-radius:8px; line-height:1.5;">${u?.cleanliness_standard||""}</textarea>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Jelaskan kriteria fisik: bebas debu kasat mata, steril, tidak licin, tidak berbau, dll.</span>
        </div>
      </form>
    `,ae({title:w?"\u270F\uFE0F Edit Parameter Kebersihan":"\u2795 Tambah Parameter Kebersihan",content:C,confirmText:w?"Simpan Perubahan":"Tambah Parameter",size:"md",onConfirm:async(_,p)=>{let T=C.querySelector("#modal-room-name").value.trim(),D=C.querySelector("#modal-item-name").value.trim(),A=C.querySelector("#modal-standard").value.trim();if(!T||!D||!A){R("Semua kolom wajib diisi");return}try{let O={room_name:T,item_name:D,cleanliness_standard:A},P=w?`/api/hygiene-standards/${u.id}`:"/api/hygiene-standards",B=await F(P,{method:w?"PUT":"POST",body:JSON.stringify(O)});if(!B.ok)throw new Error(B.data?.error||"Gagal menyimpan parameter");j(w?"Parameter berhasil diperbarui":"Parameter berhasil ditambahkan"),p(),o()}catch(O){R(O.message)}}})}function k(){let u=window.open("","_blank");if(!u){R("Pop-up printer diblokir oleh browser. Izinkan pop-up untuk mencetak SOP.");return}let b=`
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

        ${Object.entries(Ne).map(([C,_])=>{let p=n.filter(T=>_.rooms.includes(T.room_name));return p.length===0?"":`
            <div class="zone-title">${_.icon} ${C.toUpperCase()} (${p.length} PARAMETER)</div>
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
                ${p.map((T,D)=>`
                  <tr>
                    <td class="td-no">${D+1}</td>
                    <td class="td-room">${T.room_name}</td>
                    <td class="td-obj">${T.item_name}</td>
                    <td>${T.cleanliness_standard}</td>
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
    `;u.document.open(),u.document.write(b),u.document.close()}function g(){let u=n.map((b,C)=>{let _=Lt(b.room_name);return{No:C+1,"Kategori Zona":_.zoneName,"Ruangan / Lokasi":b.room_name,"Item / Objek Inspeksi":b.item_name,"Standar Kebersihan Fisik":b.cleanliness_standard}}),w=new Date().toISOString().slice(0,10);K(u,`Master_Standar_Kebersihan_FCMS_${w}`),j("Data master kebersihan berhasil diekspor ke Excel")}e.querySelector("#hygiene-search")?.addEventListener("input",()=>v()),e.querySelector("#hygiene-room-filter")?.addEventListener("change",()=>v()),e.querySelector("#btn-reset-filters")?.addEventListener("click",()=>{let u=e.querySelector("#hygiene-search"),w=e.querySelector("#hygiene-room-filter");u&&(u.value=""),w&&(w.value=""),v()}),e.querySelectorAll(".zone-tab-btn").forEach(u=>{u.addEventListener("click",()=>{t=u.dataset.zone,e.querySelectorAll(".zone-tab-btn").forEach(w=>{let b=w.dataset.zone===t;if(w.classList.toggle("active",b),b){w.style.background="#2563EB",w.style.color="#fff";let C=w.querySelector(".badge");C&&(C.style.background="rgba(255,255,255,0.25)",C.style.color="#fff")}else{w.style.background="#F1F5F9",w.style.color="var(--text-2)";let C=w.querySelector(".badge");C&&(C.style.background="#CBD5E1",C.style.color="#334155")}}),a(),v()})});let c=e.querySelector("#view-table-btn"),f=e.querySelector("#view-zone-btn");c?.addEventListener("click",()=>{i="table",c.style.background="#fff",c.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",f.style.background="transparent",f.style.boxShadow="none",v()}),f?.addEventListener("click",()=>{i="zone",f.style.background="#fff",f.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",c.style.background="transparent",c.style.boxShadow="none",v()}),e.querySelector("#btn-print-hygiene")?.addEventListener("click",k),e.querySelector("#btn-export-hygiene")?.addEventListener("click",g),e.querySelector("#btn-add-hygiene")?.addEventListener("click",()=>S(null)),o()}H();he();G();async function ua(e,t="supply"){return ln(e)}async function ln(e){let i=((await F("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name}));q({container:e,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:n=>n?new Date(n).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(n,r)=>r.branch_name_ref||r.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:n=>{try{let r=JSON.parse(n);return Array.isArray(r)?r.join(", "):n}catch{return n||"-"}}},{key:"chemical_items",label:"Chemical",render:n=>{try{let r=JSON.parse(n);return Array.isArray(r)?r.join(", "):n}catch{return n||"-"}}},{key:"additional_notes",label:"Catatan",render:n=>n?.length>40?n.slice(0,40)+"\u2026":n||"-"},{key:"status",label:"Status",render:n=>ee(n)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:n=>{let r=n?.tools_items;try{r=Array.isArray(JSON.parse(r))?JSON.parse(r).join(", "):r}catch{}let d=n?.chemical_items;try{d=Array.isArray(JSON.parse(d))?JSON.parse(d).join(", "):d}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:n?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:n?.branch_id&&!i.find(o=>o.value==n.branch_id)?[...i,{value:n.branch_id,label:n.branch_name||n.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:r},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:n?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:d},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:n?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:n?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:n?.status||""},{name:"processed_by",label:"Diproses Oleh",value:n?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async n=>{let r=new URLSearchParams(n||{}).toString(),d=await F(`/api/reports/supply?limit=10000&${r}`);if(d.ok){let o=d.data.data.map(s=>{let a=s.tools_items;try{a=Array.isArray(JSON.parse(a))?JSON.parse(a).join(", "):a}catch{}let l=s.chemical_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}return{Waktu:s.submitted_at||"",Pengirim:s.submitter_name||"",Cabang:s.branch_name_ref||s.branch_name||"","Alat/Barang":a||"",Chemical:l||"",Catatan:s.additional_notes||"",Status:s.status||"","Diproses Oleh":s.processed_by||""}});K(o,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async n=>{let d=(await F("/api/branches?all=1")).data?.data||[],o=v=>{if(!v)return null;let x=String(v||"").toLowerCase(),h=d.find(m=>String(m.full_name||"").toLowerCase()===x||String(m.code||"").toLowerCase()===x||String(m.name||"").toLowerCase()===x);return h?h.id:null},s=v=>{if(v==null||v==="")return"";if(v instanceof Date&&!isNaN(v.getTime()))return v.toISOString().slice(0,10);let x=String(v).trim();if(x===""||x==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(x))return x.slice(0,10);if(/^\d{4,5}$/.test(x)){let m=Number(x);if(m>2e4&&m<99999){let S=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}let h=x.split(/[\/\-\.]/);if(h.length===3){let[m,S,k]=h.map(g=>g.trim());if(m.length===4&&S.length<=2&&k.length<=2)return`${m}-${S.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&S.length<=2&&m.length<=2)return`${k}-${S.padStart(2,"0")}-${m.padStart(2,"0")}`}return x},a=n.map(v=>({submitted_at:s(v.Waktu||v.Tanggal),submitter_name:String(v.Pengirim||"").trim(),branch_id:o(String(v.Cabang||"").trim()),tools_items:String(v["Alat/Barang"]||v.Alat||"").trim(),chemical_items:String(v.Chemical||"").trim(),additional_notes:String(v.Catatan||v.Keterangan||"").trim(),status:String(v.Status||"").trim(),processed_by:String(v["Diproses Oleh"]||v.PIC||"").trim()})).filter(v=>v.submitted_at&&v.submitter_name&&v.branch_id),l=await F("/api/reports/supply/import",{method:"POST",body:JSON.stringify(a)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(n,r)=>{let d=ae({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(o,s)=>{let a=o.querySelector("#supply-status").value,l=o.querySelector("#supply-processed-by").value;(await F(`/api/reports/supply/${n.id}`,{method:"PUT",body:JSON.stringify({status:a,processed_by:l})})).ok?(j("Status diperbarui."),s(),r()):R("Gagal update status.")}})}}]})}H();G();async function ga(e){let t=ie();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}q({container:e,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",editor_khusus:"badge-warning",input_lembur:"badge-info",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let n=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!n,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!n,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"editor_khusus",label:"Editor (Hanya Lembur & Masalah)"},{value:"input_lembur",label:"Input Lembur (Hanya Isi Lembur)"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:n?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!n,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:n?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await F(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let n=i.data.data.map(r=>({"Nama Lengkap":r.full_name||"",Username:r.username||"",Email:r.email||"",Role:r.role||"",Status:r.is_active?"Aktif":"Nonaktif"}));K(n,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let n=i.map(d=>({full_name:String(d["Nama Lengkap"]||"").trim(),username:String(d.Username||"").trim(),email:String(d.Email||"").trim(),role:String(d.Role||"").trim()||"viewer",password:String(d.Password||"").trim()})).filter(d=>d.username&&d.password&&d.email&&d.full_name),r=await F("/api/users/import",{method:"POST",body:JSON.stringify(n)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}H();G();async function ba(e){q({container:e,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:t=>t?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:t=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:t?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:t?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:t?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:t?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:t?.is_active!==void 0?t.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let t=await F(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(t.ok)K(t.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async t=>{let i=t.map(r=>({code:String(r["Kode Cabang"]||"").trim(),name:String(r["Nama Pendek"]||"").trim(),full_name:String(r["Nama Lengkap"]||"").trim(),city:String(r.Kota||"").trim()})).filter(r=>r.code&&r.name),n=await F("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}H();async function ha(e){let t=new Date,i=[];e.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{t.setMonth(t.getMonth()-1),r()}),document.getElementById("cal-next").addEventListener("click",()=>{t.setMonth(t.getMonth()+1),r()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(o=>o.addEventListener("change",r));async function n(){try{let o=`${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,"0")}`;i=(await F(`/api/dashboard/calendar?month=${o}`)).data?.data||[]}catch(o){console.warn("[Calendar] Failed to load events, rendering empty grid:",o),i=[]}}async function r(){let o=document.getElementById("calendar-grid");if(o){o.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await n();try{let s=t.getFullYear(),a=t.getMonth(),l=t.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),v=document.getElementById("cal-month-label");v&&(v.textContent=l);let x=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(b=>b.value)),h=i.filter(b=>x.has(b.type)),m={};h.forEach(b=>{let C=dn(b.event_date);C&&(m[C]||(m[C]=[]),m[C].push(b))});let S=new Date(s,a,1).getDay(),k=new Date(s,a+1,0).getDate(),g=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],c=new Date().toISOString().slice(0,10),f='<div class="calendar-grid">';g.forEach(b=>{f+=`<div class="cal-day-header">${b}</div>`});for(let b=0;b<S;b++)f+='<div class="cal-cell cal-cell-empty"></div>';for(let b=1;b<=k;b++){let C=`${s}-${String(a+1).padStart(2,"0")}-${String(b).padStart(2,"0")}`,_=m[C]||[],p=C===c;f+=`
          <div class="cal-cell ${p?"cal-today":""} ${_.length?"cal-has-events":""}"
               data-date="${C}" tabindex="0" role="button" aria-label="${C}">
            <div class="cal-day-num ${p?"today-num":""}">${b}</div>
            <div class="cal-events-preview">
              ${_.slice(0,3).map(T=>`
                <div class="cal-event-dot cal-color-${T.color||"gray"}" title="${we(T.title||T.type)}">
                  <span class="cal-event-dot-label">${cn(T.title||T.branch_name||T.type,18)}</span>
                </div>
              `).join("")}
              ${_.length>3?`<div class="cal-more">+${_.length-3} lagi</div>`:""}
            </div>
          </div>`}let w=(S+k)%7;if(w!==0)for(let b=0;b<7-w;b++)f+='<div class="cal-cell cal-cell-empty"></div>';f+="</div>",o.innerHTML=f,o.querySelectorAll(".cal-cell[data-date]").forEach(b=>{b.addEventListener("click",()=>{let C=b.dataset.date,_=m[C]||[];if(!_.length)return;let p=document.getElementById("cal-event-list"),T=new Date(C+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=T,document.getElementById("cal-event-items").innerHTML=_.map(D=>`
            <div class="cal-event-item cal-color-border-${D.color||"gray"}">
              <div class="cal-event-type">
                ${pn(D.type)}
                ${D.date_type?` <span class="badge" style="font-size:0.75rem;padding:2px 6px;background:rgba(0,0,0,0.06);margin-left:4px;">${we(D.date_type)}</span>`:""}
              </div>
              <div class="cal-event-title" style="font-weight:600;font-size:0.95rem;margin-top:2px;">${we(D.title||"-")}</div>
              ${D.branch_name?`<div class="cal-event-branch" style="color:var(--text-2);font-size:0.85rem;margin-top:2px;">\u{1F4CD} ${we(D.branch_name)}</div>`:""}
              ${D.pic||D.period||D.opening_date||D.target_date||D.completion_date||D.notes?`
                <div class="cal-event-meta" style="margin-top:6px;padding-top:6px;border-top:1px dashed var(--border);font-size:0.8rem;color:var(--text-2);display:flex;flex-direction:column;gap:3px;">
                  ${D.pic?`<div>\u{1F464} <b>PIC:</b> ${we(D.pic)}</div>`:""}
                  ${D.period?`<div>\u{1F4C5} <b>Periode:</b> ${we(D.period)}</div>`:""}
                  ${D.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${we(Bt(D.opening_date))}</div>`:""}
                  ${D.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${we(Bt(D.target_date))}</div>`:""}
                  ${D.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${we(Bt(D.completion_date))}</div>`:""}
                  ${D.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${we(D.notes)}</div>`:""}
                </div>
              `:""}
              ${D.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${D.status==="Done"?"badge-success":"badge-warning"}">${we(D.status)}</span></div>`:""}
              ${D.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${D.days_remaining} hari</div>`:""}
            </div>
          `).join(""),p.style.display="block"})})}catch(s){console.error("[Calendar] Render error:",s),o&&(o.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let d=()=>{n().then(()=>r())};nt.on("data:changed",d),r()}function dn(e){if(!e||e==="-"||String(e).trim()==="")return"";let t=String(e).trim();if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);let i=t.split(/[\/\-\.]/);if(i.length===3){let[n,r,d]=i.map(o=>o.trim());if(d.length===4)return`${d}-${r.padStart(2,"0")}-${n.padStart(2,"0")}`;if(n.length===4)return`${n}-${r.padStart(2,"0")}-${d.padStart(2,"0")}`}return t.slice(0,10)}function Bt(e){if(!e||e==="-"||String(e).trim()==="")return"";let i=String(e).trim().split(/[\/\-\.]/);if(i.length===3){let[n,r,d]=i.map(o=>o.trim());if(n.length===4)return`${d.padStart(2,"0")}-${r.padStart(2,"0")}-${n}`;if(d.length===4)return`${n.padStart(2,"0")}-${r.padStart(2,"0")}-${d}`}return e}function cn(e,t){return e?e.length>t?e.slice(0,t)+"\u2026":e:""}function we(e){return e?String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function pn(e){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging"}[e]||e}H();async function fa(e){let t=ie(),i=(t?.full_name||t?.username||"U")[0].toUpperCase(),r={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[t?.role]||"#64748B";e.innerHTML=`
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
  `;let d=localStorage.getItem("fm_token"),o=document.getElementById("session-info");if(d&&o)try{let s=JSON.parse(atob(d.split(".")[1])),a=new Date(s.exp*1e3);o.textContent=`Berakhir: ${a.toLocaleString("id-ID")}`}catch{o.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async s=>{s.preventDefault();let a=document.getElementById("pwd-error"),l=document.getElementById("pwd-success"),v=document.getElementById("btn-save-pwd");a.style.display="none",l.style.display="none";let x=s.target,h=x.current_password.value,m=x.new_password.value,S=x.confirm_password.value;if(m!==S){a.textContent="\u274C Konfirmasi password tidak cocok.",a.style.display="block";return}if(m.length<6){a.textContent="\u274C Password baru minimal 6 karakter.",a.style.display="block";return}v.disabled=!0,v.textContent="\u23F3 Menyimpan...";let k=await F("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:h,new_password:m})});v.disabled=!1,v.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',k.ok?(l.textContent="\u2705 Password berhasil diubah.",l.style.display="block",x.reset(),j("Password berhasil diubah.")):(a.textContent=k.data?.error||"Gagal mengubah password.",a.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}H();var mt={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function re(e){if(e==null||e==="")return null;if(e instanceof Date)return isNaN(e.getTime())?null:e.toISOString().slice(0,10);let t=String(e).trim();if(t===""||t==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);if(/^\d{4,5}$/.test(t)){let r=Number(t);if(r>2e4&&r<99999){let d=new Date(Date.UTC(1899,11,30)+r*864e5);return isNaN(d.getTime())?null:d.toISOString().slice(0,10)}}let i=t.split(/[\/\-\.]/);if(i.length===3){let[r,d,o]=i.map(v=>v.trim()),s=Number(r),a=Number(d),l=Number(o);if(r.length===4&&s>1900)return`${r}-${d.padStart(2,"0")}-${o.padStart(2,"0")}`;if(o.length===4&&l>1900)return s>12?`${o}-${d.padStart(2,"0")}-${r.padStart(2,"0")}`:a>12?`${o}-${r.padStart(2,"0")}-${d.padStart(2,"0")}`:`${o}-${d.padStart(2,"0")}-${r.padStart(2,"0")}`;if(o.length===2&&!isNaN(l)){let v=l>=50?`19${o}`:`20${o}`;return s>12?`${v}-${d.padStart(2,"0")}-${r.padStart(2,"0")}`:`${v}-${d.padStart(2,"0")}-${r.padStart(2,"0")}`}}let n=new Date(t);return isNaN(n.getTime())?null:n.toISOString().slice(0,10)}function ya(e){return Object.values(e).every(t=>t==null||String(t).trim()==="")}var mn={validation:{required:[],map:e=>({cabang:e.CABANG,pic:e.PIC,kegiatan:e.KEGIATAN,quartal:e.QUARTAL,masa_pkwt:e["MASA PKWT"],pic_pelapor:e["PIC PELAPOR"],kontrak:e.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:e=>({name:e["Nama SOP"],category:e.Kategori||"Umum",document_link:e["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({full_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||e.Divisi||"FACILITY CARE",phone:e["No. Hp"]||e["No. HP"],join_date:re(e["Tanggal Masuk"]||e["Tgl Masuk"]),status:e.Status||"",target_pindah_os:re(e["Target Pindah OS"]),target_selesai_os:re(e["Target Selesai OS"]),notes:e.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({employee_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||"FACILITY CARE",start_date:re(e["Tanggal Mulai"]),end_date:re(e["Tanggal Selesai"]),contract_type:e["Tipe Kontrak"]||"",pkwt_number:e.PKWT||"",status:e.Status||"",notes:e.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:e=>({report_date:re(e["Tanggal Info"]),branch_name:e.Cabang,category:e.Kategori,source:e["Sumber Laporan"],complaint:e.Keluhan,employee_name:e["Nama FC"],fc_specialist:e["FC Spesialis"],solution:e.Solusi,status:e.Status||"",completion_date:re(e["Tanggal Selesai"])})},one_on_one:{required:[],map:e=>({meeting_date:re(e.Tanggal),branch_name:e.Cabang,employee_name:e["Nama Karyawan"],pic:e.Pic,problem:e.Masalah,solution:e.Solusi,status:e.Status||"",completion_date:re(e["Tanggal Selesai"]),document_link:e["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:e=>({branch_name:e.Cabang,activity_type:e.Kegiatan,period:e.Periode,pic:e.Pic||e.PIC,opening_date:re(e["Tanggal Opening"]||e["Tgl Opening"]),target_date:re(e["Tanggal Target"]||e["Tgl Target"]),completion_date:re(e["Tanggal Selesai"]||e["Tgl Selesai"]),status:e.Status||"",notes:e.Keterangan||e.Catatan})},inspection:{required:[],map:e=>({inspection_date:re(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",fc_score:e["Point FC SP"]!==void 0&&e["Point FC SP"]!==null?parseFloat(String(e["Point FC SP"]).replace(",",".")):null,spv_score:e["Point SPV"]!==void 0&&e["Point SPV"]!==null?parseFloat(String(e["Point SPV"]).replace(",",".")):null,document_link:e.Link,notes:""})},cleaning:{required:[],map:e=>({activity_date:re(e.Tanggal),branch_name:e.Cabang,activity_type:e["Jenis Kegiatan"]||"General Cleaning",period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},fogging:{required:[],map:e=>({activity_date:re(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:e=>({info_date:re(e["Tgl Info"]),branch_name:e.Cabang,problem:e.Permasalahan,pic:e.PIC,done_date:re(e["Tgl Done"]),status:e.Status||"",notes:e.Ket})},relievers:{required:[],map:e=>({branch_name:e.Cabang,original_fc_name:e["Nama Facility care"],period:e.Periode,reliever_name:e.Relifer,backup_date:re(e["Tanggal Back Up"]),completion_date:re(e["Tanggal Selesai"]),reason:e.Keterangan,shift:e.Shift,status:e.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:e=>({training_date:re(e.Tanggal),batch:e.Batch,subject:e.Materi,participants:e.Peserta,branch_name:e.Cabang,trainer:e.Trainer,score:e.Nilai!==void 0&&e.Nilai!==null?parseFloat(String(e.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:e=>({name:e["Master Checklist"],category:"Umum",document_link:e["Link Document"],description:""})},supply:{required:[],map:e=>({submitted_at:re(e.Timestamp),submitter_name:e["Nama Lengkap"],branch_name:e["Kebutuhan Untuk Cabang"],tools_items:e["Alat - Alat / Barang"],tools_quantity:e["Jumlah Permintaan Alat / Barang"],chemical_items:e.Chemical,chemical_quantity:e["Jumlah Permintaan Chemical"],additional_notes:e["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:e.Status||""})}};function un(e,t){let i=mt[e];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let n=mn[i.module];if(!n)return{valid:[],errors:[],mapped:[],skipped:!0};let r=[],d=[],o=[];return t.filter(a=>!ya(a)).forEach((a,l)=>{let v=t.indexOf(a)+2,x=[];n.required.forEach(({key:m,label:S})=>{let k=a[m];if(k==null||String(k).trim()===""){let g=Object.keys(a).filter(c=>c.trim()).join(", ");x.push({column:S,originalValue:k||"",reason:`Kolom "${S}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${g.slice(0,120)}`})}});let h=n.map(a);x.length>0?d.push({row:v,data:h,raw:a,errors:x}):(r.push(a),o.push(h))}),{valid:r,errors:d,mapped:o}}function va(e){let t=[];return e.SheetNames.forEach(i=>{let n=mt[i];if(!n)return;let r=e.Sheets[i],d=window.XLSX.utils.sheet_to_json(r,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),o=un(i,d),s=d.filter(a=>!ya(a));t.push({sheetName:i,module:n.module,label:n.label,total:s.length,valid:o.mapped.length,errorCount:o.errors.length,errors:o.errors,mapped:o.mapped,skipped:!1})}),t}function ka(){let e=window.XLSX,t=e.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([n,r])=>{e.utils.book_append_sheet(t,e.utils.json_to_sheet(r),n)}),e.writeFile(t,"Template_Import_Data_Awal_FCMS.xlsx")}function xa(e){let t=window.XLSX,i=t.utils.book_new(),n=!1;return e.forEach(r=>{if(!r.errors||r.errors.length===0)return;n=!0;let d=r.errors.map(s=>({"No. Baris":s.row,"Kolom Gagal":(s.errors||[]).map(a=>a.column||a).join("; "),"Alasan Error":(s.errors||[]).map(a=>a.reason||a).join("; "),...Object.fromEntries(Object.entries(s.data||{}).map(([a,l])=>[a,l??""]))})),o=t.utils.json_to_sheet(d);t.utils.book_append_sheet(i,o,r.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),n?(t.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var gn=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","sop","inspection","cleaning","fogging","basecamp","supply"];function wa(e){e.innerHTML=`
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
              ${Object.entries(mt).map(([k,{label:g}])=>`<span class="import-sheet-tag">\u{1F4C4} ${k} \u2192 ${g}</span>`).join("")}
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
  `;let t=null,i=null,n=0,r={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function d(k){Object.entries(r).forEach(([g,c])=>{c.style.display=g===k?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let k=document.getElementById("btn-backup-db");k.disabled=!0,k.textContent="\u23F3 Memproses Backup...";try{let g=await F("/api/import/backup");if(g.ok){if(!window.XLSX){R("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let c=window.XLSX,f=c.utils.book_new();Object.entries(g.data.database).forEach(([u,w])=>{let b=w.length>0?w:[{}],C=c.utils.json_to_sheet(b);c.utils.book_append_sheet(f,C,u.substring(0,31))}),c.writeFile(f,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),j("Backup berhasil diunduh!")}else R("Gagal memproses backup: "+(g.data?.error||"Unknown error"))}catch(g){R("Gagal memproses backup: "+g.message)}finally{k.disabled=!1,k.textContent="\u{1F4E6} Backup Database"}});let o=document.getElementById("btn-sync-google");o&&o.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let k=o.innerHTML;o.innerHTML='<span class="spinner"></span> Menyinkronkan...',o.disabled=!0;try{let g=await F("/api/sync/google-sheets",{method:"POST"});g.ok?alert("Sinkronisasi Berhasil: "+(g.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(g.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{o.innerHTML=k,o.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{ka(),j("Template Excel berhasil didownload!")});let s=document.getElementById("file-input"),a=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",k=>{k.stopPropagation(),s.click()}),s.addEventListener("change",k=>{k.target.files[0]&&l(k.target.files[0])}),a.addEventListener("dragover",k=>{k.preventDefault(),a.classList.add("drag-over")}),a.addEventListener("dragleave",()=>a.classList.remove("drag-over")),a.addEventListener("drop",k=>{k.preventDefault(),a.classList.remove("drag-over");let g=k.dataTransfer.files[0];g&&g.name.match(/\.xlsx?$/i)?l(g):R("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{t=null,s.value="",document.getElementById("file-info").style.display="none",a.style.display="",d("upload")});async function l(k){t=k,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${k.name} (${(k.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",a.style.display="none",await v(k)}async function v(k){d("validating");let g=document.getElementById("validation-status"),c=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");g.textContent="Membaca file Excel...",c.style.width="20%",await We(200);let f=await k.arrayBuffer(),u=window.XLSX.read(f,{type:"array",cellDates:!0});g.textContent=`Memvalidasi ${u.SheetNames.length} sheet...`,c.style.width="50%",await We(100),i=va(u),c.style.width="100%",g.textContent="Validasi selesai!",await We(300),x()}catch(f){d("upload"),R("Gagal memproses file: "+f.message),document.getElementById("file-info").style.display="flex",a.style.display="none"}}function x(){d("preview");let k=i.filter(p=>!p.skipped).length,g=i.reduce((p,T)=>p+T.total,0),c=i.reduce((p,T)=>p+T.valid,0),f=i.reduce((p,T)=>p+T.errorCount,0),u=g>0?Math.round(c/g*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${k} sheet</span>
      <span class="badge badge-secondary">${g} baris</span>
      <span class="badge badge-success">${c} valid (${u}%)</span>
      ${f>0?`<span class="badge badge-danger">${f} error</span>`:""}
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
          ${i.map((p,T)=>`
            <tr class="${p.errorCount>0?"row-error":p.skipped?"row-skipped":"row-ok"}">
              <td><strong>${p.sheetName}</strong></td>
              <td>${p.label}</td>
              <td style="text-align:center">${p.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${p.valid}</span></td>
              <td style="text-align:center">${p.errorCount>0?`<span class="badge badge-danger">${p.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${p.skipped?'<span class="badge badge-neutral">Dilewati</span>':p.errorCount>0&&p.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':p.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':p.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${p.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${T}">\u{1F50D} ${p.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,w.querySelectorAll(".btn-detail-error").forEach(p=>{p.addEventListener("click",()=>{let T=i[Number(p.dataset.idx)];h(T)})});let b=document.getElementById("error-detail-section"),C=document.getElementById("error-detail-container");C.innerHTML="",b.style.display="none";let _=document.getElementById("btn-start-import");c===0?(_.disabled=!0,_.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(_.disabled=!1,f>0?(_.innerHTML=`\u{1F680} Import ${c} Data Valid (${f} dilewati)`,_.title="Baris error akan dilewati, baris valid tetap diimport"):_.innerHTML=`\u{1F680} Mulai Import ${c} Data`)}function h(k){let g=document.getElementById("error-detail-section"),c=document.getElementById("error-detail-container");g.style.display="";let f=k.errors.slice(0,100).map(u=>(Array.isArray(u.errors)?u.errors:[]).map(b=>{let C=typeof b=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${u.row}</span></td>
            <td><strong>${C?b.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${C&&b.originalValue!==void 0?b.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${C?b.reason:b}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${C&&b.aliases?`Gunakan salah satu nama kolom:<br><em>${b.aliases}</em>`:C&&b.hint?b.hint:""}
            </td>
          </tr>
        `}).join("")).join("");c.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${k.sheetName} \u2014 ${k.errorCount} baris error dari ${k.total} total
          ${k.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
            <tbody>${f||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${k.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,g.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{d("upload"),document.getElementById("file-info").style.display="none",a.style.display="",t=null,s.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;xa(i)?j("Log error berhasil didownload."):j("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let k=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";m(k)});async function m(k){d("importing"),n=Date.now();let g=[];gn.forEach(b=>{let C=i?.find(_=>_.module===b&&_.mapped?.length>0);C&&g.push(C)});let c=document.getElementById("import-steps-list");c.innerHTML=g.map(b=>`
      <div class="import-step-item" id="step-item-${b.module}">
        <span class="step-item-icon" id="step-icon-${b.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${b.label} <span class="step-item-count">(${b.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${b.module}"></span>
      </div>
    `).join("");let f=document.getElementById("import-bar"),u=document.getElementById("import-current-status"),w={totalSheets:g.length,totalRows:g.reduce((b,C)=>b+C.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let b=0;b<g.length;b++){let C=g[b],_=document.getElementById(`step-icon-${C.module}`),p=document.getElementById(`step-status-${C.module}`);_.textContent="\u{1F504}",p.textContent="Mengimport...",u.textContent=`Mengimport ${C.label}...`,f.style.width=`${Math.round(b/g.length*100)}%`;try{let T=await F(`/api/import/${C.module}`,{method:"POST",body:JSON.stringify({rows:C.mapped,onDuplicate:k})});if(T.ok){let D=T.data;w.inserted+=D.inserted||0,w.skipped+=D.skipped||0,w.moduleResults.push({label:C.label,inserted:D.inserted||0,skipped:D.skipped||0,status:"ok"}),_.textContent="\u2705",p.innerHTML=`<span class="badge badge-success">${D.inserted||0} berhasil</span>${D.skipped>0?` <span class="badge badge-neutral">${D.skipped} skip</span>`:""}`}else w.failed++,w.moduleResults.push({label:C.label,inserted:0,skipped:0,status:"error",error:T.data?.error}),_.textContent="\u274C",p.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(T){w.failed++,w.moduleResults.push({label:C.label,inserted:0,skipped:0,status:"error",error:T.message}),_.textContent="\u274C",p.innerHTML='<span class="badge badge-danger">Gagal</span>'}await We(150)}f.style.width="100%",u.textContent="Selesai!",await We(400),S(w)}function S(k){d("summary");let g=((Date.now()-n)/1e3).toFixed(1),c=k.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${c?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${c?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${k.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${k.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${k.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${k.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${k.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${k.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
      <div class="summary-stat-card info">
        <div class="stat-value">${g}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${k.moduleResults.map(f=>`
            <tr>
              <td>${f.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${f.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${f.skipped}</span></td>
              <td style="text-align:center">
                ${f.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${f.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{t=null,i=null,s.value="",document.getElementById("file-info").style.display="none",a.style.display="",d("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function We(e){return new Promise(t=>setTimeout(t,e))}H();var ut=[],Sa=[];async function _a(e){ut=await W(),Sa=await se(),q({container:e,title:"Data SP (Surat Peringatan)",icon:"\u{1F4DC}",apiPath:"/api/sp",enableMobileFilterSheet:!0,itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:t=>t?`<span class="badge badge-info">${t}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:t=>t?new Date(t).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:t=>t?new Date(t).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:t=>`<span class="badge badge-warning">${t||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:t=>t?`<a href="${t}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:ut}],exportOptions:{moduleName:"sp_data",onExport:async t=>{let i=new URLSearchParams(t||{}).toString(),n=await F(`/api/sp?limit=10000&${i}`);if(n.ok){let r=n.data.data.map(o=>({"Nama Karyawan":o.employee_name||"",Divisi:o.division||"",Cabang:o.branch_name||"","Tanggal Sp":o.tanggal||"","Akhir Sp":o.akhir_sp||"","Jenis Sp":o.sp_type||"","Link Document / Foto":o.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(G(),me));d(r,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let t=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(G(),me));i(t,"Template_Import_SP")},onImport:async t=>{let i=o=>{if(!o)return null;let s=String(o||"").toLowerCase(),a=ut.find(l=>String(l.label||"").toLowerCase()===s);return a?a.value:null},n=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let s=String(o).trim();if(/^\d{4,5}$/.test(s)){let l=Number(s);if(l>2e4&&l<99999){let v=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);let a=s.split(/[\/\-\.]/);if(a.length===3){let[l,v,x]=a.map(h=>h.trim());if(l.length===4&&v.length<=2&&x.length<=2)return`${l}-${v.padStart(2,"0")}-${x.padStart(2,"0")}`;if(x.length===4&&v.length<=2&&l.length<=2)return`${x}-${v.padStart(2,"0")}-${l.padStart(2,"0")}`}return s},r=t.map(o=>({employee_name:String(o["Nama Karyawan"]||"").trim(),division:String(o.Divisi||"").trim(),branch_id:i(String(o.Cabang||"").trim()),tanggal:n(o["Tanggal Sp"]),akhir_sp:n(o["Akhir Sp"]),sp_type:String(o["Jenis Sp"]||"").trim(),document_link:String(o["Link Document / Foto"]||"").trim()})).filter(o=>o.employee_name&&o.branch_id),d=await F("/api/import/sp",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:[{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:Sa},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:ut,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}H();var Oe=[],Ca=[];async function Ea(e){Oe=await W(),Ca=await se(),q({container:e,title:"Data Mutasi",icon:"\u{1F501}",apiPath:"/api/mutasi",enableMobileFilterSheet:!0,itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:t=>t?new Date(t).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:t=>`<span class="badge ${t==="Selesai"?"badge-success":"badge-warning"}">${t||"-"}</span>`},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:Oe},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:Oe}],exportOptions:{moduleName:"mutasi_data",onExport:async t=>{let i=new URLSearchParams(t||{}).toString(),n=await F(`/api/mutasi?limit=10000&${i}`);if(n.ok){let r=n.data.data.map(o=>({Tanggal:o.tanggal||"","Nama Karyawan":o.employee_name||"","Cabang Asal":o.from_branch_name||"","Cabang Tujuan":o.to_branch_name||"",Status:o.status||"",Dokumen:o.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(G(),me));d(r,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let t=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(G(),me));i(t,"Template_Import_Mutasi")},onImport:async t=>{let i=o=>{if(!o)return null;let s=String(o||"").toLowerCase(),a=Oe.find(l=>String(l.label||"").toLowerCase()===s);return a?a.value:null},n=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let s=String(o).trim();if(/^\d{4,5}$/.test(s)){let l=Number(s);if(l>2e4&&l<99999){let v=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);let a=s.split(/[\/\-\.]/);if(a.length===3){let[l,v,x]=a.map(h=>h.trim());if(l.length===4&&v.length<=2&&x.length<=2)return`${l}-${v.padStart(2,"0")}-${x.padStart(2,"0")}`;if(x.length===4&&v.length<=2&&l.length<=2)return`${x}-${v.padStart(2,"0")}-${l.padStart(2,"0")}`}return s},r=t.map(o=>({tanggal:n(o.Tanggal),employee_name:String(o["Nama Karyawan"]||"").trim(),from_branch_id:i(String(o["Cabang Asal"]||"").trim()),to_branch_id:i(String(o["Cabang Tujuan"]||"").trim()),status:String(o.Status||"").trim(),document_link:String(o.Dokumen||"").trim()})).filter(o=>o.tanggal&&o.employee_name&&o.from_branch_id&&o.to_branch_id),d=await F("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:Ca},{type:"select",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Oe,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Oe,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}H();var Pt=[],$a=[];async function Ta(e){Pt=await W(),$a=await Te(),q({container:e,title:"Data Lembur",icon:"\u23F1\uFE0F",apiPath:"/api/overtime",enableMobileFilterSheet:!0,itemLabel:"Lembur",bulkDelete:!0,columns:[{key:"date",label:"Tanggal",render:t=>t?new Date(t).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"branch_name",label:"Cabang Lembur"},{key:"employee_name",label:"Nama Karyawan"},{key:"start_time",label:"Jam Mulai"},{key:"end_time",label:"Jam Selesai"},{key:"break_hours",label:"Istirahat",render:t=>t!=null?`${t} Jam`:"0 Jam"},{key:"total_hours",label:"Total Jam",render:t=>t!=null?`${t} Jam`:"0 Jam"},{key:"reason",label:"Alasan / Ket Lembur"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:Pt}],exportOptions:{moduleName:"overtime_records",onExport:async t=>{let i=new URLSearchParams(t||{}).toString(),n=await F(`/api/overtime?limit=10000&${i}`);if(n.ok){let r=n.data.data.map(o=>({Tanggal:o.date||"","Cabang Lembur":o.branch_name||"","Nama Karyawan":o.employee_name||"","Jam Mulai":o.start_time||"","Jam Selesai":o.end_time||"","Istirahat (Jam)":o.break_hours||0,"Total Jam":o.total_hours||0,"Alasan / Ket Lembur":o.reason||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(G(),me));d(r,`Data_Lembur_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")}},formFields:[{type:"date",name:"date",label:"Tanggal",required:!0},{type:"select",name:"branch_id",label:"Cabang Lembur",required:!0,options:Pt},{type:"select",name:"employee_id",label:"Nama Karyawan",required:!0,options:$a},{type:"time",name:"start_time",label:"Jam Mulai",required:!0},{type:"time",name:"end_time",label:"Jam Selesai",required:!0},{type:"number",name:"break_hours",label:"Istirahat (Jam)",required:!0},{type:"number",name:"total_hours",label:"Total Jam (Jam)",required:!0},{type:"select",name:"reason",label:"Alasan / Ket Lembur",required:!0,options:["Cover manpower","Longshift","Pekerjaan urgent","General Cleaning","Deep Cleaning"]}]})}window._overtimeListenerAttached||(window._overtimeListenerAttached=!0,document.body.addEventListener("input",e=>{if(e.target.name==="start_time"||e.target.name==="end_time"||e.target.name==="break_hours"){let t=e.target.closest("form");if(!t)return;let i=t.querySelector('[name="start_time"]'),n=t.querySelector('[name="end_time"]'),r=t.querySelector('[name="break_hours"]'),d=t.querySelector('[name="total_hours"]');if(!i||!n||!r||!d)return;let o=i.value,s=n.value,a=parseFloat(r.value)||0;if(o&&s){let[l,v]=o.split(":").map(Number),[x,h]=s.split(":").map(Number),m=x*60+h-(l*60+v);m<0&&(m+=24*60);let S=m/60-a;S<0&&(S=0),d.value=Number.isInteger(S)?S:S.toFixed(2)}}}));H();async function Fa(e){let t=ie();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}q({container:e,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:i=>new Date(i).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(i,n)=>`<strong>${i||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${n.user_role||""}</small>`},{key:"action",label:"Aksi",render:i=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[i]||"badge-neutral"}">${i}</span>`},{key:"module",label:"Modul",render:i=>`<span style="text-transform:capitalize">${(i||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(i,n)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${i}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async i=>{try{let d=((await(await fetch(`/api/audit-logs?search=${i}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(l=>String(l.id)===String(i));if(!d)return alert("Data tidak ditemukan");let o=l=>{if(!l)return"Tidak ada data";try{return JSON.stringify(JSON.parse(l),null,2)}catch{return l}},s=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${o(d.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${o(d.new_data)}</pre>
           </div>
         </div>
       `,{createModal:a}=await Promise.resolve().then(()=>(he(),vt));a({title:`Detail Audit Log #${i}`,content:s,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}H();he();G();var Mt={"2026-09":"September 2026","2026-10":"Oktober 2026","2026-11":"November 2026","2026-12":"Desember 2026"},Qe={scheduled:{label:"Terjadwal",badgeClass:"badge-warning",icon:"\u{1F7E1}",bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},in_progress:{label:"Sedang Proses",badgeClass:"badge-info",icon:"\u{1F535}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},completed:{label:"Selesai",badgeClass:"badge-success",icon:"\u{1F7E2}",bg:"#DCFCE7",color:"#166534",border:"#86EFAC"}};async function Da(e){e.innerHTML=`
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
  `;let t="board",i=null,n=[],r=ie(),d=r&&["superadmin","admin","manager","spv"].includes(r.role);F("/api/branches?all=1").then(c=>{c.ok&&c.data?.data&&(n=c.data.data)}).catch(()=>{});async function o(){try{let c=await F("/api/os-migrations/summary");if(!c.ok)throw new Error("Gagal memuat data migrasi");i=c.data?.data||c.data,s(i),l()}catch(c){document.getElementById("migration-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${c.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function s(c){let f=document.getElementById("migration-summary-cards");if(!f||!c)return;let u=c.total_branches||0,w=c.vendors?.AEON?.total||0,b=c.vendors?.["Triple S"]?.total||0,C=c.by_status?.completed||0,_=c.by_status?.in_progress||0,p=c.by_status?.scheduled||0,T=u>0?Math.round(C/u*100):0;f.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${u} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.8rem; font-weight:800; color:#1E40AF; line-height:1.2;">${w} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 9 \u2022 Nov: 3 \u2022 Des: 8</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #0D9488; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#0F766E; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor Triple S</div>
        <div style="font-size:1.8rem; font-weight:800; color:#0F766E; line-height:1.2;">${b} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 5 \u2022 Okt: 5 \u2022 Nov: 5 \u2022 Des: 27</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #16A34A; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <span style="font-size:0.8rem; font-weight:600; color:#15803D; text-transform:uppercase; letter-spacing:0.5px;">Progres Realisasi</span>
          <span style="font-size:0.85rem; font-weight:700; color:#15803D;">${T}%</span>
        </div>
        <div style="width:100%; height:8px; background:#E2E8F0; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${T}%; height:100%; background:#16A34A; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.75rem; color:var(--text-2); font-weight:500;">
          <span>\u{1F7E2} Selesai: <strong>${C}</strong></span>
          <span>\u{1F535} Proses: <strong>${_}</strong></span>
          <span>\u{1F7E1} Terjadwal: <strong>${p}</strong></span>
        </div>
      </div>
    `}function a(){if(!i||!i.matrix)return[];let c=(document.getElementById("migration-search")?.value||"").toLowerCase().trim(),f=document.getElementById("filter-vendor")?.value||"",u=document.getElementById("filter-month")?.value||"",w=document.getElementById("filter-status")?.value||"";return i.matrix.filter(b=>!(f&&b.vendor!==f||u&&b.target_month!==u||w&&b.status!==w||c&&!`${b.branch_code} ${b.branch_name} ${b.branch_full_name} ${b.vendor} ${b.notes||""} ${b.pic_vendor||""}`.toLowerCase().includes(c)))}function l(){let c=a(),f=document.getElementById("filter-stats-text");f&&(f.innerHTML=`Menampilkan <strong>${c.length}</strong> dari <strong>${i?.total_branches||0}</strong> cabang`);let u=document.getElementById("migration-view-container");t==="board"?v(u,c):x(u,c)}function v(c,f){if(!f.length){c.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;return}let u=["AEON","Triple S"],w=["2026-09","2026-10","2026-11","2026-12"],b="";u.forEach(C=>{let _=f.filter(P=>P.vendor===C);if(document.getElementById("filter-vendor")?.value&&document.getElementById("filter-vendor").value!==C||_.length===0)return;let p=_.length,T=C==="AEON"?"#1E40AF":"#0F766E",D=C==="AEON"?"#EFF6FF":"#F0FDFA",A=C==="AEON"?"#BFDBFE":"#99F6E4",O=w.filter(P=>_.some(I=>I.target_month===P));b+=`
        <div class="vendor-board-section" style="margin-bottom:32px; background:#fff; border:1px solid var(--border); border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <!-- Vendor Section Header -->
          <div style="padding:14px 20px; background:${D}; border-bottom:1px solid ${A}; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.1rem; font-weight:800; color:${T};">${C==="AEON"?"\u{1F3E2} AEON":"\u{1F6E1}\uFE0F Triple S"}</span>
              <span class="badge" style="background:#fff; color:${T}; border:1px solid ${A}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px;">
                ${p} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
              Format Timeline Berdasarkan Dokumen Migrasi
            </div>
          </div>

          <!-- Months Grid (Hanya menampilkan bulan yang memiliki jadwal) -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px; padding:18px;">
            ${O.map(P=>{let I=_.filter(N=>N.target_month===P);return`
                <div class="month-column" style="background:#F8FAFC; border:1px solid var(--border); border-radius:10px; display:flex; flex-direction:column; min-height:220px;">
                  <!-- Month Header -->
                  <div style="padding:12px 14px; background:#F1F5F9; border-bottom:1px solid var(--border); border-radius:9px 9px 0 0; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">
                      \u{1F4C5} ${(Mt[P]||P).split(" ")[0]}
                    </div>
                    <span class="badge" style="background:#E2E8F0; color:var(--text-2); font-weight:700; font-size:0.75rem; padding:2px 7px; border-radius:4px;">
                      ${I.length}
                    </span>
                  </div>

                  <!-- Branch Cards Container -->
                  <div style="padding:10px; display:flex; flex-direction:column; gap:8px; flex:1; max-height:480px; overflow-y:auto;">
                    ${I.map(N=>{let y=Qe[N.status]||Qe.scheduled;return`
                        <div class="branch-card" style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 12px; box-shadow:0 1px 2px rgba(0,0,0,0.03); transition:transform 0.15s, box-shadow 0.15s;">
                          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px; margin-bottom:6px;">
                            <div style="font-weight:700; font-size:0.9rem; color:var(--text-1); line-height:1.3;">
                              <span style="color:#2563EB; font-family:monospace; margin-right:4px;">${N.branch_code}.</span>${N.branch_name}
                            </div>
                            ${d?`
                              <div style="display:flex; gap:2px;">
                                <button class="btn-icon-mini edit-card-btn" data-id="${N.id}" title="Edit" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--text-3);">\u270F\uFE0F</button>
                                <button class="btn-icon-mini delete-card-btn" data-id="${N.id}" data-name="${N.branch_name}" title="Hapus" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--danger);">\u{1F5D1}\uFE0F</button>
                              </div>
                            `:""}
                          </div>

                          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px; margin-top:8px;">
                            <!-- Status Selector Dropdown -->
                            <div style="position:relative;">
                              <button class="status-toggle-btn" data-id="${N.id}" data-status="${N.status}" style="background:${y.bg}; color:${y.color}; border:1px solid ${y.border}; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; display:inline-flex; align-items:center; gap:4px; cursor:pointer;" title="Klik untuk ubah status">
                                ${y.icon} ${y.label} \u25BE
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
      `}),c.innerHTML=b,h(c)}function x(c,f){if(!f.length){c.innerHTML=`
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
                ${d?'<th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${f.map((u,w)=>{let b=Qe[u.status]||Qe.scheduled;return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3);">${w+1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">${u.branch_full_name}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${u.vendor==="AEON"?"#DBEAFE":"#CCFBF1"}; color:${u.vendor==="AEON"?"#1E40AF":"#0F766E"}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${u.vendor}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:var(--text-2);">
                      ${Mt[u.target_month]||u.target_month}
                    </td>
                    <td style="padding:12px 16px;">
                      <button class="status-toggle-btn" data-id="${u.id}" data-status="${u.status}" style="background:${b.bg}; color:${b.color}; border:1px solid ${b.border}; padding:4px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; cursor:pointer;" title="Klik untuk ubah status">
                        ${b.icon} ${b.label} \u25BE
                      </button>
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-2);">
                      ${u.actual_date||"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; font-weight:600; color:var(--text-2);">
                      ${u.employee_count>0?`\u{1F464} ${u.employee_count}`:"-"}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3); max-width:200px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${u.notes||"-"}
                    </td>
                    ${d?`
                      <td style="padding:12px 16px; text-align:right;">
                        <button class="btn btn-outline btn-sm edit-card-btn" data-id="${u.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-card-btn" data-id="${u.id}" data-name="${u.branch_name}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,h(c)}function h(c){c.querySelectorAll(".status-toggle-btn").forEach(f=>{f.addEventListener("click",u=>{u.stopPropagation();let w=f.dataset.id,b=f.dataset.status;m(f,w,b)})}),c.querySelectorAll(".edit-card-btn").forEach(f=>{f.addEventListener("click",()=>{let u=f.dataset.id,w=i?.matrix?.find(b=>String(b.id)===String(u));w&&S(w)})}),c.querySelectorAll(".delete-card-btn").forEach(f=>{f.addEventListener("click",()=>{let u=f.dataset.id,w=f.dataset.name||"Cabang";pe(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${w}</strong>?`,async()=>{try{if(!(await F(`/api/os-migrations/${u}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus");j(`Jadwal migrasi ${w} berhasil dihapus`),o()}catch(b){R(b.message)}})})})}function m(c,f,u){document.querySelectorAll(".status-popup-menu").forEach(p=>p.remove());let w=c.getBoundingClientRect(),b=document.createElement("div");b.className="status-popup-menu",b.style.cssText=`
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
    `;let C=[{key:"scheduled",label:"\u{1F7E1} Terjadwal"},{key:"in_progress",label:"\u{1F535} Sedang Proses"},{key:"completed",label:"\u{1F7E2} Selesai"}];b.innerHTML=C.map(p=>`
      <button class="status-option-item" data-key="${p.key}" style="
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
        color: ${p.key===u?"#2563EB":"var(--text-1)"};
        font-weight: ${p.key===u?"700":"500"};
      ">
        ${p.label} ${p.key===u?"\u2713":""}
      </button>
    `).join("");let _=p=>{!b.contains(p.target)&&p.target!==c&&(b.remove(),document.removeEventListener("click",_))};setTimeout(()=>document.addEventListener("click",_),50),b.querySelectorAll(".status-option-item").forEach(p=>{p.addEventListener("click",async()=>{let T=p.dataset.key;if(b.remove(),document.removeEventListener("click",_),T!==u)try{if(!(await F(`/api/os-migrations/${f}/status`,{method:"PATCH",body:JSON.stringify({status:T})})).ok)throw new Error("Gagal memperbarui status");j("Status berhasil diubah"),o()}catch(D){R(D.message)}})}),document.body.appendChild(b)}function S(c=null){let f=!!c,u=f?"Edit Jadwal Migrasi OS":"Tambah Jadwal Migrasi OS",b=`
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${n.map(C=>`
      <option value="${C.id}" ${c&&c.branch_id===C.id?"selected":""}>
        ${C.full_name}
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
    `;ae({title:u,content:b,confirmText:f?"Simpan Perubahan":"Tambah Cabang",onConfirm:async(C,_)=>{let p=C.querySelector("#migration-form");if(!p.checkValidity()){p.reportValidity();return}let T=new FormData(p),D={branch_id:parseInt(T.get("branch_id"),10),vendor:T.get("vendor"),target_month:T.get("target_month"),status:T.get("status"),actual_date:T.get("actual_date")||null,pic_vendor:T.get("pic_vendor")||null,notes:T.get("notes")||null};try{let A=f?`/api/os-migrations/${c.id}`:"/api/os-migrations",P=await F(A,{method:f?"PUT":"POST",body:JSON.stringify(D)});if(!P.ok){let I=await P.json().catch(()=>({}));throw new Error(I.error||"Gagal menyimpan data")}j(f?"Jadwal migrasi berhasil diupdate":"Jadwal migrasi berhasil ditambahkan"),_(),o()}catch(A){R(A.message)}}})}let k=document.getElementById("view-board-btn"),g=document.getElementById("view-table-btn");k.addEventListener("click",()=>{t="board",k.style.background="#fff",k.classList.remove("btn-ghost"),g.style.background="transparent",g.classList.add("btn-ghost"),l()}),g.addEventListener("click",()=>{t="table",g.style.background="#fff",g.classList.remove("btn-ghost"),k.style.background="transparent",k.classList.add("btn-ghost"),l()}),["migration-search","filter-vendor","filter-month","filter-status"].forEach(c=>{let f=document.getElementById(c);f&&(f.addEventListener("input",l),f.addEventListener("change",l))}),document.getElementById("btn-add-migration").addEventListener("click",()=>{S()}),document.getElementById("btn-export-migrations").addEventListener("click",()=>{let c=a();if(!c.length){R("Tidak ada data untuk diexport");return}let f=c.map((u,w)=>({No:w+1,"Kode Cabang":u.branch_code,"Nama Cabang":u.branch_name,"Cabang Lengkap":u.branch_full_name,"Vendor OS":u.vendor,"Target Bulan":Mt[u.target_month]||u.target_month,Status:Qe[u.status]?.label||u.status,"Tanggal Realisasi":u.actual_date||"","Jml Personil FC":u.employee_count||0,"PIC Vendor":u.pic_vendor||"",Catatan:u.notes||""}));K(f,`Timeline_Migrasi_OS_${new Date().toISOString().slice(0,10)}`),j("File Excel berhasil diunduh")}),o()}H();he();G();var Ia={Persiapan:{bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE"},"Area Umum":{bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD"},"Ruang Tindakan":{bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},"Ruang Penunjang":{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},Toilet:{bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4"},"Area Luar":{bg:"#FEF3C7",color:"#92400E",border:"#FDE68A"},"Kontrol & Check":{bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA"},Istirahat:{bg:"#E2E8F0",color:"#475569",border:"#CBD5E1"},"Layanan Pasien":{bg:"#DCFCE7",color:"#15803D",border:"#86EFAC"},"Sampah Umum":{bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},"Sampah Medis":{bg:"#FEE2E2",color:"#991B1B",border:"#FCA5A5"},Penutup:{bg:"#E0E7FF",color:"#3730A3",border:"#C7D2FE"}};async function Aa(e){e.innerHTML=`
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
  `;let t="pagi",i="timeline",n=[],r=ie(),d=r&&["superadmin","admin","manager","spv"].includes(r.role);async function o(){try{let c=await F("/api/dac?all=1");if(!c.ok)throw new Error("Gagal memuat standar DAC");n=c.data?.data||c.data||[];let f=n.filter(C=>C.shift==="pagi").length,u=n.filter(C=>C.shift==="siang").length;document.getElementById("badge-count-pagi").textContent=f,document.getElementById("badge-count-siang").textContent=u,document.getElementById("badge-count-all").textContent=n.length;let w=document.getElementById("dac-category-filter"),b=Array.from(new Set(n.map(C=>C.category))).sort();w.innerHTML='<option value="">Semua Kategori</option>'+b.map(C=>`<option value="${C}">${C}</option>`).join(""),a(),l()}catch(c){document.getElementById("dac-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${c.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function s(){let c=(document.getElementById("dac-search")?.value||"").toLowerCase().trim(),f=document.getElementById("dac-category-filter")?.value||"",u=document.getElementById("dac-period-filter")?.value||"";return n.filter(w=>{if(t!=="all"&&w.shift!==t||f&&w.category!==f)return!1;if(u==="morning"){let b=parseInt(w.start_time.split(":")[0],10);if(w.shift==="pagi"&&b>=12||w.shift==="siang"&&b>=17)return!1}else if(u==="afternoon"){let b=parseInt(w.start_time.split(":")[0],10);if(w.shift==="pagi"&&b<12||w.shift==="siang"&&b<17)return!1}return!(c&&!`${w.start_time} ${w.end_time} ${w.category} ${w.location} ${w.activity}`.toLowerCase().includes(c))})}function a(){let c=document.getElementById("dac-summary-cards");if(!c)return;let f=n.filter(_=>_.shift==="pagi"),u=n.filter(_=>_.shift==="siang"),w=t==="pagi"?f:t==="siang"?u:n,b=t==="pagi"?"08:00 \u2013 17:00":t==="siang"?"13:00 \u2013 22:00":"08:00 \u2013 22:00",C=t==="pagi"?"Shift Pagi":t==="siang"?"Shift Siang":"Semua Shift";c.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Aktivitas (${C})</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${w.length} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Aktivitas</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Urutan Terjadwal dari Awal s/d Penutup</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Jam Kerja Standar</div>
        <div style="font-size:1.5rem; font-weight:800; color:#1E40AF; line-height:1.2;">${b}</div>
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
    `}function l(){let c=s(),f=document.getElementById("dac-stats-text");f&&(f.innerHTML=`Menampilkan <strong>${c.length}</strong> aktivitas`);let u=document.getElementById("dac-view-container");i==="timeline"?v(u,c):x(u,c)}function v(c,f){if(!f.length){c.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;return}c.innerHTML=`
      <div class="timeline-container" style="display:flex; flex-direction:column; gap:12px; max-width:960px; margin:0 auto;">
        ${f.map(u=>{let w=Ia[u.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},b=u.shift==="pagi"?"\u{1F305} Pagi":"\u{1F307} Siang",C=u.shift==="pagi"?"#FEF3C7":"#E0E7FF",_=u.shift==="pagi"?"#92400E":"#3730A3";return`
            <div class="dac-timeline-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:16px 20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; gap:16px; align-items:flex-start; position:relative;">
              <!-- Time Column -->
              <div style="min-width:130px; text-align:center; background:#F8FAFC; border:1px solid var(--border); border-radius:8px; padding:10px 8px;">
                <div style="font-size:1.1rem; font-weight:800; color:var(--text-1); font-family:monospace;">
                  ${u.start_time} - ${u.end_time}
                </div>
                <div style="font-size:0.75rem; color:var(--text-3); font-weight:600; margin-top:4px;">
                  Item #${u.item_no}
                </div>
                ${t==="all"?`
                  <div style="margin-top:6px;">
                    <span class="badge" style="background:${C}; color:${_}; font-size:0.7rem; font-weight:700; padding:2px 6px; border-radius:4px;">
                      ${b}
                    </span>
                  </div>
                `:""}
              </div>

              <!-- Content Column -->
              <div style="flex:1;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                  <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span class="badge" style="background:${w.bg}; color:${w.color}; border:1px solid ${w.border}; font-weight:700; font-size:0.75rem; padding:3px 10px; border-radius:6px;">
                      ${u.category}
                    </span>
                    <span style="font-weight:700; font-size:0.95rem; color:var(--text-1);">
                      \u{1F4CD} ${u.location}
                    </span>
                  </div>
                  ${d?`
                    <div style="display:flex; gap:4px;">
                      <button class="btn btn-ghost btn-sm edit-dac-btn" data-id="${u.id}" title="Edit" style="padding:4px 8px; font-size:0.8rem; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-ghost btn-sm delete-dac-btn" data-id="${u.id}" data-name="${u.location}" title="Hapus" style="padding:4px 8px; font-size:0.8rem; color:var(--danger); cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </div>
                  `:""}
                </div>

                <div style="font-size:0.9rem; color:var(--text-2); line-height:1.5; background:#FAFBFD; padding:10px 14px; border-radius:8px; border-left:3px solid ${w.color}; margin-top:8px;">
                  ${u.activity}
                </div>
              </div>
            </div>
          `}).join("")}
      </div>
    `,h(c)}function x(c,f){if(!f.length){c.innerHTML=`
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
                ${d?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right; width:100px;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${f.map(u=>{let w=Ia[u.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"};return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 14px; font-size:0.85rem; color:var(--text-3); font-weight:600;">${u.item_no}</td>
                    ${t==="all"?`
                      <td style="padding:12px 14px;">
                        <span class="badge" style="background:${u.shift==="pagi"?"#FEF3C7":"#E0E7FF"}; color:${u.shift==="pagi"?"#92400E":"#3730A3"}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:4px;">
                          ${u.shift==="pagi"?"Pagi":"Siang"}
                        </span>
                      </td>
                    `:""}
                    <td style="padding:12px 14px; font-family:monospace; font-weight:700; font-size:0.85rem; color:var(--text-1); white-space:nowrap;">
                      ${u.start_time} - ${u.end_time}
                    </td>
                    <td style="padding:12px 14px;">
                      <span class="badge" style="background:${w.bg}; color:${w.color}; border:1px solid ${w.border}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px; white-space:nowrap;">
                        ${u.category}
                      </span>
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; font-weight:600; color:var(--text-1);">
                      ${u.location}
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; color:var(--text-2); line-height:1.4;">
                      ${u.activity}
                    </td>
                    ${d?`
                      <td style="padding:12px 14px; text-align:right; white-space:nowrap;">
                        <button class="btn btn-outline btn-sm edit-dac-btn" data-id="${u.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-dac-btn" data-id="${u.id}" data-name="${u.location}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
                      </td>
                    `:""}
                  </tr>
                `}).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `,h(c)}function h(c){c.querySelectorAll(".edit-dac-btn").forEach(f=>{f.addEventListener("click",()=>{let u=f.dataset.id,w=n.find(b=>String(b.id)===String(u));w&&m(w)})}),c.querySelectorAll(".delete-dac-btn").forEach(f=>{f.addEventListener("click",()=>{let u=f.dataset.id,w=f.dataset.name||"item";pe(`Apakah Anda yakin ingin menghapus aktivitas pada lokasi <strong>${w}</strong>?`,async()=>{try{if(!(await F(`/api/dac/${u}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus item");j("Aktivitas DAC berhasil dihapus"),o()}catch(b){R(b.message)}})})})}function m(c=null){let f=!!c,u=f?"Edit Aktivitas DAC":"Tambah Aktivitas DAC",w=["Persiapan","Area Umum","Ruang Tindakan","Ruang Penunjang","Toilet","Area Luar","Kontrol & Check","Istirahat","Layanan Pasien","Sampah Umum","Sampah Medis","Penutup"],b=`
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
            ${w.map(C=>`<option value="${C}">`).join("")}
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
    `;ae({title:u,content:b,confirmText:f?"Simpan Perubahan":"Tambah Aktivitas",onConfirm:async(C,_)=>{let p=C.querySelector("#dac-form");if(!p.checkValidity()){p.reportValidity();return}let T=new FormData(p),D={shift:T.get("shift"),item_no:T.get("item_no")?parseInt(T.get("item_no"),10):void 0,start_time:T.get("start_time"),end_time:T.get("end_time"),category:T.get("category"),location:T.get("location"),activity:T.get("activity")};try{let A=f?`/api/dac/${c.id}`:"/api/dac",P=await F(A,{method:f?"PUT":"POST",body:JSON.stringify(D)});if(!P.ok){let I=await P.json().catch(()=>({}));throw new Error(I.error||"Gagal menyimpan aktivitas")}j(f?"Aktivitas berhasil diperbarui":"Aktivitas berhasil ditambahkan"),_(),o()}catch(A){R(A.message)}}})}function S(){let c=s();if(!c.length){R("Tidak ada data untuk dicetak");return}let f=t==="pagi"?"SHIFT PAGI (08:00 - 17:00)":t==="siang"?"SHIFT SIANG (13:00 - 22:00)":"SELURUH SHIFT (PAGI & SIANG)",u=window.open("","_blank");if(!u){R("Popup diblokir browser, mohon izinkan popup untuk mencetak");return}u.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>STANDAR ACUAN DAC - ${f}</title>
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
          <div class="subtitle">STANDAR RUTINITAS DAILY ACTIVITY CLEANING (DAC) \u2014 ${f}</div>
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
            ${c.map(w=>`
              <tr>
                <td style="text-align:center; font-weight:bold;">${w.item_no}</td>
                ${t==="all"?`<td>${w.shift==="pagi"?"Pagi":"Siang"}</td>`:""}
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
    `),u.document.close(),u.focus(),setTimeout(()=>{u.print()},400)}document.querySelectorAll(".shift-tab-btn").forEach(c=>{c.addEventListener("click",()=>{document.querySelectorAll(".shift-tab-btn").forEach(f=>{f.style.background="#F1F5F9",f.style.color="var(--text-2)"}),c.style.background="#2563EB",c.style.color="#fff",t=c.dataset.shift,a(),l()})});let k=document.getElementById("view-timeline-btn"),g=document.getElementById("view-table-btn");k.addEventListener("click",()=>{i="timeline",k.style.background="#fff",k.classList.remove("btn-ghost"),g.style.background="transparent",g.classList.add("btn-ghost"),l()}),g.addEventListener("click",()=>{i="table",g.style.background="#fff",g.classList.remove("btn-ghost"),k.style.background="transparent",k.classList.add("btn-ghost"),l()}),["dac-search","dac-category-filter","dac-period-filter"].forEach(c=>{let f=document.getElementById(c);f&&(f.addEventListener("input",l),f.addEventListener("change",l))}),document.getElementById("btn-add-dac").addEventListener("click",()=>{m()}),document.getElementById("btn-print-dac").addEventListener("click",()=>{S()}),document.getElementById("btn-export-dac").addEventListener("click",()=>{let c=s();if(!c.length){R("Tidak ada data untuk diexport");return}let f=c.map(u=>({No:u.item_no,Shift:u.shift==="pagi"?"Shift Pagi":"Shift Siang","Jam Mulai":u.start_time,"Jam Selesai":u.end_time,Kategori:u.category,"Lokasi / Area":u.location,"Rincian Aktivitas":u.activity}));K(f,`DAC_Daily_Activity_Cleaning_${t}_${new Date().toISOString().slice(0,10)}`),j("File Excel berhasil diunduh")}),o()}window.parseFlexibleDate=e=>{if(!e||e==="-")return"";if(e=String(e).trim(),/^\d{5}$/.test(e)){let i=Math.floor(Number(e)-25569);return new Date(i*86400*1e3).toISOString().split("T")[0]}if(e.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let i=e.split(/[\/\-]/);return`${i[2]}-${i[1]}-${i[0]}`}let t=e.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(t){let i=t[1],n=parseInt(t[2],10),r=parseInt(t[3],10);if(n>12&&r<=12)return`${i}-${t[3]}-${t[2]}`}return e.split("T")[0]};window.formatDate=e=>{let t=window.parseFlexibleDate(e);if(!t)return"";let i=t.split("-");if(i.length===3&&i[0].length===4){let n=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],r=parseInt(i[2],10),d=n[parseInt(i[1],10)-1];return`${r} ${d} ${i[0]}`}return t};function Z(e){return async t=>{if(!je()){$e("/login");return}return e(t)}}var Ye=null;function bn(){Ye&&clearInterval(Ye);let e=()=>{let t=new Date,i=t.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),n=t.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),r=document.getElementById("header-clock-time"),d=document.getElementById("header-clock-date");r&&(r.textContent=i),d&&(d.textContent=n)};e(),Ye=setInterval(e,1e3)}async function hn(){try{let e=await F("/api/dashboard/kpi");if(!e.ok)return;let t=e.data?.data||e.data||{},i=(n,r)=>{let d=document.getElementById(n);d&&(d.textContent=r>0?r:"",d.style.display=r>0?"inline-flex":"none")};i("badge-issues",t.issues?.current||0),i("badge-contracts",t.expiring30?.current||0),i("badge-oo1",t.one_on_one?.current||0),i("badge-schedule",t.schedule?.current||0),i("badge-supply",t.supply?.current||0)}catch{}}var Ke=[];async function fn(){try{let e=await F("/api/dashboard/notifications");if(!e.ok)return;Ke=e.data?.data||e.data||[];let t=document.getElementById("notif-dot");t&&(t.style.display=Ke.length>0?"block":"none",t.textContent=Ke.length)}catch{}}function yn(){if(!Ke.length){ae({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(t,i)=>i()});return}let e=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${Ke.map(t=>`
        <div class="notif-item notif-severity-${t.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${t.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${t.date}</span>
            <span class="badge badge-${t.severity==="danger"?"danger":t.severity==="warning"?"warning":"info"}">${t.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;ae({title:`Notifikasi (${Ke.length})`,content:e,confirmText:"Tutup",onConfirm:(t,i)=>i()})}function La(){let e=ie(),t=(e?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
            <a href="#/sop" class="nav-item" data-route="/sop">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M4 19.5A2.5 2.5 0 016.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 014 19.5v-15A2.5 2.5 0 016.5 2z"/></svg>
              </span>
              <span class="nav-label">SOP</span>
            </a>
            <a href="#/sop-fc" class="nav-item" data-route="/sop-fc">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>
              </span>
              <span class="nav-label">SOP Facility Care</span>
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
  `;let i=document.getElementById("sidebar"),n=document.getElementById("sidebar-overlay"),r=document.getElementById("topbar-menu-btn"),d=document.getElementById("sidebar-close"),o=()=>{i.classList.add("open"),n.classList.add("show")},s=()=>{i.classList.remove("open"),n.classList.remove("show")};r?.addEventListener("click",o),d?.addEventListener("click",s),n?.addEventListener("click",s),document.querySelectorAll(".nav-item").forEach(l=>l.addEventListener("click",s)),e&&e.role==="input_lembur"&&(document.querySelectorAll(".nav-item").forEach(l=>{l.dataset.route!=="/overtime"&&(l.style.display="none")}),document.querySelectorAll(".nav-section").forEach(l=>{Array.from(l.querySelectorAll(".nav-item")).some(x=>x.style.display!=="none")||(l.style.display="none")}));function a(){let l=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(h=>{let m=h.dataset.route;h.classList.toggle("active",l===m||m!=="/dashboard"&&l.startsWith(m))});let v=document.getElementById("topbar-title"),x=document.querySelector(".nav-item.active .nav-label");v&&x&&(v.textContent=x.textContent)}window.addEventListener("hashchange",a),a(),bn(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await F("/api/auth/logout",{method:"POST"}),Re(),Ye&&clearInterval(Ye),$e("/login")}),hn(),fn(),document.getElementById("btn-notif")?.addEventListener("click",l=>{l.preventDefault(),yn()})}async function vn(){X("/login",({main:t})=>Gt(t)),X("/dashboard",Z(({main:t})=>Rt(t))),X("/calendar",Z(({main:t})=>ha(t))),X("/employees",Z(({main:t,params:i})=>Vt(t,i))),X("/contracts",Z(({main:t,params:i})=>dt(t,i))),X("/sp",Z(({main:t})=>_a(t))),X("/mutasi",Z(({main:t})=>Ea(t))),X("/overtime",Z(({main:t})=>Ta(t))),X("/timeline",Z(({main:t,params:i})=>Qt(t,i))),X("/os-migrations",Z(({main:t})=>Da(t))),X("/issues",Z(({main:t,params:i})=>Yt(t,i))),X("/one-on-one",Z(({main:t,params:i})=>Xt(t,i))),X("/training",Z(({main:t})=>Zt(t))),X("/relievers",Z(({main:t,params:i})=>ta(t,i))),X("/reports/inspection",Z(({main:t})=>aa(t))),X("/reports/cleaning",Z(({main:t})=>na(t))),X("/reports/fogging",Z(({main:t})=>ia(t))),X("/reports/basecamp",Z(({main:t})=>ra(t))),X("/reports/supply",Z(({main:t})=>ua(t,"supply"))),X("/sop",Z(({main:t})=>oa(t))),X("/sop-fc",Z(({main:t})=>ca(t))),X("/checklist",Z(({main:t})=>pa(t))),X("/hygiene-standards",Z(({main:t})=>ma(t))),X("/dac",Z(({main:t})=>Aa(t))),X("/users",Z(({main:t})=>ga(t))),X("/branches",Z(({main:t})=>ba(t))),X("/profile",Z(({main:t})=>fa(t))),X("/settings/import",Z(({main:t})=>wa(t))),X("/audit-logs",Z(({main:t})=>Fa(t)));let e=je();if(!e&&window.location.hash!=="#/login"&&$e("/login"),e){let t=await F("/api/auth/me");if(t.ok){qe(t.data.data),La();let i=ie(),n=window.location.hash;i&&i.role==="input_lembur"&&(n==="#/dashboard"||n===""||n==="#/")&&$e("/overtime")}else Re(),$e("/login")}window.addEventListener("fm:login",()=>{La();let t=ie();t&&t.role==="input_lembur"?$e("/overtime"):$e("/dashboard")}),Ot()}vn();
