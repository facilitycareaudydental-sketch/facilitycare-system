var Za=Object.defineProperty;var bt=(e,t)=>()=>(e&&(t=e(e=0)),t);var ht=(e,t)=>{for(var i in t)Za(e,i,{get:t[i],enumerable:!0})};var Gt={};ht(Gt,{API:()=>Jt,CLIENT_SIDE_MAX_ROWS:()=>Te,IS_DEVELOPMENT:()=>Ke,apiFetch:()=>I,clearToken:()=>Re,getToken:()=>Oe,getUser:()=>ae,setToken:()=>ft,setUser:()=>qe});function Oe(){return localStorage.getItem("fm_token")}function ft(e){localStorage.setItem("fm_token",e)}function Re(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ae(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function qe(e){localStorage.setItem("fm_user",JSON.stringify(e))}async function I(e,t={}){let i=Oe(),a={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...t.headers||{}};try{let o=`cb=${Date.now()}`,b=e.includes("?")?"&":"?",u=`${Jt}${e}${b}${o}`,d=await fetch(u,{...t,headers:a}),n;try{let h=await d.text();try{n=JSON.parse(h)}catch{n={error:`Server Error (${d.status}): ${h.substring(0,80)}...`}}}catch{n={error:"Gagal membaca respon dari server"}}return d.status===401&&(Re(),window.location.hash="#/login"),{ok:d.ok,status:d.status,data:n}}catch(o){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${o.message})`}}}}var Ke,en,Jt,Te,V=bt(()=>{Ke=!1,en="https://fm-operations-api.facilitycare-audydental.workers.dev",Jt=en,Te=1e4});var kt={};ht(kt,{confirmDialog:()=>ee,createModal:()=>Y});function Y({title:e,content:t,onConfirm:i,onCancel:a,confirmText:o="Simpan",cancelText:b="Batal",size:u="md",confirmClass:d="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},h=document.createElement("div");h.className="modal-overlay",h.innerHTML=`
    <div class="modal" style="max-width:${n[u]||n.md}">
      <div class="modal-header">
        <h3 class="modal-title">${e}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof t=="string"?t:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${b}</button>
        ${i?`<button class="btn ${d} modal-confirm">${o}</button>`:""}
      </div>
    </div>
  `,t instanceof HTMLElement&&h.querySelector(".modal-body").appendChild(t);let F=()=>{h.classList.remove("show"),setTimeout(()=>h.remove(),250)};return h.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),F()}),h.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),F()}),i&&h.querySelector(".modal-confirm").addEventListener("click",()=>i(h,F)),h.addEventListener("click",$=>{$.target===h&&(a&&a(),F())}),document.body.appendChild(h),requestAnimationFrame(()=>h.classList.add("show")),{overlay:h,close:F}}function ee(e,t,i="Konfirmasi"){return Y({title:i,content:`<p>${e}</p>`,onConfirm:(a,o)=>{t(),o()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var ue=bt(()=>{});var Se={};ht(Se,{downloadExcel:()=>G,parseExcel:()=>Ge,renderExcelButtons:()=>an});function Ge(e){return new Promise((t,i)=>{let a=new FileReader;a.onload=o=>{try{let b=new Uint8Array(o.target.result),u=XLSX.read(b,{type:"array"}),d=u.SheetNames[0],n=u.Sheets[d];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${e.name}`),console.log(`File Size: ${(e.size/1024).toFixed(2)} KB`),console.log(`File Type: ${e.type||"unknown"}`),console.log(`Sheets Found: ${u.SheetNames.join(", ")}`),console.log(`Sheet Used: ${d}`);let h=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),F=h.e.r-h.s.r+1,$=h.e.c-h.s.c+1;console.log(`Total Rows (including empty): ${F}`),console.log(`Total Columns: ${$}`);let E=[];for(let w=h.s.c;w<=h.e.c;++w){let k=n[XLSX.utils.encode_cell({c:w,r:h.s.r})];k&&k.v&&E.push(k.v)}console.log(`Headers Found: ${E.join(", ")}`),console.log("---------------------------");let C=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(C,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(C,"__headers",{value:E,enumerable:!1}),t(C)}catch(b){i(b)}},a.onerror=o=>i(o),a.readAsArrayBuffer(e)})}function G(e,t){try{let i=XLSX.utils.json_to_sheet(e),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,i,"Data"),XLSX.writeFile(a,`${t}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function an(e){return`
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
  `}var X=bt(()=>{});V();var yt={},Xe=null;function Q(e,t){yt[e]=t}function $e(e){window.location.hash=e}function Ut(){async function e(){let t=window.location.hash.replace("#","")||"/dashboard",[i,...a]=t.split("?"),o=yt[i];if(!o){for(let[u,d]of Object.entries(yt))if(u.endsWith("/*")&&i.startsWith(u.slice(0,-2))){o=d;break}}Xe&&(Xe(),Xe=null);let b=document.getElementById("main-content");if(b&&(b.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),o){let u=new URLSearchParams(a.join("?")),d=i.split("/").filter(Boolean),n=await o({path:i,params:u,segments:d,main:b});n&&(Xe=n)}else{let u=b||document.getElementById("app");u&&(u.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",e),e()}var He;function tn(){return He||(He=document.createElement("div"),He.id="toast-container",document.body.appendChild(He)),He}function xt(e,t="info",i=3500){let a=tn(),o=document.createElement("div");o.className=`toast toast-${t}`;let b={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};o.innerHTML=`<span class="toast-icon">${b[t]||"\u2139"}</span><span class="toast-msg">${e}</span>`,a.appendChild(o),requestAnimationFrame(()=>o.classList.add("show")),setTimeout(()=>{o.classList.remove("show"),setTimeout(()=>o.remove(),350)},i)}var j=e=>xt(e,"success"),N=e=>xt(e,"error");var vt=e=>xt(e,"info");ue();V();V();ue();function Ze({columns:e,data:t,onEdit:i,onDelete:a,onView:o,actions:b=[],emptyText:u="Tidak ada data",bulkSelect:d=null}){let n=document.createElement("div");if(n.className="table-wrapper",!t||t.length===0)return n.innerHTML=`<div class="empty-state"><p>${u}</p></div>`,n;let h=document.createElement("table");h.className="data-table";let F=document.createElement("thead"),$=document.createElement("tr");if(d){let C=document.createElement("th");C.style.width="40px",C.style.textAlign="center";let w=document.createElement("input");w.type="checkbox",w.id="select-all-checkbox",w.title="Pilih semua",w.addEventListener("change",()=>{t.forEach(k=>{w.checked?d.selectedIds.add(k.id):d.selectedIds.delete(k.id)}),n.querySelectorAll(".row-checkbox").forEach(k=>k.checked=w.checked),d.onToggle()}),C.appendChild(w),$.appendChild(C)}if(e.forEach(C=>{let w=document.createElement("th");w.textContent=C.label,C.width&&(w.style.width=C.width),$.appendChild(w)}),i||a||o||b.length>0){let C=document.createElement("th");C.textContent="Aksi",C.style.width="120px",$.appendChild(C)}F.appendChild($),h.appendChild(F);let E=document.createElement("tbody");return t.forEach(C=>{let w=document.createElement("tr");if(d){let k=document.createElement("td");k.style.textAlign="center",k.style.width="40px";let y=document.createElement("input");y.type="checkbox",y.className="row-checkbox",y.checked=d.selectedIds.has(C.id),y.addEventListener("change",()=>{if(y.checked)d.selectedIds.add(C.id);else{d.selectedIds.delete(C.id);let g=document.getElementById("select-all-checkbox");g&&(g.checked=!1)}d.onToggle()}),k.appendChild(y),w.appendChild(k)}if(e.forEach(k=>{let y=document.createElement("td");if(k.render){let g=k.render(C[k.key],C);g instanceof HTMLElement?y.appendChild(g):y.innerHTML=g||""}else y.textContent=C[k.key]!==null&&C[k.key]!==void 0&&C[k.key]!==""?C[k.key]:"";k.nowrap&&(y.style.whiteSpace="nowrap"),w.appendChild(y)}),i||a||o||b.length>0){let k=document.createElement("td");k.className="actions-cell";let y=document.createElement("div");if(y.className="btn-group",o){let g=document.createElement("button");g.className="btn btn-xs btn-ghost",g.innerHTML="\u{1F441}",g.title="Lihat",g.addEventListener("click",()=>o(C)),y.appendChild(g)}if(i){let g=document.createElement("button");g.className="btn btn-xs btn-secondary",g.innerHTML="\u270F\uFE0F",g.title="Edit",g.addEventListener("click",()=>i(C)),y.appendChild(g)}b.forEach(g=>{let r=document.createElement("button");r.className=`btn btn-xs ${g.class||"btn-ghost"}`,r.innerHTML=g.icon||g.label,r.title=g.label,r.addEventListener("click",()=>g.handler(C)),y.appendChild(r)}),k.appendChild(y),w.appendChild(k)}E.appendChild(w)}),h.appendChild(E),n.appendChild(h),n}function et({page:e,pages:t,total:i,limit:a,onPage:o}){if(t<=1)return null;let b=document.createElement("div");b.className="pagination";let u=document.createElement("span");u.className="pagination-info",u.textContent=`Total: ${i} data`,b.appendChild(u);let d=document.createElement("div");d.className="pagination-btns";let n=($,E,C=!1,w=!1)=>{let k=document.createElement("button");k.className=`btn btn-sm ${w?"btn-primary":"btn-ghost"} pagination-btn`,k.textContent=$,k.disabled=C,k.addEventListener("click",()=>o(E)),d.appendChild(k)};n("\xAB",1,e===1),n("\u2039",e-1,e===1);let h=Math.max(1,e-2),F=Math.min(t,e+2);for(let $=h;$<=F;$++)n($,$,!1,$===e);return n("\u203A",e+1,e===t),n("\xBB",t,e===t),b.appendChild(d),b}ue();function Je(e){return e.map(t=>{if(t.type==="hidden")return`<input type="hidden" name="${t.name}" value="${t.value||""}">`;if(t.type==="html")return t.html||"";if(t.type==="row")return`<div class="form-row">${Je(t.fields)}</div>`;let i=t.required?"required":"",a=t.label?`<label class="form-label">${t.label}${t.required?' <span class="required">*</span>':""}</label>`:"",o="";switch(t.type){case"textarea":o=`<textarea name="${t.name}" class="form-control" placeholder="${t.placeholder||""}" ${i} rows="${t.rows||3}">${t.value||""}</textarea>`;break;case"select":let u=(t.options||[]).map(C=>{let w=typeof C=="object"?C.value:C,k=typeof C=="object"?C.label:C,y=t.value==w?"selected":"";return`<option value="${w}" ${y}>${k}</option>`}).join("");o=`<select name="${t.name}" class="form-control" ${i}><option value="">-- Pilih ${t.label||""} --</option>${u}</select>`;break;case"combobox":let d=`dl-${t.name}-${Math.random().toString(36).substring(7)}`,n=(t.options||[]).map(C=>{let w=typeof C=="object"?C.value:C,k=typeof C=="object"?C.label||C.value||"":C||"";return(k==="undefined"||k==="[object Object]"||k==="null")&&(k=""),k?`<option value="${k}"></option>`:""}).join(""),h=t.value||"";if(t.value){let C=(t.options||[]).find(w=>(typeof w=="object"?w.value:w)==t.value);if(C){let w=typeof C=="object"?C.label||C.value||"":C||"";w&&w!=="undefined"&&w!=="[object Object]"&&w!=="null"&&(h=w)}}o=`
          <input type="text" name="${t.name}" list="${d}" class="form-control" value="${h}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${d}">${n}</datalist>
        `;break;case"checkbox":o=`<label class="checkbox-label"><input type="checkbox" name="${t.name}" value="1" ${t.value?"checked":""}> ${t.checkLabel||t.label}</label>`;break;case"date":let F=window.parseFlexibleDate&&t.value?window.parseFlexibleDate(t.value):t.value||"";o=`<input type="date" name="${t.name}" class="form-control" value="${F}" ${i}>`;break;case"month":let $=t.value?String(t.value).slice(0,7):"";o=`<input type="month" name="${t.name}" class="form-control" value="${$}" ${i}>`;break;case"number":let E=t.readonly?"readonly":"";o=`<input type="number" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" min="${t.min||""}" max="${t.max||""}" step="${t.step||"1"}" ${i} ${E}>`;break;case"email":o=`<input type="email" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${i}>`;break;case"url":o=`<input type="url" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||"https://..."}" ${i}>`;break;default:o=`<input type="${t.type||"text"}" name="${t.name}" class="form-control" value="${t.value||""}" placeholder="${t.placeholder||""}" ${i} autocomplete="off">`}let b=t.hint?`<div class="form-hint">${t.hint}</div>`:"";return`<div class="form-group ${t.class||""}">${a}${o}${b}</div>`}).join("")}function tt(e){let t={},i=new FormData(e);for(let[a,o]of i.entries())t[a]=o===""?null:o;return e.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(t[a.name]=null)}),t}function at(e,t){t&&Object.entries(t).forEach(([i,a])=>{let o=e.querySelector(`[name="${i}"]`);o&&(o.hasAttribute("list")||(o.type==="checkbox"?o.checked=!!a:o.type==="date"&&a&&window.parseFlexibleDate?o.value=window.parseFlexibleDate(a):o.type==="month"&&a?o.value=String(a).slice(0,7):o.value=a??""))})}X();var Ce={},nt={on(e,t){Ce[e]||(Ce[e]=new Set),Ce[e].add(t)},off(e,t){Ce[e]&&Ce[e].delete(t)},emit(e,t){Ce[e]&&Ce[e].forEach(i=>{try{i(t)}catch(a){console.warn("[calendarBus] Handler error:",a)}})},clear(){Object.keys(Ce).forEach(e=>delete Ce[e])}},nn=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function be(e){if(!e){nt.emit("data:changed",{module:"unknown"});return}let t=String(e).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");nt.emit("data:changed",{module:t,relevant:nn.has(t)})}function Z({container:e,title:t,icon:i,apiPath:a,columns:o,formFields:b,filterFields:u,defaultFilters:d={},itemLabel:n="Data",canCreate:h=!0,canEdit:F=!0,canDelete:$=!0,onBeforeSubmit:E,onAfterLoad:C,onDataLoaded:w,extraActions:k=[],initialSearch:y="",exportOptions:g=null,bulkDelete:r=!1,paginationMode:s="server"}){let m=ae();m&&m.role==="viewer"&&(h=!1,F=!1,$=!1,r=!1,g=null),m&&m.role==="editor_khusus"&&a!=="/api/issues"&&a!=="/api/overtime"&&(h=!1,F=!1,$=!1,r=!1,g=null),m&&m.role==="input_lembur"&&(a==="/api/overtime"?($=!1,r=!1,g=null):(h=!1,F=!1,$=!1,r=!1,g=null));let p=1,f={...d};y&&(f.search=y);let l=new Set;e.innerHTML=`
    ${r?`
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
        ${h?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
        ${g?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${g.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${g.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${g.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${g.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
    

    ${u&&u.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${u.filter(x=>x.type==="search").map(x=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${x.placeholder||"Cari..."}" id="filter-search" value="${f.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${u.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let A=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select" name="${x.name}" id="filter-${x.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${A}</option>${(x.options||[]).map(L=>`<option value="${typeof L=="object"?L.value:L}" ${f[x.name]===(typeof L=="object"?L.value:L)?"selected":""}>${typeof L=="object"?L.label:L}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${u.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let A=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select filter-select-sheet" name="${x.name}-sheet" id="filter-sheet-${x.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${A}</option>${(x.options||[]).map(L=>`<option value="${typeof L=="object"?L.value:L}" ${f[x.name]===(typeof L=="object"?L.value:L)?"selected":""}>${typeof L=="object"?L.label:L}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function c(){let x=document.getElementById("bulk-toolbar");if(!x)return;let A=document.getElementById("bulk-count"),L=document.getElementById("btn-bulk-delete"),q=document.getElementById("btn-bulk-cancel");A.textContent=`${l.size} item dipilih`,l.size>0?(x.style.display="flex",L.disabled=!1,q.disabled=!1):(x.style.display="none",L.disabled=!0,q.disabled=!0);let O=document.getElementById("select-all-checkbox");if(O){let H=document.querySelectorAll(".row-checkbox");if(H.length>0){let K=[...H].every(ie=>ie.checked),R=[...H].some(ie=>ie.checked);O.checked=K,O.indeterminate=R&&!K}else O.checked=!1,O.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{l.clear(),document.querySelectorAll(".row-checkbox").forEach(A=>A.checked=!1);let x=document.getElementById("select-all-checkbox");x&&(x.checked=!1),c()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(l.size===0)return;let x=[...l],A=document.createElement("div");A.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",A.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${x.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${x.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(A),A.querySelector("#bulk-cancel-btn").addEventListener("click",()=>A.remove()),A.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let L=A.querySelector("#bulk-confirm-btn");L.disabled=!0,L.textContent="Menghapus...";let q=await I(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:x})});A.remove(),q.ok?(j(`${x.length} ${n} berhasil dihapus.`),l.clear(),c(),be(a),_()):N(q.data?.error||"Gagal menghapus data.")})});let v=document.getElementById("filter-search"),S;v?.addEventListener("input",x=>{clearTimeout(S),S=setTimeout(()=>{f.search=x.target.value,p=1,l.clear(),c(),_()},400)}),u?.forEach(x=>{(x.type==="select"||x.type==="combobox")&&(document.getElementById(`filter-${x.name}`)?.addEventListener("change",A=>{f[x.name]=A.target.value;let L=document.getElementById(`filter-sheet-${x.name}`);L&&(L.value=A.target.value),p=1,l.clear(),c(),_()}),document.getElementById(`filter-sheet-${x.name}`)?.addEventListener("change",A=>{f[x.name]=A.target.value;let L=document.getElementById(`filter-${x.name}`);L&&(L.value=A.target.value),p=1,l.clear(),c(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{f={},v&&(v.value=""),u?.forEach(x=>{let A=document.getElementById(`filter-${x.name}`);A&&(A.value="");let L=document.getElementById(`filter-sheet-${x.name}`);L&&(L.value="")}),p=1,l.clear(),c(),_()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{f={},v&&(v.value=""),u?.forEach(x=>{let A=document.getElementById(`filter-${x.name}`);A&&(A.value="");let L=document.getElementById(`filter-sheet-${x.name}`);L&&(L.value="")}),p=1,l.clear(),c(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>M(null)),g&&document.addEventListener("click",function(x){let A=document.getElementById("aksi-menu-main"),L=document.getElementById("btn-aksi-main");A&&L&&!L.contains(x.target)&&!A.contains(x.target)&&A.classList.remove("show-aksi-menu")});let B=document.getElementById("btn-mobile-filter"),T=document.getElementById("filter-options-wrapper"),D=document.getElementById("btn-close-filter-sheet");if(B&&T&&(B.addEventListener("click",x=>{x.preventDefault(),T.classList.add("sheet-open")}),D&&D.addEventListener("click",x=>{x.preventDefault(),T.classList.remove("sheet-open")})),g){document.getElementById(`btn-export-${g.moduleName}`)?.addEventListener("click",async A=>{let L=A.target,q=L.innerHTML;L.innerHTML="\u23F3 Loading...",L.disabled=!0;try{await g.onExport()}catch{N("Gagal export data")}finally{L.innerHTML=q,L.disabled=!1}}),document.getElementById(`btn-template-${g.moduleName}`)?.addEventListener("click",()=>{g.onTemplate()});let x=document.getElementById(`input-import-${g.moduleName}`);x?.addEventListener("change",async A=>{let L=A.target.files[0];if(!L)return;let q=document.getElementById(`label-import-${g.moduleName}`),O=q?q.querySelector(".import-text"):null,H=O?O.innerText:"";O&&(O.innerText="\u231B Memproses..."),q&&(q.style.pointerEvents="none"),x.disabled=!0;try{let K=await Ge(L);if(K.length===0)throw new Error("File kosong atau format salah");await g.onImport(K),j("Import berhasil!"),be(a),_()}catch(K){N(K.message||"Gagal import data")}finally{O&&(O.innerText=H),q&&(q.style.pointerEvents="auto"),x.disabled=!1,x.value=""}})}async function _(){c();let x=document.getElementById("table-container");if(!x)return;x.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let A=s==="client",L=A?1:p,q=A?Te:20,O=new URLSearchParams({page:L,limit:q,...Object.fromEntries(Object.entries(f).filter(([,U])=>U))}),H=await I(`${a}?${O}`);if(!H.ok){x.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${H.data?.error||"Error"}</p></div>`;return}let K=H.data?.data||H.data||[],R=H.data?.pagination,ie=K.length;if(A){K=w(K);let U=K.length,re=20,oe=Math.ceil(U/re);p>oe&&oe>0&&(p=oe);let J=(p-1)*re,pe=p*re;K=K.slice(J,pe),R={page:p,limit:re,total:U,pages:oe}}!1,C&&C(K);let de=Ze({columns:o,data:K,onEdit:F?U=>M(U):null,actions:k.map(U=>({...U,handler:re=>U.handler(re,_)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:r?{selectedIds:l,onToggle:c}:null});x.innerHTML="",x.appendChild(de);let me=document.getElementById("pagination-container");if(me&&(me.innerHTML="",R&&R.pages>1)){let U=et({page:R.page,pages:R.pages,total:R.total,limit:R.limit,onPage:re=>{p=re,_()}});U&&me.appendChild(U)}}function P(x){let A=typeof b=="function"?b(x):b;return Je(A)}function M(x){let A=!!x,L=document.createElement("form");if(L.noValidate=!0,L.innerHTML=P(x),A){let O=typeof b=="function"?b(x):b;at(L,x)}let{close:q}=Y({title:A?`Edit ${n}`:`Tambah ${n}`,content:L,size:"lg",confirmText:A?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(O,H)=>{if(!L.reportValidity())return;let K=O.querySelector(".modal-confirm");K.disabled=!0,K.textContent="Menyimpan...";let R=tt(L),ie=typeof b=="function"?b(x):b,de=async oe=>{for(let J of oe)if(J.type==="row")await de(J.fields);else if(J.type==="combobox"&&R[J.name]){let pe=R[J.name],fe=(J.options||[]).find(se=>{let ge=String(typeof se=="object"?se.value:se),mt=String(typeof se=="object"?se.label:se);return ge===pe||mt===pe});if(fe)R[J.name]=typeof fe=="object"?fe.value:fe;else if(J.createApi){let se={};se[J.createApi.field]=pe,J.createApi.extra&&Object.assign(se,J.createApi.extra);let ge=await I(J.createApi.path,{method:"POST",body:JSON.stringify(se)});if(ge.ok&&ge.data?.id)R[J.name]=ge.data.id;else if(ge.ok&&!ge.data?.id)R[J.name]=pe;else throw new Error(`Gagal membuat master data: ${ge.data?.error||"Unknown error"}`)}}};try{await de(ie)}catch(oe){N(oe.message),K.disabled=!1,K.textContent=A?"Simpan Perubahan":`Tambah ${n}`;return}E&&(R=await E(R,x));let me=A?"PUT":"POST",U=A?`${a}/${x.id}`:a,re=await I(U,{method:me,body:JSON.stringify(R)});re.ok?(j(A?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),H(),be(a),_()):(N(re.data?.error||"Gagal menyimpan data."),K.disabled=!1,K.textContent=A?"Simpan Perubahan":`Tambah ${n}`)}})}function z(x){ee(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let A=await I(`${a}/${x.id}`,{method:"DELETE"});A.ok?(j(`${n} berhasil dihapus.`),be(a),_()):N(A.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return _(),_}V();V();var De=null,it=null;async function _e(e=!1){if(De&&!e)return console.log("Employees Raw (Cache Hit)",De.slice(0,5)),De;let t=await I(`/api/employees?limit=${Te}&status=Aktif`);return De=(t.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",t.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",De.slice(0,5)),De}async function ce(e=!1){let i=(await _e(e)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",i.slice(0,5)),i}async function W(e=!1){return it&&!e||(it=((await I("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),it}function ne(e){let t={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!e||e==="-"||String(e).trim()===""?"":`<span class="badge ${t[e]||"badge-neutral"}">${e}</span>`}function wt(e){return e==null?'<span class="badge badge-neutral">-</span>':e<0?`<span class="badge badge-danger">Expired (${Math.abs(e)}h)</span>`:e<=14?`<span class="badge badge-danger">${e} hari</span>`:e<=30?`<span class="badge badge-warning">${e} hari</span>`:`<span class="badge badge-success">${e} hari</span>`}function Be(e){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[e]||"badge-neutral"}">${e||"-"}</span>`}function St(e){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[e]||"badge-neutral"}">${e||"-"}</span>`}function ye(e){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[e]||"badge-neutral"}">${e||"-"}</span>`}X();V();X();V();X();V();X();function rt(e){if(e.target.name==="report_date"||e.target.name==="completion_date"){let t=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),a=document.querySelector('input[name="day_count"]');if(t&&i&&a)if(t.value&&i.value){let o=new Date(t.value),b=new Date(i.value),u=Math.floor((b-o)/864e5);a.value=isNaN(u)?"":u}else a.value=""}}document.body.removeEventListener("input",rt);document.body.addEventListener("input",rt);document.body.removeEventListener("change",rt);document.body.addEventListener("change",rt);V();var Fe={};function Ie(e){if(Fe[e]){try{Fe[e].destroy()}catch{}delete Fe[e]}}function rn(){Object.keys(Fe).forEach(Ie)}var he=(e,t=0)=>{let i=Number(e);return isNaN(i)||e===null||e===void 0?t:i},Pe=(e,t="\u2014")=>{if(e==null||e==="")return t;let i=String(e).trim();return i===""||i==="[object Object]"?t:i};function ot(e,t,i=900){if(!e)return;e._animFrame&&cancelAnimationFrame(e._animFrame);let a=Math.max(0,Math.round(he(t))),o=e.textContent?e.textContent.replace(/[^\d]/g,""):"",b=o!==""&&parseInt(o)||0;if(a===b){e.textContent=a.toLocaleString("id-ID"),e._animFrame=null;return}let u=Date.now(),d=()=>{let n=Math.min((Date.now()-u)/i,1),h=1-Math.pow(1-n,3);e.textContent=Math.round(b+h*(a-b)).toLocaleString("id-ID"),n<1?e._animFrame=requestAnimationFrame(d):(e.textContent=a.toLocaleString("id-ID"),e._animFrame=null)};e._animFrame=requestAnimationFrame(d)}var on={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},sn=e=>{let t=Pe(e,"\u2014");return`<span class="status-pill ${on[t]||"pill-neutral"}">${t}</span>`};var ve={family:"Inter",size:11},Ee="#94A3B8",Le="#F1F5F9",Et=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],ln=()=>window.innerWidth<768;function st(e={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:ln()?"bottom":"top",labels:{font:ve,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ve,titleFont:{...ve,weight:"700"}}},scales:{x:{grid:{color:Le},ticks:{font:ve,color:Ee,maxRotation:0}},y:{grid:{color:Le},ticks:{font:ve,color:Ee},beginAtZero:!0}},...e}}var dn=()=>Array(6).fill(0).map(()=>`
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
  </div>`).join(""),pn=()=>Array(6).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Vt(e=3){return Array(e).fill(0).map((t,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<e-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function xe(e,t,i=8e3){try{let a=new AbortController,o=setTimeout(()=>a.abort(),i),b=await I(e,{signal:a.signal}).catch(()=>null);if(clearTimeout(o),!b||!b.ok)return t;let u=b.data;return u?u.data!==void 0?u.data??t:u:t}catch{return t}}function cn(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(a=>{let o=document.getElementById(a);o&&(o.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(a=>{let o=document.getElementById(a);if(o&&o.style.display==="none"){o.style.display="block";let b=o.parentElement;if(b&&!b.querySelector(".chart-empty")){let u=document.createElement("div");u.className="chart-empty",u.textContent="Belum ada data",o.style.display="none",b.appendChild(u)}}});let e=document.getElementById("kpi-row");e&&e.querySelector(".skeleton")&&Qt({});let t=document.getElementById("mini-stats-row");t&&t.querySelector(".skeleton")&&Xt({}),["table-contracts","table-issues"].forEach(a=>{let o=document.getElementById(a);o&&o.querySelector(".skeleton")&&(o.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Wt(e){rn(),e._dashRefresh&&clearInterval(e._dashRefresh),e._skelTimeout&&clearTimeout(e._skelTimeout);let t=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});e.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">
      
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;">
        <h2 style="margin:0; font-size:1.2rem; color:var(--text-1);">Dashboard Overview</h2>
        <button id="btn-dash-refresh" class="btn-primary" style="padding:6px 12px; font-size:0.8rem; display:flex; align-items:center; gap:6px; cursor:pointer; background:var(--primary); color:#fff; border:none; border-radius:6px;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
          Refresh Data
        </button>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${dn()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${pn()}</div>

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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>Yt(e)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async i=>{let a=i.target.value,o=document.getElementById("jadwal-year-label");o&&(o.textContent=a);let b=document.getElementById("skel-jadwal"),u=document.getElementById("chart-jadwal");b&&(b.style.display="block",b.style.position="absolute"),u&&(u.style.display="none");let d=await xe(`/api/dashboard/schedule-chart?year=${a}`,{},8e3);try{Zt(d)}catch(n){console.warn("ScheduleChart render:",n),ke("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let a=i.target.value,o=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",b=document.getElementById("skel-insp"),u=document.getElementById("chart-insp");b&&(b.style.display="block",b.style.position="absolute"),u&&(u.style.display="none");let d=await xe(o,{},8e3);try{ea(d)}catch(n){console.warn("InspBar render:",n),ke("skel-insp","chart-insp")}}),e._skelTimeout=setTimeout(()=>cn(),5e3),await Yt(e)}async function Yt(e){e._skelTimeout&&(clearTimeout(e._skelTimeout),e._skelTimeout=null);let[t,i,a,o,b,u,d,n]=await Promise.all([xe("/api/dashboard/kpi",{},8e3),xe("/api/dashboard/issues-trend",{},8e3),xe("/api/dashboard/issues-summary",{},8e3),xe("/api/dashboard/stats",{},8e3),xe("/api/dashboard/calendar",[],8e3),xe("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),xe(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),xe("/api/dashboard/dropdown-aggregates",{},8e3)]),h=document.getElementById("filter-insp-month"),F=h?h.value:"",$=F?`/api/dashboard/inspection-bar?month=${F}`:"/api/dashboard/inspection-bar",E=await xe($,{},8e3);window.dashboardAggregates=n||{};try{Qt(t)}catch(C){console.warn("KPI render:",C)}try{Xt(t)}catch(C){console.warn("MiniStats render:",C)}try{Zt(d)}catch(C){console.warn("ScheduleChart render:",C),ke("skel-jadwal","chart-jadwal")}try{un(Array.isArray(a?.by_category)?a.by_category:[])}catch(C){console.warn("Donut render:",C),ke("skel-donut","chart-donut")}try{gn(i)}catch(C){console.warn("Trend render:",C),ke("skel-trend","chart-trend")}try{ea(E)}catch(C){console.warn("InspBar render:",C),ke("skel-insp","chart-insp")}try{let C=Array.isArray(o)?o:Array.isArray(o?.recent_issues)?o.recent_issues:[];bn(C)}catch(C){console.warn("IssuesTable render:",C)}try{let C=Array.isArray(o?.expiring_contracts)?o.expiring_contracts:[];mn(u)}catch(C){console.warn("ContractsTable render:",C)}try{hn(Array.isArray(b)?b:[])}catch(C){console.warn("Agenda render:",C)}try{fn()}catch(C){console.warn("Quick Actions render:",C)}}function Qt(e){let t=document.getElementById("kpi-row");if(!t)return;e=e||{};let i=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,a=window.dashboardAggregates||{},o=n=>{let h=window.dashboardAggregates?.schedule_quarter_details?.[n];return h?`Insp: ${h.insp} \u2022 GC: ${h.gc} \u2022 DC: ${h.dc} \u2022 Fog: ${h.fog}`:""},b=n=>{let h=window.dashboardAggregates?.schedule_quarter_details?.[n];return h?`Rincian ${n}: Inspeksi: ${h.insp} | General Cleaning: ${h.gc} | Deep Cleaning: ${h.dc} | Fogging: ${h.fog}`:`Total Kegiatan ${n}`},u=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
          <option value="Q1" ${i==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${i==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${i==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${i==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:a.schedule_by_quarter?.[i]??e.schedule?.current??0,href:`#/timeline?dash_filter=period_${i.toLowerCase()}`,color:"kpi-blue",sub:o(i),subId:"mini-jadwal-detail",tooltip:b(i)}];t.innerHTML=u.map(n=>{let h=n.val!==void 0?he(n.val,0):he(e[n.key]?.current,0);return n.dropdown?`
        <div class="kpi-card ${n.color}" ${n.id?`id="${n.id}"`:""} ${n.tooltip?`title="${n.tooltip}"`:""} data-href="${n.href}" tabindex="0" role="link" style="cursor:pointer;padding:10px 12px">
          <div style="display:flex; gap:10px; align-items:center;">
            <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${n.icon}</span></div>
            <div style="flex:1;min-width:0;">
              <div style="display:flex; align-items:baseline; gap:3px;">
                <div class="kpi-value" data-target="${h}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${h}</div>
                ${n.dropdown}
              </div>
              <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${n.label}</div>
              <div class="kpi-subtitle" ${n.subId?`id="${n.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${n.tooltip||""}">${n.sub}</div>
            </div>
          </div>
        </div>`:`
      <a href="${n.href}" class="kpi-card ${n.color}" ${n.id?`id="${n.id}"`:""} ${n.tooltip?`title="${n.tooltip}"`:""} style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${n.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="kpi-value" data-target="${h}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${h}</div>
              ${n.dropdown?n.dropdown:""}
            </div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${n.label}</div>
            <div class="kpi-subtitle" ${n.subId?`id="${n.subId}"`:""} style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis" title="${n.tooltip||""}">${n.sub}</div>
          </div>
        </div>
      </a>`}).join(""),t.querySelectorAll(".kpi-card[data-href]").forEach(n=>{n.addEventListener("click",h=>{h.target.closest("select")||n.dataset.href&&(window.location.hash=n.dataset.href)}),n.addEventListener("keydown",h=>{h.key==="Enter"&&!h.target.closest("select")&&n.dataset.href&&(window.location.hash=n.dataset.href)})}),t.querySelectorAll(".kpi-value").forEach(n=>{ot(n,parseInt(n.dataset.target)||0)});let d=document.getElementById("dash-jadwal-period");if(d){let n=h=>{let $=(window.dashboardAggregates?.schedule_by_quarter||{})[h]||0,E=document.querySelector("#mini-jadwal .kpi-value");E&&(E.dataset.target=$,ot(E,$,400));let C=document.getElementById("mini-jadwal-detail");C&&(C.textContent=o(h),C.title=b(h));let w=document.getElementById("mini-jadwal");if(w){let k=`#/timeline?dash_filter=period_${h.toLowerCase()}`;w.tagName==="A"?w.href=k:w.dataset.href=k,w.title=b(h)}};d.addEventListener("change",h=>n(h.target.value))}}function Xt(e){let t=document.getElementById("mini-stats-row");if(!t)return;e=e||{};let i=new Date().getFullYear(),a=String(new Date().getMonth()+1).padStart(2,"0"),o=`${i}-${a}`,b=h=>`
    <select id="${h}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.stopPropagation();" onmousedown="event.stopPropagation();">
      <option value="${i}-01" ${o===`${i}-01`?"selected":""}>Jan</option>
      <option value="${i}-02" ${o===`${i}-02`?"selected":""}>Feb</option>
      <option value="${i}-03" ${o===`${i}-03`?"selected":""}>Mar</option>
      <option value="${i}-04" ${o===`${i}-04`?"selected":""}>Apr</option>
      <option value="${i}-05" ${o===`${i}-05`?"selected":""}>Mei</option>
      <option value="${i}-06" ${o===`${i}-06`?"selected":""}>Jun</option>
      <option value="${i}-07" ${o===`${i}-07`?"selected":""}>Jul</option>
      <option value="${i}-08" ${o===`${i}-08`?"selected":""}>Agu</option>
      <option value="${i}-09" ${o===`${i}-09`?"selected":""}>Sep</option>
      <option value="${i}-10" ${o===`${i}-10`?"selected":""}>Okt</option>
      <option value="${i}-11" ${o===`${i}-11`?"selected":""}>Nov</option>
      <option value="${i}-12" ${o===`${i}-12`?"selected":""}>Des</option>
    </select>
  `,u=window.dashboardAggregates||{},d=[{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:b("dash-inspeksi-month"),val:u.inspeksi_by_month?.[o]??e.inspection_month?.current??0,href:`#/timeline?dash_filter=inspeksi&month=${o}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:b("dash-gcdc-month"),val:u.gcdc_by_month?.[o]??e.cleaning_month?.current??0,href:`#/timeline?dash_filter=gcdc&month=${o}`,color:"mini-green"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:b("dash-reliefer-month"),val:u.relievers_by_month?.[o]??e.reliever_completed?.current??0,href:`#/relievers?dash_filter=reliever&month=${o}`,color:"mini-teal"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:b("dash-fogging-month"),val:u.fogging_by_month?.[o]??e.fogging_month?.current??0,href:`#/reports/fogging?dash_filter=fogging&month=${o}`,color:"mini-purple"},{icon:"\u{1F393}",label:"Training",val:e.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F3E2}",label:"Cabang",val:e.branches?.current,href:"#/branches",color:"mini-teal"}];t.innerHTML=d.map(h=>h.dropdown?`
        <div class="mini-stat ${h.color}" style="cursor:pointer;" id="${h.id||""}" data-href="${h.href}" tabindex="0" role="link">
          <div class="mini-stat-icon">${h.icon}</div>
          <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
            <div style="display:flex; align-items:baseline; gap:3px;">
              <div class="mini-stat-value" data-target="${he(h.val)}">0</div>
              ${h.dropdown}
            </div>
            <div class="mini-stat-text">${h.label}</div>
            ${h.extra||""}
          </div>
        </div>`:`
      <a href="${h.href}" class="mini-stat ${h.color}" style="text-decoration:none" id="${h.id||""}">
        <div class="mini-stat-icon">${h.icon}</div>
        <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
          <div style="display:flex; align-items:baseline; gap:3px;">
            <div class="mini-stat-value" data-target="${he(h.val)}">0</div>
            ${h.dropdown?h.dropdown:""}
          </div>
          <div class="mini-stat-text">${h.label}</div>
          ${h.extra||""}
        </div>
      </a>`).join(""),t.querySelectorAll(".mini-stat[data-href]").forEach(h=>{h.addEventListener("click",F=>{F.target.closest("select")||h.dataset.href&&(window.location.hash=h.dataset.href)}),h.addEventListener("keydown",F=>{F.key==="Enter"&&!F.target.closest("select")&&h.dataset.href&&(window.location.hash=h.dataset.href)})}),t.querySelectorAll(".mini-stat-value").forEach(h=>ot(h,parseInt(h.dataset.target)||0,700));let n=(h,F,$,E)=>{let C=document.getElementById(h);if(C){let w=k=>{let g=(window.dashboardAggregates?.[$]||{})[k]||0,r=document.querySelector(`#${F} .mini-stat-value`);r&&(r.dataset.target=g,ot(r,g,400));let s=document.getElementById(F);if(s){let m=`${E}&month=${k}`;s.tagName==="A"?s.href=m:s.dataset.href=m}};C.addEventListener("change",k=>w(k.target.value))}};n("dash-reliefer-month","mini-reliefer","relievers_by_month","#/relievers?dash_filter=reliever"),n("dash-inspeksi-month","mini-inspeksi","inspeksi_by_month","#/timeline?dash_filter=inspeksi"),n("dash-gcdc-month","mini-gcdc","gcdc_by_month","#/timeline?dash_filter=gcdc"),n("dash-fogging-month","mini-fogging","fogging_by_month","#/reports/fogging?dash_filter=fogging")}function un(e){ke("skel-donut","chart-donut");let t=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!t||!i)return;Ie("donut");let a=(e||[]).filter(n=>he(n.count)>0);if(!a.length){Ue(t,"Belum ada data permasalahan");return}let o=a.map(n=>`${Pe(n.category,"Lainnya")}`),b=a.map(n=>he(n.count)),u=b.reduce((n,h)=>n+h,0);i.innerHTML=a.map((n,h)=>{let F=Et[h%Et.length],$=u>0?Math.round(n.count/u*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${F}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${$}%)</span></div>
          <div class="donut-legend-label">${o[h]}</div>
        </div>
      </div>
    `}).join("");let d={id:"centerText",beforeDraw:function(n){let h=n.width,F=n.height,$=n.ctx;$.restore();let E=(F/80).toFixed(2);$.font="bold "+E+"em Inter",$.textBaseline="middle",$.fillStyle="#1E293B";let C=u.toString(),w=Math.round((h-$.measureText(C).width)/2),k=F/2;$.fillText(C,w,k-4),$.font="600 "+(E*.35).toFixed(2)+"em Inter",$.fillStyle="#64748B";let y="Total",g=Math.round((h-$.measureText(y).width)/2);$.fillText(y,g,k+10),$.save()}};Fe.donut=new Chart(t,{type:"doughnut",data:{labels:o,datasets:[{data:b,backgroundColor:Et,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ve,titleFont:{...ve,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutout:"75%"},plugins:[d]})}function gn(e){ke("skel-trend","chart-trend");let t=document.getElementById("chart-trend");if(!t)return;Ie("trend"),e=e||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(e.labels||[]).map(u=>{if(!u||typeof u!="string")return"";try{let[d,n]=u.split("-");return(i[Number(n)-1]||n)+" "+String(d).slice(-2)}catch{return u}}),o=(e.open||[]).map(u=>he(u)),b=(e.closed||[]).map(u=>he(u));if(!a.length){Ue(t,"Belum ada data trend");return}Fe.trend=new Chart(t,{type:"line",data:{labels:a,datasets:[{label:"Open",data:o,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:b,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:st({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Ee,maxRotation:45,autoSkip:!0}},y:{grid:{color:Le},ticks:{font:{family:"Inter",size:9},color:Ee},beginAtZero:!0}}})})}function Zt(e){ke("skel-jadwal","chart-jadwal");let t=document.getElementById("chart-jadwal");if(!t)return;Ie("jadwal"),e=e||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(e).some(n=>Array.isArray(n)&&n.some(h=>h>0))){Ue(t,"Belum ada data jadwal");return}let o=e["Inspeksi Hygiene"]||Array(12).fill(0),b=e["General Cleaning"]||Array(12).fill(0),u=e["Deep Cleaning"]||Array(12).fill(0),d=e.Fogging||Array(12).fill(0);Fe.jadwal=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Inspeksi",data:o,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:b,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:u,backgroundColor:"#F59E0B"},{label:"Fogging",data:d,backgroundColor:"#EF4444"}]},options:st({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:Ee,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Le},ticks:{font:{family:"Inter",size:9},color:Ee},min:0}}})})}function ea(e){ke("skel-insp","chart-insp");let t=document.getElementById("chart-insp");if(!t)return;Ie("inspBar"),e=e||{};let i=e.labels||[],a=(e.fc||[]).map(b=>he(b)),o=(e.spv||[]).map(b=>he(b));if(!i.length){Ue(t,"Belum ada data inspeksi");return}Fe.inspBar=new Chart(t,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:o,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:st({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ve,color:Ee,maxRotation:45,minRotation:45}},y:{grid:{color:Le},ticks:{font:ve,color:Ee},min:0,max:100}}})})}function mn(e){ke("skel-contract-mini","chart-contract-mini");let t=document.getElementById("chart-contract-mini");if(!t)return;Ie("contractMiniBar"),e=e||{};let i={"01":"Jan","02":"Feb","03":"Mar","04":"Apr","05":"Mei","06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(e.labels||[]).map(u=>{let d=u.split("-")[1];return i[d]||u}),o=(e.data||[]).map(u=>he(u));if(!a.length){Ue(t,"Belum ada data");return}let b=t.getContext("2d");Fe.contractMiniBar=new Chart(t,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:o,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:st({onClick:(u,d)=>{if(d&&d.length>0){let n=d[0].index,h=(e.labels||[])[n];h&&(window.location.hash="#/contracts?month_expiry="+h)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ve,color:Ee,maxRotation:0,autoSkip:!1}},y:{grid:{color:Le,borderDash:[4,4],drawBorder:!1},ticks:{font:ve,color:Ee,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function bn(e){let t=document.getElementById("table-issues");if(!t)return;let i=(e||[]).slice(0,8);if(!i.length){t.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}t.innerHTML=`
    <div class="dash-list">
      ${i.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${sn(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Pe(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Pe(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function hn(e){let t=document.getElementById("widget-agenda");if(!t)return;let i=new Date,a=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,b=(e||[]).filter(u=>(u.event_date||"").startsWith(a)).slice(0,10);if(!b.length){t.innerHTML="";return}t.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${b.map(u=>{let d="#3B82F6",n="#EFF6FF",h="Agenda",F=(u.title||"").toLowerCase();return F.includes("inspeksi")?(d="#10B981",n="#ECFDF5",h="Inspeksi"):F.includes("cleaning")||F.includes("gcdc")?(d="#3B82F6",n="#EFF6FF",h="Cleaning"):F.includes("reliefer")?(d="#F59E0B",n="#FFFBEB",h="Reliefer"):F.includes("fogging")&&(d="#8B5CF6",n="#F5F3FF",h="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(u.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${d};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Pe(u.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Pe(u.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${d}">${h}</div>
        </div>
      `}).join("")}
    </div>
  `}function fn(){let e=document.getElementById("quick-actions");if(!e)return;let t=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];e.innerHTML=t.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function ke(e,t){let i=document.getElementById(e),a=document.getElementById(t);if(i&&(i.style.display="none",i.style.position=""),a){a.style.display="block";let o=a.parentElement;if(o){let b=o.querySelector(".chart-empty");b&&b.remove()}}}function Ue(e,t="Belum ada data"){if(!e)return;e.style.display="none";let i=e.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent=t,i.appendChild(o)}}V();async function ta(e){document.getElementById("app").innerHTML=`
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
  `;let t=document.getElementById("login-form"),i=document.getElementById("login-error"),a=document.getElementById("login-btn"),o=document.getElementById("toggle-password"),b=document.getElementById("login-password");o?.addEventListener("click",()=>{let u=b.type==="text";b.type=u?"password":"text",o.style.color=u?"":"var(--primary)"}),t?.addEventListener("submit",async u=>{u.preventDefault(),i.style.display="none";let d=t.username.value.trim(),n=t.password.value;if(!d||!n){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let h=await I("/api/auth/login",{method:"POST",body:JSON.stringify({username:d,password:n})});h.ok&&h.data.success?(ft(h.data.data.token),qe(h.data.data.user),j("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=h.data.error||"Username atau password salah.",i.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}V();X();async function yn(){return await W()}function lt(e){if(!e||e==="-"||String(e).trim()==="")return"-";let t=String(e).trim(),i=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],a=t.match(/^(\d{4})-(\d{1,2})/);if(a){let b=a[1],u=parseInt(a[2],10)-1;if(u>=0&&u<12)return`${i[u]} ${b}`}let o=t.split(/[\/\-\.]/);if(o.length===2&&o[1].length===4){let b=parseInt(o[0],10)-1;if(b>=0&&b<12)return`${i[b]} ${o[1]}`}if(o.length===3&&o[2].length===4){let b=parseInt(o[1],10)-1;if(b>=0&&b<12)return`${i[b]} ${o[2]}`}for(let b=0;b<12;b++)if(t.toLowerCase().includes(i[b].toLowerCase()))return t;return t}function xn(e,t){let i=String(e.status||"").toLowerCase();return t==="active"?i==="aktif":t==="reliefer"?e.division==="FC - RELIEFER"&&i==="aktif":!1}async function aa(e,t){let i=await yn(),a=t?t.get("dash_filter"):null;Z({container:e,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",enableMobileFilterSheet:!0,itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:o=>a?o.filter(b=>xn(b,a)):o,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:o=>Be(o)},{key:"phone",label:"No. HP",render:o=>o?`<a href="tel:${o}">${o}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>ne(o)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:o=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:o?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:o?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:i,value:o?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:o?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:o?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:o?.status||""}]},{type:"row",fields:[{name:"target_pindah_os",label:"Target Pindah OS (Bulan)",type:"month",value:o?.target_pindah_os?String(o.target_pindah_os).slice(0,7):""},{name:"target_selesai_os",label:"Target Selesai OS (Bulan)",type:"month",value:o?.target_selesai_os?String(o.target_selesai_os).slice(0,7):""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let o=await I(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(o.ok){let b=o.data.data.map(u=>({"Nama Lengkap":u.full_name,Cabang:u.branch_name||"",Divisi:u.division||"","No. HP":u.phone||"","Tgl Masuk":u.join_date||"",Status:u.status||"","Target Pindah OS":lt(u.target_pindah_os)==="-"?"":lt(u.target_pindah_os),"Target Selesai OS":lt(u.target_selesai_os)==="-"?"":lt(u.target_selesai_os)}));G(b,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{G([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif","Target Pindah OS":"September 2026","Target Selesai OS":"Oktober 2026"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif","Target Pindah OS":"November 2026","Target Selesai OS":"Desember 2026"}],"Template_Import_Karyawan")},onImport:async o=>{let b=n=>{if(!n)return null;let h=String(n||"").toLowerCase(),F=i.find($=>String($.label||"").toLowerCase()===h);return F?F.value:null},u=o.map(n=>({full_name:String(n["Nama Lengkap"]||"").trim(),branch_id:b(String(n.Cabang||"").trim()),division:String(n.Divisi||n["Div / Bagian"]||"").trim()||"FACILITY CARE",phone:String(n["No. HP"]||n["No. Hp"]||"").trim(),join_date:String(n["Tgl Masuk"]||n["Tanggal Masuk"]||"").trim(),status:String(n.Status||"").trim(),target_pindah_os:String(n["Target Pindah OS"]||"").trim(),target_selesai_os:String(n["Target Selesai OS"]||"").trim(),notes:String(n.Catatan||"").trim()})).filter(n=>n.full_name),d=await I("/api/import/employees",{method:"POST",body:JSON.stringify({rows:u,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}}})}V();X();var Ft=[],na=[];async function vn(){Ft=await W(),na=await _e()}var Ct=async e=>{let t=[],i=1;for(;;){let o=await(await Promise.resolve().then(()=>(V(),Gt))).apiFetch(`${e}${e.includes("?")?"&":"?"}limit=100&page=${i}`);if(!o.ok)break;let b=o.data?.data||o.data||[],u=Array.isArray(b)?b:[];if(t=t.concat(u),u.length<100||o.data?.pagination&&i>=o.data.pagination.pages)break;i++}return t};async function dt(e,t){await vn(),Z({container:e,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",defaultFilters:{},onDataLoaded:a=>a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>Be(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,o)=>o.end_date&&String(o.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':wt(a)},{key:"status",label:"Status",render:a=>ne(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:Ft},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[b,u]=await Promise.all([Ct("/api/employees?status=Aktif"),Ct("/api/contracts")]);if(b.length>0){let d=u.filter($=>$.status==="Aktif"),n=new Set(d.map($=>$.employee_id)),h=b.filter($=>!n.has($.id)),F=`<p style="margin-bottom:12px">Data yang terbaca: <b>${b.length}</b> Karyawan Aktif, dan <b>${d.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${h.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;h.forEach($=>{let E=u.filter(w=>w.employee_id===$.id),C='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(E.length>0){let w=E[0];C=`Pernah ada kontrak (Status: <b style="color:#EF4444">${w.status}</b>, Selesai: ${window.formatDate(w.end_date)})`}F+=`<li style="margin-bottom:8px"><b>${$.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${$.branch_name||"-"} | ${C}</span></li>`}),F+="</ul>",Promise.resolve().then(()=>(ue(),kt)).then($=>$.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:F,cancelText:"Tutup"}))}}catch(b){console.error(b)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let o=document.querySelector(".page-actions");o&&o.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:na,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:Ft,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await I(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let o=a.data.data.map(b=>({"Nama Lengkap":b.employee_name,Cabang:b.branch_name||"","Div / Bagian":b.division||"","Tanggal Mulai":b.start_date||"","Tanggal Selesai":b.end_date&&String(b.end_date).startsWith("2099")?"":b.end_date||"","Sisa Kontrak":b.end_date&&String(b.end_date).startsWith("2099")?"Tetap":b.days_remaining!==null&&b.days_remaining!==void 0?`${b.days_remaining} Hari`:"",Status:b.status||""}));G(o,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{G([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[o,b]=await Promise.all([I("/api/branches?limit=10000"),Ct("/api/employees")]),u=o.data?.data||[],d=b||[];console.log(`Total employee yang berhasil dimuat dari database : ${d.length}`),d.length>0&&(console.log("Contoh 5 employee pertama:"),d.slice(0,5).forEach((y,g)=>{console.log(`${g+1}. ID: ${y.id}, Name: ${y.full_name}, Status: ${y.status}`)}));let n=y=>{if(!y)return null;let g=String(y||"").replace(/\s+/g," ").toLowerCase().trim(),r=u.find(s=>String(s.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===g||String(s.code||"").replace(/\s+/g," ").toLowerCase().trim()===g||String(s.name||"").replace(/\s+/g," ").toLowerCase().trim()===g);return r?r.id:null},h=(y,g)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${g}`),console.log(`Nama dari Excel : "${y}"`),!y)return console.log("Alasan gagal mapping : Nama kosong"),null;let r=String(y||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${r}"`),console.log(`Jumlah employee di database : ${d.length}`);let s=d.find(m=>String(m.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===r);return s?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${s.id}`),s.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},F=y=>{if(!y)return"";if(y instanceof Date&&!isNaN(y.getTime()))return y.toISOString().slice(0,10);let g=String(y).trim();if(/^\d{4,5}(\.\d+)?$/.test(g)){let s=Math.floor(Number(g));if(s>2e4&&s<99999){let m=new Date(Date.UTC(1899,11,30)+s*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);let r=g.split(/[\/\-\.]/);if(r.length===3){let[s,m,p]=r.map(f=>f.trim());if(s.length===4&&m.length<=2&&p.length<=2)return`${s}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&m.length<=2&&s.length<=2)return`${p}-${m.padStart(2,"0")}-${s.padStart(2,"0")}`}return g},$=a.map((y,g)=>{let r=g+2,s=String(y["Nama Lengkap"]||"").trim(),m=y["Tanggal Mulai"],p=F(m);if(!p){let c=a.__worksheet,v=a.__headers||[],S=v.indexOf("Tanggal Mulai"),B="N/A",T="N/A",D="N/A";if(S!==-1&&c&&window.XLSX){let P=window.XLSX.utils.encode_cell({c:S,r:r-1});D=P;let M=c[P];M?(B=M.t||"undefined",T=M.w||"undefined"):B="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let _="Unknown";m==null||m===""?_="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":m instanceof Date&&isNaN(m.getTime())?_="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":_="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${r}`),console.log(`Employee Name : ${s}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${S})`),console.log(`Raw Cell Value : "${m}"`),console.log(`JavaScript Type : ${typeof m}`),console.log(`SheetJS Cell Type : ${B}`),console.log(`SheetJS Formatted Value : "${T}"`),console.log(`Value After Trim : "${String(m||"").trim()}"`),console.log(`Value After Date Parser : "${p}"`),console.log(`Is Empty : ${!m}`),console.log(`Is Invalid Date : ${m instanceof Date?isNaN(m.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${_}`),console.log(`Workbook Sheet : ${c?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${D}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(y,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(v)),console.log(`==========================
`)}let f=h(s,r),l=null;return f?p||(l="Tanggal Mulai kosong atau tidak berformat tanggal"):l="Karyawan tidak ditemukan di Database",{isValid:!!(f&&p),invalidReason:l,rowNum:r,data:{employee_id:f,branch_id:n(String(y.Cabang||"").trim()),division:String(y["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:p,end_date:F(y["Tanggal Selesai"])||"2099-12-31",status:String(y.Status||"").trim(),_rawName:s}}}),E=[],C=[];if($.forEach(y=>{y.isValid?E.push(y.data):C.push({rowNum:y.rowNum,name:y.data._rawName,reason:y.invalidReason})}),console.log(`Split Validation - Valid: ${E.length}, Invalid: ${C.length}`),E.length===0){let y=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${a.length}
Valid: 0
Invalid: ${C.length}

Daftar Kegagalan (Contoh):
`;C.slice(0,10).forEach(g=>{y+=`- Row ${g.rowNum} | Nama: ${g.name} | Alasan: ${g.reason}
`}),C.length>10&&(y+=`- ... dan ${C.length-10} lainnya.
`),alert(y);return}let w=await I("/api/contracts/import",{method:"POST",body:JSON.stringify(E)}),k=`IMPORT SUMMARY
======================
`;k+=`Total Baris Excel : ${a.length}
`,k+=`Baris Valid       : ${E.length}
`,k+=`Baris Invalid     : ${C.length}

`,w&&w.data&&w.data.metrics?(k+=`Berhasil INSERT   : ${w.data.metrics.inserted}
`,k+=`Berhasil UPDATE   : ${w.data.metrics.updated}
`):k+=`Berhasil diproses : ${E.length}
`,C.length>0&&(k+=`
DAFTAR DATA DILEWATI:
`,C.forEach(y=>{k+=`- Row ${y.rowNum} | ${y.name} | ${y.reason}
`})),alert(k),typeof dt=="function"&&dt()}}})}V();X();var $t=[],Ve=[];function kn(e){if(!Array.isArray(e))return"Q3";let t=["Q4","Q3","Q2","Q1"];for(let i of t)if(e.some(a=>a.period===i))return i;return"Q3"}async function ia(e,t){$t=await W();let i=await ce();Ve=["Berlin Ariansyah","Ade Surahman"];let a=y=>y&&!Ve.find(g=>String(typeof g=="object"?g.value:g).toLowerCase()===String(y).toLowerCase())?[...Ve,y]:Ve,o=await I(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),b=y=>{if(!y||y==="-"||String(y).trim()==="")return"";let g=String(y).split("-");return g.length===3&&g[0].length===4?`${g[2]}-${g[1]}-${g[0]}`:y},u=o.data?.data||[],d=kn(u),n=t?t.get("dash_filter"):null,h=new Date,F=`${h.getFullYear()}-${String(h.getMonth()+1).padStart(2,"0")}`,$={},E=t&&t.get("month")?t.get("month"):null;n==="inspeksi"?$={status:"Done",activity_type:"Inspeksi Hygiene",month:E}:n==="gcdc"?$={status:"Done",activity_type:"GCDC",month:E}:n&&n.startsWith("period_")&&($={period:n.replace("period_","").toUpperCase()});let C=new Date().getFullYear(),k=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((y,g)=>{let r=String(g+1).padStart(2,"0");return{value:`${C}-${r}`,label:`${y} ${C}`}});Z({container:e,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:$,onDataLoaded:y=>y.sort((g,r)=>{let s=g.opening_date?new Date(g.opening_date).getTime():0;return(r.opening_date?new Date(r.opening_date).getTime():0)-s}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:y=>St(y)},{key:"period",label:"Periode",render:y=>ye(y)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:y=>b(y)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:y=>b(y)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:y=>b(y)},{key:"status",label:"Status",render:y=>ne(y)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:$t},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:k},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Ve}],formFields:y=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:$t,value:y?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:y?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:y?.period},{name:"pic",label:"PIC",type:"combobox",options:a(y?.pic),value:y?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:y?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:y?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:y?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:y?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:y?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let y=await I(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(y.ok){let g=y.data.data.map(r=>({Cabang:r.branch_name||"",Kegiatan:r.activity_type||"",Periode:r.period||"",PIC:r.pic||"","Tgl Opening":r.opening_date||"","Tgl Target":r.target_date||"","Tgl Selesai":r.completion_date||"",Status:r.status||""}));G(g,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{G([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async y=>{let r=(await I("/api/branches?all=1")).data?.data||[],s=l=>{if(!l)return null;let c=String(l||"").toLowerCase().trim(),v=c.match(/^(\d{3})/);if(v){let B=r.find(T=>String(T.code||"")===v[1]);if(B)return B.id}let S=r.find(B=>String(B.full_name||"").toLowerCase()===c||String(B.code||"").toLowerCase()===c||String(B.name||"").toLowerCase()===c);return S?S.id:null},m=l=>{if(l==null||l==="")return"";if(l instanceof Date&&!isNaN(l.getTime()))return l.toISOString().slice(0,10);let c=String(l).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let S=Number(c);if(S>2e4&&S<99999){let B=new Date(Date.UTC(1899,11,30)+S*864e5);return isNaN(B.getTime())?"":B.toISOString().slice(0,10)}}let v=c.split(/[\/\-\.]/);if(v.length===3){let[S,B,T]=v.map(D=>D.trim());if(S.length===4&&B.length<=2&&T.length<=2)return`${S}-${B.padStart(2,"0")}-${T.padStart(2,"0")}`;if(T.length===4&&B.length<=2&&S.length<=2)return`${T}-${B.padStart(2,"0")}-${S.padStart(2,"0")}`}return c},p=y.map(l=>({branch_id:s(String(l.Cabang||"").trim()),branch_name:String(l.Cabang||"").trim(),activity_type:String(l.Kegiatan||"").trim(),period:String(l.Periode||"").trim(),pic:String(l.PIC||l.Pic||"").trim(),opening_date:m(l["Tgl Opening"]||l["Tanggal Opening"]||l["Tgl Openir"]),target_date:m(l["Tgl Target"]||l["Tanggal Target"]),completion_date:m(l["Tgl Selesai"]||l["Tanggal Selesai"]),status:String(l.Status||"").trim(),notes:String(l.Catatan||l.Keterangan||"").trim()})).filter(l=>l.activity_type&&l.period),f=await I("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:p,onDuplicate:"update"})});if(!f.ok)throw new Error(f.data?.error||"Import gagal");return be("schedule"),f.data}}})}V();X();var _t=[],pt=[];function ct(e){if(e.target.name==="report_date"||e.target.name==="completion_date"){let t=document.querySelector('input[name="report_date"]'),i=document.querySelector('input[name="completion_date"]'),a=document.querySelector('input[name="day_count"]');if(t&&i&&a)if(t.value&&i.value){let o=new Date(t.value),b=new Date(i.value),u=Math.floor((b-o)/864e5);a.value=isNaN(u)?"":u}else a.value=""}}document.body.removeEventListener("input",ct);document.body.addEventListener("input",ct);document.body.removeEventListener("change",ct);document.body.addEventListener("change",ct);function wn(e,t){let i=String(e.status||"").toLowerCase();return t==="open"?i==="open":!1}async function ra(e,t){let i=t?t.get("dash_filter"):null;_t=await W(),pt=await ce();let a=n=>n&&!pt.find(h=>h.value===n)?[...pt,{value:n,label:n}]:pt,o=new Date().getFullYear(),b=["2025","2026","2027","2028","2029","2030"],d=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((n,h)=>{let F=String(h+1).padStart(2,"0");return{value:`${o}-${F}`,label:`${n} ${o}`}});Z({container:e,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:n=>i?n.filter(h=>wn(h,i)):n,columns:[{key:"report_date",label:"Tanggal Info",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:n=>`<span class="badge badge-secondary">${n}</span>`},{key:"source",label:"Sumber Laporan"},{key:"complaint",label:"Keluhan",render:n=>`<span title="${n}">${n?.length>50?n.slice(0,50)+"...":n}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:n=>`<span title="${n||""}">${n?.length>40?n.slice(0,40)+"...":n||"-"}</span>`},{key:"status",label:"Status",render:n=>ne(n)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"day_count",label:"Day",render:n=>n??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:_t},{type:"select",name:"month",label:"Bulan",options:d},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:b}],formFields:n=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:n?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:_t,value:n?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:n?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:n?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:n?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(n?.employee_name),value:n?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(n?.fc_specialist),value:n?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:n?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:n?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date},{name:"day_count",label:"Day",type:"number",value:n?.day_count,readonly:!0}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let n=await I(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let h=n.data.data.map(F=>({"Tanggal Info":F.report_date||"",Cabang:F.branch_name||"",Kategori:F.category||"","Sumber Laporan":F.source||"",Keluhan:F.complaint||"","Nama FC":F.employee_name||"","FC Spesialis":F.fc_specialist||"",Solusi:F.solution||"",Status:F.status||"","Tanggal Selesai":F.completion_date||"",Day:F.day_count!==null?F.day_count:""}));G(h,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{G([{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan",Status:"Done","Tanggal Selesai":"2024-03-02",Day:1}],"Template_Import_Permasalahan")},onImport:async n=>{let F=(await I("/api/branches?all=1")).data?.data||[],$=w=>{if(!w)return null;let k=String(w||"").toLowerCase(),y=F.find(g=>String(g.full_name||"").toLowerCase()===k||String(g.code||"").toLowerCase()===k||String(g.name||"").toLowerCase()===k);return y?y.id:null},E=n.map(w=>({branch_id:$(String(w.Cabang||"").trim()),report_date:String(w["Tanggal Info"]||w.Tanggal||"").trim(),category:String(w.Kategori||"").trim(),source:String(w["Sumber Laporan"]||w.Sumber||"").trim(),complaint:String(w.Keluhan||"").trim(),employee_name:String(w["Nama FC"]||"").trim(),fc_specialist:String(w["FC Spesialis"]||"").trim(),solution:String(w.Solusi||"").trim(),completion_date:String(w["Tanggal Selesai"]||w["Tgl Selesai"]||"").trim(),status:String(w.Status||"").trim(),day_count:w.Day||w.Hari||null})).filter(w=>w.report_date&&w.complaint&&w.category),C=await I("/api/import/issues",{method:"POST",body:JSON.stringify({rows:E,onDuplicate:"update"})});if(!C.ok)throw new Error(C.data?.error||"Import gagal");return C.data}}})}V();var Me=[];function Sn(e,t){let i=String(e.status||"").toLowerCase();return t==="pending"?i==="pending":!1}async function oa(e,t){let i=t?t.get("dash_filter"):null;Me=await W();let a=await ce(),o=["Ade","Berlin"],b=d=>d&&!a.find(n=>n.value===d)?[...a,{value:d,label:d}]:a,u=d=>d&&!o.find(n=>(typeof n=="object"?n.value:n)===d)?[...o,d]:o;Z({container:e,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:d=>i?d.filter(n=>Sn(n,i)):d,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:d=>window.formatDate(d)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:d=>`<span title="${d||""}">${d?.length>50?d.slice(0,50)+"\u2026":d||"-"}</span>`},{key:"solution",label:"Solusi",render:d=>`<span title="${d||""}">${d?.length>40?d.slice(0,40)+"\u2026":d||"-"}</span>`},{key:"status",label:"Status",render:d=>ne(d)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:d=>window.formatDate(d)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:d=>d?`<a href="${d}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:Me},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async d=>{let n=new URLSearchParams(d||{}).toString(),h=await I(`/api/one-on-one?limit=10000&${n}`);if(h.ok){let F=h.data.data.map(E=>({Tanggal:E.meeting_date||"",Cabang:E.branch_name||"","Nama Karyawan":E.employee_name||"",PIC:E.pic||"",Masalah:E.problem||"",Solusi:E.solution||"",Status:E.status||"","Tgl Selesai":E.completion_date||"",Dokumen:E.document_link||""})),{downloadExcel:$}=await Promise.resolve().then(()=>(X(),Se));$(F,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let d=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(X(),Se));n(d,"Template_Import_OneOnOne")},onImport:async d=>{let n=E=>{if(!E)return null;let C=String(E||"").toLowerCase(),w=Me.find(k=>String(k.label||"").toLowerCase()===C);return w?w.value:null},h=E=>{if(!E)return"";if(E instanceof Date&&!isNaN(E.getTime()))return E.toISOString().slice(0,10);let C=String(E).trim();if(/^\d{4,5}$/.test(C)){let k=Number(C);if(k>2e4&&k<99999){let y=new Date(Date.UTC(1899,11,30)+k*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(C))return C.slice(0,10);let w=C.split(/[\/\-\.]/);if(w.length===3){let[k,y,g]=w.map(r=>r.trim());if(k.length===4&&y.length<=2&&g.length<=2)return`${k}-${y.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&y.length<=2&&k.length<=2)return`${g}-${y.padStart(2,"0")}-${k.padStart(2,"0")}`}return C},F=d.map(E=>({meeting_date:h(E.Tanggal),employee_name:String(E["Nama Karyawan"]||"").trim(),branch_id:n(String(E.Cabang||"").trim()),pic:String(E.PIC||"").trim(),problem:String(E.Masalah||"").trim(),solution:String(E.Solusi||"").trim(),status:String(E.Status||"").trim(),completion_date:h(E["Tgl Selesai"]),document_link:String(E.Dokumen||"").trim()})).filter(E=>E.meeting_date&&E.employee_name&&E.branch_id),$=await I("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:F,onDuplicate:"update"})});if(!$.ok)throw new Error($.data?.error||"Import gagal");return $.data}},formFields:d=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:d?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:d?.branch_id&&!Me.find(n=>n.value==d.branch_id)?[...Me,{value:d.branch_id,label:d.branch_name||d.branch_id}]:Me,createApi:{path:"/api/branches",field:"full_name"},value:d?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:b(d?.employee_name),value:d?.employee_name},{name:"pic",label:"PIC",type:"select",options:u(d?.pic),createApi:{path:"/api/pic",field:"name"},value:d?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:d?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:d?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:d?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:d?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:d?.document_link}]})}V();ue();var Tt=[{id:1,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:6,branch_name:"006. Kelapa Gading MOI",trainer:"Berlin Ariansyah",participants:"MUHAMMAD PRAKAS AGAZI",score:90,document_link:""},{id:2,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:11,branch_name:"011. Kuningan",trainer:"Berlin Ariansyah",participants:"KAMALLUDIN",score:88,document_link:""},{id:3,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:42,branch_name:"042. Muara Karang",trainer:"Berlin Ariansyah",participants:"SAHRUL ICHSANI",score:85,document_link:""},{id:4,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:51,branch_name:"051. Bintaro Sektor 9",trainer:"Berlin Ariansyah",participants:"MUHAMAD VICKRI FATHURROHMAN",score:92,document_link:""},{id:5,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:45,branch_name:"045. FX Mall",trainer:"Berlin Ariansyah",participants:"DEDE AHDORI",score:86,document_link:""},{id:6,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:54,branch_name:"054. Bogor Gunung Batu",trainer:"Berlin Ariansyah",participants:"ADITYA PRATAMA",score:95,document_link:""},{id:7,training_date:"2026-06-29",batch:"Batch 5",subject:"Basic Knowledge & Practical Skills",branch_id:55,branch_name:"055. Gajah Mada",trainer:"Berlin Ariansyah",participants:"ELVAN JANGGA PRAMANA",score:89,document_link:""}],sa=["Batch 5","Batch 4","Batch 3","Batch 2","Batch 1"],la=["Berlin Ariansyah","Ade Surahman","Mizwar","Fajar","Ade","Berlin"];function En(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}function Ae(e){if(!e)return"-";try{let t=JSON.parse(e);if(Array.isArray(t))return t.join(", ")}catch{}return String(e)}async function da(e){let t=[],i=[],a=[],o="all",b="all",u="all",d="",n=!1,F=ae()?.role!=="viewer";try{[i,a]=await Promise.all([W(),ce()])}catch(r){console.error("Error fetching cache:",r)}async function $(){n=!0,y();try{let r=await I("/api/training?limit=10000");if(r.ok){let s=r.data?.data||r.data||[];t=Array.isArray(s)&&s.length>0?s:Tt}else t=Tt}catch(r){console.warn("Gagal load training, menggunakan data lokal:",r),t=Tt}finally{n=!1,y()}}async function E(){try{let{downloadExcel:r}=await Promise.resolve().then(()=>(X(),Se)),s=t.map((m,p)=>({No:p+1,Tanggal:m.training_date||"",Batch:m.batch||"",Materi:m.subject||"",Cabang:m.branch_name||"",Trainer:m.trainer||"",Peserta:Ae(m.participants),Nilai:m.score!=null?m.score:"",Dokumen:m.document_link||"",Catatan:m.notes||""}));r(s,`Data_Training_${new Date().toISOString().slice(0,10)}`),j("Data training berhasil diekspor ke Excel!")}catch(r){N("Gagal ekspor Excel: "+r.message)}}async function C(r){if(!F)return;let s=t.find(m=>m.id===r);ee({title:"Hapus Data Training",message:`Apakah Anda yakin ingin menghapus data sesi training "${s?.subject||"ini"}" untuk ${Ae(s?.participants)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let m=await I(`/api/training/${r}`,{method:"DELETE"});m.ok?(j("Data training berhasil dihapus"),t=t.filter(p=>p.id!==r),y()):N(m.data?.error||"Gagal menghapus data training")}catch(m){N("Gagal menghapus training: "+m.message)}}})}function w(r=null){if(!F)return;let s=!!r,m=r?.training_date?r.training_date.slice(0,10):new Date().toISOString().slice(0,10),p=r?.batch||"Batch 5",f=r?.trainer||"Berlin Ariansyah",l=Ae(r?.participants)!=="-"?Ae(r?.participants):"",c=document.createElement("div");c.innerHTML=`
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
              placeholder="Pilih atau ketik batch..." value="${p}" required />
            <datalist id="tr-batch-options">
              ${sa.map(v=>`<option value="${v}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Materi / Topik Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="tr-modal-subject" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Basic Knowledge & Practical Skills" value="${r?.subject?String(r.subject).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="tr-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${i.map(v=>`
                <option value="${v.value}" ${String(r?.branch_id)===String(v.value)?"selected":""}>
                  ${v.label}
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
              ${la.map(v=>`<option value="${v}">`).join("")}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Peserta Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="tr-modal-participants" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Ketik nama karyawan peserta training..." required>${l}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nilai / Score Evaluasi (0 - 100)
            </label>
            <input type="number" min="0" max="100" step="1" id="tr-modal-score" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: 90" value="${r?.score!=null?r.score:""}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Link Dokumen / Sertifikat (Opsional)
            </label>
            <input type="url" id="tr-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="https://drive.google.com/..." value="${r?.document_link?String(r.document_link).replace(/"/g,"&quot;"):""}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan Pelatihan
          </label>
          <textarea id="tr-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Catatan hasil pelatihan, rekomendasi materi lanjutan...">${r?.notes||""}</textarea>
        </div>
      </div>
    `,Y({title:s?"\u270F\uFE0F Edit Data Sesi Training":"\u2795 Tambah Data Sesi Training",content:c,size:"md",confirmText:s?"Simpan Perubahan":"Simpan Training",confirmClass:"btn-primary",onConfirm:async(v,S)=>{let B=c.querySelector("#tr-modal-date").value,T=(c.querySelector("#tr-modal-batch").value||"").trim(),D=(c.querySelector("#tr-modal-subject").value||"").trim(),_=c.querySelector("#tr-modal-branch").value,P=(c.querySelector("#tr-modal-trainer").value||"").trim(),M=(c.querySelector("#tr-modal-participants").value||"").trim(),z=c.querySelector("#tr-modal-score").value,x=z!==""?Number(z):null,A=(c.querySelector("#tr-modal-link").value||"").trim(),L=(c.querySelector("#tr-modal-notes").value||"").trim();if(!B){N("Tanggal training wajib diisi");return}if(!D){N("Materi / topik pelatihan wajib diisi");return}if(!_){N("Cabang klinik wajib dipilih");return}if(!P){N("Trainer wajib diisi");return}if(!M){N("Nama peserta training wajib diisi");return}let q=i.find(H=>String(H.value)===String(_)),O={training_date:B,batch:T,subject:D,branch_id:Number(_),branch_name:q?q.label:"Cabang",trainer:P,participants:M,score:x,document_link:A,notes:L};try{if(s){let H=await I(`/api/training/${r.id}`,{method:"PUT",body:JSON.stringify(O)});if(H.ok){j("Data training berhasil diperbarui");let K=t.findIndex(R=>R.id===r.id);K!==-1&&(t[K]={...t[K],...O}),S(),y()}else N(H.data?.error||"Gagal memperbarui data training")}else{let H=await I("/api/training",{method:"POST",body:JSON.stringify(O)});if(H.ok){j("Sesi training baru berhasil ditambahkan");let K=H.data?.id||Date.now();t.unshift({id:K,...O}),S(),y()}else N(H.data?.error||"Gagal menambahkan training")}}catch(H){N("Terjadi kesalahan: "+H.message)}}})}function k(){return t.filter(r=>{if(d){let s=d.toLowerCase(),m=(r.subject||"").toLowerCase().includes(s),p=(r.trainer||"").toLowerCase().includes(s),f=Ae(r.participants).toLowerCase().includes(s),l=(r.branch_name||"").toLowerCase().includes(s),c=(r.batch||"").toLowerCase().includes(s);if(!m&&!p&&!f&&!l&&!c)return!1}return!(o!=="all"&&r.batch!==o||b!=="all"&&String(r.branch_id)!==String(b)||u!=="all"&&r.trainer!==u)})}function y(){let r=k(),s=t.length,m=new Set;t.forEach(T=>{Ae(T.participants).split(",").forEach(_=>{let P=_.trim();P&&P!=="-"&&m.add(P)})});let p=m.size,l=new Set(t.map(T=>T.batch).filter(Boolean)).size,c=new Date,v=c.getFullYear(),S=c.getMonth(),B=t.filter(T=>{if(!T.training_date)return!1;let D=new Date(T.training_date);return!isNaN(D.getTime())&&D.getFullYear()===v&&D.getMonth()===S}).length;e.innerHTML=`
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
            gap: 8px !important;
            flex-wrap: nowrap !important;
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
            grid-template-columns: 1fr;
          }
        }
      </style>

      <div class="training-page">

        <!-- Hero Header Card -->
        <div class="training-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                <span style="font-size:1.6rem;">\u{1F393}</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Training Karyawan
                </h1>
              </div>
              <p style="margin:0; font-size:0.88rem; color:#CBD5E1; max-width:650px; line-height:1.5;">
                Sistem evaluasi pengembangan kompetensi, sertifikasi keahlian, dan arsip riwayat pelatihan personil PT Audy Mandiri Indonesia.
              </p>
            </div>

            <div class="training-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${F?`
                <button id="tr-btn-add" class="btn btn-primary training-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Training</span>
                </button>
              `:""}
              
              <button id="tr-btn-excel" class="btn btn-ghost training-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="tr-btn-refresh" class="btn btn-ghost training-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
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
                <div class="training-stat-val">${s}</div>
                <div class="training-stat-lbl">Sesi Training</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                \u{1F465}
              </div>
              <div>
                <div class="training-stat-val">${p}</div>
                <div class="training-stat-lbl">Peserta Terlatih</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4DA}
              </div>
              <div>
                <div class="training-stat-val">${l} Batch</div>
                <div class="training-stat-lbl">Total Angkatan</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="training-stat-val">${B}</div>
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
              <input type="text" id="tr-search-input" class="training-search-input" placeholder="Cari materi, trainer, peserta, atau cabang..." value="${d}" />
            </div>

            <select id="tr-filter-branch" class="training-select">
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(T=>`<option value="${T.value}" ${String(b)===String(T.value)?"selected":""}>${T.label}</option>`).join("")}
            </select>

            <select id="tr-filter-trainer" class="training-select">
              <option value="all">Semua Trainer</option>
              ${la.map(T=>`<option value="${T}" ${u===T?"selected":""}>${T}</option>`).join("")}
            </select>
          </div>

          <!-- Batch Filter Pills (Swipeable on Mobile) -->
          <div class="training-filter-pills">
            <div class="training-pill ${o==="all"?"active":""}" data-batch="all">
              <span>Semua Batch</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${s}</span>
            </div>
            ${sa.map(T=>{let D=t.filter(_=>_.batch===T).length;return`
                <div class="training-pill ${o===T?"active":""}" data-batch="${T}">
                  <span>\u{1F393}</span>
                  <span>${T}</span>
                  ${D>0?`<span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${D}</span>`:""}
                </div>
              `}).join("")}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="training-list-container">
          ${r.length===0?`
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
              ${r.map(T=>{let D=Ae(T.participants),P=D.split(",")[0].trim().split(" ").map(M=>M[0]).filter(Boolean).slice(0,2).join("").toUpperCase()||"TR";return`
                  <div class="training-card" data-id="${T.id}">
                    <div>
                      <!-- Card Header: Batch & Nilai/Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="training-batch-badge">
                          <span>\u{1F393}</span>
                          <span>${T.batch||"Reguler"}</span>
                        </span>
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${T.score!=null?`
                            <span class="training-score-badge">
                              <span>\u2B50</span>
                              <span>${T.score}</span>
                            </span>
                          `:""}
                          <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                            \u{1F5D3}\uFE0F ${En(T.training_date)}
                          </span>
                        </div>
                      </div>

                      <!-- Subject / Materi Title -->
                      <h4 style="margin:0 0 10px 0; font-size:1.05rem; font-weight:700; color:#0F172A; line-height:1.4;">
                        ${T.subject||"Materi Pelatihan"}
                      </h4>

                      <!-- Participant & Branch Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:10px;">
                        <div class="training-avatar-box">
                          ${P}
                        </div>
                        <div style="min-width:0;">
                          <div style="font-size:0.72rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:1px;">Peserta Training</div>
                          <div style="font-size:0.95rem; font-weight:700; color:#1E293B; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${D}">
                            ${D}
                          </div>
                        </div>
                      </div>

                      <!-- Detail Box: Cabang & Trainer -->
                      <div class="training-info-box">
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#334155;">
                          <span>\u{1F4CD}</span>
                          <span style="font-weight:600;">${T.branch_name||"Cabang Belum Ditentukan"}</span>
                        </div>
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#64748B;">
                          <span>\u{1F468}\u200D\u{1F3EB}</span>
                          <span>Trainer: <strong style="color:#0F172A;">${T.trainer||"-"}</strong></span>
                        </div>
                      </div>

                      ${T.notes?`
                        <div style="font-size:0.78rem; color:#64748B; font-style:italic; line-height:1.4; margin-bottom:8px;">
                          \u{1F4DD} ${T.notes}
                        </div>
                      `:""}
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:8px;">
                      <div>
                        ${T.document_link?`
                          <a href="${T.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>\u{1F4C4} Materi / Sertifikat</span>
                          </a>
                        `:`
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${F?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          <button class="training-icon-btn training-edit-btn" title="Edit Data Training" data-id="${T.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                          <button class="training-icon-btn delete training-delete-btn" title="Hapus Training" data-id="${T.id}">
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
    `,g()}function g(){let r=e.querySelector("#tr-btn-add");r&&(r.onclick=()=>w());let s=e.querySelector("#tr-btn-excel");s&&(s.onclick=()=>E());let m=e.querySelector("#tr-btn-refresh");m&&(m.onclick=()=>$());let p=e.querySelector("#tr-btn-reset-filter");p&&(p.onclick=()=>{d="",b="all",o="all",u="all",y()});let f=e.querySelector("#tr-search-input");f&&(f.oninput=v=>{d=v.target.value,y();let S=e.querySelector("#tr-search-input");S&&(S.focus(),S.setSelectionRange(S.value.length,S.value.length))});let l=e.querySelector("#tr-filter-branch");l&&(l.onchange=v=>{b=v.target.value,y()});let c=e.querySelector("#tr-filter-trainer");c&&(c.onchange=v=>{u=v.target.value,y()}),e.querySelectorAll(".training-pill").forEach(v=>{v.onclick=()=>{o=v.getAttribute("data-batch")||"all",y()}}),e.querySelectorAll(".training-edit-btn").forEach(v=>{v.onclick=S=>{S.stopPropagation();let B=Number(v.getAttribute("data-id")),T=t.find(D=>D.id===B);T&&w(T)}}),e.querySelectorAll(".training-delete-btn").forEach(v=>{v.onclick=S=>{S.stopPropagation();let B=Number(v.getAttribute("data-id"));C(B)}})}await $()}V();ue();X();function pa({container:e,title:t,icon:i,apiPath:a,columns:o,formFields:b,filterFields:u,defaultFilters:d={},itemLabel:n="Data",canCreate:h=!0,canEdit:F=!0,canDelete:$=!0,onBeforeSubmit:E,onAfterLoad:C,onDataLoaded:w,extraActions:k=[],initialSearch:y="",exportOptions:g=null,bulkDelete:r=!1,paginationMode:s="server"}){let m=ae();m&&m.role==="viewer"&&(h=!1,F=!1,$=!1,r=!1,g=null),m&&m.role==="editor_khusus"&&a!=="/api/issues"&&a!=="/api/overtime"&&(h=!1,F=!1,$=!1,r=!1,g=null),m&&m.role==="input_lembur"&&(a==="/api/overtime"?($=!1,r=!1,g=null):(h=!1,F=!1,$=!1,r=!1,g=null));let p=1,f={...d};y&&(f.search=y);let l=new Set;e.innerHTML=`
    ${r?`
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
        ${h?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
        ${g?`
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              \u22EE Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${g.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4E5} Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${g.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                \u{1F4C4} Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${g.moduleName}">
                \u{1F4E4} Import Excel
                <input type="file" id="input-import-${g.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
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
    

    ${u&&u.length>0?`
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${u.filter(x=>x.type==="search").map(x=>`<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${x.placeholder||"Cari..."}" id="filter-search" value="${f.search||""}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`).join("")}
        
        <div class="filter-dropdowns-desktop">
          ${u.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let A=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select" name="${x.name}" id="filter-${x.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${A}</option>${(x.options||[]).map(L=>`<option value="${typeof L=="object"?L.value:L}" ${f[x.name]===(typeof L=="object"?L.value:L)?"selected":""}>${typeof L=="object"?L.label:L}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">\u2699 Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${u.filter(x=>x.type!=="search").map(x=>{if(x.type==="select"||x.type==="combobox"){let A=(x.label||"").startsWith("Pilih")?x.label:`Pilih ${x.label||""}`;return`<select class="filter-select filter-select-sheet" name="${x.name}-sheet" id="filter-sheet-${x.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${A}</option>${(x.options||[]).map(L=>`<option value="${typeof L=="object"?L.value:L}" ${f[x.name]===(typeof L=="object"?L.value:L)?"selected":""}>${typeof L=="object"?L.label:L}</option>`).join("")}</select>`}return""}).join("")}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function c(){let x=document.getElementById("bulk-toolbar");if(!x)return;let A=document.getElementById("bulk-count"),L=document.getElementById("btn-bulk-delete"),q=document.getElementById("btn-bulk-cancel");A.textContent=`${l.size} item dipilih`,l.size>0?(x.style.display="flex",L.disabled=!1,q.disabled=!1):(x.style.display="none",L.disabled=!0,q.disabled=!0);let O=document.getElementById("select-all-checkbox");if(O){let H=document.querySelectorAll(".row-checkbox");if(H.length>0){let K=[...H].every(ie=>ie.checked),R=[...H].some(ie=>ie.checked);O.checked=K,O.indeterminate=R&&!K}else O.checked=!1,O.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{l.clear(),document.querySelectorAll(".row-checkbox").forEach(A=>A.checked=!1);let x=document.getElementById("select-all-checkbox");x&&(x.checked=!1),c()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(l.size===0)return;let x=[...l],A=document.createElement("div");A.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",A.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${x.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${x.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(A),A.querySelector("#bulk-cancel-btn").addEventListener("click",()=>A.remove()),A.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let L=A.querySelector("#bulk-confirm-btn");L.disabled=!0,L.textContent="Menghapus...";let q=await I(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:x})});A.remove(),q.ok?(j(`${x.length} ${n} berhasil dihapus.`),l.clear(),c(),be(a),_()):N(q.data?.error||"Gagal menghapus data.")})});let v=document.getElementById("filter-search"),S;v?.addEventListener("input",x=>{clearTimeout(S),S=setTimeout(()=>{f.search=x.target.value,p=1,l.clear(),c(),_()},400)}),u?.forEach(x=>{(x.type==="select"||x.type==="combobox")&&(document.getElementById(`filter-${x.name}`)?.addEventListener("change",A=>{f[x.name]=A.target.value;let L=document.getElementById(`filter-sheet-${x.name}`);L&&(L.value=A.target.value),p=1,l.clear(),c(),_()}),document.getElementById(`filter-sheet-${x.name}`)?.addEventListener("change",A=>{f[x.name]=A.target.value;let L=document.getElementById(`filter-${x.name}`);L&&(L.value=A.target.value),p=1,l.clear(),c(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}))}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{f={},v&&(v.value=""),u?.forEach(x=>{let A=document.getElementById(`filter-${x.name}`);A&&(A.value="");let L=document.getElementById(`filter-sheet-${x.name}`);L&&(L.value="")}),p=1,l.clear(),c(),_()}),document.getElementById("btn-reset-filter-sheet")?.addEventListener("click",()=>{f={},v&&(v.value=""),u?.forEach(x=>{let A=document.getElementById(`filter-${x.name}`);A&&(A.value="");let L=document.getElementById(`filter-sheet-${x.name}`);L&&(L.value="")}),p=1,l.clear(),c(),_(),document.getElementById("filter-options-wrapper")?.classList.remove("sheet-open")}),document.getElementById("btn-create")?.addEventListener("click",()=>M(null)),g&&document.addEventListener("click",function(x){let A=document.getElementById("aksi-menu-main"),L=document.getElementById("btn-aksi-main");A&&L&&!L.contains(x.target)&&!A.contains(x.target)&&A.classList.remove("show-aksi-menu")});let B=document.getElementById("btn-mobile-filter"),T=document.getElementById("filter-options-wrapper"),D=document.getElementById("btn-close-filter-sheet");if(B&&T&&(B.addEventListener("click",x=>{x.preventDefault(),T.classList.add("sheet-open")}),D&&D.addEventListener("click",x=>{x.preventDefault(),T.classList.remove("sheet-open")})),g){document.getElementById(`btn-export-${g.moduleName}`)?.addEventListener("click",async A=>{let L=A.target,q=L.innerHTML;L.innerHTML="\u23F3 Loading...",L.disabled=!0;try{await g.onExport()}catch{N("Gagal export data")}finally{L.innerHTML=q,L.disabled=!1}}),document.getElementById(`btn-template-${g.moduleName}`)?.addEventListener("click",()=>{g.onTemplate()});let x=document.getElementById(`input-import-${g.moduleName}`);x?.addEventListener("change",async A=>{let L=A.target.files[0];if(!L)return;let q=document.getElementById(`label-import-${g.moduleName}`),O=q?q.querySelector(".import-text"):null,H=O?O.innerText:"";O&&(O.innerText="\u231B Memproses..."),q&&(q.style.pointerEvents="none"),x.disabled=!0;try{let K=await Ge(L);if(K.length===0)throw new Error("File kosong atau format salah");await g.onImport(K),j("Import berhasil!"),be(a),_()}catch(K){N(K.message||"Gagal import data")}finally{O&&(O.innerText=H),q&&(q.style.pointerEvents="auto"),x.disabled=!1,x.value=""}})}async function _(){c();let x=document.getElementById("table-container");if(!x)return;x.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let A=s==="client",L=A?1:p,q=A?Te:20,O=new URLSearchParams({page:L,limit:q,...Object.fromEntries(Object.entries(f).filter(([,U])=>U))}),H=await I(`${a}?${O}`);if(!H.ok){x.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${H.data?.error||"Error"}</p></div>`;return}let K=H.data?.data||H.data||[],R=H.data?.pagination,ie=K.length;if(A){K=w(K);let U=K.length,re=20,oe=Math.ceil(U/re);p>oe&&oe>0&&(p=oe);let J=(p-1)*re,pe=p*re;K=K.slice(J,pe),R={page:p,limit:re,total:U,pages:oe}}!1,C&&C(K);let de=Ze({columns:o,data:K,onEdit:F?U=>M(U):null,actions:k.map(U=>({...U,handler:re=>U.handler(re,_)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:r?{selectedIds:l,onToggle:c}:null});x.innerHTML="",x.appendChild(de);let me=document.getElementById("pagination-container");if(me&&(me.innerHTML="",R&&R.pages>1)){let U=et({page:R.page,pages:R.pages,total:R.total,limit:R.limit,onPage:re=>{p=re,_()}});U&&me.appendChild(U)}}function P(x){let A=typeof b=="function"?b(x):b;return Je(A)}function M(x){let A=!!x,L=document.createElement("form");if(L.noValidate=!0,L.innerHTML=P(x),A){let O=typeof b=="function"?b(x):b;at(L,x)}let{close:q}=Y({title:A?`Edit ${n}`:`Tambah ${n}`,content:L,size:"lg",confirmText:A?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(O,H)=>{if(!L.reportValidity())return;let K=O.querySelector(".modal-confirm");K.disabled=!0,K.textContent="Menyimpan...";let R=tt(L),ie=typeof b=="function"?b(x):b,de=async oe=>{for(let J of oe)if(J.type==="row")await de(J.fields);else if(J.type==="combobox"&&R[J.name]){let pe=R[J.name],fe=(J.options||[]).find(se=>{let ge=String(typeof se=="object"?se.value:se),mt=String(typeof se=="object"?se.label:se);return ge===pe||mt===pe});if(fe)R[J.name]=typeof fe=="object"?fe.value:fe;else if(J.createApi){let se={};se[J.createApi.field]=pe,J.createApi.extra&&Object.assign(se,J.createApi.extra);let ge=await I(J.createApi.path,{method:"POST",body:JSON.stringify(se)});if(ge.ok&&ge.data?.id)R[J.name]=ge.data.id;else if(ge.ok&&!ge.data?.id)R[J.name]=pe;else throw new Error(`Gagal membuat master data: ${ge.data?.error||"Unknown error"}`)}}};try{await de(ie)}catch(oe){N(oe.message),K.disabled=!1,K.textContent=A?"Simpan Perubahan":`Tambah ${n}`;return}E&&(R=await E(R,x));let me=A?"PUT":"POST",U=A?`${a}/${x.id}`:a,re=await I(U,{method:me,body:JSON.stringify(R)});re.ok?(j(A?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),H(),be(a),_()):(N(re.data?.error||"Gagal menyimpan data."),K.disabled=!1,K.textContent=A?"Simpan Perubahan":`Tambah ${n}`)}})}function z(x){ee(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let A=await I(`${a}/${x.id}`,{method:"DELETE"});A.ok?(j(`${n} berhasil dihapus.`),be(a),_()):N(A.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return _(),_}V();X();async function ca(e,t){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await W(),a=await ce(),o=t?t.get("dash_filter"):null,b={};if(o==="reliever"){let E=new Date,C=`${E.getFullYear()}-${String(E.getMonth()+1).padStart(2,"0")}`;b={status:"Done",month:t&&t.get("month")?t.get("month"):C}}console.log("RAW",await _e()),console.log("OPTIONS",a);let u=E=>E&&!a.find(C=>C.value===E)?[...a,{value:E,label:E}]:a,d=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],n=E=>E&&!d.includes(E)?[...d,E]:d,h=new Date().getFullYear(),$=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((E,C)=>{let w=String(C+1).padStart(2,"0");return{value:`${h}-${w}`,label:`${E} ${h}`}});pa({container:e,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:b,onDataLoaded:E=>E.sort((C,w)=>{let k=C.backup_date?new Date(C.backup_date).getTime():0;return(w.backup_date?new Date(w.backup_date).getTime():0)-k}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:E=>ye(E)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:E=>window.formatDate(E)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:E=>window.formatDate(E)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:E=>E?`<span class="badge badge-info">${E}</span>`:"-"},{key:"status",label:"Status",render:E=>ne(E)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:d},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"month",label:"Bulan",options:$},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:E=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:E?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:E?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:u(E?.original_fc_name),value:E?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:n(E?.reliever_name),value:E?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:E?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:E?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:E?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:E?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:E?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let E=await I(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(E.ok){let C=E.data.data.map(w=>({Cabang:w.branch_name||"","Nama Facility care":w.original_fc_name||"",Periode:w.period||"",Relifer:w.reliever_name||"","Tanggal Back Up":w.backup_date||"","Tanggal Selesai":w.completion_date||"",Keterangan:w.reason||"",Shift:w.shift||"",Status:w.status||""}));C.length===0&&C.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),G(C,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{G([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async E=>{let w=(await I("/api/branches?all=1")).data?.data||[],k=r=>{if(!r)return null;let s=String(r||"").toLowerCase(),m=w.find(p=>String(p.full_name||"").toLowerCase()===s||String(p.code||"").toLowerCase()===s||String(p.name||"").toLowerCase()===s);return m?m.id:null},y=E.map(r=>({branch_name:String(r.Cabang||"").trim(),backup_date:String(r["Tanggal Back Up"]||r["Tanggal Backup"]||"").trim(),original_fc_name:String(r["Nama Facility care"]||r["FC Digantikan"]||"").trim(),reliever_name:String(r.Relifer||r.Reliefer||"").trim(),period:String(r.Periode||"").trim(),reason:String(r.Keterangan||"").trim(),shift:String(r.Shift||"").trim(),completion_date:String(r["Tanggal Selesai"]||"").trim(),status:String(r.Status||"").trim()})).filter(r=>r.reliever_name&&r.backup_date),g=await I("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:y,onDuplicate:"update"})});if(!g.ok)throw new Error(g.data?.error||"Import gagal");return g.data}}})}V();X();async function ua(e){let t=await W(),i=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));Z({container:e,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>ye(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>ne(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),b=await I(`/api/reports/inspection?limit=10000&${o}`);if(b.ok){let u=b.data.data.map(d=>({Cabang:d.branch_name||"",Periode:d.period||"",Tanggal:d.inspection_date||"","Point FC":d.fc_score!==null&&d.fc_score!==void 0?d.fc_score:"","Point SPV":d.spv_score!==null&&d.spv_score!==void 0?d.spv_score:"",Status:d.status||"","Link Dokumen":d.document_link||""}));G(u,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{G([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let o=n=>{if(!n)return null;let h=String(n||"").toLowerCase(),F=t.find($=>String($.label||"").toLowerCase()===h);return F?F.value:null},b=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let h=String(n).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let $=Number(h);if($>2e4&&$<99999){let E=new Date(Date.UTC(1899,11,30)+$*864e5);return isNaN(E.getTime())?"":E.toISOString().slice(0,10)}}let F=h.split(/[\/\-\.]/);if(F.length===3){let[$,E,C]=F.map(w=>w.trim());if($.length===4&&E.length<=2&&C.length<=2)return`${$}-${E.padStart(2,"0")}-${C.padStart(2,"0")}`;if(C.length===4&&E.length<=2&&$.length<=2)return`${C}-${E.padStart(2,"0")}-${$.padStart(2,"0")}`}return h},u=a.map(n=>({branch_id:o(String(n.Cabang||"").trim()),period:String(n.Periode||"").trim(),inspection_date:b(n.Tanggal),fc_score:n["Point FC"]!==void 0&&n["Point FC"]!==""?Number(n["Point FC"]):null,spv_score:n["Point SPV"]!==void 0&&n["Point SPV"]!==""?Number(n["Point SPV"]):null,status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.inspection_date),d=await I("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:u,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}}})}V();X();async function ga(e){let t=await W(),i=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));Z({container:e,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>ye(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>ne(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),b=await I(`/api/reports/cleaning?limit=10000&${o}`);if(b.ok){let u=b.data.data.map(d=>({Cabang:d.branch_name||"",Jenis:d.activity_type||"",Periode:d.period||"",Tanggal:d.activity_date||"",Status:d.status||"","Link Dokumen":d.document_link||""}));G(u,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{G([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let o=n=>{if(!n)return null;let h=String(n||"").toLowerCase(),F=t.find($=>String($.label||"").toLowerCase()===h);return F?F.value:null},b=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let h=String(n).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let $=Number(h);if($>2e4&&$<99999){let E=new Date(Date.UTC(1899,11,30)+$*864e5);return isNaN(E.getTime())?"":E.toISOString().slice(0,10)}}let F=h.split(/[\/\-\.]/);if(F.length===3){let[$,E,C]=F.map(w=>w.trim());if($.length===4&&E.length<=2&&C.length<=2)return`${$}-${E.padStart(2,"0")}-${C.padStart(2,"0")}`;if(C.length===4&&E.length<=2&&$.length<=2)return`${C}-${E.padStart(2,"0")}-${$.padStart(2,"0")}`}return h},u=a.map(n=>({branch_id:o(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:b(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),d=await I("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:u,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}}})}V();X();async function ma(e,t){let i=await W(),a=Array.from({length:4},(u,d)=>String(new Date().getFullYear()-d)),o=t?t.get("dash_filter"):null,b={};if(o==="fogging"){let u=new Date,d=String(u.getMonth()+1).padStart(2,"0"),n=String(u.getFullYear()),h=t?t.get("month"):null;h&&h.length===7&&(n=h.split("-")[0],d=h.split("-")[1]),b={status:"Done",month:d,year:n}}Z({container:e,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:b,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:u=>`<span class="badge badge-warning">${u}</span>`},{key:"period",label:"Periode",render:u=>ye(u)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:u=>window.formatDate(u)},{key:"status",label:"Status",render:u=>ne(u)},{key:"document_link",label:"Dokumen",render:u=>u?`<a href="${u}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:u=>u||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:u=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:i,value:u?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:u?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:u?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:u?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:u?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:u?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async u=>{let d=new URLSearchParams(u||{}).toString(),n=await I(`/api/reports/fogging?limit=10000&${d}`);if(n.ok){let h=n.data.data.map(F=>({Cabang:F.branch_name||"",Jenis:F.activity_type||"Fogging",Periode:F.period||"",Tanggal:F.activity_date||"",Status:F.status||"","Link Dokumen":F.document_link||""}));G(h,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{G([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async u=>{let d=$=>{if(!$)return null;let E=String($||"").toLowerCase(),C=i.find(w=>String(w.label||"").toLowerCase()===E);return C?C.value:null},n=$=>{if($==null||$==="")return"";if($ instanceof Date&&!isNaN($.getTime()))return $.toISOString().slice(0,10);let E=String($).trim();if(E===""||E==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(E))return E.slice(0,10);if(/^\d{4,5}$/.test(E)){let w=Number(E);if(w>2e4&&w<99999){let k=new Date(Date.UTC(1899,11,30)+w*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}let C=E.split(/[\/\-\.]/);if(C.length===3){let[w,k,y]=C.map(g=>g.trim());if(w.length===4&&k.length<=2&&y.length<=2)return`${w}-${k.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&k.length<=2&&w.length<=2)return`${y}-${k.padStart(2,"0")}-${w.padStart(2,"0")}`}return E},h=u.map($=>({branch_id:d(String($.Cabang||"").trim()),activity_type:String($.Jenis||$.Kegiatan||"Fogging").trim(),period:String($.Periode||"").trim(),activity_date:n($.Tanggal),status:String($.Status||"").trim(),document_link:String($["Link Dokumen"]||"").trim(),notes:String($.Catatan||$.Keterangan||"").trim()})).filter($=>$.branch_id&&$.period&&$.activity_date),F=await I("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(h)});if(!F.ok)throw new Error(F.data?.error||"Import gagal");return F.data}}})}V();X();async function ba(e){let t=await W(),i=await ce(),a=[{value:"Berlin",label:"Berlin"},{value:"Ade",label:"Ade"}],o=Array.from({length:4},(d,n)=>String(new Date().getFullYear()-n)),b=d=>d&&!i.find(n=>n.value===d)?[...i,{value:d,label:d}]:i,u=d=>d&&!a.find(n=>n.value===d)?[...a,{value:d,label:d}]:a;Z({container:e,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:d=>window.formatDate(d)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:d=>`<span title="${d||""}">${d?.length>60?d.slice(0,60)+"\u2026":d||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:d=>window.formatDate(d)},{key:"status",label:"Status",render:d=>ne(d)},{key:"notes",label:"Keterangan",render:d=>d?.length>40?d.slice(0,40)+"\u2026":d||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade"]},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:o}],formFields:d=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:d?.branch_id},{name:"pic",label:"PIC",type:"select",options:u(d?.pic),value:d?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:d?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:d?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:d?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:d?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:d?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async d=>{let n=new URLSearchParams(d||{}).toString(),h=await I(`/api/reports/basecamp?limit=10000&${n}`);if(h.ok){let F=h.data.data.map($=>({"Tgl Info":$.info_date||"",Cabang:$.branch_name||"",Permasalahan:$.problem||"",PIC:$.pic||"","Tgl Done":$.done_date||"",Status:$.status||"",Keterangan:$.notes||""}));G(F,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{G([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async d=>{let n=E=>{if(!E)return null;let C=String(E||"").toLowerCase(),w=t.find(k=>String(k.label||"").toLowerCase()===C);return w?w.value:null},h=E=>{if(E==null||E==="")return"";if(E instanceof Date&&!isNaN(E.getTime()))return E.toISOString().slice(0,10);let C=String(E).trim();if(C===""||C==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(C))return C.slice(0,10);if(/^\d{4,5}$/.test(C)){let k=Number(C);if(k>2e4&&k<99999){let y=new Date(Date.UTC(1899,11,30)+k*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}let w=C.split(/[\/\-\.]/);if(w.length===3){let[k,y,g]=w.map(r=>r.trim());if(k.length===4&&y.length<=2&&g.length<=2)return`${k}-${y.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&y.length<=2&&k.length<=2)return`${g}-${y.padStart(2,"0")}-${k.padStart(2,"0")}`}return C},F=d.map(E=>({info_date:h(E["Tgl Info"]||E["Tanggal Info"]),branch_id:n(String(E.Cabang||"").trim()),problem:String(E.Permasalahan||"").trim(),pic:String(E.PIC||"").trim(),done_date:h(E["Tgl Done"]||E["Tanggal Done"]),status:String(E.Status||"").trim(),notes:String(E.Keterangan||E.Catatan||"").trim()})).filter(E=>E.info_date&&E.branch_id&&E.problem),$=await I("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(F)});if(!$.ok)throw new Error($.data?.error||"Import gagal");return $.data}}})}ue();var Cn=[{id:1,title:"Dusting (Pembersihan Debu)",category:"Pembersihan Rutin",icon:"\u{1F9F9}",color:"#2563EB",bg:"#EFF6FF",target:"Meja resepsionis, meja pasien, meja dokter, furniture, kaca, monitor & keyboard",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Kain Lap Microfiber Biru (Area Umum)","Kanebo (Kaca & Layar Monitor)","Kemoceng","Botol Spray Glass Cleaner","Kuas 4 Inch (Keyboard & Sela Sempit)"],steps:["Siapkan larutan glass cleaner di botol spray khusus kaca dan perabot.","Lipat kain microfiber menjadi 8 bidang kerja agar efisien dan bersih merata.","Semprotkan cairan pembersih ke permukaan kain lap, bukan langsung ke monitor atau bodi elektronik.","Usap permukaan perabot dari bagian paling atas ke bagian bawah secara searah atau pola S.","Gunakan kuas cat 4 inch untuk membersihkan sela-sela keyboard dan lubang ventilasi.","Keringkan kaca atau monitor menggunakan lap microfiber kering atau kanebo bersih."],notes:"Khusus layar monitor medis & dental unit, dilarang menyemprotkan cairan langsung ke layar."},{id:2,title:"Sweeping (Lobby Duster & Sapu Nylon)",category:"Pembersihan Rutin",icon:"\u{1F33E}",color:"#D97706",bg:"#FFFBEB",target:"Lobby utama, lorong, ruang tindakan, balkon, pantry, teras",chemical_ratio:"-",tools:["Lobby Duster Set 60cm (Area Lobby & Lorong)","Lobby Duster Set 40cm (Area Ruang Tindakan)","Sapu & Pengki Nagata 742s","Kapi Scrapper (Kerak/Permen Karet)"],steps:["Pastikan kain lobby duster dalam keadaan kering, bersih, dan bebas pasir kasar.","Dorong lobby duster menyusuri tepi dinding ruangan terlebih dahulu membentuk perimeter.","Lanjutkan mendorong dari sudut terjauh menuju arah pintu keluar tanpa mengangkat kain pel lobi dari lantai.","Kumpulkan debu dan kotoran di dekat pintu keluar, lalu angkat menggunakan sapu dan pengki Nagata.","Jika ada noda kerak atau bekas stiker/permen karet menempel di lantai, kikis perlahan menggunakan kapi scrapper.","Kibaskan kain duster di luar ruangan atau sedot sisa debu pada kain duster."],notes:"Jangan pernah menggunakan lobby duster di area lantai yang masih basah atau tergenang air."},{id:3,title:"Mopping (Pengepelan Lantai)",category:"Pembersihan Rutin",icon:"\u{1FAA3}",color:"#059669",bg:"#ECFDF5",target:"Seluruh lantai ruangan, koridor, dan area tunggu klinik",chemical_ratio:"25 ml Floor Cleaner per 1 Liter air bersih",tools:["Single Bucket Mop Wringer 20L / Double Bucket 46L","Mopping Stick / Tongkat & Kain Pel","Chemical Floor Cleaner Prokleen Lavender","Warning Sign Lantai Licin"],steps:['Pasang papan peringatan "Caution Wet Floor" di titik strategis yang mudah terlihat.',"Campurkan cairan Floor Cleaner ke dalam ember air bersih dengan takaran 25 ml per 1 Liter air.","Celupkan kain pel, masukkan ke dalam alat pemeras (wringer), dan peras hingga kondisi kain pel lembap (tidak menetes).","Mulai mengepel dari sudut terjauh ruangan mundur ke arah pintu keluar dengan gerakan membentuk pola angka delapan (8).","Bilas kain pel secara berkala apabila air pada ember bilas sudah terlihat keruh.","Biarkan lantai mengering sempurna secara alami sebelum papan peringatan wet floor dilepas."],notes:'Wajib pasang papan peringatan "Caution Wet Floor" sebelum mulai mengepel untuk mencegah risiko terpeleset.'},{id:4,title:"Membersihkan Toilet & Wastafel",category:"Sanitasi Toilet",icon:"\u{1F6BD}",color:"#DC2626",bg:"#FEF2F2",target:"Toilet pasien, toilet tim, closet, wastafel, dinding keramik, floor drain",chemical_ratio:"Karbol 25 ml:1 L Air / Bowl Cleaner 100 ml spray (kerak) / Bubuk KIFA",tools:["Sarung Tangan Karet (Wajib)","Sikat Kloset Scotch-Brite 3M","Kain Lap Microfiber Merah (Khusus Toilet)","Sponge Scotch Brite & Sikat Lantai ID-51","Floor Squeegee 75cm","Chemical Karbol Pine, Bowl Cleaner, Bubuk KIFA"],steps:["Kenakan sarung tangan karet dan siapkan perlengkapan sanitasi toilet.","Flush kloset, tuangkan Toilet Bowl Cleaner ke bagian dalam dinding kloset dan leher angsa, diamkan 5\u201310 menit.","Sikat bagian dalam kloset secara menyeluruh dengan sikat kloset, lalu bilas (flush) kembali hingga bersih.","Seka dudukan, penutup, dan tabung luar kloset menggunakan kain lap microfiber merah yang dibasahi larutan disinfektan karbol.","Bersihkan wastafel dan kran stainless menggunakan sponge tapas dengan bubuk pembersih KIFA (jangan gunakan Bowl Cleaner pada kran stainless).","Sikat lantai kamar mandi dengan sikat lantai ID-51 dan cairan karbol disinfektan.","Tarik genangan air ke arah floor drain menggunakan floor squeegee hingga lantai tidak licin dan cepat kering.","Pasang kamper/urinal block, isi ulang sabun cuci tangan, dan pasang plastik sampah baru."],notes:"Kain lap warna MERAH mutlak hanya untuk toilet dan mangkuk spittoon. Dilarang mencampur cairan kimia sembarangan."},{id:5,title:"Pembersihan Area Parkir & Luar",category:"Pembersihan Rutin",icon:"\u{1F17F}\uFE0F",color:"#4F46E5",bg:"#EEF2FF",target:"Halaman parkir mobil & motor, saluran air luar, teras depan",chemical_ratio:"-",tools:["Sapu Lidi Gagang Panjang","Pengki Nagata","Plastik Sampah Hitam Besar (90x120)","Selang Air"],steps:["Sapu daun kering, puntung rokok, plastik, dan pasir dari batas trotoar/pagar menuju satu titik.","Kumpulkan dan angkat sampah ke dalam pengki, masukkan langsung ke plastik sampah hitam besar.","Periksa dan bersihkan saluran drainase air di area parkir dari sumbatan sampah padat.","Bila ada tumpahan oli atau kotoran tanah membandel, siram dengan air dan sikat.","Rapikan pembatas parkir dan pastikan area depan klinik terlihat bersih dan rapi menyambut pasien."],notes:"Lakukan sebelum jam operasional klinik dibuka (pukul 07.30 - 08.00)."},{id:6,title:"Pembersihan Kaca & Jendela",category:"Pembersihan Rutin",icon:"\u{1FA9F}",color:"#0891B2",bg:"#ECFEFF",target:"Pintu kaca utama, jendela kaca klinik, partisi kaca, cermin toilet",chemical_ratio:"30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray",tools:["Botol Spray Glass Cleaner Prokleen","Window Wiper / Window Squeegee","Kain Microfiber Kering / Kanebo","Tongkat Teleskopik (Kaca Tinggi)"],steps:["Semprotkan larutan glass cleaner merata pada permukaan kaca yang hendak dibersihkan.","Tarik wiper karet dari arah atas ke bawah secara lurus vertikal atau zigzag horizontal tanpa terputus.","Lap bilah karet wiper setiap selesai satu tarikan menggunakan kain lap bersih agar tidak meninggalkan garis bekas air.","Gunakan kain microfiber kering untuk membersihkan sisa tetesan air di sudut-sudut lis bingkai kaca.","Untuk kaca bagian luar yang tinggi, pasang wiper pada tongkat sambungan teleskopik."],notes:"Hindari membersihkan kaca saat terpapar terik sinar matahari langsung karena cairan akan cepat mengering dan membekas noda."},{id:7,title:"Mencuci Bola Mandi Bola (Playground)",category:"Pembersihan Khusus",icon:"\u{1F939}",color:"#7C3AED",bg:"#F5F3FF",target:"Area bermain anak (Kids Playground) dan bola-bola mandi bola",chemical_ratio:"30 ml Dishwashing Liquid per 1 Liter Air + Karbol Disinfektan",tools:["Baskom Plastik Besar 55cm","Dishwashing Liquid Prokleen","Karbol Disinfektan Pine","Lap Microfiber Bersih","Karung / Wadah Pengering Jaring"],steps:["Kumpulkan seluruh bola mandi bola ke dalam wadah atau karung bersih.","Siapkan baskom plastik besar berisi air hangat kuku dicampur Dishwashing Liquid dan sedikit karbol disinfektan.","Rendam bola dan cuci secara bertahap sambil digosok perlahan menggunakan sponge/lap bersih.","Bilas bola dengan air bersih mengalir hingga tidak ada sisa busa sabun yang menempel.","Tiriskan dan angin-anginkan bola di tempat bersih beralaskan terpal/kain kering hingga benar-benar kering.","Seka dan sterilkan area kolam mandi bola sebelum bola dimasukkan kembali."],notes:"Jadwal pencucian berkala dilakukan minimal 1 kali seminggu atau sesuai kebutuhan kebersihan arena anak."},{id:8,title:"Membersihkan Filter Air Purifier",category:"Elektronik & Utility",icon:"\u{1F4A8}",color:"#0D9488",bg:"#F0FDFA",target:"Unit Air Purifier di ruang tunggu, ruang dokter, dan koridor",chemical_ratio:"-",tools:["Vacuum Cleaner (Nozzle Sikat Halus)","Kain Microfiber Kering","Kuas Cat Halus"],steps:["Matikan saklar power unit air purifier dan cabut kabel daya dari stop kontak dinding.","Buka penutup bodi (casing) depan atau belakang air purifier secara hati-hati.","Keluarkan pre-filter jaring nilon dari unit.","Bersihkan debu tebal pada pre-filter menggunakan vacuum cleaner nozzle sikat halus atau cuci dengan air mengalir (jika tipe wash-able).","Untuk filter HEPA dan Karbon Aktif: HANYA sedot debu permukaannya secara perlahan dengan vacuum, JANGAN DICUCI DENGAN AIR.","Lap bodi luar dan kisi-kisi udara masuk menggunakan lap microfiber kering.","Pastikan pre-filter benar-benar kering sebelum dipasang kembali ke posisinya semula."],notes:"Jangan pernah mencuci filter HEPA dengan air karena akan merusak serat mikro penyaring bakteri/virus."},{id:9,title:"Membersihkan Filter & Bodi AC",category:"Elektronik & Utility",icon:"\u2744\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Filter udara dan penutup indoor unit AC split di seluruh ruangan",chemical_ratio:"-",tools:["Tangga Besi Lipat","Kain Lap Microfiber Lembap","Kuas Cat Halus","Kran Air Mengalir"],steps:["Matikan AC melalui remote dan matikan MCB/saklar AC bila diperlukan.","Gunakan tangga besi lipat yang stabil di lantai datar.","Buka tutup cover depan unit indoor AC ke atas hingga terkunci.","Lepaskan kedua kisi filter saringan udara plastik secara perlahan agar debu tidak berjatuhan.","Cuci saringan filter di bawah kran air mengalir dari arah bagian dalam ke bagian luar.","Keringkan filter di tempat teduh (jangan dijemur langsung di bawah terik matahari).","Lap bodi casing dan louver/swing AC dengan kain microfiber lembap.","Pasang kembali filter yang sudah kering, tutup cover AC, dan nyalakan AC untuk pengetesan."],notes:"Pembersihan rutin mandiri dilakukan 2 minggu sekali. Deep cleaning / cuci evaporator dilakukan oleh teknisi AC vendor."},{id:10,title:"Membersihkan Dental Unit",category:"Klinis & Medis",icon:"\u{1F9B7}",color:"#E11D48",bg:"#FFF1F2",target:"Kursi dental pasien, spittoon bowl, meja instrumen, lampu dental, cover bawah",chemical_ratio:"Glass Cleaner (netral) / Karbol Disinfektan (mangkuk) / Bowl Cleaner (kerak)",tools:["Sarung Tangan Karet (Wajib)","Kain Lap Microfiber Biru (Khusus Bodi Luar)","Kain Lap Microfiber Merah (Khusus Mangkuk Spittoon)","Botol Spray Glass Cleaner (25ml + Air)","Botol Spray Karbol (20ml + Air)","Botol Spray Bowl Cleaner (100ml + Air, bila berkerak)"],steps:["Pastikan saklar utama dental unit dalam posisi OFF dan kursi dinaikkan ke posisi tertinggi untuk memudahkan akses.","Lap seluruh jok kulit, sandaran kepala, sandaran tangan, dan tiang lampu menggunakan lap microfiber biru dan cairan pembersih netral/glass cleaner.","Bersihkan mangkuk kumur (spittoon bowl) menggunakan kain lap merah dan cairan karbol disinfektan.","Jika ditemukan kerak sisa darah/plak membandel di mangkuk spittoon, gunakan sedikit bowl cleaner, sikat perlahan, lalu bilas hingga bersih.","Lap permukaan meja instrumen meja dokter. Bila ada instrumen steril, geser dengan hati-hati dan jangan mengotori instrumen.","Buka penutup cover bawah unit dental, lap debu yang terjangkau dan pastikan kondisi area selang dalam keadaan kering."],notes:"Kain lap warna MERAH hanya untuk mangkuk spittoon. Dilarang keras menggosok jok kulit dental unit dengan bahan kimia asam keras."},{id:11,title:"Membersihkan Wastafel & Keran Air",category:"Sanitasi Toilet",icon:"\u{1F6B0}",color:"#0284C7",bg:"#F0F9FF",target:"Wastafel keramik ruang tindakan, wastafel lorong, kran stainless steel, afur sink",chemical_ratio:"Bubuk Pembersih KIFA / Dishwashing Liquid Prokleen",tools:["Sponge Scotch Brite Hijau/Kuning","Kain Microfiber Hijau (Area Basah)","Bubuk Pembersih KIFA","Sikat Gigi Bekas / Sikat Kecil"],steps:["Basahi seluruh mangkuk wastafel keramik dan kran air stainless steel.","Taburkan sedikit Bubuk KIFA pada tapas sponge yang basah.","Gosok perlahan noda sabun dan kerak air putih pada kran stainless steel searah guratan logam hingga mengkilap kembali.","Gunakan sikat kecil untuk membersihkan sela leher kran dan sudut sambungan afur pembuangan wastafel.","Bilas mangkuk wastafel dan kran dengan air bersih mengalir hingga seluruh sisa residu bubuk KIFA hilang.","Keringkan kran air stainless steel dan cermin wastafel menggunakan lap microfiber hijau kering."],notes:"DILARANG KERAS menggunakan Toilet Bowl Cleaner pada kran air stainless steel karena menyebabkan karat dan warna menghitam permanen."},{id:12,title:"Membersihkan Kulkas Klinik & Pantry",category:"Elektronik & Utility",icon:"\u{1F9CA}",color:"#475569",bg:"#F8FAFC",target:"Kulkas pantry staf dan kulkas penyimpan obat/material medis (eksterior & rak)",chemical_ratio:"30 ml Dishwashing Liquid : 1 Liter Air Hangat + Lap Kering",tools:["Kain Microfiber Biru & Hijau","Dishwashing Liquid","Ember Air Hangat","Sponge Scotch Brite Halus"],steps:["Keluarkan makanan/minuman yang telah kedaluwarsa atau berbau tidak sedap.","Lepas rak-rak kaca/plastik interior kulkas yang dapat dilepas, cuci dengan sabun cuci piring dan bilas air bersih.","Seka dinding interior kulkas menggunakan kain microfiber yang dicelupkan ke dalam air hangat bersabun lembut.","Lap kering seluruh permukaan dalam kulkas dengan microfiber kering bersih.","Pasang kembali rak interior dan susun kembali bahan/makanan secara tertib dan rapi.","Bersihkan karet pintu kulkas dari noda hitam/jamur menggunakan sikat halus dan lap kering.","Lap bodi luar kulkas dan pegangan pintu menggunakan cairan pembersih netral."],notes:"Untuk kulkas penyimpanan obat/material dental, koordinasikan terlebih dahulu dengan Perawat/PIC Obat sebelum membersihkan rak interior."},{id:13,title:"Membersihkan Sofa & Kursi Ruang Tunggu",category:"Pembersihan Khusus",icon:"\u{1F6CB}\uFE0F",color:"#9333EA",bg:"#FAF5FF",target:"Sofa ruang tunggu pasien, kursi reception, kursi kerja tim",chemical_ratio:"Baking Soda (penghilang bau) / Pembersih Upholstery secukupnya",tools:["Vacuum Cleaner (Nozzle Kecil & Sikat Lembut)","Kain Microfiber Bersih","Sikat Bulu Halus","Ember Air Hangat","Cairan Pembersih Khusus Sofa (Upholstery Cleaner)","Baking Soda (Penghilang Bau Tidak Sedap)","Sarung Tangan Karet"],steps:["Kenali bahan sofa (kain fabric, beludru, kulit asli, atau kulit sintetis/oscar).","Sedot debu di seluruh permukaan, lipatan jahitan, dan celah bantal sofa menggunakan vacuum cleaner nozzle kecil.","Gunakan sikat bulu halus untuk mengangkat rambut atau kotoran yang menempel kuat di serat kain.","Jika ada noda bercak (kopi, makanan, tinta), semprotkan pembersih upholstery secukupnya, gosok memutar perlahan dengan microfiber.","Celupkan kain microfiber ke air hangat suam-suam kuku, peras hingga lembap, lalu seka seluruh permukaan sofa.","Lap kembali dengan kain kering bersih untuk menyerap sisa kelembapan.","Jika sofa berbau apek: taburkan baking soda tipis di permukaan, diamkan 20 menit, lalu sedot kembali dengan vacuum cleaner.","Biarkan sofa mengering secara alami di dalam ruangan berventilasi atau ber-AC."],notes:"Jangan menjemur sofa langsung di bawah terik matahari karena dapat menyebabkan bahan kulit pecah-pecah dan warna kain memudar."}],Fn=[{no:1,name:"Sapu Lidi Gagang",qty:"1 unit",area:"Area parkir mobil & motor klinik"},{no:2,name:"Sapu & Pengki Nagata 742s",qty:"1 set",area:"Area pantry, balkon, dak lantai atas, atau teras"},{no:3,name:"Lobby Duster Set 60 cm",qty:"1 set",area:"Lobby utama dan koridor/lorong klinik"},{no:4,name:"Lobby Duster Set 40 cm / Refill",qty:"1 unit",area:"Area ruang tindakan dokter"},{no:5,name:"Single Bucket Mop Wringer 20L",qty:"1 paket",area:"Ember pel dengan pemeras + tongkat pel"},{no:6,name:"Double Bucket Wringer Trolley 46L",qty:"1 paket",area:"Khusus klinik besar / klinik dengan fasilitas lift"},{no:7,name:"Kemoceng Bulu Lembut",qty:"1 unit",area:"Membersihkan debu kasar meja dan perabot"},{no:8,name:"Lap Microfiber Warna Merah",qty:"2 pcs",area:"Khusus area toilet, sanitasi, dan mangkuk spittoon"},{no:9,name:"Lap Microfiber Warna Biru",qty:"2 pcs",area:"Area umum kering (meja, kursi, bodi peralatan)"},{no:10,name:"Lap Microfiber Warna Hijau",qty:"2 pcs",area:"Area umum basah (pantry, wastafel luar)"},{no:11,name:"Kanebo / Synthetic Chamois",qty:"1 pcs",area:"Area umum, cermin, dan pengering kaca"},{no:12,name:"Caddy Bag Cleaning",qty:"1 unit",area:"Tas perlengkapan pinggang petugas cleaning"},{no:13,name:"Sponge Scotch Brite 3M",qty:"2 pcs",area:"Pembersihan noda lantai dan kerak ringan (spotting)"},{no:14,name:"Scotch-Brite 3M Sikat Toilet",qty:"1 unit",area:"Sikat khusus kloset dan leher angsa"},{no:15,name:"Kapi Gagang Karet / Scrapper",qty:"1 unit",area:"Pembersihan kerak lantai, noda permen karet, stiker"},{no:16,name:"Kuas Cat 4 Inch",qty:"2 pcs",area:"Pembersih sela keyboard, ventilasi, dan sudut sempit"},{no:17,name:"Warning Sign (Lantai Licin)",qty:"1 unit",area:"Digunakan saat mopping dan general cleaning"},{no:18,name:"Teleskopik + Pembersih Kaca/Langit",qty:"1 set",area:"Membersihkan jendela tinggi, plafon, dan sarang laba-laba"},{no:19,name:"Sarung Tangan Karet 3M",qty:"1 pasang",area:"Perlindungan tangan area toilet dan bahan kimia"},{no:20,name:"Botol Spray 500ml",qty:"5 unit",area:"Wadah larutan chemical (diberi label masing-masing)"},{no:21,name:"Pengharum Ruangan Stella Aerosol",qty:"1 kaleng",area:"Digunakan darurat apabila timbul bau menyengat"},{no:22,name:"Kleaner Gantungan Sapu & Pel",qty:"1 set",area:"Gantungan penyimpanan alat cleaning di janitor"},{no:23,name:"Floor Squeegee 75 cm",qty:"1 unit",area:"Pendorong genangan air lantai toilet/teras"},{no:24,name:"Sikat Botol Ikea",qty:"1 unit",area:"Di setiap ruang tindakan dokter gigi"},{no:25,name:"Scotch-Brite 3M Sikat Lantai ID-51",qty:"1 unit",area:"Sikat lantai kamar mandi dan teras"},{no:26,name:"Kamper Toilet / Urinal Block",qty:"1 pack",area:"Pengharum di toilet dan urinoir"},{no:27,name:"Plastik Sampah Hitam Sedang (60x100)",qty:"1 pack",area:"Tempat sampah indoor ruangan"},{no:28,name:"Plastik Sampah Hitam Besar (90x120)",qty:"1 pack",area:"Tempat sampah luar & penampung sampah akhir"},{no:29,name:"Plastik Sampah Medis Kuning (90x120)",qty:"1 pack",area:"Limbah medis padat B3 klinik"},{no:30,name:"Plastik Sampah Medis Kuning (40x60)",qty:"1 pack",area:"Tempat sampah medis kecil di ruang tindakan"},{no:31,name:"Tangga Besi Lipat",qty:"1 unit",area:"Alat bantu membersihkan plafon, lampu, dan AC"},{no:32,name:"Baskom Plastik Besar 55cm",qty:"1 unit",area:"Mencuci bola kids playground"}],$n=[{no:1,name:"Floor Cleaner Prokleen Lavender",size:"5 Liter",ratio:"25 ml : 1 Liter Air",function:"Pembersihan dan pengepelan rutin seluruh lantai ruangan"},{no:2,name:"Karbol Prokleen Pine Disinfektan",size:"5 Liter",ratio:"25 ml : 1 Liter Air (Mop) / 20 ml per botol spray",function:"Mengepel lantai toilet, disinfeksi area basah dan berbau tak sedap"},{no:3,name:"Glass Cleaner Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air / Siap pakai di botol spray",function:"Membersihkan pintu kaca, jendela, cermin, dan layar monitor"},{no:4,name:"Manual Dishwashing Liquid Prokleen",size:"5 Liter",ratio:"30 ml : 1 Liter Air",function:"Mencuci piring, gelas, sendok pantry, dan mencuci bola playground"},{no:5,name:"Hand Soap Prokleen Green Tea",size:"5 Liter",ratio:"Siap pakai (tanpa campur)",function:"Refill sabun pencuci tangan di wastafel pasien dan tim"},{no:6,name:"Liquid Detergent Prokleen",size:"5 Liter",ratio:"Sesuai kapasitas mesin",function:"Mencuci handuk, kain microfiber, dan seragam"},{no:7,name:"Toilet Bowl Cleaner Prokleen",size:"5 Liter",ratio:"100 ml per botol spray / Khusus kerak",function:"Pembersih kerak membandel pada kloset, wastafel keramik, dan nat keramik"},{no:8,name:"Bubuk Pembersih KIFA",size:"500 gram",ratio:"Taburkan tipis pada tapas basah",function:"Membersihkan kerak air pada kran air stainless steel, jet washer, dan sink"}],Pt="fcms_sop_fc_ik_v1",Lt="fcms_sop_fc_tools_v1",It="fcms_sop_fc_chem_v1";function ha(){try{let e=localStorage.getItem(Pt);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK",e)}return JSON.parse(JSON.stringify(Cn))}function At(e){localStorage.setItem(Pt,JSON.stringify(e))}function fa(){try{let e=localStorage.getItem(Lt);if(e)return JSON.parse(e)}catch(e){console.error("Error load tools",e)}return JSON.parse(JSON.stringify(Fn))}function Dt(e){localStorage.setItem(Lt,JSON.stringify(e))}function ya(){try{let e=localStorage.getItem(It);if(e)return JSON.parse(e)}catch(e){console.error("Error load chem",e)}return JSON.parse(JSON.stringify($n))}function Bt(e){localStorage.setItem(It,JSON.stringify(e))}async function xa(e){let t="ik",i="",a="all",o="tools",b="",u=ha(),d=fa(),n=ya();function h(){e.innerHTML=`
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
        <button class="sop-tab-btn ${t==="ik"?"active":""}" data-tab="ik" style="${F(t==="ik")}">
          \u{1F4CB} Instruksi Kerja (IK) <span class="badge" style="${$(t==="ik")}">${u.length}</span>
        </button>
        <button class="sop-tab-btn ${t==="tools"?"active":""}" data-tab="tools" style="${F(t==="tools")}">
          \u{1F6E0}\uFE0F Alat & Chemical <span class="badge" style="${$(t==="tools")}">${d.length+n.length}</span>
        </button>
        <button class="sop-tab-btn ${t==="policy"?"active":""}" data-tab="policy" style="${F(t==="policy")}">
          \u{1F4D6} Pedoman & Ketentuan
        </button>
        <button class="sop-tab-btn ${t==="links"?"active":""}" data-tab="links" style="${F(t==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sop-content-area"></div>
    `,E(),C()}function F(c){return c?"border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function $(c){return c?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function E(){e.querySelectorAll(".sop-tab-btn").forEach(B=>{B.addEventListener("click",()=>{t=B.dataset.tab,h()})});let c=e.querySelector("#btn-reset-default");c&&c.addEventListener("click",()=>{ee("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja, Alat, dan Chemical ke data standar dokumen asli?",()=>{localStorage.removeItem(Pt),localStorage.removeItem(Lt),localStorage.removeItem(It),u=ha(),d=fa(),n=ya(),j("Data SOP Facility Care berhasil di-reset ke standar asli"),h()},"Konfirmasi Reset Standar")});let v=e.querySelector("#btn-print-sop");v&&v.addEventListener("click",()=>{window.print()});let S=e.querySelector("#btn-export-tools");S&&S.addEventListener("click",()=>{f()})}function C(){let c=e.querySelector("#sop-content-area");c&&(t==="ik"?w(c):t==="tools"?g(c):t==="policy"?m(c):t==="links"&&p(c))}function w(c){c.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="ik-search" class="form-control" placeholder="Cari pekerjaan, alat, atau chemical..." value="${i}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="cat-pills-wrap">
              <button class="cat-pill ${a==="all"?"active":""}" data-cat="all">Semua (${u.length})</button>
              <button class="cat-pill ${a==="Pembersihan Rutin"?"active":""}" data-cat="Pembersihan Rutin">Rutin (${u.filter(P=>P.category==="Pembersihan Rutin").length})</button>
              <button class="cat-pill ${a==="Sanitasi Toilet"?"active":""}" data-cat="Sanitasi Toilet">Toilet & Wastafel (${u.filter(P=>P.category==="Sanitasi Toilet").length})</button>
              <button class="cat-pill ${a==="Klinis & Medis"?"active":""}" data-cat="Klinis & Medis">Klinis Dental (${u.filter(P=>P.category==="Klinis & Medis").length})</button>
              <button class="cat-pill ${a==="Elektronik & Utility"?"active":""}" data-cat="Elektronik & Utility">Elektronik & AC (${u.filter(P=>P.category==="Elektronik & Utility").length})</button>
              <button class="cat-pill ${a==="Pembersihan Khusus"?"active":""}" data-cat="Pembersihan Khusus">Khusus (${u.filter(P=>P.category==="Pembersihan Khusus").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${u.length} dari ${u.length} Instruksi Kerja
            </div>
            <button id="btn-add-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
              \u2795 Tambah IK Baru
            </button>
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="ik-cards-grid-container"></div>
    `;let v=document.createElement("style");v.innerHTML=`
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
    `,c.appendChild(v);let S=c.querySelector("#ik-cards-grid-container"),B=c.querySelector("#ik-count-label");function T(){let P=u.filter(M=>{let z=a==="all"||M.category===a,x=i.toLowerCase().trim(),A=!x||M.title.toLowerCase().includes(x)||M.target.toLowerCase().includes(x)||M.tools.some(L=>L.toLowerCase().includes(x))||M.steps.some(L=>L.toLowerCase().includes(x));return z&&A});if(B&&(B.textContent=`Menampilkan ${P.length} dari ${u.length} Instruksi Kerja`),P.length===0){S.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}S.innerHTML=`
        <div class="sop-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${P.map(M=>`
            <div class="card ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${M.bg||"#EFF6FF"}; border-radius:10px;">${M.icon||"\u{1F4CB}"}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        ${M.id}. ${M.title}
                      </h3>
                      <span class="badge" style="background:${M.bg||"#EFF6FF"}; color:${M.color||"#2563EB"}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${M.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    <button class="btn btn-sm btn-ghost btn-edit-ik" data-id="${M.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    <button class="btn btn-sm btn-ghost btn-delete-ik" data-id="${M.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${M.color||"#2563EB"};">
                  <strong>Area:</strong> ${M.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Bahan Utama:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${M.tools.slice(0,3).map(z=>`<li style="margin-bottom:2px;">${z}</li>`).join("")}
                    ${M.tools.length>3?`<li style="color:var(--text-3); font-style:italic;">+ ${M.tools.length-3} alat lainnya...</li>`:""}
                  </ul>
                </div>

                ${M.chemical_ratio&&M.chemical_ratio!=="-"?`
                  <div style="font-size:0.8rem; background:#FEF3C7; color:#92400E; padding:6px 10px; border-radius:6px; margin-bottom:14px; font-weight:600;">
                    \u{1F9EA} Takaran: ${M.chemical_ratio}
                  </div>
                `:""}
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${M.steps.length} Langkah Kerja
                </span>
                <div style="display:flex; gap:6px;">
                  <button class="btn btn-sm btn-primary btn-view-ik" data-id="${M.id}" style="padding:6px 14px; font-weight:600; border-radius:6px;">
                    Lihat Panduan \u2794
                  </button>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,S.querySelectorAll(".btn-view-ik").forEach(M=>{M.addEventListener("click",()=>{let z=parseInt(M.dataset.id,10),x=u.find(A=>A.id===z);x&&k(x)})}),S.querySelectorAll(".btn-edit-ik").forEach(M=>{M.addEventListener("click",()=>{let z=parseInt(M.dataset.id,10),x=u.find(A=>A.id===z);x&&y(x)})}),S.querySelectorAll(".btn-delete-ik").forEach(M=>{M.addEventListener("click",()=>{let z=parseInt(M.dataset.id,10),x=u.find(A=>A.id===z);x&&ee(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${x.title}"?`,()=>{u=u.filter(A=>A.id!==z),At(u),j("Instruksi Kerja berhasil dihapus"),h()},"Hapus Instruksi Kerja")})})}T();let D=c.querySelector("#ik-search");D&&D.addEventListener("input",P=>{i=P.target.value,T()}),c.querySelectorAll(".cat-pill").forEach(P=>{P.addEventListener("click",()=>{c.querySelectorAll(".cat-pill").forEach(M=>M.classList.remove("active")),P.classList.add("active"),a=P.dataset.cat,T()})});let _=c.querySelector("#btn-add-ik");_&&_.addEventListener("click",()=>{y()})}function k(c){Y({title:`${c.icon||"\u{1F4CB}"} IK-${c.id}: ${c.title}`,size:"lg",content:`
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
                ${c.tools.map(v=>`<li>${v}</li>`).join("")}
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
                ${c.steps.map(v=>`<li style="margin-bottom:6px;">${v}</li>`).join("")}
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
      `,confirmText:"Tutup Panduan",onConfirm:(v,S)=>{S()}})}function y(c=null){let v=!!c,S=c?c.tools.join(`
`):"",B=c?c.steps.join(`
`):"",T=`
      <form id="ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-ik-title" class="form-control" required value="${c?l(c.title):""}" placeholder="Contoh: Membersihkan Dental Unit" style="width:100%; border-radius:6px; height:36px;">
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
            <input type="text" id="form-ik-ratio" class="form-control" value="${c?l(c.chemical_ratio||""):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-ik-target" class="form-control" required value="${c?l(c.target):""}" placeholder="Contoh: Kursi pasien, meja instrumen, mangkuk spittoon" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Bahan yang Digunakan (1 item per baris) *</label>
          <textarea id="form-ik-tools" class="form-control" rows="4" placeholder="Kain Microfiber Biru&#10;Sarung Tangan Karet&#10;Botol Spray Glass Cleaner" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${l(S)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah-Langkah Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-ik-steps" class="form-control" rows="6" placeholder="1. Pastikan posisi saklar OFF.&#10;2. Lap jok dental unit dengan kain microfiber biru.&#10;3. Sikat mangkuk spittoon..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${l(B)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Penting / Hal yang Perlu Diperhatikan (Safety)</label>
          <textarea id="form-ik-notes" class="form-control" rows="2" placeholder="Catatan larangan atau peringatan keselamatan..." style="width:100%; border-radius:6px; font-size:0.85rem;">${c?l(c.notes||""):""}</textarea>
        </div>
      </form>
    `;Y({title:v?`\u270F\uFE0F Edit IK-${c.id}: ${c.title}`:"\u2795 Tambah Instruksi Kerja Baru",content:T,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(D,_)=>{let P=D.querySelector("#form-ik-title").value.trim(),M=D.querySelector("#form-ik-icon").value.trim()||"\u{1F4CB}",z=D.querySelector("#form-ik-cat").value,x=D.querySelector("#form-ik-ratio").value.trim(),A=D.querySelector("#form-ik-target").value.trim(),L=D.querySelector("#form-ik-tools").value.trim(),q=D.querySelector("#form-ik-steps").value.trim(),O=D.querySelector("#form-ik-notes").value.trim();if(!P||!A||!L||!q){N("Harap lengkapi seluruh field wajib (*)");return}let H=L.split(`
`).map(de=>de.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),K=q.split(`
`).map(de=>de.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),ie={"Pembersihan Rutin":{color:"#2563EB",bg:"#EFF6FF"},"Sanitasi Toilet":{color:"#DC2626",bg:"#FEF2F2"},"Klinis & Medis":{color:"#E11D48",bg:"#FFF1F2"},"Elektronik & Utility":{color:"#0D9488",bg:"#F0FDFA"},"Pembersihan Khusus":{color:"#7C3AED",bg:"#F5F3FF"}}[z]||{color:"#2563EB",bg:"#EFF6FF"};if(v)c.title=P,c.icon=M,c.category=z,c.chemical_ratio=x,c.target=A,c.tools=H,c.steps=K,c.notes=O,c.color=ie.color,c.bg=ie.bg,At(u),j(`Instruksi Kerja "${P}" berhasil diperbarui`);else{let de=u.length>0?Math.max(...u.map(U=>U.id))+1:1,me={id:de,title:P,category:z,icon:M,color:ie.color,bg:ie.bg,target:A,chemical_ratio:x||"-",tools:H,steps:K,notes:O};u.push(me),At(u),j(`Instruksi Kerja baru IK-${de} berhasil ditambahkan`)}_(),h()}})}function g(c){c.innerHTML=`
      <!-- Sub-tab Switcher & Actions -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-tools-btn" class="btn btn-sm ${o==="tools"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9F9} Alat Kebersihan (${d.length} Item)
          </button>
          <button id="view-chem-btn" class="btn btn-sm ${o==="chemical"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F9EA} Bahan Chemical (${n.length} Item)
          </button>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="tool-search" class="form-control" placeholder="Cari nama / peruntukan / rasio..." value="${b}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
          <button id="btn-add-tool-chem" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            \u2795 ${o==="tools"?"Tambah Alat Baru":"Tambah Chemical Baru"}
          </button>
        </div>
      </div>

      <!-- Table Container -->
      <div id="tool-table-container"></div>
    `;let v=c.querySelector("#tool-table-container"),S=c.querySelector("#tool-search");function B(){let T=b.toLowerCase().trim();if(o==="tools"){let D=d.filter(_=>!T||_.name.toLowerCase().includes(T)||_.area.toLowerCase().includes(T));v.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Standar Peralatan Kebersihan (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-primary">${D.length} Item</span>
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
                  ${D.map((_,P)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${P%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${_.no||P+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${l(_.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#E0F2FE; color:#0369A1; font-weight:600;">${l(_.qty)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${l(_.area)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                        <button class="btn btn-sm btn-ghost btn-edit-tool" data-no="${_.no}" title="Edit Alat" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                        <button class="btn btn-sm btn-ghost btn-delete-tool" data-no="${_.no}" title="Hapus Alat" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,v.querySelectorAll(".btn-edit-tool").forEach(_=>{_.addEventListener("click",()=>{let P=parseInt(_.dataset.no,10),M=d.find(z=>z.no===P);M&&r(M)})}),v.querySelectorAll(".btn-delete-tool").forEach(_=>{_.addEventListener("click",()=>{let P=parseInt(_.dataset.no,10),M=d.find(z=>z.no===P);M&&ee(`Apakah Anda yakin ingin menghapus alat: "${M.name}"?`,()=>{d=d.filter(z=>z.no!==P),d.forEach((z,x)=>{z.no=x+1}),Dt(d),j("Alat kebersihan berhasil dihapus"),h()},"Hapus Alat Kebersihan")})})}else{let D=n.filter(_=>!T||_.name.toLowerCase().includes(T)||_.function.toLowerCase().includes(T));v.innerHTML=`
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Bahan Kimia & Standar Takaran (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-success">${D.length} Chemical</span>
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
                  ${D.map((_,P)=>`
                    <tr style="border-bottom:1px solid #F1F5F9; background:${P%2===0?"#fff":"#FAFAFA"};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${_.no||P+1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${l(_.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#FEF3C7; color:#92400E; font-weight:600;">${l(_.size)}</span></td>
                      <td style="padding:10px 14px;"><span style="background:#EFF6FF; color:#1E40AF; padding:3px 8px; border-radius:6px; font-weight:600; font-size:0.82rem;">${l(_.ratio)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${l(_.function)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                        <button class="btn btn-sm btn-ghost btn-edit-chem" data-no="${_.no}" title="Edit Chemical" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                        <button class="btn btn-sm btn-ghost btn-delete-chem" data-no="${_.no}" title="Hapus Chemical" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                      </td>
                    </tr>
                  `).join("")}
                </tbody>
              </table>
            </div>
          </div>
        `,v.querySelectorAll(".btn-edit-chem").forEach(_=>{_.addEventListener("click",()=>{let P=parseInt(_.dataset.no,10),M=n.find(z=>z.no===P);M&&s(M)})}),v.querySelectorAll(".btn-delete-chem").forEach(_=>{_.addEventListener("click",()=>{let P=parseInt(_.dataset.no,10),M=n.find(z=>z.no===P);M&&ee(`Apakah Anda yakin ingin menghapus chemical: "${M.name}"?`,()=>{n=n.filter(z=>z.no!==P),n.forEach((z,x)=>{z.no=x+1}),Bt(n),j("Bahan chemical berhasil dihapus"),h()},"Hapus Bahan Chemical")})})}}B(),c.querySelector("#view-tools-btn").addEventListener("click",()=>{o="tools",g(c)}),c.querySelector("#view-chem-btn").addEventListener("click",()=>{o="chemical",g(c)}),S&&S.addEventListener("input",T=>{b=T.target.value,B()}),c.querySelector("#btn-add-tool-chem").addEventListener("click",()=>{o==="tools"?r():s()})}function r(c=null){let v=!!c,S=`
      <form id="tool-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Alat / Item *</label>
          <input type="text" id="form-tool-name" class="form-control" required value="${c?l(c.name):""}" placeholder="Contoh: Lobby Duster Set 60 cm" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kuantitas Standar *</label>
          <input type="text" id="form-tool-qty" class="form-control" required value="${c?l(c.qty):"1 unit"}" placeholder="Contoh: 1 set, 2 pcs, 1 paket" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Peruntukan Area & Kegunaan *</label>
          <textarea id="form-tool-area" class="form-control" rows="3" required placeholder="Contoh: Lobby utama dan koridor/lorong klinik" style="width:100%; border-radius:6px; font-size:0.85rem;">${c?l(c.area):""}</textarea>
        </div>
      </form>
    `;Y({title:v?`\u270F\uFE0F Edit Alat: ${c.name}`:"\u2795 Tambah Alat Kebersihan Baru",content:S,confirmText:"Simpan",onConfirm:(B,T)=>{let D=B.querySelector("#form-tool-name").value.trim(),_=B.querySelector("#form-tool-qty").value.trim(),P=B.querySelector("#form-tool-area").value.trim();if(!D||!_||!P){N("Harap lengkapi semua kolom");return}if(v)c.name=D,c.qty=_,c.area=P,Dt(d),j(`Alat "${D}" berhasil diperbarui`);else{let M=d.length>0?Math.max(...d.map(z=>z.no))+1:1;d.push({no:M,name:D,qty:_,area:P}),Dt(d),j(`Alat baru "${D}" berhasil ditambahkan`)}T(),h()}})}function s(c=null){let v=!!c,S=`
      <form id="chem-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Produk Chemical *</label>
          <input type="text" id="form-chem-name" class="form-control" required value="${c?l(c.name):""}" placeholder="Contoh: Floor Cleaner Prokleen Lavender" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kemasan Standar *</label>
            <input type="text" id="form-chem-size" class="form-control" required value="${c?l(c.size):"5 Liter"}" placeholder="Contoh: 5 Liter, 500 gram" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Rasio & Takaran Standar *</label>
            <input type="text" id="form-chem-ratio" class="form-control" required value="${c?l(c.ratio):""}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Fungsi & Cara Penggunaan *</label>
          <textarea id="form-chem-function" class="form-control" rows="3" required placeholder="Contoh: Pembersihan dan pengepelan rutin seluruh lantai ruangan" style="width:100%; border-radius:6px; font-size:0.85rem;">${c?l(c.function):""}</textarea>
        </div>
      </form>
    `;Y({title:v?`\u270F\uFE0F Edit Chemical: ${c.name}`:"\u2795 Tambah Chemical Baru",content:S,confirmText:"Simpan",onConfirm:(B,T)=>{let D=B.querySelector("#form-chem-name").value.trim(),_=B.querySelector("#form-chem-size").value.trim(),P=B.querySelector("#form-chem-ratio").value.trim(),M=B.querySelector("#form-chem-function").value.trim();if(!D||!_||!P||!M){N("Harap lengkapi semua kolom");return}if(v)c.name=D,c.size=_,c.ratio=P,c.function=M,Bt(n),j(`Chemical "${D}" berhasil diperbarui`);else{let z=n.length>0?Math.max(...n.map(x=>x.no))+1:1;n.push({no:z,name:D,size:_,ratio:P,function:M}),Bt(n),j(`Chemical baru "${D}" berhasil ditambahkan`)}T(),h()}})}function m(c){c.innerHTML=`
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
    `}function p(c){c.innerHTML=`
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
    `}async function f(){try{let{downloadExcel:c}=await Promise.resolve().then(()=>(X(),Se)),v=u.map(B=>({"No IK":`IK-${B.id}`,"Judul Pekerjaan":B.title,Kategori:B.category,"Sasaran Target Area":B.target,"Rasio Chemical":B.chemical_ratio||"-","Alat & Bahan":B.tools.join(", "),"Langkah Kerja":B.steps.map((T,D)=>`${D+1}. ${T}`).join(" | "),"Catatan Safety":B.notes||"-"})),S=new Date().toISOString().slice(0,10);c(v,`SOP_Facility_Care_Instruksi_Kerja_${S}`),j("Data Excel Instruksi Kerja berhasil diunduh")}catch(c){console.error(c),N("Gagal mendownload data Excel: "+c.message)}}function l(c){return c?String(c).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}h()}ue();var _n=[{id:1,title:"Pengamanan Area & Aset Perusahaan",category:"Pengamanan Fisik & Patroli",icon:"\u{1F6E1}\uFE0F",color:"#2563EB",bg:"#EFF6FF",target:"Seluruh area dalam & luar Head Quarter dan Klinik Cabang Audy Dental",tools:["Google Form Checklist Monitoring Klinik: https://forms.gle/1EAvkqDJVdctXEGn7","Senter Patroli","Komputer dengan Akses Monitor CCTV (Live View)","Alat Komunikasi / Handphone / HT"],steps:["Lakukan Safety Patrol secara rutin mengelilingi seluruh area sesuai jadwal yang telah ditentukan (Patrol 1, 2, dan 3).","Catat hasil temuan dan kondisi keamanan secara berkala ke dalam Google Form Safety Patrol resmi.","Awasi kendaraan yang terparkir di area luar klinik; pastikan khusus sepeda motor dalam keadaan terkunci ganda.","Pantau situasi lingkungan dan area sensitif secara berkala melalui layar monitor CCTV saat berada di pos security.","Segera lakukan tindakan persuasif ramah apabila ditemukan potensi gangguan keamanan, dan segera hubungi pihak Kepolisian terdekat bila situasi darurat tidak dapat diatasi.","Bila ditemukan permasalahan fisik, kerusakan fasilitas, atau kendala di area klinik, segera dokumentasikan (foto) dan laporkan langsung ke SPV Klinik dan Leader Facility Care.","Disaat malam hari atau saat klinik telah tutup operasional, pastikan seluruh AC dan lampu ruangan dalam keadaan mati, kecuali lampu keamanan tertentu yang memang harus dibiarkan menyala terang."],notes:"Prioritaskan keselamatan diri dan tindakan persuasif. Dilarang meninggalkan area pengawasan tanpa personil pengganti yang jelas."},{id:2,title:"Penertiban & Pengaturan Parkir Kendaraan",category:"Operasional Parkir & Luar",icon:"\u{1F17F}\uFE0F",color:"#D97706",bg:"#FFFBEB",target:"Area parkir mobil, motor, akses keluar masuk, dan teras depan klinik",tools:["Lampu Lalu Lintas (Lampu Lalin)","Cone / Kerucut Pembatas Parkir","Peluit Parkir"],steps:["Berdiri di posisi strategis dan bantu mengarahkan setiap kendaraan pasien atau tamu yang datang secara tepat dan rapi ke slot parkir yang tersedia.","Berikan informasi dan solusi dengan sikap ramah serta santun apabila seluruh slot parkir klinik dalam kondisi penuh.","Dalam kondisi tertentu (bila ada mobil yang terhalang mobil lain dan harus dipindahkan, sementara pasien sedang dalam tindakan perawatan): bantu memindahkan kendaraan HANYA JIKA petugas security memiliki kemampuan mengemudi dan memiliki SIM aktif.","Apabila petugas security tidak memiliki kemampuan mengemudikan mobil: sampaikan permohonan maaf dengan tulus kepada pasien/tamu, minta untuk menunggu sejenak, dan segera minta bantuan kepada staf Resepsionis untuk menghubungi pemilik mobil yang menghalangi.","Bantu pasien atau tamu yang akan keluar parkir dengan memberikan aba-aba dan arahan lalu lintas jalan raya yang aman dan tepat."],notes:"Dilarang memaksakan memindahkan mobil pasien jika tidak memiliki lisensi mengemudi demi menghindari risiko gesekan atau kerusakan kendaraan."},{id:3,title:"Penerimaan Tamu & Pasien Klinik",category:"Layanan & Hospitalitas",icon:"\u{1F91D}",color:"#059669",bg:"#ECFDF5",target:"Pintu masuk utama klinik, lobby, area resepsionis, dan pos penerimaan tamu",tools:["Alat Tulis Kantor (Pulpen)","Buku Tamu / Log Book Keamanan","Alat Komunikasi / Handphone"],steps:['Berdiri dengan tegap dan sigap, bukakan pintu untuk setiap pasien/tamu yang datang dengan menerapkan budaya 5S (Senyum, Salam, Sapa, Sopan, Santun) dan menyapa: "Selamat pagi / siang / sore, ada yang bisa kami bantu?".',"Setelah pasien atau tamu memberitahukan tujuannya, persilakan dan arahkan pasien langsung menuju meja Resepsionis.","Apabila yang datang adalah tamu instansi, vendor, atau pihak Manajemen/BOD: persilakan mengisi Buku Tamu (Log Book) terlebih dahulu dan segera hubungi staf/karyawan yang ingin ditemui.",'Apabila pasien langsung melangkah ke arah tujuannya, berikan gestur tubuh hormat menyilangkan kedua tangan di depan, dan jika pasien/tamu mengucapkan terima kasih, jawab dengan ramah: "Dengan Senang Hati".',"Dalam kondisi khusus (pasien lansia, ibu hamil, atau pengguna kursi roda): petugas security wajib mendampingi, mengawal, dan membantu mendorong kursi roda sampai ke depan meja resepsionis.",'Buka kembali pintu untuk setiap pasien yang telah selesai melakukan perawatan saat hendak keluar klinik sambil mengucapkan ucapan hangat: "Terima Kasih, semoga lekas sehat kembali".'],notes:"Jaga senyum tulus dan keramahan sebagai garda terdepan representasi citra Audy Dental."}],Tn=[{no:1,type:"Aset Tetap",item:"Rumah & Bangunan",description:"Gedung klinik, pos satpam, pagar, kanopi, teras beton",pic:"Building Maintenance & Security"},{no:2,type:"Aset Tetap",item:"Genset & Lift",description:"Mesin generator listrik cadangan dan elevator penumpang tetap",pic:"Teknisi & Security"},{no:3,type:"Aset Tetap",item:"AC Central & Split Tetap",description:"Outdoor dan indoor AC yang terpasang permanen di dinding/plafon",pic:"Teknisi AC & Security"},{no:4,type:"Aset Tetap",item:"Dental Unit Permanen",description:"Kursi dental unit terpasang instalasi air dan kompresor",pic:"Tim Klinik & Security"},{no:5,type:"Aset Tetap",item:"Instalasi Listrik, Pipa & Jaringan IT Permanen",description:"Panel MCB, kabel fiber optik, pipa drainase, server rack",pic:"IT & Teknisi"},{no:6,type:"Aset Tetap",item:"Furniture Melekat (Built-in)",description:"Meja resepsionis cor/melekat, partisi kaca mati, kitchen set pantry",pic:"General Support"},{no:7,type:"Aset Bergerak",item:"Komputer, Laptop & Printer",description:"PC kasir/resepsionis, laptop staf, printer rontgen/dokumen",pic:"Staf Terkait & IT"},{no:8,type:"Aset Bergerak",item:"Kendaraan Operasional",description:"Mobil operasional dan sepeda motor inventaris perusahaan",pic:"Security & Driver"},{no:9,type:"Aset Bergerak",item:"Furniture Lepasan",description:"Meja kerja, kursi resepsionis, kursi tunggu pasien, lemari file",pic:"Facility Care & Staf"},{no:10,type:"Aset Bergerak",item:"Hand Tools, Alat Kerja & Alat Dental Portabel",description:"Handpiece dental, alat rontgen portable, toolbox teknisi",pic:"Dokter/Perawat & Teknisi"},{no:11,type:"Low Value Asset",item:"Peralatan Kantor (Telepon, ATK)",description:"Pesawat telepon, kalkulator, stapler, perlengkapan meja",pic:"Resepsionis"},{no:12,type:"Low Value Asset",item:"Alat-alat Kebersihan",description:"Lobby duster, ember wringer pel, sapu, tangga, botol spray",pic:"Facility Care"},{no:13,type:"Low Value Asset",item:"Perlengkapan Security",description:"Lampu lalin, cone parkir, peluit, senter patroli, jas hujan",pic:"Security"}],va={cabang:[{zone:"Zona Hijau",level:"Akses Publik Bebas",color:"#059669",bg:"#ECFDF5",rooms:["Area Parkir Kendaraan","Lobby Lantai 1","Teras Luar Klinik"],desc:"Area dengan pengamanan umum yang bebas diakses pasien dan tamu setelah pengarahan security."},{zone:"Zona Kuning",level:"Akses Terbatas / Pendampingan",color:"#D97706",bg:"#FFFBEB",rooms:["Koridor Ruang Tindakan","Tangga & Lift Klinik","Ruang Tunggu Lantai 2 & 3","Kids Playground","Toilet Lantai 1 & 2","Mushola Staf & Pasien"],desc:"Area terbatas, tamu wajib didampingi petugas atau staf klinik terkait."},{zone:"Zona Merah",level:"Area Terlarang / Restricted",color:"#DC2626",bg:"#FEF2F2",rooms:["Area Meja Receptionist / Kasir","Seluruh Ruang Dental Unit (DU 101, 102, 201, 202, 301, 302)","Ruang Sterilisasi Alat","Ruang Rontgen Radiologi","Ruang Suction & Kompresor","Ruang Tim Staf & Gudang Material Medis"],desc:"Area sangat sensitif, hanya staf medis bertugas dan atasan pemilik area. Tamu wajib izin Manajemen/BOD dan didampingi penuh."}],hq:[{floor:"Lantai 1",hijau:["Area Parkir Mobil/Motor","Lobby Resepsionis Lantai 1"],kuning:["Seluruh Ruang Meeting Lantai 1","Ruang Training Lantai 1","Koridor Lift & Toilet"],merah:["Ruang Customer Services","Gudang Utama Lantai 1","Gudang Facility Care","Gudang Teknisi & Ruang Teknisi","Pos Satpam & Ruang Genset"]},{floor:"Lantai 2",hijau:["-"],kuning:["Ruang Meeting Lantai 2","Koridor Lift & Toilet","Pantry"],merah:["Ruang CEO","Ruang Sekretaris CEO","Ruang Human Capital","Ruang Dentist Relation","Ruang General Support","Ruang Eks. Patrel"]},{floor:"Lantai 3",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry","Mushola & Tempat Wudhu","Ruang Santai Karyawan"],merah:["Ruang Marketing","Studio Konten Marketing","Gudang Marketing","Ruang Patient Relation","Ruang OPS","Ruang Training Lantai 3","Gudang Lantai 3"]},{floor:"Lantai 4",hijau:["-"],kuning:["Koridor Lift & Toilet","Pantry"],merah:["Ruang Finance & Accounting","Ruang VP Finance & Accounting","Ruang Audit Internal","Ruang Training Lantai 4","Gudang Lantai 4"]},{floor:"Lantai 5",hijau:["-"],kuning:["Lobby Lantai 5","Pantry Lantai 5","Toilet Lantai 5","Ruang Training Besar Lantai 5"],merah:["Area Dak Toren Air & Rooftop"]}]},Nt="fcms_sop_sec_ik_v1",jt="fcms_sop_sec_assets_v1";function ka(){try{let e=localStorage.getItem(Nt);if(e)return JSON.parse(e)}catch(e){console.error("Error load IK Sec",e)}return JSON.parse(JSON.stringify(_n))}function Mt(e){localStorage.setItem(Nt,JSON.stringify(e))}function wa(){try{let e=localStorage.getItem(jt);if(e)return JSON.parse(e)}catch(e){console.error("Error load assets sec",e)}return JSON.parse(JSON.stringify(Tn))}function zt(e){localStorage.setItem(jt,JSON.stringify(e))}async function Sa(e){let t="ik",i="",a="all",o="",b="cabang",u=ka(),d=wa();function n(){e.innerHTML=`
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
        <button class="sec-tab-btn ${t==="ik"?"active":""}" data-tab="ik" style="${h(t==="ik")}">
          \u{1F6E1}\uFE0F Instruksi Kerja (IK) <span class="badge" style="${F(t==="ik")}">${u.length}</span>
        </button>
        <button class="sec-tab-btn ${t==="zonasi"?"active":""}" data-tab="zonasi" style="${h(t==="zonasi")}">
          \u{1F5FA}\uFE0F Zonasi Akses & Denah
        </button>
        <button class="sec-tab-btn ${t==="assets"?"active":""}" data-tab="assets" style="${h(t==="assets")}">
          \u{1F4E6} Klasifikasi Aset <span class="badge" style="${F(t==="assets")}">${d.length}</span>
        </button>
        <button class="sec-tab-btn ${t==="policy"?"active":""}" data-tab="policy" style="${h(t==="policy")}">
          \u{1F4CB} Shift, Patrol & Insiden
        </button>
        <button class="sec-tab-btn ${t==="links"?"active":""}" data-tab="links" style="${h(t==="links")}">
          \u{1F517} Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sec-content-area"></div>
    `,$(),E()}function h(l){return l?"border:none; background:#1E40AF; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;":"border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;"}function F(l){return l?"background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;":"background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;"}function $(){e.querySelectorAll(".sec-tab-btn").forEach(S=>{S.addEventListener("click",()=>{t=S.dataset.tab,n()})});let l=e.querySelector("#btn-reset-sec-default");l&&l.addEventListener("click",()=>{ee("Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja dan Klasifikasi Aset Security ke standar dokumen asli?",()=>{localStorage.removeItem(Nt),localStorage.removeItem(jt),u=ka(),d=wa(),j("Data SOP Security berhasil di-reset ke standar dokumen"),n()},"Reset ke Standar Dokumen")});let c=e.querySelector("#btn-print-sec");c&&c.addEventListener("click",()=>{window.print()});let v=e.querySelector("#btn-export-sec");v&&v.addEventListener("click",()=>{p()})}function E(){let l=e.querySelector("#sec-content-area");l&&(t==="ik"?C(l):t==="zonasi"?y(l):t==="assets"?g(l):t==="policy"?s(l):t==="links"&&m(l))}function C(l){l.innerHTML=`
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="sec-ik-search" class="form-control" placeholder="Cari instruksi kerja security..." value="${i}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">\u{1F50D}</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="sec-cat-pills-wrap">
              <button class="sec-cat-pill ${a==="all"?"active":""}" data-cat="all">Semua (${u.length})</button>
              <button class="sec-cat-pill ${a==="Pengamanan Fisik & Patroli"?"active":""}" data-cat="Pengamanan Fisik & Patroli">Patroli & Fisik (${u.filter(_=>_.category==="Pengamanan Fisik & Patroli").length})</button>
              <button class="sec-cat-pill ${a==="Operasional Parkir & Luar"?"active":""}" data-cat="Operasional Parkir & Luar">Parkir & Luar (${u.filter(_=>_.category==="Operasional Parkir & Luar").length})</button>
              <button class="sec-cat-pill ${a==="Layanan & Hospitalitas"?"active":""}" data-cat="Layanan & Hospitalitas">Tamu & Pasien (${u.filter(_=>_.category==="Layanan & Hospitalitas").length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="sec-ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${u.length} Instruksi Kerja
            </div>
            <button id="btn-add-sec-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
              \u2795 Tambah IK Baru
            </button>
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
    `,l.appendChild(c);let v=l.querySelector("#sec-ik-cards-grid-container"),S=l.querySelector("#sec-ik-count-label");function B(){let _=u.filter(P=>{let M=a==="all"||P.category===a,z=i.toLowerCase().trim(),x=!z||P.title.toLowerCase().includes(z)||P.target.toLowerCase().includes(z)||P.tools.some(A=>A.toLowerCase().includes(z))||P.steps.some(A=>A.toLowerCase().includes(z));return M&&x});if(S&&(S.textContent=`Menampilkan ${_.length} dari ${u.length} Instruksi Kerja`),_.length===0){v.innerHTML=`
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">\u{1F50D}</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;return}v.innerHTML=`
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
                    <button class="btn btn-sm btn-ghost btn-edit-sec-ik" data-id="${P.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      \u270F\uFE0F
                    </button>
                    <button class="btn btn-sm btn-ghost btn-delete-sec-ik" data-id="${P.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      \u{1F5D1}\uFE0F
                    </button>
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${P.color||"#1E40AF"};">
                  <strong>Area Sasaran:</strong> ${P.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">\u{1F9F0} Alat & Perlengkapan:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${P.tools.slice(0,3).map(M=>`<li style="margin-bottom:2px;">${f(M)}</li>`).join("")}
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
      `,v.querySelectorAll(".btn-view-sec-ik").forEach(P=>{P.addEventListener("click",()=>{let M=parseInt(P.dataset.id,10),z=u.find(x=>x.id===M);z&&w(z)})}),v.querySelectorAll(".btn-edit-sec-ik").forEach(P=>{P.addEventListener("click",()=>{let M=parseInt(P.dataset.id,10),z=u.find(x=>x.id===M);z&&k(z)})}),v.querySelectorAll(".btn-delete-sec-ik").forEach(P=>{P.addEventListener("click",()=>{let M=parseInt(P.dataset.id,10),z=u.find(x=>x.id===M);z&&ee(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${z.title}"?`,()=>{u=u.filter(x=>x.id!==M),Mt(u),j("Instruksi Kerja Security berhasil dihapus"),n()},"Hapus Instruksi Kerja")})})}B();let T=l.querySelector("#sec-ik-search");T&&T.addEventListener("input",_=>{i=_.target.value,B()}),l.querySelectorAll(".sec-cat-pill").forEach(_=>{_.addEventListener("click",()=>{l.querySelectorAll(".sec-cat-pill").forEach(P=>P.classList.remove("active")),_.classList.add("active"),a=_.dataset.cat,B()})});let D=l.querySelector("#btn-add-sec-ik");D&&D.addEventListener("click",()=>{k()})}function w(l){Y({title:`${l.icon||"\u{1F6E1}\uFE0F"} IK-${String(l.id).padStart(2,"0")}: ${l.title}`,size:"lg",content:`
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${l.bg||"#EFF6FF"}; color:${l.color||"#1E40AF"}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${l.category}</span>
            <span>Kode: IK-SEC-${String(l.id).padStart(2,"0")}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${l.color||"#1E40AF"};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">\u{1F4CD} Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${l.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F9F0} Alat & Perlengkapan yang Disiapkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${l.tools.map(c=>{if(c.includes("http")){let v=c.split(": ");return`<li>${v[0]}: <a href="${v.slice(1).join(": ")}" target="_blank" rel="noopener" style="color:#2563EB; font-weight:600; text-decoration:underline;">Buka Tautan Form \u2794</a></li>`}return`<li>${f(c)}</li>`}).join("")}
              </ul>
            </div>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">\u{1F4DD} Langkah & Prosedur Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${l.steps.map(c=>`<li style="margin-bottom:6px;">${f(c)}</li>`).join("")}
              </ol>
            </div>
          </div>

          ${l.notes?`
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">\u26A0\uFE0F Catatan Khusus & Ketentuan Penting:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${f(l.notes)}</span>
            </div>
          `:""}
        </div>
      `,confirmText:"Tutup Panduan",onConfirm:(c,v)=>{v()}})}function k(l=null){let c=!!l,v=l?l.tools.join(`
`):"",S=l?l.steps.join(`
`):"",B=`
      <form id="sec-ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-sec-ik-title" class="form-control" required value="${l?f(l.title):""}" placeholder="Contoh: Pengamanan Area & Aset Perusahaan" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-sec-ik-icon" class="form-control" value="${l?l.icon:"\u{1F6E1}\uFE0F"}" placeholder="\u{1F6E1}\uFE0F" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
          <select id="form-sec-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Pengamanan Fisik & Patroli" ${l?.category==="Pengamanan Fisik & Patroli"?"selected":""}>Pengamanan Fisik & Patroli</option>
            <option value="Operasional Parkir & Luar" ${l?.category==="Operasional Parkir & Luar"?"selected":""}>Operasional Parkir & Luar</option>
            <option value="Layanan & Hospitalitas" ${l?.category==="Layanan & Hospitalitas"?"selected":""}>Layanan & Hospitalitas</option>
          </select>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-sec-ik-target" class="form-control" required value="${l?f(l.target):""}" placeholder="Contoh: Seluruh area luar klinik dan akses parkir" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Perlengkapan (1 per baris) *</label>
          <textarea id="form-sec-ik-tools" class="form-control" rows="3" placeholder="Senter Patroli&#10;Lampu Lalin&#10;Peluit" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${f(v)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah & Prosedur Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-sec-ik-steps" class="form-control" rows="5" placeholder="1. Lakukan Safety Patrol sesuai jadwal.&#10;2. Arahkan kendaraan pasien dengan ramah..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${f(S)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Khusus / Ketentuan Penting</label>
          <textarea id="form-sec-ik-notes" class="form-control" rows="2" placeholder="Peringatan keselamatan atau pedoman khusus..." style="width:100%; border-radius:6px; font-size:0.85rem;">${l?f(l.notes||""):""}</textarea>
        </div>
      </form>
    `;Y({title:c?`\u270F\uFE0F Edit IK-${l.id}: ${l.title}`:"\u2795 Tambah Instruksi Kerja Security Baru",content:B,size:"lg",confirmText:"Simpan Perubahan",onConfirm:(T,D)=>{let _=T.querySelector("#form-sec-ik-title").value.trim(),P=T.querySelector("#form-sec-ik-icon").value.trim()||"\u{1F6E1}\uFE0F",M=T.querySelector("#form-sec-ik-cat").value,z=T.querySelector("#form-sec-ik-target").value.trim(),x=T.querySelector("#form-sec-ik-tools").value.trim(),A=T.querySelector("#form-sec-ik-steps").value.trim(),L=T.querySelector("#form-sec-ik-notes").value.trim();if(!_||!z||!x||!A){N("Harap lengkapi semua kolom bertanda bintang (*)");return}let q=x.split(`
`).map(R=>R.trim().replace(/^[-•*]\s*/,"")).filter(Boolean),O=A.split(`
`).map(R=>R.trim().replace(/^\d+[\.\)]\s*/,"")).filter(Boolean),K={"Pengamanan Fisik & Patroli":{color:"#2563EB",bg:"#EFF6FF"},"Operasional Parkir & Luar":{color:"#D97706",bg:"#FFFBEB"},"Layanan & Hospitalitas":{color:"#059669",bg:"#ECFDF5"}}[M]||{color:"#1E40AF",bg:"#EFF6FF"};if(c)l.title=_,l.icon=P,l.category=M,l.target=z,l.tools=q,l.steps=O,l.notes=L,l.color=K.color,l.bg=K.bg,Mt(u),j(`Instruksi Kerja "${_}" berhasil diperbarui`);else{let R=u.length>0?Math.max(...u.map(de=>de.id))+1:1,ie={id:R,title:_,category:M,icon:P,color:K.color,bg:K.bg,target:z,tools:q,steps:O,notes:L};u.push(ie),Mt(u),j(`Instruksi Kerja baru IK-${R} berhasil ditambahkan`)}D(),n()}})}function y(l){l.innerHTML=`
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:20px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-cabang-btn" class="btn btn-sm ${b==="cabang"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            \u{1F3E5} Zonasi Klinik / Cabang
          </button>
          <button id="view-hq-btn" class="btn btn-sm ${b==="hq"?"btn-primary":"btn-ghost"}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
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
    `;let c=l.querySelector("#zonasi-content-container");b==="cabang"?c.innerHTML=`
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
          ${va.cabang.map(v=>`
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; border-bottom:1px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.05rem; font-weight:700; color:${v.color}; display:flex; align-items:center; gap:8px;">
                  <span style="width:12px; height:12px; border-radius:50%; background:${v.color}; display:inline-block;"></span>
                  ${v.zone}
                </h3>
                <span class="badge" style="background:${v.bg}; color:${v.color}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:6px;">
                  ${v.level}
                </span>
              </div>
              <p style="font-size:0.82rem; color:var(--text-3); margin-bottom:14px; line-height:1.5;">${v.desc}</p>
              <div>
                <strong style="font-size:0.8rem; text-transform:uppercase; color:var(--text-2); display:block; margin-bottom:6px;">Daftar Ruangan Terkait:</strong>
                <ul style="margin:0 0 0 16px; padding:0; font-size:0.85rem; color:var(--text-1); line-height:1.7;">
                  ${v.rooms.map(S=>`<li>${S}</li>`).join("")}
                </ul>
              </div>
            </div>
          `).join("")}
        </div>
      `:c.innerHTML=`
        <div style="display:flex; flex-direction:column; gap:16px;">
          ${va.hq.map(v=>`
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; border-bottom:2px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1); display:flex; align-items:center; gap:8px;">
                  <span>\u{1F3E2}</span> Head Quarter \u2014 ${v.floor}
                </h3>
              </div>
              <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:14px;">
                <!-- Zona Hijau -->
                <div style="background:#ECFDF5; border:1px solid #A7F3D0; border-radius:8px; padding:12px;">
                  <strong style="color:#065F46; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E2} ZONA HIJAU (Publik):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#064E3B; line-height:1.6;">
                    ${v.hijau.map(S=>`<li>${S}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Kuning -->
                <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:8px; padding:12px;">
                  <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F7E1} ZONA KUNING (Terbatas):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#78350F; line-height:1.6;">
                    ${v.kuning.map(S=>`<li>${S}</li>`).join("")}
                  </ul>
                </div>
                <!-- Zona Merah -->
                <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:8px; padding:12px;">
                  <strong style="color:#991B1B; font-size:0.85rem; display:block; margin-bottom:6px;">\u{1F534} ZONA MERAH (Restricted):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#7F1D1D; line-height:1.6;">
                    ${v.merah.map(S=>`<li>${S}</li>`).join("")}
                  </ul>
                </div>
              </div>
            </div>
          `).join("")}
        </div>
      `,l.querySelector("#view-cabang-btn").addEventListener("click",()=>{b="cabang",y(l)}),l.querySelector("#view-hq-btn").addEventListener("click",()=>{b="hq",y(l)})}function g(l){l.innerHTML=`
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
          <input type="text" id="sec-asset-search" class="form-control" placeholder="Cari nama aset / tipe / PIC..." value="${o}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
          <button id="btn-add-sec-asset" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
            \u2795 Tambah Aset Baru
          </button>
        </div>
      </div>

      <div id="sec-asset-table-container"></div>
    `;let c=l.querySelector("#sec-asset-table-container"),v=l.querySelector("#sec-asset-search");function S(){let T=o.toLowerCase().trim(),D=d.filter(_=>!T||_.item.toLowerCase().includes(T)||_.type.toLowerCase().includes(T)||_.description.toLowerCase().includes(T)||_.pic&&_.pic.toLowerCase().includes(T));c.innerHTML=`
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
            <strong style="color:var(--text-1); font-size:0.95rem;">Rekap Klasifikasi Aset Perusahaan</strong>
            <span class="badge badge-primary">${D.length} Aset</span>
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
                ${D.map((_,P)=>`
                  <tr style="border-bottom:1px solid #F1F5F9; background:${P%2===0?"#fff":"#FAFAFA"};">
                    <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${_.no||P+1}</td>
                    <td style="padding:10px 14px;">
                      <span class="badge" style="${B(_.type)}">${f(_.type)}</span>
                    </td>
                    <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${f(_.item)}</td>
                    <td style="padding:10px 14px; color:var(--text-2);">${f(_.description)}</td>
                    <td style="padding:10px 14px; font-weight:600; color:#1E40AF;">${f(_.pic||"-")}</td>
                    <td style="padding:10px 14px; text-align:center;">
                      <button class="btn btn-sm btn-ghost btn-edit-asset" data-no="${_.no}" title="Edit Aset" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-sm btn-ghost btn-delete-asset" data-no="${_.no}" title="Hapus Aset" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
        </div>
      `,c.querySelectorAll(".btn-edit-asset").forEach(_=>{_.addEventListener("click",()=>{let P=parseInt(_.dataset.no,10),M=d.find(z=>z.no===P);M&&r(M)})}),c.querySelectorAll(".btn-delete-asset").forEach(_=>{_.addEventListener("click",()=>{let P=parseInt(_.dataset.no,10),M=d.find(z=>z.no===P);M&&ee(`Apakah Anda yakin ingin menghapus aset: "${M.item}"?`,()=>{d=d.filter(z=>z.no!==P),d.forEach((z,x)=>{z.no=x+1}),zt(d),j("Aset berhasil dihapus"),n()},"Hapus Aset")})})}function B(T){return T==="Aset Tetap"?"background:#EFF6FF; color:#1E40AF; font-weight:700; padding:3px 8px; border-radius:6px;":T==="Aset Bergerak"?"background:#ECFDF5; color:#047857; font-weight:700; padding:3px 8px; border-radius:6px;":"background:#FEF3C7; color:#92400E; font-weight:700; padding:3px 8px; border-radius:6px;"}S(),v&&v.addEventListener("input",T=>{o=T.target.value,S()}),l.querySelector("#btn-add-sec-asset").addEventListener("click",()=>{r()})}function r(l=null){let c=!!l,v=`
      <form id="asset-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Tipe / Klasifikasi Aset *</label>
          <select id="form-asset-type" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Aset Tetap" ${l?.type==="Aset Tetap"?"selected":""}>Aset Tetap</option>
            <option value="Aset Bergerak" ${l?.type==="Aset Bergerak"?"selected":""}>Aset Bergerak</option>
            <option value="Low Value Asset" ${l?.type==="Low Value Asset"?"selected":""}>Low Value Asset</option>
          </select>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Aset / Kelompok *</label>
          <input type="text" id="form-asset-item" class="form-control" required value="${l?f(l.item):""}" placeholder="Contoh: Genset & Elevator" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Deskripsi / Contoh Item *</label>
          <textarea id="form-asset-desc" class="form-control" rows="3" required placeholder="Contoh: Generator daya darurat dan panel listrik utama" style="width:100%; border-radius:6px; font-size:0.85rem;">${l?f(l.description):""}</textarea>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">PIC Pengawas / Penanggung Jawab *</label>
          <input type="text" id="form-asset-pic" class="form-control" required value="${l?f(l.pic):""}" placeholder="Contoh: Teknisi & Security" style="width:100%; border-radius:6px; height:36px;">
        </div>
      </form>
    `;Y({title:c?`\u270F\uFE0F Edit Aset: ${l.item}`:"\u2795 Tambah Klasifikasi Aset Baru",content:v,confirmText:"Simpan",onConfirm:(S,B)=>{let T=S.querySelector("#form-asset-type").value,D=S.querySelector("#form-asset-item").value.trim(),_=S.querySelector("#form-asset-desc").value.trim(),P=S.querySelector("#form-asset-pic").value.trim();if(!D||!_||!P){N("Harap lengkapi semua kolom");return}if(c)l.type=T,l.item=D,l.description=_,l.pic=P,zt(d),j(`Aset "${D}" berhasil diperbarui`);else{let M=d.length>0?Math.max(...d.map(z=>z.no))+1:1;d.push({no:M,type:T,item:D,description:_,pic:P}),zt(d),j(`Aset baru "${D}" berhasil ditambahkan`)}B(),n()}})}function s(l){l.innerHTML=`
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
    `}function m(l){l.innerHTML=`
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
    `}async function p(){try{let{downloadExcel:l}=await Promise.resolve().then(()=>(X(),Se)),c=u.map(S=>({"Kode IK":`IK-SEC-${String(S.id).padStart(2,"0")}`,"Judul Prosedur":S.title,Kategori:S.category,"Sasaran Target":S.target,"Alat & Perlengkapan":S.tools.join(", "),"Langkah Pelaksanaan":S.steps.map((B,T)=>`${T+1}. ${B}`).join(" | "),"Catatan Khusus":S.notes||"-"})),v=new Date().toISOString().slice(0,10);l(c,`SOP_Security_Instruksi_Kerja_${v}`),j("Data Excel SOP Security berhasil diunduh")}catch(l){console.error(l),N("Gagal mendownload data Excel: "+l.message)}}function f(l){return l?String(l).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}n()}V();ue();var Ea=[{id:1,name:"Berita Acara Kerusakan Alat",category:"Berita Acara",document_link:"https://docs.google.com/document/d/1example_ba_kerusakan",description:"Format berita acara resmi pencatatan dan eskalasi kerusakan peralatan klinik dan aset fasilitas."},{id:2,name:"Checklist Apar",category:"Apar & K3",document_link:"https://docs.google.com/spreadsheets/d/1example_apar",description:"Lembar inspeksi berkala tabung pemadam api ringan (APAR), tekanan manometer, dan masa kadaluarsa."},{id:3,name:"Checklist Deep Cleaning",category:"Kegiatan Deep Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_deep_cleaning",description:"Panduan dan verifikasi checklist pembersihan mendalam (deep cleaning) area dental unit dan klinik."},{id:4,name:"Checklist Fulldraine",category:"Drain Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_fulldraine",description:"Checklist monitoring pembersihan saluran pembuangan air kotor, floor drain, dan sedotan suction."},{id:5,name:"Checklist General Cleaning",category:"Kegiatan General Cleaning",document_link:"https://docs.google.com/spreadsheets/d/1example_general_cleaning",description:"Standar checklist kebersihan harian/mingguan seluruh ruangan, lobby, toilet, dan koridor."},{id:6,name:"Form Aktivity Reliefer",category:"Form Operasional",document_link:"https://forms.gle/1example_reliefer",description:"Formulir pencatatan absensi dan penugasan harian reliefer Facility Care di cabang pengganti."},{id:7,name:"Form Aktivity Security HQ",category:"Form Operasional",document_link:"https://forms.gle/1example_security_hq",description:"Formulir monitoring shift dan pengawasan keamanan harian di Kantor Pusat (Head Quarter)."}];function Ca(e=""){let t=(e||"").toLowerCase();return t.includes("berita acara")?{label:e,bg:"#F3E8FF",color:"#7E22CE",border:"#D8B4FE",icon:"\u2696\uFE0F"}:t.includes("apar")||t.includes("k3")||t.includes("safety")?{label:e,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F9EF}"}:t.includes("deep cleaning")?{label:e,bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2728"}:t.includes("drain")?{label:e,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u{1F6BF}"}:t.includes("general cleaning")||t.includes("kebersihan")?{label:e,bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F9F9}"}:t.includes("form")?{label:e,bg:"#E0E7FF",color:"#4338CA",border:"#C7D2FE",icon:"\u{1F4DD}"}:{label:e||"Dokumen",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4C1}"}}function Fa(e=""){if(!e)return{type:"Tidak Ada Tautan",icon:"\u{1F4CE}",badgeClass:"bg-slate-100 text-slate-500"};let t=e.toLowerCase();return t.includes("forms.gle")||t.includes("google.com/forms")?{type:"Google Form",icon:"\u{1F7E3}",badgeClass:"badge-form",color:"#673AB7"}:t.includes("docs.google.com/spreadsheets")||t.includes("sheet")?{type:"Google Sheets",icon:"\u{1F7E2}",badgeClass:"badge-sheet",color:"#0F9D58"}:t.includes("docs.google.com/document")||t.includes("doc")?{type:"Google Docs",icon:"\u{1F535}",badgeClass:"badge-doc",color:"#4285F4"}:t.includes("drive.google.com")?{type:"Google Drive",icon:"\u{1F7E1}",badgeClass:"badge-drive",color:"#F4B400"}:t.endsWith(".pdf")?{type:"PDF Document",icon:"\u{1F534}",badgeClass:"badge-pdf",color:"#EA4335"}:{type:"Web Link",icon:"\u{1F310}",badgeClass:"badge-link",color:"#2563EB"}}async function $a(e){let t=[],i="all",a="",o="fcms-master-doc-styles";if(!document.getElementById(o)){let w=document.createElement("style");w.id=o,w.textContent=`
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
        font-size: 1.65rem;
        font-weight: 700;
        margin: 0 0 6px 0;
        display: flex;
        align-items: center;
        gap: 12px;
        letter-spacing: -0.02em;
      }
      .doc-hero-subtitle {
        color: #94A3B8;
        font-size: 0.92rem;
        margin: 0;
        max-width: 720px;
        line-height: 1.5;
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
        .doc-hero-header > div:last-child {
          display: flex !important;
          width: 100% !important;
          gap: 8px !important;
        }
        #doc-btn-add {
          flex: 2 !important;
          justify-content: center !important;
          font-size: 0.82rem !important;
          padding: 9px 12px !important;
        }
        #doc-btn-excel {
          flex: 1.2 !important;
          justify-content: center !important;
          font-size: 0.82rem !important;
          padding: 9px 10px !important;
        }
        #doc-btn-refresh {
          flex: 0 0 38px !important;
          justify-content: center !important;
          padding: 9px 0 !important;
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
          grid-template-columns: 1fr !important;
          gap: 12px !important;
        }
        .doc-card {
          padding: 16px !important;
          border-radius: 14px !important;
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
    `,document.head.appendChild(w)}async function b(){try{let w=await I("/api/checklist?limit=1000");if(w.ok&&w.data?.data&&Array.isArray(w.data.data)&&w.data.data.length>0)t=w.data.data,localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t));else{let k=localStorage.getItem("fcms_master_doc_cache");k?t=JSON.parse(k):(t=[...Ea],localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)))}}catch(w){console.warn("Fallback loading documents from cache/seed:",w);let k=localStorage.getItem("fcms_master_doc_cache");t=k?JSON.parse(k):[...Ea]}E()}function u(){return t.filter(w=>{let k=i==="all"||w.category===i,y=a.toLowerCase().trim(),g=!y||(w.name||"").toLowerCase().includes(y)||(w.category||"").toLowerCase().includes(y)||(w.description||"").toLowerCase().includes(y);return k&&g})}function d(){let w={all:t.length};return t.forEach(k=>{let y=k.category||"Lainnya";w[y]=(w[y]||0)+1}),w}function n(w=null){let k=!!w,y=Array.from(new Set(t.map(r=>r.category).filter(Boolean))),g=document.createElement("div");g.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Nama Dokumen / Checklist <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Checklist Pembersihan Wastafel & Toilet" value="${w?.name?String(w.name).replace(/"/g,"&quot;"):""}" required />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Kategori Dokumen <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-category" list="doc-categories-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Pilih atau ketik kategori baru..." value="${w?.category?String(w.category).replace(/"/g,"&quot;"):""}" required />
          <datalist id="doc-categories-list">
            ${y.map(r=>`<option value="${r}">${r}</option>`).join("")}
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
            placeholder="https://docs.google.com/..." value="${w?.document_link?String(w.document_link).replace(/"/g,"&quot;"):""}" />
          <span style="font-size:0.75rem; color:#64748B; margin-top:4px; display:block;">
            Bisa berupa tautan Google Forms untuk inspeksi harian, atau link Google Sheets / Docs sebagai acuan.
          </span>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Deskripsi / Petunjuk Penggunaan
          </label>
          <textarea id="doc-modal-desc" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan keterangan singkat mengenai tujuan dan instruksi pengisian dokumen ini...">${w?.description||""}</textarea>
        </div>
      </div>
    `,Y({title:k?"\u270F\uFE0F Edit Master Dokumen":"\u2795 Tambah Master Dokumen Baru",content:g,size:"md",confirmText:k?"Simpan Perubahan":"Tambahkan Dokumen",confirmClass:"btn-primary",onConfirm:async(r,s)=>{let m=g.querySelector("#doc-modal-name"),p=g.querySelector("#doc-modal-category"),f=g.querySelector("#doc-modal-link"),l=g.querySelector("#doc-modal-desc"),c=(m.value||"").trim(),v=(p.value||"").trim(),S=(f.value||"").trim(),B=(l.value||"").trim();if(!c){N("Nama Dokumen wajib diisi"),m.focus();return}if(!v){N("Kategori Dokumen wajib diisi"),p.focus();return}let T={name:c,category:v,document_link:S,description:B};try{if(k)if((await I(`/api/checklist/${w.id}`,{method:"PUT",body:JSON.stringify(T)})).ok){let _=t.findIndex(P=>P.id===w.id);_!==-1&&(t[_]={...t[_],...T}),j("Dokumen berhasil diperbarui")}else{let _=t.findIndex(P=>P.id===w.id);_!==-1&&(t[_]={...t[_],...T}),j("Dokumen disimpan secara lokal")}else{let D=await I("/api/checklist",{method:"POST",body:JSON.stringify(T)});if(D.ok&&D.data?.id)t.unshift({id:D.data.id,...T}),j("Dokumen baru berhasil ditambahkan");else{let _=Date.now();t.unshift({id:_,...T}),j("Dokumen baru ditambahkan secara lokal")}}localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),s(),E()}catch(D){console.error("Error saving document:",D),N("Gagal menyimpan dokumen: "+D.message)}}})}function h(w){ee(`Apakah Anda yakin ingin menghapus dokumen <strong>"${w.name}"</strong>?`,async()=>{try{let k=await I(`/api/checklist/${w.id}`,{method:"DELETE"});k.ok||console.warn("API delete warning:",k),t=t.filter(y=>y.id!==w.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),j(`Dokumen "${w.name}" berhasil dihapus`),E()}catch{t=t.filter(y=>y.id!==w.id),localStorage.setItem("fcms_master_doc_cache",JSON.stringify(t)),j(`Dokumen "${w.name}" berhasil dihapus`),E()}},"Hapus Dokumen")}function F(w,k){if(!w){vt("Dokumen belum memiliki tautan digital.");return}navigator.clipboard.writeText(w).then(()=>{j(`Tautan "${k}" disalin ke clipboard!`)}).catch(()=>{vt(`Tautan: ${w}`)})}async function $(){try{let{downloadExcel:w}=await Promise.resolve().then(()=>(X(),Se)),k=t.map((y,g)=>({No:g+1,"Nama Dokumen":y.name||"",Kategori:y.category||"","Platform / Format":Fa(y.document_link).type,"Link Dokumen":y.document_link||"","Deskripsi / Keterangan":y.description||""}));w(k,`Master_Dokumen_FCMS_${new Date().toISOString().slice(0,10)}`),j("Berhasil mengunduh data dokumen ke Excel.")}catch(w){N("Gagal mengekspor data: "+w.message)}}function E(){let w=u(),k=d(),y=Object.keys(k).filter(r=>r!=="all"),g=t.filter(r=>!!r.document_link).length;e.innerHTML=`
      <div class="doc-page-container">
        
        <!-- Hero Header -->
        <div class="doc-hero-card">
          <div class="doc-hero-header">
            <div>
              <h1 class="doc-hero-title">
                <span>\u{1F4CB}</span> Master Document & Template Standar
              </h1>
              <p class="doc-hero-subtitle">
                Pusat repositori formulir digital, checklist inspeksi kebersihan, berita acara, dan dokumen kerja standar PT Audy Mandiri Indonesia.
              </p>
            </div>
            <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
              <button id="doc-btn-add" class="btn btn-primary" style="background:#2563EB; border:none; border-radius:10px; padding:10px 18px; font-weight:600; display:flex; align-items:center; gap:8px;">
                <span>\u2795</span> Tambah Dokumen Baru
              </button>
              <button id="doc-btn-excel" class="btn btn-ghost" style="background:rgba(255,255,255,0.1); color:#FFFFFF; border:1px solid rgba(255,255,255,0.2); border-radius:10px; padding:10px 16px; font-weight:600; display:flex; align-items:center; gap:8px;">
                <span>\u{1F4E5}</span> Ekspor Excel
              </button>
              <button id="doc-btn-refresh" class="btn btn-ghost" style="background:rgba(255,255,255,0.1); color:#FFFFFF; border:1px solid rgba(255,255,255,0.2); border-radius:10px; padding:10px 14px; font-weight:600;">
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
                <div class="doc-stat-value">${y.length}</div>
                <div class="doc-stat-label">Kategori Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon" style="background:rgba(16,185,129,0.18); color:#34D399;">\u{1F310}</div>
              <div>
                <div class="doc-stat-value">${g} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">/ ${t.length}</span></div>
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
              <input type="text" id="doc-search-input" class="doc-search-input" placeholder="Cari nama dokumen, kategori, atau deskripsi..." value="${a?a.replace(/"/g,"&quot;"):""}" />
            </div>
          </div>

          <!-- Category Pills -->
          <div class="doc-filter-pills">
            <button class="doc-pill ${i==="all"?"active":""}" data-cat="all">
              <span>Semua Dokumen</span>
              <span class="doc-pill-count">${k.all||0}</span>
            </button>
            ${y.map(r=>{let s=Ca(r);return`
                <button class="doc-pill ${i===r?"active":""}" data-cat="${r}">
                  <span>${s.icon}</span>
                  <span>${r}</span>
                  <span class="doc-pill-count">${k[r]||0}</span>
                </button>
              `}).join("")}
          </div>
        </div>

        <!-- Documents View Container -->
        <div id="doc-list-view">
          ${w.length===0?`
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
              ${w.map(r=>{let s=Ca(r.category),m=Fa(r.document_link);return`
                  <div class="doc-card" data-id="${r.id}">
                    <div>
                      <div class="doc-card-top">
                        <span class="doc-category-badge" style="background:${s.bg}; color:${s.color}; border:1px solid ${s.border};">
                          <span>${s.icon}</span>
                          <span>${s.label}</span>
                        </span>
                        <span class="doc-platform-badge">
                          <span>${m.icon}</span>
                          <span>${m.type}</span>
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
                        <button class="doc-icon-btn edit-doc-btn" title="Edit Dokumen" data-id="${r.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        <button class="doc-icon-btn delete delete-doc-btn" title="Hapus Dokumen" data-id="${r.id}">
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
    `,C()}function C(){let w=e.querySelector("#doc-btn-add");w&&(w.onclick=()=>n());let k=e.querySelector("#doc-btn-excel");k&&(k.onclick=()=>$());let y=e.querySelector("#doc-btn-refresh");y&&(y.onclick=()=>b());let g=e.querySelector("#doc-btn-reset-filter");g&&(g.onclick=()=>{a="",i="all",E()});let r=e.querySelector("#doc-search-input");r&&(r.oninput=s=>{a=s.target.value;let m=u();E();let p=e.querySelector("#doc-search-input");p&&(p.focus(),p.setSelectionRange(p.value.length,p.value.length))}),e.querySelectorAll(".doc-pill").forEach(s=>{s.onclick=()=>{i=s.getAttribute("data-cat")||"all",E()}}),e.querySelectorAll(".copy-link-btn").forEach(s=>{s.onclick=m=>{m.stopPropagation();let p=s.getAttribute("data-link"),f=s.getAttribute("data-title");F(p,f)}}),e.querySelectorAll(".edit-doc-btn").forEach(s=>{s.onclick=m=>{m.stopPropagation();let p=Number(s.getAttribute("data-id")),f=t.find(l=>l.id===p);f&&n(f)}}),e.querySelectorAll(".delete-doc-btn").forEach(s=>{s.onclick=m=>{m.stopPropagation();let p=Number(s.getAttribute("data-id")),f=t.find(l=>l.id===p);f&&h(f)}})}await b()}V();ue();X();var ze={"Standar Umum":{icon:"\u{1F4CC}",bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE",rooms:["Lantai & Sudut Ruangan","Dinding, Partisi & Plint","Kaca, Cermin & Partisi Kaca","Sanitair & Area Basah Toilet","Plafon, Lampu & Kisi Ventilasi","Mebel, Meja & Kursi Kerja/Pasien","Peralatan Khusus & Medis Non-Steril","Wadah Sampah & Utilitas"]},"Area Depan & Pelayanan":{icon:"\u{1F6CB}\uFE0F",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",rooms:["Teras","Lobby 1","Receptionist","Snack Corner"]},"Area Klinis & Medis":{icon:"\u{1FA7A}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD",rooms:["Ruang Tindakan","Ronsen","Ruang Dokter","Ruang Steril"]},"Toilet & Sanitasi":{icon:"\u{1F6BF}",bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4",rooms:["Toilet 1","Janitor","Ruang Limbah"]},Sirkulasi:{icon:"\u{1F6B6}",bg:"#F1F5F9",color:"#334155",border:"#CBD5E1",rooms:["Area Tangga","Lift","Area Tangga Exit"]},"Area Anak & Fasilitas":{icon:"\u{1F9F8}",bg:"#FEF3C7",color:"#92400E",border:"#FDE68A",rooms:["Lobby 2","Playground","Musholla","Tempat Wudhu","Ruang Team"]},"Gudang & Teknis":{icon:"\u2699\uFE0F",bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA",rooms:["Gudang Farmasi & Bahan","RoofTop","Ruang Kompressor","Area Genset"]},Eksterior:{icon:"\u{1F333}",bg:"#DCFCE7",color:"#15803D",border:"#86EFAC",rooms:["Taman","Pos Security","Toilet Security","Halaman Parkir"]}};function Kt(e){for(let[t,i]of Object.entries(ze))if(i.rooms.includes(e))return{zoneName:t,...i};return{zoneName:"Lainnya",icon:"\u{1F4CD}",bg:"#F1F5F9",color:"#475569",border:"#E2E8F0"}}async function _a(e){e.innerHTML=`
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
        ${Object.entries(ze).map(([s,m])=>`
          <button class="zone-tab-btn" data-zone="${s}" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
            <span>${m.icon}</span> ${s}
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
  `;let t="all",i="table",a=[],o=ae(),b=o&&["superadmin","admin","manager","spv"].includes(o.role);async function u(){try{let s=await I("/api/hygiene-standards?all=1");if(!s.ok)throw new Error("Gagal memuat standar hygiene");a=Array.isArray(s.data)?s.data:s.data?.data||[],n(),d(),F()}catch(s){console.error(s),N(s.message||"Terjadi kesalahan memuat data");let m=e.querySelector("#hygiene-view-container");m&&(m.innerHTML=`
          <div class="empty-state" style="text-align:center; padding:40px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:12px;">\u26A0\uFE0F</div>
            <h3 style="margin:0 0 8px 0; color:var(--text-1);">Gagal Memuat Data</h3>
            <p style="color:var(--text-3); margin-bottom:16px;">${s.message}</p>
            <button id="btn-retry-hygiene" class="btn btn-primary" style="padding:8px 16px; border-radius:8px;">Coba Lagi</button>
          </div>
        `,m.querySelector("#btn-retry-hygiene")?.addEventListener("click",u))}}function d(){let s=e.querySelector("#badge-count-all");s&&(s.textContent=a.length),Object.entries(ze).forEach(([f,l])=>{let c=f.replace(/[\s&]+/g,"-"),v=e.querySelector(`#badge-count-${c}`);if(v){let S=a.filter(B=>l.rooms.includes(B.room_name)).length;v.textContent=S}});let m=new Set(a.map(f=>f.room_name)),p=e.querySelector("#card-total-rooms");p&&(p.textContent=m.size)}function n(){let s=e.querySelector("#hygiene-room-filter");if(!s)return;let m=s.value,p=[];if(t==="all")p=Array.from(new Set(a.map(f=>f.room_name))).filter(Boolean).sort();else{let f=ze[t];f&&(p=f.rooms.filter(l=>a.some(c=>c.room_name===l)))}s.innerHTML=`<option value="">Semua Ruangan (${p.length})</option>`+p.map(f=>`<option value="${f}" ${f===m?"selected":""}>${f}</option>`).join("")}function h(){let s=(e.querySelector("#hygiene-search")?.value||"").toLowerCase().trim(),m=e.querySelector("#hygiene-room-filter")?.value||"";return a.filter(p=>{if(t!=="all"){let f=ze[t];if(!f||!f.rooms.includes(p.room_name))return!1}if(m&&p.room_name!==m)return!1;if(s){let f=(p.room_name||"").toLowerCase(),l=(p.item_name||"").toLowerCase(),c=(p.cleanliness_standard||"").toLowerCase();if(!f.includes(s)&&!l.includes(s)&&!c.includes(s))return!1}return!0})}function F(){let s=h(),m=e.querySelector("#hygiene-stats-text");m&&(m.innerHTML=`Menampilkan <strong>${s.length}</strong> dari <strong>${a.length}</strong> parameter`);let p=e.querySelector("#hygiene-view-container");if(p){if(s.length===0){p.innerHTML=`
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
      `,p.querySelector("#btn-clear-search-empty")?.addEventListener("click",()=>{let f=e.querySelector("#hygiene-search"),l=e.querySelector("#hygiene-room-filter");f&&(f.value=""),l&&(l.value=""),t="all",e.querySelectorAll(".zone-tab-btn").forEach(c=>c.classList.toggle("active",c.dataset.zone==="all")),e.querySelectorAll(".zone-tab-btn").forEach(c=>{c.dataset.zone==="all"?(c.style.background="#2563EB",c.style.color="#fff"):(c.style.background="#F1F5F9",c.style.color="var(--text-2)")}),n(),F()});return}i==="table"?$(p,s):E(p,s)}}function $(s,m){s.innerHTML=`
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
              ${m.map((p,f)=>{let l=Kt(p.room_name);return`
                  <tr style="border-bottom:1px solid #F1F5F9; transition:background-color 0.15s ease;" onmouseover="this.style.backgroundColor='#F8FAFC'" onmouseout="this.style.backgroundColor='transparent'">
                    <td style="padding:14px 16px; text-align:center; color:var(--text-3); font-weight:600; font-size:0.85rem;">
                      ${f+1}
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="margin-bottom:4px;">
                        <span style="display:inline-flex; align-items:center; gap:6px; background:${l.bg}; color:${l.color}; border:1px solid ${l.border}; padding:4px 10px; border-radius:6px; font-weight:700; font-size:0.83rem;">
                          <span>${l.icon}</span> ${p.room_name}
                        </span>
                      </div>
                      <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">
                        ${l.zoneName}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="font-weight:700; color:var(--text-1); line-height:1.45; font-size:0.9rem;">
                        ${p.item_name}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-left:3px solid #10B981; border-radius:6px; padding:10px 12px; line-height:1.55; color:#334155; font-size:0.875rem;">
                        <span style="color:#059669; font-weight:700; margin-right:4px;">\u2713</span> ${p.cleanliness_standard}
                      </div>
                    </td>
                    ${b?`
                      <td style="padding:14px 16px; text-align:center; vertical-align:middle;">
                        <div style="display:inline-flex; gap:6px;">
                          <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${p.id}" title="Edit Parameter" style="padding:6px 8px; border-radius:6px; color:#2563EB; cursor:pointer;">
                            \u270F\uFE0F
                          </button>
                          <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${p.id}" data-name="${p.item_name}" title="Hapus Parameter" style="padding:6px 8px; border-radius:6px; color:#DC2626; cursor:pointer;">
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
    `,C(s)}function E(s,m){let p={};m.forEach(f=>{let l=Kt(f.room_name);p[l.zoneName]||(p[l.zoneName]={info:l,rooms:{}}),p[l.zoneName].rooms[f.room_name]||(p[l.zoneName].rooms[f.room_name]=[]),p[l.zoneName].rooms[f.room_name].push(f)}),s.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:24px;">
        ${Object.entries(p).map(([f,l])=>{let c=Object.values(l.rooms).reduce((v,S)=>v+S.length,0);return`
            <div class="card" style="background:#fff; border:1px solid ${l.info.border}; border-radius:14px; overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.03);">
              <!-- Zone Card Header -->
              <div style="background:${l.info.bg}; border-bottom:1px solid ${l.info.border}; padding:14px 20px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="font-size:1.4rem;">${l.info.icon}</span>
                  <div>
                    <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:${l.info.color};">
                      ${f}
                    </h3>
                    <div style="font-size:0.78rem; color:${l.info.color}; opacity:0.85;">
                      ${Object.keys(l.rooms).length} Ruangan \xB7 ${c} Parameter Standar
                    </div>
                  </div>
                </div>
                <span class="badge" style="background:#fff; color:${l.info.color}; border:1px solid ${l.info.border}; font-weight:700; padding:4px 10px; border-radius:12px; font-size:0.8rem;">
                  ${c} Parameter
                </span>
              </div>

              <!-- Zone Rooms List -->
              <div style="padding:16px 20px;">
                ${Object.entries(l.rooms).map(([v,S])=>`
                  <div style="margin-bottom:20px; last-child:margin-bottom:0;">
                    <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                      <h4 style="margin:0; font-size:0.95rem; font-weight:700; color:var(--text-1);">
                        \u{1F6AA} ${v}
                      </h4>
                      <span class="badge badge-sm" style="background:#F1F5F9; color:var(--text-3); font-size:0.75rem;">
                        ${S.length} item
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
                          ${S.map((B,T)=>`
                            <tr style="border-bottom:1px solid #F1F5F9;">
                              <td style="padding:10px 12px; text-align:center; color:var(--text-3); font-size:0.8rem;">
                                ${T+1}
                              </td>
                              <td style="padding:10px 12px; font-weight:600; color:var(--text-1); vertical-align:top;">
                                ${B.item_name}
                              </td>
                              <td style="padding:10px 12px; vertical-align:top; color:#334155;">
                                <div style="display:flex; align-items:flex-start; gap:6px;">
                                  <span style="color:#10B981; font-weight:700;">\u2713</span>
                                  <span>${B.cleanliness_standard}</span>
                                </div>
                              </td>
                              ${b?`
                                <td style="padding:10px 12px; text-align:center; vertical-align:middle;">
                                  <div style="display:inline-flex; gap:4px;">
                                    <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${B.id}" title="Edit" style="padding:4px 6px; cursor:pointer;">\u270F\uFE0F</button>
                                    <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${B.id}" data-name="${B.item_name}" title="Hapus" style="padding:4px 6px; color:#DC2626; cursor:pointer;">\u{1F5D1}\uFE0F</button>
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
    `,C(s)}function C(s){s.querySelectorAll(".btn-edit-item").forEach(m=>{m.addEventListener("click",()=>{let p=m.dataset.id,f=a.find(l=>String(l.id)===String(p));f&&w(f)})}),s.querySelectorAll(".btn-delete-item").forEach(m=>{m.addEventListener("click",()=>{let p=m.dataset.id,f=m.dataset.name;ee(`Yakin ingin menghapus parameter standar <strong>"${f}"</strong>?`,async()=>{try{let l=await I(`/api/hygiene-standards/${p}`,{method:"DELETE"});if(!l.ok)throw new Error(l.data?.error||"Gagal menghapus parameter");j("Parameter kebersihan berhasil dihapus"),u()}catch(l){N(l.message)}},"Hapus Parameter Standar")})})}function w(s=null){let m=!!s,p=Array.from(new Set(a.map(l=>l.room_name))).filter(Boolean).sort(),f=document.createElement("div");f.innerHTML=`
      <form id="form-hygiene-item" style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Ruangan / Lokasi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-room-name" list="room-suggestions" class="form-control" placeholder="Contoh: Ruang Tindakan, Lobby 1, Toilet 1..." value="${s?.room_name||""}" required style="width:100%; height:40px; border-radius:8px;">
          <datalist id="room-suggestions">
            ${p.map(l=>`<option value="${l}">`).join("")}
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
    `,Y({title:m?"\u270F\uFE0F Edit Parameter Kebersihan":"\u2795 Tambah Parameter Kebersihan",content:f,confirmText:m?"Simpan Perubahan":"Tambah Parameter",size:"md",onConfirm:async(l,c)=>{let v=f.querySelector("#modal-room-name").value.trim(),S=f.querySelector("#modal-item-name").value.trim(),B=f.querySelector("#modal-standard").value.trim();if(!v||!S||!B){N("Semua kolom wajib diisi");return}try{let T={room_name:v,item_name:S,cleanliness_standard:B},D=m?`/api/hygiene-standards/${s.id}`:"/api/hygiene-standards",P=await I(D,{method:m?"PUT":"POST",body:JSON.stringify(T)});if(!P.ok)throw new Error(P.data?.error||"Gagal menyimpan parameter");j(m?"Parameter berhasil diperbarui":"Parameter berhasil ditambahkan"),c(),u()}catch(T){N(T.message)}}})}function k(){let s=window.open("","_blank");if(!s){N("Pop-up printer diblokir oleh browser. Izinkan pop-up untuk mencetak SOP.");return}let p=`
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
            <div>Total: <strong>${a.length} Parameter Fisik</strong></div>
          </div>
        </div>

        <div class="title">
          <h1>STANDAR ACUAN KEBERSIHAN FISIK SELURUH CABANG</h1>
          <p>Buku Panduan Kriteria Kebersihan Fisik 8 Zona Area (Cleaning Service & Facility Care)</p>
        </div>

        ${Object.entries(ze).map(([f,l])=>{let c=a.filter(v=>l.rooms.includes(v.room_name));return c.length===0?"":`
            <div class="zone-title">${l.icon} ${f.toUpperCase()} (${c.length} PARAMETER)</div>
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
                ${c.map((v,S)=>`
                  <tr>
                    <td class="td-no">${S+1}</td>
                    <td class="td-room">${v.room_name}</td>
                    <td class="td-obj">${v.item_name}</td>
                    <td>${v.cleanliness_standard}</td>
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
    `;s.document.open(),s.document.write(p),s.document.close()}function y(){let s=a.map((p,f)=>{let l=Kt(p.room_name);return{No:f+1,"Kategori Zona":l.zoneName,"Ruangan / Lokasi":p.room_name,"Item / Objek Inspeksi":p.item_name,"Standar Kebersihan Fisik":p.cleanliness_standard}}),m=new Date().toISOString().slice(0,10);G(s,`Master_Standar_Kebersihan_FCMS_${m}`),j("Data master kebersihan berhasil diekspor ke Excel")}e.querySelector("#hygiene-search")?.addEventListener("input",()=>F()),e.querySelector("#hygiene-room-filter")?.addEventListener("change",()=>F()),e.querySelector("#btn-reset-filters")?.addEventListener("click",()=>{let s=e.querySelector("#hygiene-search"),m=e.querySelector("#hygiene-room-filter");s&&(s.value=""),m&&(m.value=""),F()}),e.querySelectorAll(".zone-tab-btn").forEach(s=>{s.addEventListener("click",()=>{t=s.dataset.zone,e.querySelectorAll(".zone-tab-btn").forEach(m=>{let p=m.dataset.zone===t;if(m.classList.toggle("active",p),p){m.style.background="#2563EB",m.style.color="#fff";let f=m.querySelector(".badge");f&&(f.style.background="rgba(255,255,255,0.25)",f.style.color="#fff")}else{m.style.background="#F1F5F9",m.style.color="var(--text-2)";let f=m.querySelector(".badge");f&&(f.style.background="#CBD5E1",f.style.color="#334155")}}),n(),F()})});let g=e.querySelector("#view-table-btn"),r=e.querySelector("#view-zone-btn");g?.addEventListener("click",()=>{i="table",g.style.background="#fff",g.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",r.style.background="transparent",r.style.boxShadow="none",F()}),r?.addEventListener("click",()=>{i="zone",r.style.background="#fff",r.style.boxShadow="0 1px 2px rgba(0,0,0,0.05)",g.style.background="transparent",g.style.boxShadow="none",F()}),e.querySelector("#btn-print-hygiene")?.addEventListener("click",k),e.querySelector("#btn-export-hygiene")?.addEventListener("click",y),e.querySelector("#btn-add-hygiene")?.addEventListener("click",()=>w(null)),u()}V();ue();X();async function Ta(e,t="supply"){return An(e)}async function An(e){let i=((await I("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));Z({container:e,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,o)=>o.branch_name_ref||o.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let o=JSON.parse(a);return Array.isArray(o)?o.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let o=JSON.parse(a);return Array.isArray(o)?o.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>ne(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let o=a?.tools_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}let b=a?.chemical_items;try{b=Array.isArray(JSON.parse(b))?JSON.parse(b).join(", "):b}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:a?.branch_id&&!i.find(u=>u.value==a.branch_id)?[...i,{value:a.branch_id,label:a.branch_name||a.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:o},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:b},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),b=await I(`/api/reports/supply?limit=10000&${o}`);if(b.ok){let u=b.data.data.map(d=>{let n=d.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let h=d.chemical_items;try{h=Array.isArray(JSON.parse(h))?JSON.parse(h).join(", "):h}catch{}return{Waktu:d.submitted_at||"",Pengirim:d.submitter_name||"",Cabang:d.branch_name_ref||d.branch_name||"","Alat/Barang":n||"",Chemical:h||"",Catatan:d.additional_notes||"",Status:d.status||"","Diproses Oleh":d.processed_by||""}});G(u,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{G([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let b=(await I("/api/branches?all=1")).data?.data||[],u=F=>{if(!F)return null;let $=String(F||"").toLowerCase(),E=b.find(C=>String(C.full_name||"").toLowerCase()===$||String(C.code||"").toLowerCase()===$||String(C.name||"").toLowerCase()===$);return E?E.id:null},d=F=>{if(F==null||F==="")return"";if(F instanceof Date&&!isNaN(F.getTime()))return F.toISOString().slice(0,10);let $=String(F).trim();if($===""||$==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test($))return $.slice(0,10);if(/^\d{4,5}$/.test($)){let C=Number($);if(C>2e4&&C<99999){let w=new Date(Date.UTC(1899,11,30)+C*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}let E=$.split(/[\/\-\.]/);if(E.length===3){let[C,w,k]=E.map(y=>y.trim());if(C.length===4&&w.length<=2&&k.length<=2)return`${C}-${w.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&w.length<=2&&C.length<=2)return`${k}-${w.padStart(2,"0")}-${C.padStart(2,"0")}`}return $},n=a.map(F=>({submitted_at:d(F.Waktu||F.Tanggal),submitter_name:String(F.Pengirim||"").trim(),branch_id:u(String(F.Cabang||"").trim()),tools_items:String(F["Alat/Barang"]||F.Alat||"").trim(),chemical_items:String(F.Chemical||"").trim(),additional_notes:String(F.Catatan||F.Keterangan||"").trim(),status:String(F.Status||"").trim(),processed_by:String(F["Diproses Oleh"]||F.PIC||"").trim()})).filter(F=>F.submitted_at&&F.submitter_name&&F.branch_id),h=await I("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!h.ok)throw new Error(h.data?.error||"Import gagal");return h.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,o)=>{let b=Y({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(u,d)=>{let n=u.querySelector("#supply-status").value,h=u.querySelector("#supply-processed-by").value;(await I(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:h})})).ok?(j("Status diperbarui."),d(),o()):N("Gagal update status.")}})}}]})}V();X();async function Aa(e){let t=ae();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}Z({container:e,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",editor_khusus:"badge-warning",input_lembur:"badge-info",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let a=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"editor_khusus",label:"Editor (Hanya Lembur & Masalah)"},{value:"input_lembur",label:"Input Lembur (Hanya Isi Lembur)"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await I(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let a=i.data.data.map(o=>({"Nama Lengkap":o.full_name||"",Username:o.username||"",Email:o.email||"",Role:o.role||"",Status:o.is_active?"Aktif":"Nonaktif"}));G(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{G([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let a=i.map(b=>({full_name:String(b["Nama Lengkap"]||"").trim(),username:String(b.Username||"").trim(),email:String(b.Email||"").trim(),role:String(b.Role||"").trim()||"viewer",password:String(b.Password||"").trim()})).filter(b=>b.username&&b.password&&b.email&&b.full_name),o=await I("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}}})}V();X();async function Da(e){Z({container:e,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:t=>t?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:t=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:t?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:t?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:t?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:t?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:t?.is_active!==void 0?t.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let t=await I(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(t.ok)G(t.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{G([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async t=>{let i=t.map(o=>({code:String(o["Kode Cabang"]||"").trim(),name:String(o["Nama Pendek"]||"").trim(),full_name:String(o["Nama Lengkap"]||"").trim(),city:String(o.Kota||"").trim()})).filter(o=>o.code&&o.name),a=await I("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}V();async function Ba(e){let t=new Date,i=[];e.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{t.setMonth(t.getMonth()-1),o()}),document.getElementById("cal-next").addEventListener("click",()=>{t.setMonth(t.getMonth()+1),o()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(u=>u.addEventListener("change",o));async function a(){try{let u=`${t.getFullYear()}-${String(t.getMonth()+1).padStart(2,"0")}`;i=(await I(`/api/dashboard/calendar?month=${u}`)).data?.data||[]}catch(u){console.warn("[Calendar] Failed to load events, rendering empty grid:",u),i=[]}}async function o(){let u=document.getElementById("calendar-grid");if(u){u.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let d=t.getFullYear(),n=t.getMonth(),h=t.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),F=document.getElementById("cal-month-label");F&&(F.textContent=h);let $=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(p=>p.value)),E=i.filter(p=>$.has(p.type)),C={};E.forEach(p=>{let f=Dn(p.event_date);f&&(C[f]||(C[f]=[]),C[f].push(p))});let w=new Date(d,n,1).getDay(),k=new Date(d,n+1,0).getDate(),y=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],g=new Date().toISOString().slice(0,10),r='<div class="calendar-grid">';y.forEach(p=>{r+=`<div class="cal-day-header">${p}</div>`});for(let p=0;p<w;p++)r+='<div class="cal-cell cal-cell-empty"></div>';for(let p=1;p<=k;p++){let f=`${d}-${String(n+1).padStart(2,"0")}-${String(p).padStart(2,"0")}`,l=C[f]||[],c=f===g;r+=`
          <div class="cal-cell ${c?"cal-today":""} ${l.length?"cal-has-events":""}"
               data-date="${f}" tabindex="0" role="button" aria-label="${f}">
            <div class="cal-day-num ${c?"today-num":""}">${p}</div>
            <div class="cal-events-preview">
              ${l.slice(0,3).map(v=>`
                <div class="cal-event-dot cal-color-${v.color||"gray"}" title="${we(v.title||v.type)}">
                  <span class="cal-event-dot-label">${Bn(v.title||v.branch_name||v.type,18)}</span>
                </div>
              `).join("")}
              ${l.length>3?`<div class="cal-more">+${l.length-3} lagi</div>`:""}
            </div>
          </div>`}let m=(w+k)%7;if(m!==0)for(let p=0;p<7-m;p++)r+='<div class="cal-cell cal-cell-empty"></div>';r+="</div>",u.innerHTML=r,u.querySelectorAll(".cal-cell[data-date]").forEach(p=>{p.addEventListener("click",()=>{let f=p.dataset.date,l=C[f]||[];if(!l.length)return;let c=document.getElementById("cal-event-list"),v=new Date(f+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=v,document.getElementById("cal-event-items").innerHTML=l.map(S=>`
            <div class="cal-event-item cal-color-border-${S.color||"gray"}">
              <div class="cal-event-type">
                ${Pn(S.type)}
                ${S.date_type?` <span class="badge" style="font-size:0.75rem;padding:2px 6px;background:rgba(0,0,0,0.06);margin-left:4px;">${we(S.date_type)}</span>`:""}
              </div>
              <div class="cal-event-title" style="font-weight:600;font-size:0.95rem;margin-top:2px;">${we(S.title||"-")}</div>
              ${S.branch_name?`<div class="cal-event-branch" style="color:var(--text-2);font-size:0.85rem;margin-top:2px;">\u{1F4CD} ${we(S.branch_name)}</div>`:""}
              ${S.pic||S.period||S.opening_date||S.target_date||S.completion_date||S.notes?`
                <div class="cal-event-meta" style="margin-top:6px;padding-top:6px;border-top:1px dashed var(--border);font-size:0.8rem;color:var(--text-2);display:flex;flex-direction:column;gap:3px;">
                  ${S.pic?`<div>\u{1F464} <b>PIC:</b> ${we(S.pic)}</div>`:""}
                  ${S.period?`<div>\u{1F4C5} <b>Periode:</b> ${we(S.period)}</div>`:""}
                  ${S.opening_date?`<div>\u{1F6AA} <b>Tgl Opening:</b> ${we(Ot(S.opening_date))}</div>`:""}
                  ${S.target_date?`<div>\u{1F3AF} <b>Tgl Target:</b> ${we(Ot(S.target_date))}</div>`:""}
                  ${S.completion_date?`<div>\u2705 <b>Tgl Selesai:</b> ${we(Ot(S.completion_date))}</div>`:""}
                  ${S.notes?`<div>\u{1F4DD} <b>Catatan:</b> ${we(S.notes)}</div>`:""}
                </div>
              `:""}
              ${S.status?`<div class="cal-event-status" style="margin-top:6px;"><span class="badge ${S.status==="Done"?"badge-success":"badge-warning"}">${we(S.status)}</span></div>`:""}
              ${S.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${S.days_remaining} hari</div>`:""}
            </div>
          `).join(""),c.style.display="block"})})}catch(d){console.error("[Calendar] Render error:",d),u&&(u.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}let b=()=>{a().then(()=>o())};nt.on("data:changed",b),o()}function Dn(e){if(!e||e==="-"||String(e).trim()==="")return"";let t=String(e).trim();if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);let i=t.split(/[\/\-\.]/);if(i.length===3){let[a,o,b]=i.map(u=>u.trim());if(b.length===4)return`${b}-${o.padStart(2,"0")}-${a.padStart(2,"0")}`;if(a.length===4)return`${a}-${o.padStart(2,"0")}-${b.padStart(2,"0")}`}return t.slice(0,10)}function Ot(e){if(!e||e==="-"||String(e).trim()==="")return"";let i=String(e).trim().split(/[\/\-\.]/);if(i.length===3){let[a,o,b]=i.map(u=>u.trim());if(a.length===4)return`${b.padStart(2,"0")}-${o.padStart(2,"0")}-${a}`;if(b.length===4)return`${a.padStart(2,"0")}-${o.padStart(2,"0")}-${b}`}return e}function Bn(e,t){return e?e.length>t?e.slice(0,t)+"\u2026":e:""}function we(e){return e?String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Pn(e){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50D} Inspeksi",fogging:"\u{1F4A8} Fogging"}[e]||e}V();async function Pa(e){let t=ae(),i=(t?.full_name||t?.username||"U")[0].toUpperCase(),o={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[t?.role]||"#64748B";e.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${o},${o}99)">
            ${i}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${t?.full_name||"\u2014"}</div>
            <div class="profile-username">@${t?.username||"\u2014"}</div>
            <span class="badge badge-info" style="background:${o}18;color:${o};margin-top:6px">
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
            <span class="info-value" style="color:${o};font-weight:700">${t?.role||"\u2014"}</span>
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
  `;let b=localStorage.getItem("fm_token"),u=document.getElementById("session-info");if(b&&u)try{let d=JSON.parse(atob(b.split(".")[1])),n=new Date(d.exp*1e3);u.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{u.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async d=>{d.preventDefault();let n=document.getElementById("pwd-error"),h=document.getElementById("pwd-success"),F=document.getElementById("btn-save-pwd");n.style.display="none",h.style.display="none";let $=d.target,E=$.current_password.value,C=$.new_password.value,w=$.confirm_password.value;if(C!==w){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(C.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}F.disabled=!0,F.textContent="\u23F3 Menyimpan...";let k=await I("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:E,new_password:C})});F.disabled=!1,F.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',k.ok?(h.textContent="\u2705 Password berhasil diubah.",h.style.display="block",$.reset(),j("Password berhasil diubah.")):(n.textContent=k.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}V();var ut={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function le(e){if(e==null||e==="")return null;if(e instanceof Date)return isNaN(e.getTime())?null:e.toISOString().slice(0,10);let t=String(e).trim();if(t===""||t==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(t))return t.slice(0,10);if(/^\d{4,5}$/.test(t)){let o=Number(t);if(o>2e4&&o<99999){let b=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(b.getTime())?null:b.toISOString().slice(0,10)}}let i=t.split(/[\/\-\.]/);if(i.length===3){let[o,b,u]=i.map(F=>F.trim()),d=Number(o),n=Number(b),h=Number(u);if(o.length===4&&d>1900)return`${o}-${b.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&h>1900)return d>12?`${u}-${b.padStart(2,"0")}-${o.padStart(2,"0")}`:n>12?`${u}-${o.padStart(2,"0")}-${b.padStart(2,"0")}`:`${u}-${b.padStart(2,"0")}-${o.padStart(2,"0")}`;if(u.length===2&&!isNaN(h)){let F=h>=50?`19${u}`:`20${u}`;return d>12?`${F}-${b.padStart(2,"0")}-${o.padStart(2,"0")}`:`${F}-${b.padStart(2,"0")}-${o.padStart(2,"0")}`}}let a=new Date(t);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function La(e){return Object.values(e).every(t=>t==null||String(t).trim()==="")}var Ln={validation:{required:[],map:e=>({cabang:e.CABANG,pic:e.PIC,kegiatan:e.KEGIATAN,quartal:e.QUARTAL,masa_pkwt:e["MASA PKWT"],pic_pelapor:e["PIC PELAPOR"],kontrak:e.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:e=>({name:e["Nama SOP"],category:e.Kategori||"Umum",document_link:e["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({full_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||e.Divisi||"FACILITY CARE",phone:e["No. Hp"]||e["No. HP"],join_date:le(e["Tanggal Masuk"]||e["Tgl Masuk"]),status:e.Status||"",target_pindah_os:le(e["Target Pindah OS"]),target_selesai_os:le(e["Target Selesai OS"]),notes:e.Catatan||""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({employee_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||"FACILITY CARE",start_date:le(e["Tanggal Mulai"]),end_date:le(e["Tanggal Selesai"]),contract_type:e["Tipe Kontrak"]||"",pkwt_number:e.PKWT||"",status:e.Status||"",notes:e.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:e=>({report_date:le(e["Tanggal Info"]),branch_name:e.Cabang,category:e.Kategori,source:e["Sumber Laporan"],complaint:e.Keluhan,employee_name:e["Nama FC"],fc_specialist:e["FC Spesialis"],solution:e.Solusi,status:e.Status||"",completion_date:le(e["Tanggal Selesai"])})},one_on_one:{required:[],map:e=>({meeting_date:le(e.Tanggal),branch_name:e.Cabang,employee_name:e["Nama Karyawan"],pic:e.Pic,problem:e.Masalah,solution:e.Solusi,status:e.Status||"",completion_date:le(e["Tanggal Selesai"]),document_link:e["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:e=>({branch_name:e.Cabang,activity_type:e.Kegiatan,period:e.Periode,pic:e.Pic||e.PIC,opening_date:le(e["Tanggal Opening"]||e["Tgl Opening"]),target_date:le(e["Tanggal Target"]||e["Tgl Target"]),completion_date:le(e["Tanggal Selesai"]||e["Tgl Selesai"]),status:e.Status||"",notes:e.Keterangan||e.Catatan})},inspection:{required:[],map:e=>({inspection_date:le(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",fc_score:e["Point FC SP"]!==void 0&&e["Point FC SP"]!==null?parseFloat(String(e["Point FC SP"]).replace(",",".")):null,spv_score:e["Point SPV"]!==void 0&&e["Point SPV"]!==null?parseFloat(String(e["Point SPV"]).replace(",",".")):null,document_link:e.Link,notes:""})},cleaning:{required:[],map:e=>({activity_date:le(e.Tanggal),branch_name:e.Cabang,activity_type:e["Jenis Kegiatan"]||"General Cleaning",period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},fogging:{required:[],map:e=>({activity_date:le(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:e=>({info_date:le(e["Tgl Info"]),branch_name:e.Cabang,problem:e.Permasalahan,pic:e.PIC,done_date:le(e["Tgl Done"]),status:e.Status||"",notes:e.Ket})},relievers:{required:[],map:e=>({branch_name:e.Cabang,original_fc_name:e["Nama Facility care"],period:e.Periode,reliever_name:e.Relifer,backup_date:le(e["Tanggal Back Up"]),completion_date:le(e["Tanggal Selesai"]),reason:e.Keterangan,shift:e.Shift,status:e.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:e=>({training_date:le(e.Tanggal),batch:e.Batch,subject:e.Materi,participants:e.Peserta,branch_name:e.Cabang,trainer:e.Trainer,score:e.Nilai!==void 0&&e.Nilai!==null?parseFloat(String(e.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:e=>({name:e["Master Checklist"],category:"Umum",document_link:e["Link Document"],description:""})},supply:{required:[],map:e=>({submitted_at:le(e.Timestamp),submitter_name:e["Nama Lengkap"],branch_name:e["Kebutuhan Untuk Cabang"],tools_items:e["Alat - Alat / Barang"],tools_quantity:e["Jumlah Permintaan Alat / Barang"],chemical_items:e.Chemical,chemical_quantity:e["Jumlah Permintaan Chemical"],additional_notes:e["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:e.Status||""})}};function In(e,t){let i=ut[e];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Ln[i.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let o=[],b=[],u=[];return t.filter(n=>!La(n)).forEach((n,h)=>{let F=t.indexOf(n)+2,$=[];a.required.forEach(({key:C,label:w})=>{let k=n[C];if(k==null||String(k).trim()===""){let y=Object.keys(n).filter(g=>g.trim()).join(", ");$.push({column:w,originalValue:k||"",reason:`Kolom "${w}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${y.slice(0,120)}`})}});let E=a.map(n);$.length>0?b.push({row:F,data:E,raw:n,errors:$}):(o.push(n),u.push(E))}),{valid:o,errors:b,mapped:u}}function Ia(e){let t=[];return e.SheetNames.forEach(i=>{let a=ut[i];if(!a)return;let o=e.Sheets[i],b=window.XLSX.utils.sheet_to_json(o,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),u=In(i,b),d=b.filter(n=>!La(n));t.push({sheetName:i,module:a.module,label:a.label,total:d.length,valid:u.mapped.length,errorCount:u.errors.length,errors:u.errors,mapped:u.mapped,skipped:!1})}),t}function Ma(){let e=window.XLSX,t=e.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,o])=>{e.utils.book_append_sheet(t,e.utils.json_to_sheet(o),a)}),e.writeFile(t,"Template_Import_Data_Awal_FCMS.xlsx")}function za(e){let t=window.XLSX,i=t.utils.book_new(),a=!1;return e.forEach(o=>{if(!o.errors||o.errors.length===0)return;a=!0;let b=o.errors.map(d=>({"No. Baris":d.row,"Kolom Gagal":(d.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(d.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(d.data||{}).map(([n,h])=>[n,h??""]))})),u=t.utils.json_to_sheet(b);t.utils.book_append_sheet(i,u,o.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(t.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Mn=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","sop","inspection","cleaning","fogging","basecamp","supply"];function Na(e){e.innerHTML=`
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
              ${Object.entries(ut).map(([k,{label:y}])=>`<span class="import-sheet-tag">\u{1F4C4} ${k} \u2192 ${y}</span>`).join("")}
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
  `;let t=null,i=null,a=0,o={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function b(k){Object.entries(o).forEach(([y,g])=>{g.style.display=y===k?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let k=document.getElementById("btn-backup-db");k.disabled=!0,k.textContent="\u23F3 Memproses Backup...";try{let y=await I("/api/import/backup");if(y.ok){if(!window.XLSX){N("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let g=window.XLSX,r=g.utils.book_new();Object.entries(y.data.database).forEach(([s,m])=>{let p=m.length>0?m:[{}],f=g.utils.json_to_sheet(p);g.utils.book_append_sheet(r,f,s.substring(0,31))}),g.writeFile(r,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),j("Backup berhasil diunduh!")}else N("Gagal memproses backup: "+(y.data?.error||"Unknown error"))}catch(y){N("Gagal memproses backup: "+y.message)}finally{k.disabled=!1,k.textContent="\u{1F4E6} Backup Database"}});let u=document.getElementById("btn-sync-google");u&&u.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let k=u.innerHTML;u.innerHTML='<span class="spinner"></span> Menyinkronkan...',u.disabled=!0;try{let y=await I("/api/sync/google-sheets",{method:"POST"});y.ok?alert("Sinkronisasi Berhasil: "+(y.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(y.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{u.innerHTML=k,u.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Ma(),j("Template Excel berhasil didownload!")});let d=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",k=>{k.stopPropagation(),d.click()}),d.addEventListener("change",k=>{k.target.files[0]&&h(k.target.files[0])}),n.addEventListener("dragover",k=>{k.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",k=>{k.preventDefault(),n.classList.remove("drag-over");let y=k.dataTransfer.files[0];y&&y.name.match(/\.xlsx?$/i)?h(y):N("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{t=null,d.value="",document.getElementById("file-info").style.display="none",n.style.display="",b("upload")});async function h(k){t=k,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${k.name} (${(k.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await F(k)}async function F(k){b("validating");let y=document.getElementById("validation-status"),g=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");y.textContent="Membaca file Excel...",g.style.width="20%",await Ye(200);let r=await k.arrayBuffer(),s=window.XLSX.read(r,{type:"array",cellDates:!0});y.textContent=`Memvalidasi ${s.SheetNames.length} sheet...`,g.style.width="50%",await Ye(100),i=Ia(s),g.style.width="100%",y.textContent="Validasi selesai!",await Ye(300),$()}catch(r){b("upload"),N("Gagal memproses file: "+r.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function $(){b("preview");let k=i.filter(c=>!c.skipped).length,y=i.reduce((c,v)=>c+v.total,0),g=i.reduce((c,v)=>c+v.valid,0),r=i.reduce((c,v)=>c+v.errorCount,0),s=y>0?Math.round(g/y*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${k} sheet</span>
      <span class="badge badge-secondary">${y} baris</span>
      <span class="badge badge-success">${g} valid (${s}%)</span>
      ${r>0?`<span class="badge badge-danger">${r} error</span>`:""}
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
          ${i.map((c,v)=>`
            <tr class="${c.errorCount>0?"row-error":c.skipped?"row-skipped":"row-ok"}">
              <td><strong>${c.sheetName}</strong></td>
              <td>${c.label}</td>
              <td style="text-align:center">${c.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${c.valid}</span></td>
              <td style="text-align:center">${c.errorCount>0?`<span class="badge badge-danger">${c.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${c.skipped?'<span class="badge badge-neutral">Dilewati</span>':c.errorCount>0&&c.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':c.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':c.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${c.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${v}">\u{1F50D} ${c.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,m.querySelectorAll(".btn-detail-error").forEach(c=>{c.addEventListener("click",()=>{let v=i[Number(c.dataset.idx)];E(v)})});let p=document.getElementById("error-detail-section"),f=document.getElementById("error-detail-container");f.innerHTML="",p.style.display="none";let l=document.getElementById("btn-start-import");g===0?(l.disabled=!0,l.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(l.disabled=!1,r>0?(l.innerHTML=`\u{1F680} Import ${g} Data Valid (${r} dilewati)`,l.title="Baris error akan dilewati, baris valid tetap diimport"):l.innerHTML=`\u{1F680} Mulai Import ${g} Data`)}function E(k){let y=document.getElementById("error-detail-section"),g=document.getElementById("error-detail-container");y.style.display="";let r=k.errors.slice(0,100).map(s=>(Array.isArray(s.errors)?s.errors:[]).map(p=>{let f=typeof p=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${s.row}</span></td>
            <td><strong>${f?p.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${f&&p.originalValue!==void 0?p.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${f?p.reason:p}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${f&&p.aliases?`Gunakan salah satu nama kolom:<br><em>${p.aliases}</em>`:f&&p.hint?p.hint:""}
            </td>
          </tr>
        `}).join("")).join("");g.innerHTML=`
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
            <tbody>${r||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${k.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,y.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{b("upload"),document.getElementById("file-info").style.display="none",n.style.display="",t=null,d.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;za(i)?j("Log error berhasil didownload."):j("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let k=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";C(k)});async function C(k){b("importing"),a=Date.now();let y=[];Mn.forEach(p=>{let f=i?.find(l=>l.module===p&&l.mapped?.length>0);f&&y.push(f)});let g=document.getElementById("import-steps-list");g.innerHTML=y.map(p=>`
      <div class="import-step-item" id="step-item-${p.module}">
        <span class="step-item-icon" id="step-icon-${p.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${p.label} <span class="step-item-count">(${p.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${p.module}"></span>
      </div>
    `).join("");let r=document.getElementById("import-bar"),s=document.getElementById("import-current-status"),m={totalSheets:y.length,totalRows:y.reduce((p,f)=>p+f.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let p=0;p<y.length;p++){let f=y[p],l=document.getElementById(`step-icon-${f.module}`),c=document.getElementById(`step-status-${f.module}`);l.textContent="\u{1F504}",c.textContent="Mengimport...",s.textContent=`Mengimport ${f.label}...`,r.style.width=`${Math.round(p/y.length*100)}%`;try{let v=await I(`/api/import/${f.module}`,{method:"POST",body:JSON.stringify({rows:f.mapped,onDuplicate:k})});if(v.ok){let S=v.data;m.inserted+=S.inserted||0,m.skipped+=S.skipped||0,m.moduleResults.push({label:f.label,inserted:S.inserted||0,skipped:S.skipped||0,status:"ok"}),l.textContent="\u2705",c.innerHTML=`<span class="badge badge-success">${S.inserted||0} berhasil</span>${S.skipped>0?` <span class="badge badge-neutral">${S.skipped} skip</span>`:""}`}else m.failed++,m.moduleResults.push({label:f.label,inserted:0,skipped:0,status:"error",error:v.data?.error}),l.textContent="\u274C",c.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(v){m.failed++,m.moduleResults.push({label:f.label,inserted:0,skipped:0,status:"error",error:v.message}),l.textContent="\u274C",c.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Ye(150)}r.style.width="100%",s.textContent="Selesai!",await Ye(400),w(m)}function w(k){b("summary");let y=((Date.now()-a)/1e3).toFixed(1),g=k.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${g?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${g?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
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
        <div class="stat-value">${y}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${k.moduleResults.map(r=>`
            <tr>
              <td>${r.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${r.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${r.skipped}</span></td>
              <td style="text-align:center">
                ${r.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${r.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{t=null,i=null,d.value="",document.getElementById("file-info").style.display="none",n.style.display="",b("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Ye(e){return new Promise(t=>setTimeout(t,e))}V();ue();var ja=[{id:1,employee_name:"Teguh Ariyandi",division:"FACILITY CARE",branch_id:18,branch_name:"018. Harapan Indah",tanggal:"2025-08-19",akhir_sp:"2026-02-18",sp_type:"SP 1",status:"Aktif",document_link:"",notes:"Keterlambatan berulang dan ketidaklengkapan checklist harian lantai dasar."},{id:2,employee_name:"Andri",division:"FACILITY CARE",branch_id:1,branch_name:"001. Pondok Bambu",tanggal:"2025-04-10",akhir_sp:"2025-10-09",sp_type:"SP 1",status:"Selesai",document_link:"",notes:"Pelanggaran grooming dan standar pembersihan area tindakan tanpa APD lengkap."},{id:3,employee_name:"M. FIRMANSYAH",division:"FACILITY CARE",branch_id:25,branch_name:"025. Bogor Baru",tanggal:"2024-03-21",akhir_sp:"2024-09-20",sp_type:"SP 2",status:"Selesai",document_link:"",notes:"Meninggalkan area tugas tanpa konfirmasi supervisor selama jam operasional klinik."}];function Ne(e){if(!e)return{status:"Tidak Ditentukan",label:"Tidak Ada Tanggal",bg:"#F1F5F9",color:"#64748B",border:"#CBD5E1",icon:"\u26AA",daysLeft:0,isExpired:!1,isExpiringSoon:!1};let t=new Date;t.setHours(0,0,0,0);let i=new Date(e);i.setHours(0,0,0,0);let a=i.getTime()-t.getTime(),o=Math.ceil(a/(1e3*60*60*24));return o<0?{status:"Selesai",label:`Masa Sanksi Selesai (${Math.abs(o)} hari lalu)`,shortLabel:"Selesai",bg:"#DCFCE7",color:"#15803D",border:"#BBF7D0",icon:"\u2705",daysLeft:o,isExpired:!0,isExpiringSoon:!1}:o<=30?{status:"Mendekati Berakhir",label:`Segera Berakhir (Sisa ${o} hari)`,shortLabel:`Sisa ${o} hari`,bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F",daysLeft:o,isExpired:!1,isExpiringSoon:!0}:{status:"Aktif",label:`Aktif Berlaku (Sisa ${o} hari)`,shortLabel:`Aktif (${o} hari)`,bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F534}",daysLeft:o,isExpired:!1,isExpiringSoon:!1}}function Ka(e=""){let t=(e||"").toUpperCase().trim();return t==="SP 3"?{label:"SP 3 (Terakhir)",bg:"#7F1D1D",color:"#FFFFFF",border:"#991B1B",icon:"\u{1F6A8}"}:t==="SP 2"?{label:"SP 2",bg:"#FEE2E2",color:"#B91C1C",border:"#FECACA",icon:"\u{1F6D1}"}:t==="SP 1"?{label:"SP 1",bg:"#FEF3C7",color:"#B45309",border:"#FDE68A",icon:"\u26A0\uFE0F"}:t.includes("TEGURAN")?{label:"Teguran Lisan",bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD",icon:"\u{1F4AC}"}:{label:e||"SP",bg:"#F1F5F9",color:"#475569",border:"#CBD5E1",icon:"\u{1F4DC}"}}function gt(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Oa(e){let t=[],i=[],a=[],o="all",b="all",u="all",d="",n="fcms-sp-page-styles";if(!document.getElementById(n)){let r=document.createElement("style");r.id=n,r.textContent=`
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
        font-size: 1.65rem;
        font-weight: 700;
        margin: 0 0 6px 0;
        display: flex;
        align-items: center;
        gap: 12px;
        letter-spacing: -0.02em;
      }
      .sp-hero-subtitle {
        color: #94A3B8;
        font-size: 0.92rem;
        margin: 0;
        max-width: 720px;
        line-height: 1.5;
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
          gap: 8px !important;
          flex-wrap: nowrap !important;
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
          grid-template-columns: 1fr !important;
          gap: 12px !important;
        }
        .sp-card {
          padding: 16px !important;
          border-radius: 14px !important;
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
    `,document.head.appendChild(r)}async function h(){try{let[r,s]=await Promise.all([W(),ce()]);i=r||[],a=s||[];let m=await I("/api/sp?limit=1000");if(m.ok&&m.data?.data&&Array.isArray(m.data.data)&&m.data.data.length>0)t=m.data.data,localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t));else{let p=localStorage.getItem("fcms_sp_data_cache");p?t=JSON.parse(p):(t=[...ja],localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t)))}}catch(r){console.warn("Fallback loading SP data from cache/seed:",r);let s=localStorage.getItem("fcms_sp_data_cache");t=s?JSON.parse(s):[...ja]}y()}function F(){return t.filter(r=>{let s=d.toLowerCase().trim(),m=!s||(r.employee_name||"").toLowerCase().includes(s)||(r.branch_name||"").toLowerCase().includes(s)||(r.division||"").toLowerCase().includes(s)||(r.sp_type||"").toLowerCase().includes(s),p=b==="all"||String(r.branch_id)===String(b),f=u==="all"||r.division===u,l=Ne(r.akhir_sp),c=!0;return o==="active"?c=!l.isExpired:o==="expiring"?c=l.isExpiringSoon:o==="expired"?c=l.isExpired:o==="sp1"?c=(r.sp_type||"").toUpperCase()==="SP 1":o==="sp2"?c=(r.sp_type||"").toUpperCase()==="SP 2":o==="sp3"?c=(r.sp_type||"").toUpperCase()==="SP 3":o==="teguran"&&(c=(r.sp_type||"").toUpperCase().includes("TEGURAN")),m&&p&&f&&c})}function $(){let r=0,s=0,m=0,p=0,f=0,l=0,c=0,v=new Set;return t.forEach(S=>{S.employee_name&&v.add(S.employee_name.toLowerCase().trim());let B=Ne(S.akhir_sp);B.isExpired?m++:r++,B.isExpiringSoon&&s++;let T=(S.sp_type||"").toUpperCase();T==="SP 1"?p++:T==="SP 2"?f++:T==="SP 3"?l++:T.includes("TEGURAN")&&c++}),{total:t.length,active:r,expiring:s,expired:m,sp1:p,sp2:f,sp3:l,teguran:c,uniqueEmployees:v.size}}function E(r=null){let s=!!r,m=r?.tanggal?r.tanggal.slice(0,10):new Date().toISOString().slice(0,10),p=r?.akhir_sp?r.akhir_sp.slice(0,10):"";if(!p){let v=new Date(m);v.setMonth(v.getMonth()+6),p=v.toISOString().slice(0,10)}let f=document.createElement("div");f.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <input type="text" id="sp-modal-emp" list="sp-emp-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih atau ketik nama..." value="${r?.employee_name?String(r.employee_name).replace(/"/g,"&quot;"):""}" required />
            <datalist id="sp-emp-list">
              ${a.map(v=>`<option value="${v.label}">${v.label}</option>`).join("")}
            </datalist>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-div" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="FACILITY CARE" ${r?.division==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${r?.division==="SECURITY"?"selected":""}>SECURITY</option>
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
              ${i.map(v=>`<option value="${v.value}" ${String(r?.branch_id)===String(v.value)?"selected":""}>${v.label}</option>`).join("")}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tingkat / Jenis SP <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="SP 1" ${r?.sp_type==="SP 1"?"selected":""}>SP 1 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 2" ${r?.sp_type==="SP 2"?"selected":""}>SP 2 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 3" ${r?.sp_type==="SP 3"?"selected":""}>SP 3 (Masa Berlaku 6 Bulan / Terakhir)</option>
              <option value="Teguran Lisan" ${r?.sp_type==="Teguran Lisan"?"selected":""}>Teguran Lisan (Pembinaan)</option>
            </select>
          </div>
        </div>

        <div class="sp-form-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Terbit SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${m}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Akhir Masa Berlaku SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-akhir" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${p}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Link Berkas / Foto BAP / Scan Surat SP
          </label>
          <input type="url" id="sp-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://drive.google.com/..." value="${r?.document_link?String(r.document_link).replace(/"/g,"&quot;"):""}" />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Kronologi Pelanggaran / Catatan Pembinaan
          </label>
          <textarea id="sp-modal-notes" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan ringkasan pelanggaran disiplin kerja yang terjadi...">${r?.notes||r?.description||""}</textarea>
        </div>
      </div>
    `;let l=f.querySelector("#sp-modal-tanggal"),c=f.querySelector("#sp-modal-akhir");l.addEventListener("change",()=>{if(l.value){let v=new Date(l.value);v.setMonth(v.getMonth()+6),c.value=v.toISOString().slice(0,10)}}),Y({title:s?"\u270F\uFE0F Edit Data Surat Peringatan (SP)":"\u2795 Terbitkan Surat Peringatan (SP) Baru",content:f,size:"md",confirmText:s?"Simpan Perubahan":"Terbitkan SP",confirmClass:"btn-primary",onConfirm:async(v,S)=>{let B=(f.querySelector("#sp-modal-emp").value||"").trim(),T=f.querySelector("#sp-modal-div").value,D=f.querySelector("#sp-modal-branch").value,_=f.querySelector("#sp-modal-type").value,P=f.querySelector("#sp-modal-tanggal").value,M=f.querySelector("#sp-modal-akhir").value,z=(f.querySelector("#sp-modal-link").value||"").trim(),x=(f.querySelector("#sp-modal-notes").value||"").trim();if(!B){N("Nama Karyawan wajib diisi");return}if(!D){N("Cabang penempatan wajib dipilih");return}if(!P||!M){N("Tanggal SP dan Akhir SP wajib diisi");return}let A=i.find(O=>String(O.value)===String(D)),L=A?A.label:"Cabang",q={employee_name:B,division:T,branch_id:Number(D),branch_name:L,sp_type:_,tanggal:P,akhir_sp:M,document_link:z,notes:x,status:Ne(M).status};try{if(s){let O=await I(`/api/sp/${r.id}`,{method:"PUT",body:JSON.stringify(q)}),H=t.findIndex(K=>K.id===r.id);H!==-1&&(t[H]={...t[H],...q}),j("Data SP berhasil diperbarui")}else{let O=await I("/api/sp",{method:"POST",body:JSON.stringify(q)}),H=O.ok&&O.data?.id?O.data.id:Date.now();t.unshift({id:H,...q}),j("Surat Peringatan baru berhasil diterbitkan")}localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t)),S(),y()}catch(O){console.error("Error saving SP:",O),N("Gagal menyimpan SP: "+O.message)}}})}function C(r){ee(`Hapus data Surat Peringatan untuk karyawan <strong>${r.employee_name}</strong> (${r.sp_type})?`,async()=>{try{await I(`/api/sp/${r.id}`,{method:"DELETE"})}catch(s){console.warn("API delete error fallback to local:",s)}t=t.filter(s=>s.id!==r.id),localStorage.setItem("fcms_sp_data_cache",JSON.stringify(t)),j(`Data SP ${r.employee_name} berhasil dihapus`),y()},"Hapus Data SP")}function w(r){let s=Ne(r.akhir_sp),m=Ka(r.sp_type),p=document.createElement("div");p.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div style="display:flex; align-items:center; gap:14px; padding-bottom:14px; border-bottom:1px solid #E2E8F0;">
          <div class="sp-avatar-box" style="width:54px; height:54px; font-size:1.2rem;">
            ${r.employee_name.slice(0,2).toUpperCase()}
          </div>
          <div>
            <h3 style="margin:0 0 4px 0; font-size:1.15rem; font-weight:700; color:#0F172A;">
              ${r.employee_name}
            </h3>
            <div style="display:flex; align-items:center; gap:8px;">
              <span class="badge" style="background:#EFF6FF; color:#1D4ED8; font-weight:700;">${r.division||"FACILITY CARE"}</span>
              <span style="color:#64748B; font-size:0.85rem;">\u{1F4CD} ${r.branch_name||"Cabang"}</span>
            </div>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Tingkat Sanksi</div>
            <span class="sp-type-badge" style="background:${m.bg}; color:${m.color}; border:1px solid ${m.border}; font-size:0.85rem;">
              <span>${m.icon}</span>
              <span>${m.label}</span>
            </span>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Status Masa Berlaku</div>
            <span class="sp-status-badge" style="background:${s.bg}; color:${s.color}; border:1px solid ${s.border}; font-size:0.85rem;">
              <span>${s.icon}</span>
              <span>${s.label}</span>
            </span>
          </div>
        </div>

        <div class="sp-detail-grid" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Tanggal Terbit SP</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${gt(r.tanggal)}</div>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Akhir Masa Berlaku</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${gt(r.akhir_sp)}</div>
          </div>
        </div>

        ${r.notes||r.description?`
          <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:10px; padding:14px;">
            <div style="font-size:0.78rem; font-weight:700; color:#92400E; margin-bottom:4px;">\u{1F4DD} Catatan Pelanggaran / Pembinaan:</div>
            <div style="font-size:0.88rem; color:#78350F; line-height:1.5;">${r.notes||r.description}</div>
          </div>
        `:""}

        ${r.document_link?`
          <div>
            <a href="${r.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-primary" style="width:100%; justify-content:center; padding:10px; font-weight:600; display:flex; align-items:center; gap:8px;">
              <span>\u{1F4C4}</span> Buka Dokumen BAP / Scan SP Fisik
            </a>
          </div>
        `:`
          <div style="text-align:center; padding:8px; font-size:0.82rem; color:#94A3B8;">
            (Belum ada scan dokumen / BAP yang diunggah)
          </div>
        `}
      </div>
    `,Y({title:"\u{1F4CB} Lembar Riwayat Surat Peringatan (SP)",content:p,size:"md",cancelText:"Tutup"})}async function k(){try{let{downloadExcel:r}=await Promise.resolve().then(()=>(X(),Se)),s=t.map((m,p)=>{let f=Ne(m.akhir_sp);return{No:p+1,"Nama Karyawan":m.employee_name||"",Divisi:m.division||"",Cabang:m.branch_name||"","Jenis SP":m.sp_type||"","Tanggal SP":m.tanggal||"","Akhir SP":m.akhir_sp||"","Status Masa Berlaku":f.status,"Sisa Hari":f.daysLeft>0?f.daysLeft:0,"Link BAP / Foto":m.document_link||"","Catatan / Kronologi":m.notes||m.description||""}});r(s,`Data_Surat_Peringatan_FCMS_${new Date().toISOString().slice(0,10)}`),j("Berhasil mengunduh data SP ke format Excel.")}catch(r){N("Gagal ekspor: "+r.message)}}function y(){let r=F(),s=$();e.innerHTML=`
      <div class="sp-page-container">
        
        <!-- Hero Header -->
        <div class="sp-hero-card">
          <div class="sp-hero-header">
            <div>
              <h1 class="sp-hero-title">
                <span>\u{1F4DC}</span> Monitoring Data SP (Surat Peringatan)
              </h1>
              <p class="sp-hero-subtitle">
                Sistem pengawasan kedisiplinan kerja, masa aktif sanksi pembinaan, dan arsip dokumen BAP karyawan PT Audy Mandiri Indonesia.
              </p>
            </div>
            <div class="sp-hero-actions" style="display:flex; align-items:center; gap:8px;">
              <button id="sp-btn-add" class="btn btn-primary sp-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                <span>Terbitkan SP</span>
              </button>
              
              <button id="sp-btn-excel" class="btn btn-ghost sp-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="sp-btn-refresh" class="btn btn-ghost sp-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- Quick Metrics -->
          <div class="sp-stats-grid">
            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(239,68,68,0.2); color:#F87171;">\u{1F534}</div>
              <div>
                <div class="sp-stat-value">${s.active}</div>
                <div class="sp-stat-label">SP Aktif Berlaku</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">\u26A0\uFE0F</div>
              <div>
                <div class="sp-stat-value">${s.expiring}</div>
                <div class="sp-stat-label">Segera Berakhir (&lt;30 Hari)</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">\u2705</div>
              <div>
                <div class="sp-stat-value">${s.expired}</div>
                <div class="sp-stat-label">Masa Sanksi Selesai</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">\u{1F465}</div>
              <div>
                <div class="sp-stat-value">${s.uniqueEmployees} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">karyawan</span></div>
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
              <input type="text" id="sp-search-input" class="sp-search-input" placeholder="Cari nama karyawan, cabang, atau divisi..." value="${d?d.replace(/"/g,"&quot;"):""}" />
            </div>

            <!-- Branch Dropdown -->
            <select id="sp-filter-branch" class="sp-filter-select">
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(m=>`<option value="${m.value}" ${String(b)===String(m.value)?"selected":""}>${m.label}</option>`).join("")}
            </select>

            <!-- Division Dropdown -->
            <select id="sp-filter-division" class="sp-filter-select">
              <option value="all" ${u==="all"?"selected":""}>Semua Divisi</option>
              <option value="FACILITY CARE" ${u==="FACILITY CARE"?"selected":""}>FACILITY CARE</option>
              <option value="SECURITY" ${u==="SECURITY"?"selected":""}>SECURITY</option>
            </select>
          </div>

          <!-- Filter Pills -->
          <div class="sp-filter-pills">
            <button class="sp-pill ${o==="all"?"active":""}" data-filter="all">
              <span>Semua SP</span>
              <span class="sp-pill-count">${s.total}</span>
            </button>
            <button class="sp-pill ${o==="active"?"active":""}" data-filter="active">
              <span>\u{1F534} SP Aktif</span>
              <span class="sp-pill-count">${s.active}</span>
            </button>
            <button class="sp-pill ${o==="expiring"?"active":""}" data-filter="expiring">
              <span>\u26A0\uFE0F &lt;30 Hari</span>
              <span class="sp-pill-count">${s.expiring}</span>
            </button>
            <button class="sp-pill ${o==="expired"?"active":""}" data-filter="expired">
              <span>\u2705 Selesai</span>
              <span class="sp-pill-count">${s.expired}</span>
            </button>
            <button class="sp-pill ${o==="sp1"?"active":""}" data-filter="sp1">
              <span>SP 1</span>
              <span class="sp-pill-count">${s.sp1}</span>
            </button>
            <button class="sp-pill ${o==="sp2"?"active":""}" data-filter="sp2">
              <span>SP 2</span>
              <span class="sp-pill-count">${s.sp2}</span>
            </button>
            <button class="sp-pill ${o==="sp3"?"active":""}" data-filter="sp3">
              <span>SP 3</span>
              <span class="sp-pill-count">${s.sp3}</span>
            </button>
          </div>
        </div>

        <!-- Content Area -->
        <div id="sp-list-container">
          ${r.length===0?`
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
              ${r.map(m=>{let p=Ne(m.akhir_sp),f=Ka(m.sp_type);return`
                  <div class="sp-card" data-id="${m.id}">
                    <div>
                      <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px;">
                        <span class="sp-type-badge" style="background:${f.bg}; color:${f.color}; border:1px solid ${f.border};">
                          <span>${f.icon}</span>
                          <span>${f.label}</span>
                        </span>
                        <span class="sp-status-badge" style="background:${p.bg}; color:${p.color}; border:1px solid ${p.border};">
                          <span>${p.icon}</span>
                          <span>${p.shortLabel}</span>
                        </span>
                      </div>

                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="sp-avatar-box" style="width:48px; height:48px; font-size:1.05rem;">
                          ${m.employee_name.slice(0,2).toUpperCase()}
                        </div>
                        <div>
                          <h3 style="margin:0 0 3px 0; font-size:1.1rem; font-weight:700; color:#0F172A;">
                            ${m.employee_name}
                          </h3>
                          <div style="font-size:0.8rem; color:#64748B;">
                            \u{1F4CD} ${m.branch_name||"Cabang"} \u2022 <span style="color:#2563EB; font-weight:600;">${m.division||"FC"}</span>
                          </div>
                        </div>
                      </div>

                      <div style="background:#F8FAFC; border-radius:10px; padding:10px 12px; margin-bottom:14px; font-size:0.82rem; display:flex; justify-content:space-between;">
                        <div>
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Tanggal Terbit</div>
                          <div style="font-weight:700; color:#1E293B;">${gt(m.tanggal)}</div>
                        </div>
                        <div style="text-align:right;">
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Akhir SP</div>
                          <div style="font-weight:700; color:#0F172A;">${gt(m.akhir_sp)}</div>
                        </div>
                      </div>

                      <p style="font-size:0.83rem; color:#64748B; line-height:1.45; margin:0 0 16px 0; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">
                        ${m.notes||m.description||"Pencatatan sanksi peringatan disiplin operasional."}
                      </p>
                    </div>

                    <div style="padding-top:12px; border-top:1px solid #F1F5F9; display:flex; align-items:center; justify-content:space-between;">
                      <button class="btn btn-sm btn-ghost sp-detail-btn" style="border:1px solid #CBD5E1; border-radius:8px; font-weight:600; font-size:0.82rem; padding:6px 12px;" data-id="${m.id}">
                        \u{1F441}\uFE0F Rincian BAP
                      </button>

                      <div style="display:flex; align-items:center; gap:6px;">
                        ${m.document_link?`
                          <a href="${m.document_link}" target="_blank" rel="noopener noreferrer" class="sp-icon-btn" title="Buka Dokumen">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                          </a>
                        `:""}
                        <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${m.id}">
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${m.id}">
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
    `,g()}function g(){let r=e.querySelector("#sp-btn-add");r&&(r.onclick=()=>E());let s=e.querySelector("#sp-btn-excel");s&&(s.onclick=()=>k());let m=e.querySelector("#sp-btn-refresh");m&&(m.onclick=()=>h());let p=e.querySelector("#sp-btn-reset-filter");p&&(p.onclick=()=>{d="",b="all",u="all",o="all",y()});let f=e.querySelector("#sp-search-input");f&&(f.oninput=v=>{d=v.target.value,y();let S=e.querySelector("#sp-search-input");S&&(S.focus(),S.setSelectionRange(S.value.length,S.value.length))});let l=e.querySelector("#sp-filter-branch");l&&(l.onchange=v=>{b=v.target.value,y()});let c=e.querySelector("#sp-filter-division");c&&(c.onchange=v=>{u=v.target.value,y()}),e.querySelectorAll(".sp-pill").forEach(v=>{v.onclick=()=>{o=v.getAttribute("data-filter")||"all",y()}}),e.querySelectorAll(".sp-detail-btn, .sp-view-detail-link").forEach(v=>{v.onclick=S=>{S.stopPropagation();let B=Number(v.getAttribute("data-id")),T=t.find(D=>D.id===B);T&&w(T)}}),e.querySelectorAll(".sp-edit-btn").forEach(v=>{v.onclick=S=>{S.stopPropagation();let B=Number(v.getAttribute("data-id")),T=t.find(D=>D.id===B);T&&E(T)}}),e.querySelectorAll(".sp-delete-btn").forEach(v=>{v.onclick=S=>{S.stopPropagation();let B=Number(v.getAttribute("data-id")),T=t.find(D=>D.id===B);T&&C(T)}})}await h()}V();ue();var Rt=[{id:1,tanggal:"2026-10-02",employee_name:"TAUFIK HIDAYAT",from_branch_id:12,from_branch_name:"012. BSD",to_branch_id:0,to_branch_name:"000. Management",status:"Selesai",document_link:""},{id:2,tanggal:"2026-09-30",employee_name:"ELVAN JANGGA PRAMANA",from_branch_id:55,from_branch_name:"055. Gajah Mada",to_branch_id:0,to_branch_name:"000. Management",status:"Proses",document_link:""},{id:3,tanggal:"2026-06-23",employee_name:"Widya Astuti",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:7,to_branch_name:"007. Bekasi",status:"Selesai",document_link:""},{id:4,tanggal:"2026-06-22",employee_name:"Gema Azan Saputra",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:21,to_branch_name:"021. Kelapa Gading Boulevard",status:"Selesai",document_link:""},{id:5,tanggal:"2026-06-22",employee_name:"Ahmad Farhan",from_branch_id:1,from_branch_name:"001. Pondok Bambu",to_branch_id:62,to_branch_name:"062. Summarecon Bekasi",status:"Selesai",document_link:""}];function zn(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Ra(e){let t=[],i=[],a=[],o="all",b="all",u="all",d="",n=!1,F=ae()?.role!=="viewer";try{[i,a]=await Promise.all([W(),ce()])}catch(r){console.error("Error fetching cache:",r)}async function $(){n=!0,y();try{let r=await I("/api/mutasi?limit=10000");if(r.ok){let s=r.data?.data||r.data||[];t=Array.isArray(s)&&s.length>0?s:Rt}else t=Rt}catch(r){console.warn("Gagal load mutasi, menggunakan data lokal:",r),t=Rt}finally{n=!1,y()}}async function E(){try{let{downloadExcel:r}=await Promise.resolve().then(()=>(X(),Se)),s=t.map((m,p)=>({No:p+1,Tanggal:m.tanggal||"","Nama Karyawan":m.employee_name||"","Cabang Asal":m.from_branch_name||"","Cabang Tujuan":m.to_branch_name||"",Status:m.status||"","Link Dokumen":m.document_link||""}));r(s,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`),j("Data mutasi berhasil diekspor ke Excel!")}catch(r){N("Gagal ekspor Excel: "+r.message)}}async function C(r){if(!F)return;let s=t.find(m=>m.id===r);ee({title:"Hapus Data Mutasi",message:`Apakah Anda yakin ingin menghapus data mutasi untuk "${s?.employee_name||"karyawan ini"}"?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let m=await I(`/api/mutasi/${r}`,{method:"DELETE"});m.ok?(j("Data mutasi berhasil dihapus"),t=t.filter(p=>p.id!==r),y()):N(m.data?.error||"Gagal menghapus mutasi")}catch(m){N("Gagal menghapus data mutasi: "+m.message)}}})}function w(r=null){if(!F)return;let s=!!r,m=r?.tanggal?r.tanggal.slice(0,10):new Date().toISOString().slice(0,10),p=document.createElement("div");p.innerHTML=`
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="mutasi-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${m}" required />
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
            ${a.map(f=>`<option value="${typeof f=="string"?f:f.label||f.name||""}">`).join("")}
          </datalist>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Asal <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-from" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang Asal --</option>
              ${i.map(f=>`
                <option value="${f.value}" ${String(r?.from_branch_id)===String(f.value)?"selected":""}>
                  ${f.label}
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
              ${i.map(f=>`
                <option value="${f.value}" ${String(r?.to_branch_id)===String(f.value)?"selected":""}>
                  ${f.label}
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
    `,Y({title:s?"\u270F\uFE0F Edit Data Mutasi Karyawan":"\u2795 Tambah Data Mutasi Karyawan",content:p,size:"md",confirmText:s?"Simpan Perubahan":"Simpan Mutasi",confirmClass:"btn-primary",onConfirm:async(f,l)=>{let c=p.querySelector("#mutasi-modal-tanggal").value,v=(p.querySelector("#mutasi-modal-emp").value||"").trim(),S=p.querySelector("#mutasi-modal-from").value,B=p.querySelector("#mutasi-modal-to").value,T=p.querySelector("#mutasi-modal-status").value,D=(p.querySelector("#mutasi-modal-link").value||"").trim();if(!c){N("Tanggal mutasi wajib diisi");return}if(!v){N("Nama karyawan wajib diisi");return}if(!S){N("Cabang asal wajib dipilih");return}if(!B){N("Cabang tujuan wajib dipilih");return}let _=i.find(z=>String(z.value)===String(S)),P=i.find(z=>String(z.value)===String(B)),M={tanggal:c,employee_name:v,from_branch_id:Number(S),from_branch_name:_?_.label:"Cabang Asal",to_branch_id:Number(B),to_branch_name:P?P.label:"Cabang Tujuan",status:T,document_link:D};try{if(s){let z=await I(`/api/mutasi/${r.id}`,{method:"PUT",body:JSON.stringify(M)});if(z.ok){j("Data mutasi berhasil diperbarui");let x=t.findIndex(A=>A.id===r.id);x!==-1&&(t[x]={...t[x],...M}),l(),y()}else N(z.data?.error||"Gagal memperbarui data mutasi")}else{let z=await I("/api/mutasi",{method:"POST",body:JSON.stringify(M)});if(z.ok){j("Data mutasi baru berhasil ditambahkan");let x=z.data?.id||Date.now();t.unshift({id:x,...M}),l(),y()}else N(z.data?.error||"Gagal menambahkan mutasi")}}catch(z){N("Terjadi kesalahan: "+z.message)}}})}function k(){return t.filter(r=>{if(d){let s=d.toLowerCase(),m=(r.employee_name||"").toLowerCase().includes(s),p=(r.from_branch_name||"").toLowerCase().includes(s),f=(r.to_branch_name||"").toLowerCase().includes(s);if(!m&&!p&&!f)return!1}return!(b!=="all"&&String(r.from_branch_id)!==String(b)||u!=="all"&&String(r.to_branch_id)!==String(u)||o!=="all"&&r.status!==o)})}function y(){let r=k(),s=t.length,m=t.filter(S=>S.status==="Proses").length,p=t.filter(S=>S.status==="Selesai").length,f=new Date,l=f.getFullYear(),c=f.getMonth(),v=t.filter(S=>{if(!S.tanggal)return!1;let B=new Date(S.tanggal);return!isNaN(B.getTime())&&B.getFullYear()===l&&B.getMonth()===c}).length;e.innerHTML=`
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
            gap: 8px !important;
            flex-wrap: nowrap !important;
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
            grid-template-columns: 1fr;
          }
        }
      </style>

      <div class="mutasi-page">

        <!-- Hero Header Card -->
        <div class="mutasi-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                <span style="font-size:1.6rem;">\u{1F501}</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Data Mutasi Karyawan
                </h1>
              </div>
              <p style="margin:0; font-size:0.88rem; color:#CBD5E1; max-width:650px; line-height:1.5;">
                Sistem pengelolaan rotasi penempatan tugas, arsip SK mutasi cabang, dan riwayat perpindahan personil PT Audy Mandiri Indonesia.
              </p>
            </div>

            <div class="mutasi-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${F?`
                <button id="mutasi-btn-add" class="btn btn-primary mutasi-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Mutasi</span>
                </button>
              `:""}
              
              <button id="mutasi-btn-excel" class="btn btn-ghost mutasi-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="mutasi-btn-refresh" class="btn btn-ghost mutasi-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
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
                <div class="mutasi-stat-val">${s}</div>
                <div class="mutasi-stat-lbl">Total Mutasi</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u23F3
              </div>
              <div>
                <div class="mutasi-stat-val">${m}</div>
                <div class="mutasi-stat-lbl">Dalam Proses</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(34,197,94,0.2); color:#4ADE80;">
                \u2705
              </div>
              <div>
                <div class="mutasi-stat-val">${p}</div>
                <div class="mutasi-stat-lbl">Mutasi Selesai</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(168,85,247,0.2); color:#C084FC;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="mutasi-stat-val">${v}</div>
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
              <input type="text" id="mutasi-search-input" class="mutasi-search-input" placeholder="Cari nama karyawan, cabang asal, atau cabang tujuan..." value="${d}" />
            </div>

            <select id="mutasi-filter-from" class="mutasi-select">
              <option value="all">Semua Cabang Asal (${i.length})</option>
              ${i.map(S=>`<option value="${S.value}" ${String(b)===String(S.value)?"selected":""}>${S.label}</option>`).join("")}
            </select>

            <select id="mutasi-filter-to" class="mutasi-select">
              <option value="all">Semua Cabang Tujuan (${i.length})</option>
              ${i.map(S=>`<option value="${S.value}" ${String(u)===String(S.value)?"selected":""}>${S.label}</option>`).join("")}
            </select>
          </div>

          <!-- Status Filter Pills (Swipeable on Mobile) -->
          <div class="mutasi-filter-pills">
            <div class="mutasi-pill ${o==="all"?"active":""}" data-status="all">
              <span>Semua Mutasi</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${s}</span>
            </div>
            <div class="mutasi-pill ${o==="Proses"?"active":""}" data-status="Proses">
              <span style="color:#D97706;">\u23F3</span>
              <span>Proses</span>
              <span class="badge" style="background:rgba(217,119,6,0.15); color:#B45309; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${m}</span>
            </div>
            <div class="mutasi-pill ${o==="Selesai"?"active":""}" data-status="Selesai">
              <span style="color:#16A34A;">\u2705</span>
              <span>Selesai</span>
              <span class="badge" style="background:rgba(22,163,74,0.15); color:#15803D; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${p}</span>
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
          `:`
            <div class="mutasi-cards-grid">
              ${r.map(S=>{let B=S.status==="Selesai",T=(S.employee_name||"Karyawan").split(" ").map(D=>D[0]).filter(Boolean).slice(0,2).join("").toUpperCase();return`
                  <div class="mutasi-card" data-id="${S.id}">
                    <div>
                      <!-- Card Header: Status & Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="mutasi-status-badge ${B?"selesai":"proses"}">
                          <span>${B?"\u2705":"\u23F3"}</span>
                          <span>${S.status||"Proses"}</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B; display:inline-flex; align-items:center; gap:4px;">
                          <span>\u{1F5D3}\uFE0F</span>
                          <span>${zn(S.tanggal)}</span>
                        </span>
                      </div>

                      <!-- Employee Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="mutasi-avatar-box">
                          ${T}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${S.employee_name}
                          </h4>
                          <div style="font-size:0.75rem; color:#64748B;">
                            ID Mutasi: #${S.id}
                          </div>
                        </div>
                      </div>

                      <!-- Route Visual Box -->
                      <div class="mutasi-route-box">
                        <div class="mutasi-route-node">
                          <div class="mutasi-route-label">Cabang Asal</div>
                          <div class="mutasi-route-branch" title="${S.from_branch_name||"-"}">
                            ${S.from_branch_name||"-"}
                          </div>
                        </div>

                        <div class="mutasi-route-arrow">
                          \u2794
                        </div>

                        <div class="mutasi-route-node" style="text-align:right;">
                          <div class="mutasi-route-label">Cabang Tujuan</div>
                          <div class="mutasi-route-branch" style="color:#2563EB;" title="${S.to_branch_name||"-"}">
                            ${S.to_branch_name||"-"}
                          </div>
                        </div>
                      </div>
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:8px;">
                      <div>
                        ${S.document_link?`
                          <a href="${S.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>\u{1F4C4} Buka SK</span>
                          </a>
                        `:`
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${F?`
                        <div style="display:flex; align-items:center; gap:6px;">
                          <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${S.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                          <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${S.id}">
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
    `,g()}function g(){let r=e.querySelector("#mutasi-btn-add");r&&(r.onclick=()=>w());let s=e.querySelector("#mutasi-btn-excel");s&&(s.onclick=()=>E());let m=e.querySelector("#mutasi-btn-refresh");m&&(m.onclick=()=>$());let p=e.querySelector("#mutasi-btn-reset-filter");p&&(p.onclick=()=>{d="",b="all",u="all",o="all",y()});let f=e.querySelector("#mutasi-search-input");f&&(f.oninput=v=>{d=v.target.value,y();let S=e.querySelector("#mutasi-search-input");S&&(S.focus(),S.setSelectionRange(S.value.length,S.value.length))});let l=e.querySelector("#mutasi-filter-from");l&&(l.onchange=v=>{b=v.target.value,y()});let c=e.querySelector("#mutasi-filter-to");c&&(c.onchange=v=>{u=v.target.value,y()}),e.querySelectorAll(".mutasi-pill").forEach(v=>{v.onclick=()=>{o=v.getAttribute("data-status")||"all",y()}}),e.querySelectorAll(".mutasi-edit-btn").forEach(v=>{v.onclick=S=>{S.stopPropagation();let B=Number(v.getAttribute("data-id")),T=t.find(D=>D.id===B);T&&w(T)}}),e.querySelectorAll(".mutasi-delete-btn").forEach(v=>{v.onclick=S=>{S.stopPropagation();let B=Number(v.getAttribute("data-id"));C(B)}})}await $()}V();ue();var qt=[{id:1,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:1,employee_name:"ADE SURAHMAN",start_time:"23:00",end_time:"08:00",break_hours:1,total_hours:8,reason:"General Cleaning"},{id:2,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:2,employee_name:"TAUFIK HIDAYAT",start_time:"00:00",end_time:"08:00",break_hours:1,total_hours:7,reason:"General Cleaning"},{id:3,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:3,employee_name:"YUDA WILDAN WIGUNA",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"},{id:4,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:4,employee_name:"REZA RIZKI",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"},{id:5,date:"2026-10-03",branch_id:25,branch_name:"025. Bogor Baru",employee_id:5,employee_name:"ADITYA PRATAMA",start_time:"22:00",end_time:"08:00",break_hours:1,total_hours:9,reason:"General Cleaning"}],qa=["General Cleaning","Deep Cleaning","Cover manpower","Longshift","Pekerjaan urgent"];function Ha(e,t,i){if(!e||!t)return 0;try{let[a,o]=e.split(":").map(Number),[b,u]=t.split(":").map(Number),d=b*60+u-(a*60+o);d<0&&(d+=24*60);let n=d/60-(parseFloat(i)||0);return n<0&&(n=0),Number.isInteger(n)?n:parseFloat(n.toFixed(2))}catch{return 0}}function Ja(e=""){let t=(e||"").toLowerCase();return t.includes("general")?{bg:"#EFF6FF",color:"#1D4ED8",border:"#BFDBFE",icon:"\u{1F9F9}"}:t.includes("deep")?{bg:"#F5F3FF",color:"#6D28D9",border:"#DDD6FE",icon:"\u2728"}:t.includes("cover")?{bg:"#ECFDF5",color:"#047857",border:"#A7F3D0",icon:"\u{1F465}"}:t.includes("longshift")?{bg:"#FFFBEB",color:"#B45309",border:"#FDE68A",icon:"\u23F3"}:t.includes("urgent")?{bg:"#FFF1F2",color:"#BE123C",border:"#FECDD3",icon:"\u{1F6A8}"}:{bg:"#F8FAFC",color:"#475569",border:"#CBD5E1",icon:"\u23F1\uFE0F"}}function Ga(e){if(!e)return"-";try{let t=new Date(e);return isNaN(t.getTime())?e:t.toLocaleDateString("id-ID",{day:"numeric",month:"short",year:"numeric"})}catch{return e}}async function Ua(e){let t=[],i=[],a=[],o="all",b="all",u="",d=!1,n=ae(),h=n?.role==="viewer",F=n?.role==="input_lembur",$=!h,E=!h,C=!h&&!F;try{[i,a]=await Promise.all([W(),_e()])}catch(p){console.error("Error fetching cache:",p)}async function w(){d=!0,s();try{let p=await I("/api/overtime?limit=10000");if(p.ok){let f=p.data?.data||p.data||[];t=Array.isArray(f)&&f.length>0?f:qt}else t=qt}catch(p){console.warn("Gagal load lembur, menggunakan data lokal:",p),t=qt}finally{d=!1,s()}}async function k(){try{let{downloadExcel:p}=await Promise.resolve().then(()=>(X(),Se)),f=t.map((l,c)=>({No:c+1,Tanggal:l.date||"","Cabang Lembur":l.branch_name||"","Nama Karyawan":l.employee_name||"","Jam Mulai":l.start_time||"","Jam Selesai":l.end_time||"","Istirahat (Jam)":l.break_hours!=null?l.break_hours:0,"Total Jam":l.total_hours!=null?l.total_hours:0,"Alasan / Keterangan":l.reason||""}));p(f,`Data_Lembur_${new Date().toISOString().slice(0,10)}`),j("Data lembur berhasil diekspor ke Excel!")}catch(p){N("Gagal ekspor Excel: "+p.message)}}async function y(p){if(!C)return;let f=t.find(l=>l.id===p);ee({title:"Hapus Catatan Lembur",message:`Apakah Anda yakin ingin menghapus catatan lembur untuk "${f?.employee_name||"karyawan ini"}" pada tanggal ${Ga(f?.date)}?`,confirmText:"Ya, Hapus",confirmClass:"btn-danger",onConfirm:async()=>{try{let l=await I(`/api/overtime/${p}`,{method:"DELETE"});l.ok?(j("Catatan lembur berhasil dihapus"),t=t.filter(c=>c.id!==p),s()):N(l.data?.error||"Gagal menghapus lembur")}catch(l){N("Gagal menghapus catatan lembur: "+l.message)}}})}function g(p=null){if(!E)return;let f=!!p,l=p?.date?p.date.slice(0,10):new Date().toISOString().slice(0,10),c=p?.start_time||"22:00",v=p?.end_time||"08:00",S=p?.break_hours!=null?p.break_hours:1,B=p?.total_hours!=null?p.total_hours:Ha(c,v,S),T=document.createElement("div");T.innerHTML=`
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
              ${i.map(x=>`
                <option value="${x.value}" ${String(p?.branch_id)===String(x.value)?"selected":""}>
                  ${x.label}
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
            ${a.map(x=>{let A=typeof x=="string"?x:x.label||x.full_name||x.name||"",L=typeof x=="object"&&x!==null?x.value!=null?x.value:x.id!=null?x.id:A:x,q=String(p?.employee_id)===String(L)||p?.employee_name===A;return`<option value="${L}" ${q?"selected":""}>${A}</option>`}).join("")}
          </select>
        </div>

        <!-- Shift & Waktu Lembur -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Mulai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${c}" required />
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
              value="${S}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Total Jam Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.1" min="0" id="ot-modal-total" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; background:#F8FAFC; font-weight:700; color:#2563EB;" 
              value="${B}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Alasan / Keterangan Penugasan Lembur <span style="color:#DC2626;">*</span>
          </label>
          <select id="ot-modal-reason" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
            ${qa.map(x=>`
              <option value="${x}" ${p?.reason===x?"selected":""}>${x}</option>
            `).join("")}
          </select>
        </div>
      </div>
    `;let D=T.querySelector("#ot-modal-start"),_=T.querySelector("#ot-modal-end"),P=T.querySelector("#ot-modal-break"),M=T.querySelector("#ot-modal-total");function z(){let x=D.value,A=_.value,L=P.value;M.value=Ha(x,A,L)}D.addEventListener("input",z),_.addEventListener("input",z),P.addEventListener("input",z),Y({title:f?"\u270F\uFE0F Edit Data Lembur Karyawan":"\u2795 Catat Penugasan Lembur Baru",content:T,size:"md",confirmText:f?"Simpan Perubahan":"Catat Lembur",confirmClass:"btn-primary",onConfirm:async(x,A)=>{let L=T.querySelector("#ot-modal-date").value,q=T.querySelector("#ot-modal-branch").value,O=T.querySelector("#ot-modal-emp").value,H=T.querySelector("#ot-modal-start").value,K=T.querySelector("#ot-modal-end").value,R=parseFloat(T.querySelector("#ot-modal-break").value)||0,ie=parseFloat(T.querySelector("#ot-modal-total").value)||0,de=T.querySelector("#ot-modal-reason").value;if(!L){N("Tanggal lembur wajib diisi");return}if(!q){N("Cabang penugasan wajib dipilih");return}if(!O){N("Nama karyawan wajib dipilih");return}if(!H||!K){N("Jam mulai dan jam selesai wajib diisi");return}let me=i.find(J=>String(J.value)===String(q)),U=a.find(J=>String(J.value)===String(O)||String(J.id)===String(O)),re=U?U.label||U.full_name||U.name||"Karyawan":typeof O=="string"&&isNaN(Number(O))?O:"Karyawan",oe={date:L,branch_id:Number(q),branch_name:me?me.label:"Cabang Lembur",employee_id:Number(O),employee_name:U?U.label:"Karyawan",start_time:H,end_time:K,break_hours:R,total_hours:ie,reason:de};try{if(f){let J=await I(`/api/overtime/${p.id}`,{method:"PUT",body:JSON.stringify(oe)});if(J.ok){j("Data lembur berhasil diperbarui");let pe=t.findIndex(fe=>fe.id===p.id);pe!==-1&&(t[pe]={...t[pe],...oe}),A(),s()}else N(J.data?.error||"Gagal memperbarui data lembur")}else{let J=await I("/api/overtime",{method:"POST",body:JSON.stringify(oe)});if(J.ok){j("Data lembur berhasil ditambahkan");let pe=J.data?.id||Date.now();t.unshift({id:pe,...oe}),A(),s()}else N(J.data?.error||"Gagal menambahkan lembur")}}catch(J){N("Terjadi kesalahan: "+J.message)}}})}function r(){return t.filter(p=>{if(u){let f=u.toLowerCase(),l=(p.employee_name||"").toLowerCase().includes(f),c=(p.branch_name||"").toLowerCase().includes(f),v=(p.reason||"").toLowerCase().includes(f);if(!l&&!c&&!v)return!1}return!(b!=="all"&&String(p.branch_id)!==String(b)||o!=="all"&&p.reason!==o)})}function s(){let p=r(),f=t.length,l=t.reduce((D,_)=>D+(parseFloat(_.total_hours)||0),0),c=f>0?(l/f).toFixed(1):0,v=new Date,S=v.getFullYear(),B=v.getMonth(),T=t.filter(D=>{if(!D.date)return!1;let _=new Date(D.date);return!isNaN(_.getTime())&&_.getFullYear()===S&&_.getMonth()===B}).length;e.innerHTML=`
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
            gap: 8px !important;
            flex-wrap: nowrap !important;
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
          .lembur-cards-grid {
            grid-template-columns: 1fr;
          }
        }
      </style>

      <div class="lembur-page">

        <!-- Hero Header Card -->
        <div class="lembur-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                <span style="font-size:1.6rem;">\u23F1\uFE0F</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Data Lembur Karyawan
                </h1>
              </div>
              <p style="margin:0; font-size:0.88rem; color:#CBD5E1; max-width:650px; line-height:1.5;">
                Pencatatan jam kerja lembur, approval penugasan overtime cabang, dan rekapitulasi durasi jam kerja PT Audy Mandiri Indonesia.
              </p>
            </div>

            <div class="lembur-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${$?`
                <button id="lembur-btn-add" class="btn btn-primary lembur-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Lembur</span>
                </button>
              `:""}
              
              <button id="lembur-btn-excel" class="btn btn-ghost lembur-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>\u{1F4E5}</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="lembur-btn-refresh" class="btn btn-ghost lembur-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
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
                <div class="lembur-stat-val">${f}</div>
                <div class="lembur-stat-lbl">Catatan Penugasan</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                \u{1F4CA}
              </div>
              <div>
                <div class="lembur-stat-val">${c} Jam</div>
                <div class="lembur-stat-lbl">Rata-rata Durasi</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                \u{1F5D3}\uFE0F
              </div>
              <div>
                <div class="lembur-stat-val">${T}</div>
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
              <input type="text" id="lembur-search-input" class="lembur-search-input" placeholder="Cari nama karyawan, cabang, atau alasan lembur..." value="${u}" />
            </div>

            <select id="lembur-filter-branch" class="lembur-select">
              <option value="all">Semua Cabang (${i.length})</option>
              ${i.map(D=>`<option value="${D.value}" ${String(b)===String(D.value)?"selected":""}>${D.label}</option>`).join("")}
            </select>
          </div>

          <!-- Reason Filter Pills (Swipeable on Mobile) -->
          <div class="lembur-filter-pills">
            <div class="lembur-pill ${o==="all"?"active":""}" data-reason="all">
              <span>Semua Lembur</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${f}</span>
            </div>
            ${qa.map(D=>{let _=t.filter(M=>M.reason===D).length,P=Ja(D);return`
                <div class="lembur-pill ${o===D?"active":""}" data-reason="${D}">
                  <span>${P.icon}</span>
                  <span>${D}</span>
                  ${_>0?`<span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${_}</span>`:""}
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
          `:`
            <div class="lembur-cards-grid">
              ${p.map(D=>{let _=(D.employee_name||"Karyawan").split(" ").map(M=>M[0]).filter(Boolean).slice(0,2).join("").toUpperCase(),P=Ja(D.reason);return`
                  <div class="lembur-card" data-id="${D.id}">
                    <div>
                      <!-- Card Header: Jam & Istirahat -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="lembur-hours-badge">
                          <span>\u23F1\uFE0F</span>
                          <span>${D.total_hours!=null?D.total_hours:0} Jam Lembur</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                          \u2615 Istirahat: ${D.break_hours!=null?D.break_hours:0} Jam
                        </span>
                      </div>

                      <!-- Employee & Branch -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="lembur-avatar-box">
                          ${_}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${D.employee_name}
                          </h4>
                          <div style="font-size:0.8rem; color:#64748B; display:flex; align-items:center; gap:4px;">
                            <span>\u{1F4CD}</span>
                            <span style="font-weight:600; color:#334155;">${D.branch_name||"Cabang"}</span>
                          </div>
                        </div>
                      </div>

                      <!-- Shift Times & Date Box -->
                      <div class="lembur-time-box">
                        <div>
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Tanggal</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#1E293B;">
                            \u{1F5D3}\uFE0F ${Ga(D.date)}
                          </div>
                        </div>

                        <div style="text-align:right;">
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Waktu Shift</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#2563EB;">
                            \u23F0 ${D.start_time||"00:00"} - ${D.end_time||"00:00"}
                          </div>
                        </div>
                      </div>

                      <!-- Reason Badge -->
                      <div>
                        <span class="lembur-reason-badge" style="background:${P.bg}; color:${P.color}; border:1px solid ${P.border};">
                          <span>${P.icon}</span>
                          <span>${D.reason||"Lembur Operasional"}</span>
                        </span>
                      </div>
                    </div>

                    <!-- Card Footer: Aksi -->
                    ${E?`
                      <div style="display:flex; justify-content:flex-end; align-items:center; gap:6px; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:12px;">
                        <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${D.id}">
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        ${C?`
                          <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${D.id}">
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
    `,m()}function m(){let p=e.querySelector("#lembur-btn-add");p&&(p.onclick=()=>g());let f=e.querySelector("#lembur-btn-excel");f&&(f.onclick=()=>k());let l=e.querySelector("#lembur-btn-refresh");l&&(l.onclick=()=>w());let c=e.querySelector("#lembur-btn-reset-filter");c&&(c.onclick=()=>{u="",b="all",o="all",s()});let v=e.querySelector("#lembur-search-input");v&&(v.oninput=B=>{u=B.target.value,s();let T=e.querySelector("#lembur-search-input");T&&(T.focus(),T.setSelectionRange(T.value.length,T.value.length))});let S=e.querySelector("#lembur-filter-branch");S&&(S.onchange=B=>{b=B.target.value,s()}),e.querySelectorAll(".lembur-pill").forEach(B=>{B.onclick=()=>{o=B.getAttribute("data-reason")||"all",s()}}),e.querySelectorAll(".lembur-edit-btn").forEach(B=>{B.onclick=T=>{T.stopPropagation();let D=Number(B.getAttribute("data-id")),_=t.find(P=>P.id===D);_&&g(_)}}),e.querySelectorAll(".lembur-delete-btn").forEach(B=>{B.onclick=T=>{T.stopPropagation();let D=Number(B.getAttribute("data-id"));y(D)}})}await w()}V();async function Va(e){let t=ae();if(!t||!["superadmin","admin"].includes(t.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}Z({container:e,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:i=>new Date(i).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(i,a)=>`<strong>${i||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${a.user_role||""}</small>`},{key:"action",label:"Aksi",render:i=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[i]||"badge-neutral"}">${i}</span>`},{key:"module",label:"Modul",render:i=>`<span style="text-transform:capitalize">${(i||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(i,a)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${i}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async i=>{try{let b=((await(await fetch(`/api/audit-logs?search=${i}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(h=>String(h.id)===String(i));if(!b)return alert("Data tidak ditemukan");let u=h=>{if(!h)return"Tidak ada data";try{return JSON.stringify(JSON.parse(h),null,2)}catch{return h}},d=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${u(b.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${u(b.new_data)}</pre>
           </div>
         </div>
       `,{createModal:n}=await Promise.resolve().then(()=>(ue(),kt));n({title:`Detail Audit Log #${i}`,content:d,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}V();ue();X();var Ht={"2026-09":"September 2026","2026-10":"Oktober 2026","2026-11":"November 2026","2026-12":"Desember 2026"},We={scheduled:{label:"Terjadwal",badgeClass:"badge-warning",icon:"\u{1F7E1}",bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},in_progress:{label:"Sedang Proses",badgeClass:"badge-info",icon:"\u{1F535}",bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},completed:{label:"Selesai",badgeClass:"badge-success",icon:"\u{1F7E2}",bg:"#DCFCE7",color:"#166534",border:"#86EFAC"}};async function Ya(e){e.innerHTML=`
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
  `;let t="board",i=null,a=[],o=ae(),b=o&&["superadmin","admin","manager","spv"].includes(o.role);I("/api/branches?all=1").then(g=>{g.ok&&g.data?.data&&(a=g.data.data)}).catch(()=>{});async function u(){try{let g=await I("/api/os-migrations/summary");if(!g.ok)throw new Error("Gagal memuat data migrasi");i=g.data?.data||g.data,d(i),h()}catch(g){document.getElementById("migration-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${g.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function d(g){let r=document.getElementById("migration-summary-cards");if(!r||!g)return;let s=g.total_branches||0,m=g.vendors?.AEON?.total||0,p=g.vendors?.["Triple S"]?.total||0,f=g.by_status?.completed||0,l=g.by_status?.in_progress||0,c=g.by_status?.scheduled||0,v=s>0?Math.round(f/s*100):0;r.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${s} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.8rem; font-weight:800; color:#1E40AF; line-height:1.2;">${m} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 9 \u2022 Nov: 3 \u2022 Des: 8</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #0D9488; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#0F766E; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor Triple S</div>
        <div style="font-size:1.8rem; font-weight:800; color:#0F766E; line-height:1.2;">${p} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 5 \u2022 Okt: 5 \u2022 Nov: 5 \u2022 Des: 27</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #16A34A; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <span style="font-size:0.8rem; font-weight:600; color:#15803D; text-transform:uppercase; letter-spacing:0.5px;">Progres Realisasi</span>
          <span style="font-size:0.85rem; font-weight:700; color:#15803D;">${v}%</span>
        </div>
        <div style="width:100%; height:8px; background:#E2E8F0; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${v}%; height:100%; background:#16A34A; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.75rem; color:var(--text-2); font-weight:500;">
          <span>\u{1F7E2} Selesai: <strong>${f}</strong></span>
          <span>\u{1F535} Proses: <strong>${l}</strong></span>
          <span>\u{1F7E1} Terjadwal: <strong>${c}</strong></span>
        </div>
      </div>
    `}function n(){if(!i||!i.matrix)return[];let g=(document.getElementById("migration-search")?.value||"").toLowerCase().trim(),r=document.getElementById("filter-vendor")?.value||"",s=document.getElementById("filter-month")?.value||"",m=document.getElementById("filter-status")?.value||"";return i.matrix.filter(p=>!(r&&p.vendor!==r||s&&p.target_month!==s||m&&p.status!==m||g&&!`${p.branch_code} ${p.branch_name} ${p.branch_full_name} ${p.vendor} ${p.notes||""} ${p.pic_vendor||""}`.toLowerCase().includes(g)))}function h(){let g=n(),r=document.getElementById("filter-stats-text");r&&(r.innerHTML=`Menampilkan <strong>${g.length}</strong> dari <strong>${i?.total_branches||0}</strong> cabang`);let s=document.getElementById("migration-view-container");t==="board"?F(s,g):$(s,g)}function F(g,r){if(!r.length){g.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;return}let s=["AEON","Triple S"],m=["2026-09","2026-10","2026-11","2026-12"],p="";s.forEach(f=>{let l=r.filter(D=>D.vendor===f);if(document.getElementById("filter-vendor")?.value&&document.getElementById("filter-vendor").value!==f||l.length===0)return;let c=l.length,v=f==="AEON"?"#1E40AF":"#0F766E",S=f==="AEON"?"#EFF6FF":"#F0FDFA",B=f==="AEON"?"#BFDBFE":"#99F6E4",T=m.filter(D=>l.some(_=>_.target_month===D));p+=`
        <div class="vendor-board-section" style="margin-bottom:32px; background:#fff; border:1px solid var(--border); border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <!-- Vendor Section Header -->
          <div style="padding:14px 20px; background:${S}; border-bottom:1px solid ${B}; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.1rem; font-weight:800; color:${v};">${f==="AEON"?"\u{1F3E2} AEON":"\u{1F6E1}\uFE0F Triple S"}</span>
              <span class="badge" style="background:#fff; color:${v}; border:1px solid ${B}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px;">
                ${c} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
              Format Timeline Berdasarkan Dokumen Migrasi
            </div>
          </div>

          <!-- Months Grid (Hanya menampilkan bulan yang memiliki jadwal) -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px; padding:18px;">
            ${T.map(D=>{let _=l.filter(z=>z.target_month===D);return`
                <div class="month-column" style="background:#F8FAFC; border:1px solid var(--border); border-radius:10px; display:flex; flex-direction:column; min-height:220px;">
                  <!-- Month Header -->
                  <div style="padding:12px 14px; background:#F1F5F9; border-bottom:1px solid var(--border); border-radius:9px 9px 0 0; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">
                      \u{1F4C5} ${(Ht[D]||D).split(" ")[0]}
                    </div>
                    <span class="badge" style="background:#E2E8F0; color:var(--text-2); font-weight:700; font-size:0.75rem; padding:2px 7px; border-radius:4px;">
                      ${_.length}
                    </span>
                  </div>

                  <!-- Branch Cards Container -->
                  <div style="padding:10px; display:flex; flex-direction:column; gap:8px; flex:1; max-height:480px; overflow-y:auto;">
                    ${_.map(z=>{let x=We[z.status]||We.scheduled;return`
                        <div class="branch-card" style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 12px; box-shadow:0 1px 2px rgba(0,0,0,0.03); transition:transform 0.15s, box-shadow 0.15s;">
                          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px; margin-bottom:6px;">
                            <div style="font-weight:700; font-size:0.9rem; color:var(--text-1); line-height:1.3;">
                              <span style="color:#2563EB; font-family:monospace; margin-right:4px;">${z.branch_code}.</span>${z.branch_name}
                            </div>
                            ${b?`
                              <div style="display:flex; gap:2px;">
                                <button class="btn-icon-mini edit-card-btn" data-id="${z.id}" title="Edit" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--text-3);">\u270F\uFE0F</button>
                                <button class="btn-icon-mini delete-card-btn" data-id="${z.id}" data-name="${z.branch_name}" title="Hapus" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--danger);">\u{1F5D1}\uFE0F</button>
                              </div>
                            `:""}
                          </div>

                          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px; margin-top:8px;">
                            <!-- Status Selector Dropdown -->
                            <div style="position:relative;">
                              <button class="status-toggle-btn" data-id="${z.id}" data-status="${z.status}" style="background:${x.bg}; color:${x.color}; border:1px solid ${x.border}; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; display:inline-flex; align-items:center; gap:4px; cursor:pointer;" title="Klik untuk ubah status">
                                ${x.icon} ${x.label} \u25BE
                              </button>
                            </div>

                            ${z.employee_count>0?`
                              <span style="font-size:0.75rem; color:var(--text-3); font-weight:600; display:inline-flex; align-items:center; gap:3px;" title="${z.employee_count} personil Facility Care aktif di cabang ini">
                                \u{1F464} ${z.employee_count} FC
                              </span>
                            `:""}
                          </div>

                          ${z.actual_date?`
                            <div style="font-size:0.7rem; color:var(--text-3); margin-top:6px;">
                              Realisasi: <strong>${z.actual_date}</strong>
                            </div>
                          `:""}
                        </div>
                      `}).join("")}
                  </div>
                </div>
              `}).join("")}
          </div>
        </div>
      `}),g.innerHTML=p,E(g)}function $(g,r){if(!r.length){g.innerHTML=`
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
                ${b?'<th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${r.map((s,m)=>{let p=We[s.status]||We.scheduled;return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3);">${m+1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">${s.branch_full_name}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${s.vendor==="AEON"?"#DBEAFE":"#CCFBF1"}; color:${s.vendor==="AEON"?"#1E40AF":"#0F766E"}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${s.vendor}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:var(--text-2);">
                      ${Ht[s.target_month]||s.target_month}
                    </td>
                    <td style="padding:12px 16px;">
                      <button class="status-toggle-btn" data-id="${s.id}" data-status="${s.status}" style="background:${p.bg}; color:${p.color}; border:1px solid ${p.border}; padding:4px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; cursor:pointer;" title="Klik untuk ubah status">
                        ${p.icon} ${p.label} \u25BE
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
                    ${b?`
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
    `,E(g)}function E(g){g.querySelectorAll(".status-toggle-btn").forEach(r=>{r.addEventListener("click",s=>{s.stopPropagation();let m=r.dataset.id,p=r.dataset.status;C(r,m,p)})}),g.querySelectorAll(".edit-card-btn").forEach(r=>{r.addEventListener("click",()=>{let s=r.dataset.id,m=i?.matrix?.find(p=>String(p.id)===String(s));m&&w(m)})}),g.querySelectorAll(".delete-card-btn").forEach(r=>{r.addEventListener("click",()=>{let s=r.dataset.id,m=r.dataset.name||"Cabang";ee(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${m}</strong>?`,async()=>{try{if(!(await I(`/api/os-migrations/${s}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus");j(`Jadwal migrasi ${m} berhasil dihapus`),u()}catch(p){N(p.message)}})})})}function C(g,r,s){document.querySelectorAll(".status-popup-menu").forEach(c=>c.remove());let m=g.getBoundingClientRect(),p=document.createElement("div");p.className="status-popup-menu",p.style.cssText=`
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
    `;let f=[{key:"scheduled",label:"\u{1F7E1} Terjadwal"},{key:"in_progress",label:"\u{1F535} Sedang Proses"},{key:"completed",label:"\u{1F7E2} Selesai"}];p.innerHTML=f.map(c=>`
      <button class="status-option-item" data-key="${c.key}" style="
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
        color: ${c.key===s?"#2563EB":"var(--text-1)"};
        font-weight: ${c.key===s?"700":"500"};
      ">
        ${c.label} ${c.key===s?"\u2713":""}
      </button>
    `).join("");let l=c=>{!p.contains(c.target)&&c.target!==g&&(p.remove(),document.removeEventListener("click",l))};setTimeout(()=>document.addEventListener("click",l),50),p.querySelectorAll(".status-option-item").forEach(c=>{c.addEventListener("click",async()=>{let v=c.dataset.key;if(p.remove(),document.removeEventListener("click",l),v!==s)try{if(!(await I(`/api/os-migrations/${r}/status`,{method:"PATCH",body:JSON.stringify({status:v})})).ok)throw new Error("Gagal memperbarui status");j("Status berhasil diubah"),u()}catch(S){N(S.message)}})}),document.body.appendChild(p)}function w(g=null){let r=!!g,s=r?"Edit Jadwal Migrasi OS":"Tambah Jadwal Migrasi OS",p=`
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${a.map(f=>`
      <option value="${f.id}" ${g&&g.branch_id===f.id?"selected":""}>
        ${f.full_name}
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
    `;Y({title:s,content:p,confirmText:r?"Simpan Perubahan":"Tambah Cabang",onConfirm:async(f,l)=>{let c=f.querySelector("#migration-form");if(!c.checkValidity()){c.reportValidity();return}let v=new FormData(c),S={branch_id:parseInt(v.get("branch_id"),10),vendor:v.get("vendor"),target_month:v.get("target_month"),status:v.get("status"),actual_date:v.get("actual_date")||null,pic_vendor:v.get("pic_vendor")||null,notes:v.get("notes")||null};try{let B=r?`/api/os-migrations/${g.id}`:"/api/os-migrations",D=await I(B,{method:r?"PUT":"POST",body:JSON.stringify(S)});if(!D.ok){let _=await D.json().catch(()=>({}));throw new Error(_.error||"Gagal menyimpan data")}j(r?"Jadwal migrasi berhasil diupdate":"Jadwal migrasi berhasil ditambahkan"),l(),u()}catch(B){N(B.message)}}})}let k=document.getElementById("view-board-btn"),y=document.getElementById("view-table-btn");k.addEventListener("click",()=>{t="board",k.style.background="#fff",k.classList.remove("btn-ghost"),y.style.background="transparent",y.classList.add("btn-ghost"),h()}),y.addEventListener("click",()=>{t="table",y.style.background="#fff",y.classList.remove("btn-ghost"),k.style.background="transparent",k.classList.add("btn-ghost"),h()}),["migration-search","filter-vendor","filter-month","filter-status"].forEach(g=>{let r=document.getElementById(g);r&&(r.addEventListener("input",h),r.addEventListener("change",h))}),document.getElementById("btn-add-migration").addEventListener("click",()=>{w()}),document.getElementById("btn-export-migrations").addEventListener("click",()=>{let g=n();if(!g.length){N("Tidak ada data untuk diexport");return}let r=g.map((s,m)=>({No:m+1,"Kode Cabang":s.branch_code,"Nama Cabang":s.branch_name,"Cabang Lengkap":s.branch_full_name,"Vendor OS":s.vendor,"Target Bulan":Ht[s.target_month]||s.target_month,Status:We[s.status]?.label||s.status,"Tanggal Realisasi":s.actual_date||"","Jml Personil FC":s.employee_count||0,"PIC Vendor":s.pic_vendor||"",Catatan:s.notes||""}));G(r,`Timeline_Migrasi_OS_${new Date().toISOString().slice(0,10)}`),j("File Excel berhasil diunduh")}),u()}V();ue();X();var Wa={Persiapan:{bg:"#F3E8FF",color:"#6B21A8",border:"#D8B4FE"},"Area Umum":{bg:"#E0F2FE",color:"#0369A1",border:"#BAE6FD"},"Ruang Tindakan":{bg:"#DBEAFE",color:"#1E40AF",border:"#93C5FD"},"Ruang Penunjang":{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},Toilet:{bg:"#CCFBF1",color:"#0F766E",border:"#99F6E4"},"Area Luar":{bg:"#FEF3C7",color:"#92400E",border:"#FDE68A"},"Kontrol & Check":{bg:"#FFEDD5",color:"#C2410C",border:"#FED7AA"},Istirahat:{bg:"#E2E8F0",color:"#475569",border:"#CBD5E1"},"Layanan Pasien":{bg:"#DCFCE7",color:"#15803D",border:"#86EFAC"},"Sampah Umum":{bg:"#FEF9C3",color:"#854D0E",border:"#FDE047"},"Sampah Medis":{bg:"#FEE2E2",color:"#991B1B",border:"#FCA5A5"},Penutup:{bg:"#E0E7FF",color:"#3730A3",border:"#C7D2FE"}};async function Qa(e){e.innerHTML=`
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
  `;let t="pagi",i="timeline",a=[],o=ae(),b=o&&["superadmin","admin","manager","spv"].includes(o.role);async function u(){try{let g=await I("/api/dac?all=1");if(!g.ok)throw new Error("Gagal memuat standar DAC");a=g.data?.data||g.data||[];let r=a.filter(f=>f.shift==="pagi").length,s=a.filter(f=>f.shift==="siang").length;document.getElementById("badge-count-pagi").textContent=r,document.getElementById("badge-count-siang").textContent=s,document.getElementById("badge-count-all").textContent=a.length;let m=document.getElementById("dac-category-filter"),p=Array.from(new Set(a.map(f=>f.category))).sort();m.innerHTML='<option value="">Semua Kategori</option>'+p.map(f=>`<option value="${f}">${f}</option>`).join(""),n(),h()}catch(g){document.getElementById("dac-view-container").innerHTML=`
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">\u26A0\uFE0F</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${g.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `}}function d(){let g=(document.getElementById("dac-search")?.value||"").toLowerCase().trim(),r=document.getElementById("dac-category-filter")?.value||"",s=document.getElementById("dac-period-filter")?.value||"";return a.filter(m=>{if(t!=="all"&&m.shift!==t||r&&m.category!==r)return!1;if(s==="morning"){let p=parseInt(m.start_time.split(":")[0],10);if(m.shift==="pagi"&&p>=12||m.shift==="siang"&&p>=17)return!1}else if(s==="afternoon"){let p=parseInt(m.start_time.split(":")[0],10);if(m.shift==="pagi"&&p<12||m.shift==="siang"&&p<17)return!1}return!(g&&!`${m.start_time} ${m.end_time} ${m.category} ${m.location} ${m.activity}`.toLowerCase().includes(g))})}function n(){let g=document.getElementById("dac-summary-cards");if(!g)return;let r=a.filter(l=>l.shift==="pagi"),s=a.filter(l=>l.shift==="siang"),m=t==="pagi"?r:t==="siang"?s:a,p=t==="pagi"?"08:00 \u2013 17:00":t==="siang"?"13:00 \u2013 22:00":"08:00 \u2013 22:00",f=t==="pagi"?"Shift Pagi":t==="siang"?"Shift Siang":"Semua Shift";g.innerHTML=`
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Aktivitas (${f})</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${m.length} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Aktivitas</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Urutan Terjadwal dari Awal s/d Penutup</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Jam Kerja Standar</div>
        <div style="font-size:1.5rem; font-weight:800; color:#1E40AF; line-height:1.2;">${p}</div>
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
    `}function h(){let g=d(),r=document.getElementById("dac-stats-text");r&&(r.innerHTML=`Menampilkan <strong>${g.length}</strong> aktivitas`);let s=document.getElementById("dac-view-container");i==="timeline"?F(s,g):$(s,g)}function F(g,r){if(!r.length){g.innerHTML=`
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">\u{1F50D}</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;return}g.innerHTML=`
      <div class="timeline-container" style="display:flex; flex-direction:column; gap:12px; max-width:960px; margin:0 auto;">
        ${r.map(s=>{let m=Wa[s.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"},p=s.shift==="pagi"?"\u{1F305} Pagi":"\u{1F307} Siang",f=s.shift==="pagi"?"#FEF3C7":"#E0E7FF",l=s.shift==="pagi"?"#92400E":"#3730A3";return`
            <div class="dac-timeline-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:16px 20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; gap:16px; align-items:flex-start; position:relative;">
              <!-- Time Column -->
              <div style="min-width:130px; text-align:center; background:#F8FAFC; border:1px solid var(--border); border-radius:8px; padding:10px 8px;">
                <div style="font-size:1.1rem; font-weight:800; color:var(--text-1); font-family:monospace;">
                  ${s.start_time} - ${s.end_time}
                </div>
                <div style="font-size:0.75rem; color:var(--text-3); font-weight:600; margin-top:4px;">
                  Item #${s.item_no}
                </div>
                ${t==="all"?`
                  <div style="margin-top:6px;">
                    <span class="badge" style="background:${f}; color:${l}; font-size:0.7rem; font-weight:700; padding:2px 6px; border-radius:4px;">
                      ${p}
                    </span>
                  </div>
                `:""}
              </div>

              <!-- Content Column -->
              <div style="flex:1;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                  <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span class="badge" style="background:${m.bg}; color:${m.color}; border:1px solid ${m.border}; font-weight:700; font-size:0.75rem; padding:3px 10px; border-radius:6px;">
                      ${s.category}
                    </span>
                    <span style="font-weight:700; font-size:0.95rem; color:var(--text-1);">
                      \u{1F4CD} ${s.location}
                    </span>
                  </div>
                  ${b?`
                    <div style="display:flex; gap:4px;">
                      <button class="btn btn-ghost btn-sm edit-dac-btn" data-id="${s.id}" title="Edit" style="padding:4px 8px; font-size:0.8rem; cursor:pointer;">\u270F\uFE0F</button>
                      <button class="btn btn-ghost btn-sm delete-dac-btn" data-id="${s.id}" data-name="${s.location}" title="Hapus" style="padding:4px 8px; font-size:0.8rem; color:var(--danger); cursor:pointer;">\u{1F5D1}\uFE0F</button>
                    </div>
                  `:""}
                </div>

                <div style="font-size:0.9rem; color:var(--text-2); line-height:1.5; background:#FAFBFD; padding:10px 14px; border-radius:8px; border-left:3px solid ${m.color}; margin-top:8px;">
                  ${s.activity}
                </div>
              </div>
            </div>
          `}).join("")}
      </div>
    `,E(g)}function $(g,r){if(!r.length){g.innerHTML=`
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
                ${t==="all"?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">SHIFT</th>':""}
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:110px;">JAM</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:140px;">KATEGORI</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:200px;">LOKASI / AREA</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">AKTIVITAS</th>
                ${b?'<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right; width:100px;">AKSI</th>':""}
              </tr>
            </thead>
            <tbody>
              ${r.map(s=>{let m=Wa[s.category]||{bg:"#F1F5F9",color:"#334155",border:"#CBD5E1"};return`
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 14px; font-size:0.85rem; color:var(--text-3); font-weight:600;">${s.item_no}</td>
                    ${t==="all"?`
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
                      <span class="badge" style="background:${m.bg}; color:${m.color}; border:1px solid ${m.border}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px; white-space:nowrap;">
                        ${s.category}
                      </span>
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; font-weight:600; color:var(--text-1);">
                      ${s.location}
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; color:var(--text-2); line-height:1.4;">
                      ${s.activity}
                    </td>
                    ${b?`
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
    `,E(g)}function E(g){g.querySelectorAll(".edit-dac-btn").forEach(r=>{r.addEventListener("click",()=>{let s=r.dataset.id,m=a.find(p=>String(p.id)===String(s));m&&C(m)})}),g.querySelectorAll(".delete-dac-btn").forEach(r=>{r.addEventListener("click",()=>{let s=r.dataset.id,m=r.dataset.name||"item";ee(`Apakah Anda yakin ingin menghapus aktivitas pada lokasi <strong>${m}</strong>?`,async()=>{try{if(!(await I(`/api/dac/${s}`,{method:"DELETE"})).ok)throw new Error("Gagal menghapus item");j("Aktivitas DAC berhasil dihapus"),u()}catch(p){N(p.message)}})})})}function C(g=null){let r=!!g,s=r?"Edit Aktivitas DAC":"Tambah Aktivitas DAC",m=["Persiapan","Area Umum","Ruang Tindakan","Ruang Penunjang","Toilet","Area Luar","Kontrol & Check","Istirahat","Layanan Pasien","Sampah Umum","Sampah Medis","Penutup"],p=`
      <form id="dac-form" style="display:flex; flex-direction:column; gap:14px;">
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Shift *</label>
            <select name="shift" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="pagi" ${g?.shift==="pagi"||!g&&t==="pagi"?"selected":""}>\u{1F305} Shift Pagi (08:00 - 17:00)</option>
              <option value="siang" ${g?.shift==="siang"||!g&&t==="siang"?"selected":""}>\u{1F307} Shift Siang (13:00 - 22:00)</option>
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
            ${m.map(f=>`<option value="${f}">`).join("")}
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
    `;Y({title:s,content:p,confirmText:r?"Simpan Perubahan":"Tambah Aktivitas",onConfirm:async(f,l)=>{let c=f.querySelector("#dac-form");if(!c.checkValidity()){c.reportValidity();return}let v=new FormData(c),S={shift:v.get("shift"),item_no:v.get("item_no")?parseInt(v.get("item_no"),10):void 0,start_time:v.get("start_time"),end_time:v.get("end_time"),category:v.get("category"),location:v.get("location"),activity:v.get("activity")};try{let B=r?`/api/dac/${g.id}`:"/api/dac",D=await I(B,{method:r?"PUT":"POST",body:JSON.stringify(S)});if(!D.ok){let _=await D.json().catch(()=>({}));throw new Error(_.error||"Gagal menyimpan aktivitas")}j(r?"Aktivitas berhasil diperbarui":"Aktivitas berhasil ditambahkan"),l(),u()}catch(B){N(B.message)}}})}function w(){let g=d();if(!g.length){N("Tidak ada data untuk dicetak");return}let r=t==="pagi"?"SHIFT PAGI (08:00 - 17:00)":t==="siang"?"SHIFT SIANG (13:00 - 22:00)":"SELURUH SHIFT (PAGI & SIANG)",s=window.open("","_blank");if(!s){N("Popup diblokir browser, mohon izinkan popup untuk mencetak");return}s.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>STANDAR ACUAN DAC - ${r}</title>
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
          <div class="subtitle">STANDAR RUTINITAS DAILY ACTIVITY CLEANING (DAC) \u2014 ${r}</div>
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
            ${g.map(m=>`
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
    `),s.document.close(),s.focus(),setTimeout(()=>{s.print()},400)}document.querySelectorAll(".shift-tab-btn").forEach(g=>{g.addEventListener("click",()=>{document.querySelectorAll(".shift-tab-btn").forEach(r=>{r.style.background="#F1F5F9",r.style.color="var(--text-2)"}),g.style.background="#2563EB",g.style.color="#fff",t=g.dataset.shift,n(),h()})});let k=document.getElementById("view-timeline-btn"),y=document.getElementById("view-table-btn");k.addEventListener("click",()=>{i="timeline",k.style.background="#fff",k.classList.remove("btn-ghost"),y.style.background="transparent",y.classList.add("btn-ghost"),h()}),y.addEventListener("click",()=>{i="table",y.style.background="#fff",y.classList.remove("btn-ghost"),k.style.background="transparent",k.classList.add("btn-ghost"),h()}),["dac-search","dac-category-filter","dac-period-filter"].forEach(g=>{let r=document.getElementById(g);r&&(r.addEventListener("input",h),r.addEventListener("change",h))}),document.getElementById("btn-add-dac").addEventListener("click",()=>{C()}),document.getElementById("btn-print-dac").addEventListener("click",()=>{w()}),document.getElementById("btn-export-dac").addEventListener("click",()=>{let g=d();if(!g.length){N("Tidak ada data untuk diexport");return}let r=g.map(s=>({No:s.item_no,Shift:s.shift==="pagi"?"Shift Pagi":"Shift Siang","Jam Mulai":s.start_time,"Jam Selesai":s.end_time,Kategori:s.category,"Lokasi / Area":s.location,"Rincian Aktivitas":s.activity}));G(r,`DAC_Daily_Activity_Cleaning_${t}_${new Date().toISOString().slice(0,10)}`),j("File Excel berhasil diunduh")}),u()}window.parseFlexibleDate=e=>{if(!e||e==="-")return"";if(e=String(e).trim(),/^\d{5}$/.test(e)){let i=Math.floor(Number(e)-25569);return new Date(i*86400*1e3).toISOString().split("T")[0]}if(e.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let i=e.split(/[\/\-]/);return`${i[2]}-${i[1]}-${i[0]}`}let t=e.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(t){let i=t[1],a=parseInt(t[2],10),o=parseInt(t[3],10);if(a>12&&o<=12)return`${i}-${t[3]}-${t[2]}`}return e.split("T")[0]};window.formatDate=e=>{let t=window.parseFlexibleDate(e);if(!t)return"";let i=t.split("-");if(i.length===3&&i[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],o=parseInt(i[2],10),b=a[parseInt(i[1],10)-1];return`${o} ${b} ${i[0]}`}return t};function te(e){return async t=>{if(!Oe()){$e("/login");return}return e(t)}}var Qe=null;function Nn(){Qe&&clearInterval(Qe);let e=()=>{let t=new Date,i=t.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=t.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),o=document.getElementById("header-clock-time"),b=document.getElementById("header-clock-date");o&&(o.textContent=i),b&&(b.textContent=a)};e(),Qe=setInterval(e,1e3)}async function jn(){try{let e=await I("/api/dashboard/kpi");if(!e.ok)return;let t=e.data?.data||e.data||{},i=(a,o)=>{let b=document.getElementById(a);b&&(b.textContent=o>0?o:"",b.style.display=o>0?"inline-flex":"none")};i("badge-issues",t.issues?.current||0),i("badge-contracts",t.expiring30?.current||0),i("badge-oo1",t.one_on_one?.current||0),i("badge-schedule",t.schedule?.current||0),i("badge-supply",t.supply?.current||0)}catch{}}var je=[];async function Kn(){try{let e=await I("/api/dashboard/notifications");if(!e.ok)return;je=e.data?.data||e.data||[];let t=document.getElementById("notif-dot");t&&(t.style.display=je.length>0?"block":"none",t.textContent=je.length)}catch{}}function On(){if(!je.length){Y({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(t,i)=>i()});return}let e=`
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
  `;Y({title:`Notifikasi (${je.length})`,content:e,confirmText:"Tutup",onConfirm:(t,i)=>i()})}function Xa(){let e=ae(),t=(e?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
                <span class="topbar-greeting-time">${(()=>{let h=new Date().getHours();return h>=4&&h<11?"Selamat Pagi":h>=11&&h<15?"Selamat Siang":h>=15&&h<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${e?.full_name||e?.username||"Guest"}</span> \u{1F44B}
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
  `;let i=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),o=document.getElementById("topbar-menu-btn"),b=document.getElementById("sidebar-close"),u=()=>{i.classList.add("open"),a.classList.add("show")},d=()=>{i.classList.remove("open"),a.classList.remove("show")};o?.addEventListener("click",u),b?.addEventListener("click",d),a?.addEventListener("click",d),document.querySelectorAll(".nav-item").forEach(h=>h.addEventListener("click",d)),e&&e.role==="input_lembur"&&(document.querySelectorAll(".nav-item").forEach(h=>{h.dataset.route!=="/overtime"&&(h.style.display="none")}),document.querySelectorAll(".nav-section").forEach(h=>{Array.from(h.querySelectorAll(".nav-item")).some($=>$.style.display!=="none")||(h.style.display="none")}));function n(){let h=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(E=>{let C=E.dataset.route;E.classList.toggle("active",h===C||C!=="/dashboard"&&h.startsWith(C))});let F=document.getElementById("topbar-title"),$=document.querySelector(".nav-item.active .nav-label");F&&$&&(F.textContent=$.textContent)}window.addEventListener("hashchange",n),n(),Nn(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await I("/api/auth/logout",{method:"POST"}),Re(),Qe&&clearInterval(Qe),$e("/login")}),jn(),Kn(),document.getElementById("btn-notif")?.addEventListener("click",h=>{h.preventDefault(),On()})}async function Rn(){Q("/login",({main:t})=>ta(t)),Q("/dashboard",te(({main:t})=>Wt(t))),Q("/calendar",te(({main:t})=>Ba(t))),Q("/employees",te(({main:t,params:i})=>aa(t,i))),Q("/contracts",te(({main:t,params:i})=>dt(t,i))),Q("/sp",te(({main:t})=>Oa(t))),Q("/mutasi",te(({main:t})=>Ra(t))),Q("/overtime",te(({main:t})=>Ua(t))),Q("/timeline",te(({main:t,params:i})=>ia(t,i))),Q("/os-migrations",te(({main:t})=>Ya(t))),Q("/issues",te(({main:t,params:i})=>ra(t,i))),Q("/one-on-one",te(({main:t,params:i})=>oa(t,i))),Q("/training",te(({main:t})=>da(t))),Q("/relievers",te(({main:t,params:i})=>ca(t,i))),Q("/reports/inspection",te(({main:t})=>ua(t))),Q("/reports/cleaning",te(({main:t})=>ga(t))),Q("/reports/fogging",te(({main:t})=>ma(t))),Q("/reports/basecamp",te(({main:t})=>ba(t))),Q("/reports/supply",te(({main:t})=>Ta(t,"supply"))),Q("/sop",()=>{window.location.hash="#/sop-fc"}),Q("/sop-fc",te(({main:t})=>xa(t))),Q("/sop-security",te(({main:t})=>Sa(t))),Q("/checklist",te(({main:t})=>$a(t))),Q("/hygiene-standards",te(({main:t})=>_a(t))),Q("/dac",te(({main:t})=>Qa(t))),Q("/users",te(({main:t})=>Aa(t))),Q("/branches",te(({main:t})=>Da(t))),Q("/profile",te(({main:t})=>Pa(t))),Q("/settings/import",te(({main:t})=>Na(t))),Q("/audit-logs",te(({main:t})=>Va(t)));let e=Oe();if(!e&&window.location.hash!=="#/login"&&$e("/login"),e){let t=await I("/api/auth/me");if(t.ok){qe(t.data.data),Xa();let i=ae(),a=window.location.hash;i&&i.role==="input_lembur"&&(a==="#/dashboard"||a===""||a==="#/")&&$e("/overtime")}else Re(),$e("/login")}window.addEventListener("fm:login",()=>{Xa();let t=ae();t&&t.role==="input_lembur"?$e("/overtime"):$e("/dashboard")}),Ut()}Rn();
