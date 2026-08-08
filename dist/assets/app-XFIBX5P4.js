var ya=Object.defineProperty;var dt=(t,e)=>()=>(t&&(e=t(t=0)),e);var ct=(t,e)=>{for(var r in e)ya(t,r,{get:e[r],enumerable:!0})};var _e={};ct(_e,{API:()=>Ct,CLIENT_SIDE_MAX_ROWS:()=>xe,IS_DEVELOPMENT:()=>Me,apiFetch:()=>S,clearToken:()=>Re,getToken:()=>Oe,getUser:()=>ce,setToken:()=>pt,setUser:()=>Ke});function Oe(){return localStorage.getItem("fm_token")}function pt(t){localStorage.setItem("fm_token",t)}function Re(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ce(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Ke(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function S(t,e={}){let r=Oe(),a={"Content-Type":"application/json",...r?{Authorization:`Bearer ${r}`}:{},...e.headers||{}};try{let l=`cb=${Date.now()}`,s=t.includes("?")?"&":"?",o=`${Ct}${t}${s}${l}`,n=await fetch(o,{...e,headers:a}),i;try{let c=await n.text();try{i=JSON.parse(c)}catch{i={error:`Server Error (${n.status}): ${c.substring(0,80)}...`}}}catch{i={error:"Gagal membaca respon dari server"}}return n.status===401&&(Re(),window.location.hash="#/login"),{ok:n.ok,status:n.status,data:i}}catch(l){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${l.message})`}}}}var Me,fa,Ct,xe,j=dt(()=>{Me=!1,fa="https://fm-operations-api.facilitycare-audydental.workers.dev",Ct=fa,xe=1e4});var mt={};ct(mt,{confirmDialog:()=>He,createModal:()=>ue});function ue({title:t,content:e,onConfirm:r,onCancel:a,confirmText:l="Simpan",cancelText:s="Batal",size:o="md",confirmClass:n="btn-primary"}){let i={sm:"400px",md:"560px",lg:"720px",xl:"900px"},c=document.createElement("div");c.className="modal-overlay",c.innerHTML=`
    <div class="modal" style="max-width:${i[o]||i.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${s}</button>
        ${r?`<button class="btn ${n} modal-confirm">${l}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&c.querySelector(".modal-body").appendChild(e);let g=()=>{c.classList.remove("show"),setTimeout(()=>c.remove(),250)};return c.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),g()}),c.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),g()}),r&&c.querySelector(".modal-confirm").addEventListener("click",()=>r(c,g)),c.addEventListener("click",p=>{p.target===c&&(a&&a(),g())}),document.body.appendChild(c),requestAnimationFrame(()=>c.classList.add("show")),{overlay:c,close:g}}function He(t,e,r="Konfirmasi"){return ue({title:r,content:`<p>${t}</p>`,onConfirm:(a,l)=>{e(),l()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var Se=dt(()=>{});var pe={};ct(pe,{downloadExcel:()=>M,parseExcel:()=>Je,renderExcelButtons:()=>Ue});function Je(t){return new Promise((e,r)=>{let a=new FileReader;a.onload=l=>{try{let s=new Uint8Array(l.target.result),o=XLSX.read(s,{type:"array"}),n=o.SheetNames[0],i=o.Sheets[n];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${o.SheetNames.join(", ")}`),console.log(`Sheet Used: ${n}`);let c=XLSX.utils.decode_range(i["!ref"]||"A1:A1"),g=c.e.r-c.s.r+1,p=c.e.c-c.s.c+1;console.log(`Total Rows (including empty): ${g}`),console.log(`Total Columns: ${p}`);let d=[];for(let y=c.s.c;y<=c.e.c;++y){let m=i[XLSX.utils.encode_cell({c:y,r:c.s.r})];m&&m.v&&d.push(m.v)}console.log(`Headers Found: ${d.join(", ")}`),console.log("---------------------------");let u=XLSX.utils.sheet_to_json(i,{defval:""});Object.defineProperty(u,"__worksheet",{value:i,enumerable:!1}),Object.defineProperty(u,"__headers",{value:d,enumerable:!1}),e(u)}catch(s){r(s)}},a.onerror=l=>r(l),a.readAsArrayBuffer(t)})}function M(t,e){try{let r=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,r,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(r){throw console.error("Error generating Excel file:",r),r}}function Ue(t){return`
    <div class="excel-actions" style="display:flex; flex-wrap:wrap; gap:0.5rem; margin-bottom:1rem;">
      <button class="btn btn-secondary btn-sm" id="btn-export-${t}">
        \u{1F4E5} Export Excel
      </button>
      <button class="btn btn-secondary btn-sm" id="btn-template-${t}">
        \u{1F4C4} Download Template
      </button>
      <label class="btn btn-primary btn-sm" style="cursor:pointer;margin:0;" id="label-import-${t}">
        <span class="import-text">\u{1F4E4} Import Excel</span>
        <input type="file" id="input-import-${t}" accept=".xlsx, .xls, .csv" style="display:none;">
      </label>
    </div>
  `}var H=dt(()=>{});j();var ut={},Ye=null;function U(t,e){ut[t]=e}function Ce(t){window.location.hash=t}function $t(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[r,...a]=e.split("?"),l=ut[r];if(!l){for(let[o,n]of Object.entries(ut))if(o.endsWith("/*")&&r.startsWith(o.slice(0,-2))){l=n;break}}Ye&&(Ye(),Ye=null);let s=document.getElementById("main-content");if(s&&(s.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),l){let o=new URLSearchParams(a.join("?")),n=r.split("/").filter(Boolean),i=await l({path:r,params:o,segments:n,main:s});i&&(Ye=i)}else{let o=s||document.getElementById("app");o&&(o.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var je;function va(){return je||(je=document.createElement("div"),je.id="toast-container",document.body.appendChild(je)),je}function Tt(t,e="info",r=3500){let a=va(),l=document.createElement("div");l.className=`toast toast-${e}`;let s={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};l.innerHTML=`<span class="toast-icon">${s[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(l),requestAnimationFrame(()=>l.classList.add("show")),setTimeout(()=>{l.classList.remove("show"),setTimeout(()=>l.remove(),350)},r)}var ee=t=>Tt(t,"success"),z=t=>Tt(t,"error");Se();j();j();Se();function Ve({columns:t,data:e,onEdit:r,onDelete:a,onView:l,actions:s=[],emptyText:o="Tidak ada data",bulkSelect:n=null}){let i=document.createElement("div");if(i.className="table-wrapper",!e||e.length===0)return i.innerHTML=`<div class="empty-state"><p>${o}</p></div>`,i;let c=document.createElement("table");c.className="data-table";let g=document.createElement("thead"),p=document.createElement("tr");if(n){let u=document.createElement("th");u.style.width="40px",u.style.textAlign="center";let y=document.createElement("input");y.type="checkbox",y.id="select-all-checkbox",y.title="Pilih semua",y.addEventListener("change",()=>{e.forEach(m=>{y.checked?n.selectedIds.add(m.id):n.selectedIds.delete(m.id)}),i.querySelectorAll(".row-checkbox").forEach(m=>m.checked=y.checked),n.onToggle()}),u.appendChild(y),p.appendChild(u)}if(t.forEach(u=>{let y=document.createElement("th");y.textContent=u.label,u.width&&(y.style.width=u.width),p.appendChild(y)}),r||a||l||s.length>0){let u=document.createElement("th");u.textContent="Aksi",u.style.width="120px",p.appendChild(u)}g.appendChild(p),c.appendChild(g);let d=document.createElement("tbody");return e.forEach(u=>{let y=document.createElement("tr");if(n){let m=document.createElement("td");m.style.textAlign="center",m.style.width="40px";let b=document.createElement("input");b.type="checkbox",b.className="row-checkbox",b.checked=n.selectedIds.has(u.id),b.addEventListener("change",()=>{if(b.checked)n.selectedIds.add(u.id);else{n.selectedIds.delete(u.id);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1)}n.onToggle()}),m.appendChild(b),y.appendChild(m)}if(t.forEach(m=>{let b=document.createElement("td");if(m.render){let v=m.render(u[m.key],u);v instanceof HTMLElement?b.appendChild(v):b.innerHTML=v||""}else b.textContent=u[m.key]!==null&&u[m.key]!==void 0&&u[m.key]!==""?u[m.key]:"";m.nowrap&&(b.style.whiteSpace="nowrap"),y.appendChild(b)}),r||a||l||s.length>0){let m=document.createElement("td");m.className="actions-cell";let b=document.createElement("div");if(b.className="btn-group",l){let v=document.createElement("button");v.className="btn btn-xs btn-ghost",v.innerHTML="\u{1F441}",v.title="Lihat",v.addEventListener("click",()=>l(u)),b.appendChild(v)}if(r){let v=document.createElement("button");v.className="btn btn-xs btn-secondary",v.innerHTML="\u270F\uFE0F",v.title="Edit",v.addEventListener("click",()=>r(u)),b.appendChild(v)}s.forEach(v=>{let k=document.createElement("button");k.className=`btn btn-xs ${v.class||"btn-ghost"}`,k.innerHTML=v.icon||v.label,k.title=v.label,k.addEventListener("click",()=>v.handler(u)),b.appendChild(k)}),m.appendChild(b),y.appendChild(m)}d.appendChild(y)}),c.appendChild(d),i.appendChild(c),i}function Xe({page:t,pages:e,total:r,limit:a,onPage:l}){if(e<=1)return null;let s=document.createElement("div");s.className="pagination";let o=document.createElement("span");o.className="pagination-info",o.textContent=`Total: ${r} data`,s.appendChild(o);let n=document.createElement("div");n.className="pagination-btns";let i=(p,d,u=!1,y=!1)=>{let m=document.createElement("button");m.className=`btn btn-sm ${y?"btn-primary":"btn-ghost"} pagination-btn`,m.textContent=p,m.disabled=u,m.addEventListener("click",()=>l(d)),n.appendChild(m)};i("\xAB",1,t===1),i("\u2039",t-1,t===1);let c=Math.max(1,t-2),g=Math.min(e,t+2);for(let p=c;p<=g;p++)i(p,p,!1,p===t);return i("\u203A",t+1,t===e),i("\xBB",e,t===e),s.appendChild(n),s}Se();function qe(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${qe(e.fields)}</div>`;let r=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",l="";switch(e.type){case"textarea":l=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${r} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let o=(e.options||[]).map(p=>{let d=typeof p=="object"?p.value:p,u=typeof p=="object"?p.label:p,y=e.value==d?"selected":"";return`<option value="${d}" ${y}>${u}</option>`}).join("");l=`<select name="${e.name}" class="form-control" ${r}><option value="">-- Pilih ${e.label||""} --</option>${o}</select>`;break;case"combobox":let n=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,i=(e.options||[]).map(p=>{let d=typeof p=="object"?p.value:p,u=typeof p=="object"?p.label||p.value||"":p||"";return(u==="undefined"||u==="[object Object]"||u==="null")&&(u=""),u?`<option value="${u}"></option>`:""}).join(""),c=e.value||"";if(e.value){let p=(e.options||[]).find(d=>(typeof d=="object"?d.value:d)==e.value);if(p){let d=typeof p=="object"?p.label||p.value||"":p||"";d&&d!=="undefined"&&d!=="[object Object]"&&d!=="null"&&(c=d)}}l=`
          <input type="text" name="${e.name}" list="${n}" class="form-control" value="${c}" placeholder="Pilih atau ketik baru..." ${r} autocomplete="off">
          <datalist id="${n}">${i}</datalist>
        `;break;case"checkbox":l=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let g=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";l=`<input type="date" name="${e.name}" class="form-control" value="${g}" ${r}>`;break;case"number":l=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${r}>`;break;case"email":l=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${r}>`;break;case"url":l=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${r}>`;break;default:l=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${r} autocomplete="off">`}let s=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${l}${s}</div>`}).join("")}function Ze(t){let e={},r=new FormData(t);for(let[a,l]of r.entries())e[a]=l===""?null:l;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function et(t,e){e&&Object.entries(e).forEach(([r,a])=>{let l=t.querySelector(`[name="${r}"]`);l&&(l.hasAttribute("list")||(l.type==="checkbox"?l.checked=!!a:l.type==="date"&&a&&window.parseFlexibleDate?l.value=window.parseFlexibleDate(a):l.value=a??""))})}H();function O({container:t,title:e,icon:r,apiPath:a,columns:l,formFields:s,filterFields:o,defaultFilters:n={},enableMobileFilterSheet:i=!1,itemLabel:c="Data",canCreate:g=!0,canEdit:p=!0,canDelete:d=!0,onBeforeSubmit:u,onAfterLoad:y,onDataLoaded:m,extraActions:b=[],initialSearch:v="",exportOptions:k=null,bulkDelete:w=!1,paginationMode:T="server"}){let $=ce();$&&typeof $=="object"&&$.role==="viewer"&&(g=!1,p=!1,d=!1,w=!1,k=null);let C=1,x={...n};v&&(x.search=v);let _=new Set;t.innerHTML=`
    <div class="crud-layout-wrapper ${i?"mobile-active":""}">
      <div class="page-header">
      <h1 class="page-title">${r} ${e}</h1>
      <div class="page-actions">
        ${g?`<button class="btn btn-primary" id="btn-create">+ Tambah ${c}</button>`:""}
        ${k?'<button class="btn btn-outline" id="btn-mobile-aksi" style="display:none; align-items:center; justify-content:center; gap:0.25rem;">\u22EE Aksi</button>':""}
      </div>
    </div>

    ${w?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:none; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${k?`
    <div class="excel-actions-wrapper" id="excel-actions-wrapper">
      <div class="bottom-sheet-header aksi-header" style="display:none;">
        <h3 style="margin:0; font-size:1rem;">Aksi</h3>
        <button class="btn-close-sheet" id="btn-close-aksi-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
      </div>
      ${Ue(k.moduleName)}
    </div>`:""}

    ${o&&o.length>0?`
    <div class="filter-bar card ${i?"has-mobile-sheet":""}" style="padding: 1rem; display: flex; gap: 0.5rem; flex-wrap: wrap;">
        ${o.filter(h=>h.type==="search"||h.type==="search-combo").map(h=>{if(h.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${h.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}"></div>`;if(h.type==="search-combo"){let f="dl-filter-search",E=(h.options||[]).map(D=>`<option value="${typeof D=="object"?D.label:D}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${f}" class="form-control" autocomplete="off" placeholder="${h.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}"><datalist id="${f}">${E}</datalist></div>`}return""}).join("")}
        
        <div class="filter-options" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(h=>h.type!=="search"&&h.type!=="search-combo").map(h=>{if(h.type==="select")return`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${h.name}" id="filter-${h.name}"><option value="">Pilih ${h.label}</option>${(h.options||[]).map(f=>`<option value="${typeof f=="object"?f.value:f}" ${x[h.name]===(typeof f=="object"?f.value:f)?"selected":""}>${typeof f=="object"?f.label:f}</option>`).join("")}</select>`;if(h.type==="combobox"){let f=`dl-filter-${h.name}`,E=(h.options||[]).map(P=>{let I=typeof P=="object"?P.label||P.value||"":P||"";return(I==="undefined"||I==="[object Object]"||I==="null")&&(I=""),I?`<option value="${I}"></option>`:""}).join(""),D=x[h.name]||"";if(x[h.name]){let P=(h.options||[]).find(I=>String(typeof I=="object"?I.value:I)==String(x[h.name]));if(P){let I=typeof P=="object"?P.label||P.value||"":P||"";I&&I!=="undefined"&&I!=="[object Object]"&&I!=="null"&&(D=I)}}return`<div class="filter-combobox" style="flex:1; min-width:120px;">
                <input type="text" name="${h.name}" id="filter-${h.name}" list="${f}" class="form-control filter-combobox-input" value="${D}" placeholder="Pilih ${h.label}..." autocomplete="off">
                <datalist id="${f}">${E}</datalist>
              </div>`}return""}).join("")}
          <div style="display:flex; gap:0.5rem; margin-top:0.5rem;">
            <button class="btn btn-outline" id="btn-reset-filter" style="flex:1;">Reset</button>
            <button class="btn btn-primary" id="btn-apply-filter" style="flex:1;">\u2713 Terapkan (OK)</button>
          </div>
        </div>
        ${i?'<button id="btn-mobile-filter" class="btn btn-outline" style="display:none; align-items:center; gap:0.25rem;">\u2699\uFE0F Filter</button>':""}
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
    </div>
  `;function N(){let h=document.getElementById("bulk-toolbar");if(!h)return;let f=document.getElementById("bulk-count"),E=document.getElementById("btn-bulk-delete"),D=document.getElementById("btn-bulk-cancel");f.textContent=`${_.size} item dipilih`,_.size>0?(h.style.display="flex",h.classList.add("has-items"),E.disabled=!1,D.disabled=!1,window.innerWidth<=768?(h.style.setProperty("top","auto","important"),h.style.setProperty("bottom","0","important")):(h.style.removeProperty("top"),h.style.removeProperty("bottom"))):(h.style.display="none",h.classList.remove("has-items"),E.disabled=!0,D.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{_.clear(),document.querySelectorAll(".row-checkbox").forEach(f=>f.checked=!1);let h=document.getElementById("select-all-checkbox");h&&(h.checked=!1),N()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(_.size===0)return;let h=[..._],f=document.createElement("div");f.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",f.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${h.length} ${c}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${h.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(f),f.querySelector("#bulk-cancel-btn").addEventListener("click",()=>f.remove()),f.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let E=f.querySelector("#bulk-confirm-btn");E.disabled=!0,E.textContent="Menghapus...";let D=await S(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:h})});f.remove(),D.ok?(ee(`${h.length} ${c} berhasil dihapus.`),_.clear(),N(),Q()):z(D.data?.error||"Gagal menghapus data.")})});let F=document.getElementById("filter-search"),le;F?.addEventListener("input",h=>{clearTimeout(le),le=setTimeout(()=>{x.search=h.target.value,C=1,_.clear(),Q()},400)}),o?.forEach(h=>{h.type==="select"&&document.getElementById(`filter-${h.name}`)?.addEventListener("change",f=>{x[h.name]=f.target.value,C=1,_.clear(),Q()}),h.type==="combobox"&&document.getElementById(`filter-${h.name}`)?.addEventListener("change",f=>{let E=f.target.value,D=(h.options||[]).find(P=>{let I=String(typeof P=="object"?P.value:P),B=String(typeof P=="object"?P.label:P);return I===E||B===E});E?x[h.name]=D?typeof D=="object"?D.value:D:E:x[h.name]="",C=1,_.clear(),Q()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{x={...n},F&&(F.value=""),o?.forEach(h=>{let f=document.getElementById(`filter-${h.name}`);f&&(f.value="")}),C=1,_.clear(),Q()});let de=document.getElementById("btn-mobile-filter"),se=document.getElementById("filter-options-wrapper"),ke=document.getElementById("btn-close-filter-sheet");if(de&&se){de.addEventListener("click",f=>{f.preventDefault(),se.classList.add("sheet-open")}),ke&&ke.addEventListener("click",f=>{f.preventDefault(),se.classList.remove("sheet-open")});let h=document.getElementById("btn-apply-filter");h&&h.addEventListener("click",f=>{f.preventDefault(),se.classList.remove("sheet-open")})}let fe=document.getElementById("btn-mobile-aksi"),we=document.getElementById("excel-actions-wrapper"),Ne=document.getElementById("btn-close-aksi-sheet");if(fe&&we&&(fe.addEventListener("click",h=>{h.preventDefault(),we.classList.add("sheet-open")}),Ne&&Ne.addEventListener("click",h=>{h.preventDefault(),we.classList.remove("sheet-open")})),document.getElementById("btn-create")?.addEventListener("click",()=>Fe(null)),k){document.getElementById(`btn-export-${k.moduleName}`)?.addEventListener("click",async f=>{let E=f.target,D=E.innerHTML;E.innerHTML="\u23F3 Loading...",E.disabled=!0;try{await k.onExport()}catch{z("Gagal export data")}finally{E.innerHTML=D,E.disabled=!1}}),document.getElementById(`btn-template-${k.moduleName}`)?.addEventListener("click",()=>{k.onTemplate()});let h=document.getElementById(`input-import-${k.moduleName}`);h?.addEventListener("change",async f=>{let E=f.target.files[0];if(!E)return;h.disabled=!0;let D=document.createElement("div");D.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",D.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(D);let P=D.querySelector("#import-progress-text"),I=D.querySelector("#import-progress-bar"),B=D.querySelector("#import-summary"),A=D.querySelector("#import-close-btn");A.addEventListener("click",()=>{D.remove(),Q()});try{let X=await Je(E);if(X.length===0)throw new Error("File kosong atau format salah");let Z=500,oe=0,ne=0,R=0,K=X.length;P.textContent=`Ditemukan ${K} baris data. Memulai import...`;for(let L=0;L<K;L+=Z){let Y=X.slice(L,L+Z);P.textContent=`Mengimport baris ${L+1} - ${Math.min(L+Z,K)} dari ${K}...`,I.style.width=`${Math.round(L/K*100)}%`;try{let J=await k.onImport(Y);J?(oe+=J.inserted||J.metrics?.inserted||Y.length,ne+=J.skipped||J.metrics?.updated||0):oe+=Y.length}catch(J){console.error("Chunk import failed:",J),R+=Y.length}}I.style.width="100%",P.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',B.style.display="block",B.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${K}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${oe}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ne}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${R}</strong></div>
        `,R>0&&(B.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),A.style.display="block",h.value=""}catch(X){P.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${X.message}`,I.style.background="var(--danger)",I.style.width="100%",A.style.display="block",h.value=""}finally{h.disabled=!1}})}async function Q(){N();let h=document.getElementById("table-container");if(!h)return;h.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let f=T==="client",E=f?1:C,D=f?xe:20,P=new URLSearchParams({page:E,limit:D,...Object.fromEntries(Object.entries(x).filter(([,R])=>R))}),I=await S(`${a}?${P}`);if(!I.ok){h.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${I.data?.error||"Error"}</p></div>`;return}let B=I.data?.data||I.data||[],A=I.data?.pagination,X=B.length,Z=B;if(f){B=m(B),Z=B;let R=B.length,K=20,L=Math.ceil(R/K);C>L&&L>0&&(C=L);let Y=(C-1)*K,J=C*K;B=B.slice(Y,J),A={page:C,limit:K,total:R,pages:L}}!1,y&&y(B);let oe=Ve({columns:l,data:B,fullData:Z,onEdit:p?R=>Fe(R):null,actions:b.map(R=>({...R,handler:K=>R.handler(K,Q)})),emptyText:`Tidak ada ${String(c||"").toLowerCase()}`,bulkSelect:w?{selectedIds:_,onToggle:N}:null});h.innerHTML="",h.appendChild(oe);let ne=document.getElementById("pagination-container");if(ne&&(ne.innerHTML="",A&&A.pages>1)){let R=Xe({page:A.page,pages:A.pages,total:A.total,limit:A.limit,onPage:K=>{C=K,Q()}});R&&ne.appendChild(R)}}function st(h){let f=typeof s=="function"?s(h):s;return qe(f)}function Fe(h){let f=!!h,E=document.createElement("form");if(E.noValidate=!0,E.innerHTML=st(h),f){let P=typeof s=="function"?s(h):s;et(E,h)}let{close:D}=ue({title:f?`Edit ${c}`:`Tambah ${c}`,content:E,size:"lg",confirmText:f?"Simpan Perubahan":`Tambah ${c}`,onConfirm:async(P,I)=>{if(!E.reportValidity())return;let B=P.querySelector(".modal-confirm");B.disabled=!0,B.textContent="Menyimpan...";let A=Ze(E),X=typeof s=="function"?s(h):s,Z=async K=>{for(let L of K)if(L.type==="row")await Z(L.fields);else if(L.type==="combobox"&&A[L.name]){let Y=A[L.name],J=(L.options||[]).find(V=>{let ie=String(typeof V=="object"?V.value:V),ot=String(typeof V=="object"?V.label:V);return ie===Y||ot===Y});if(J)A[L.name]=typeof J=="object"?J.value:J;else if(L.createApi){let V={};V[L.createApi.field]=Y,L.createApi.extra&&Object.assign(V,L.createApi.extra);let ie=await S(L.createApi.path,{method:"POST",body:JSON.stringify(V)});if(ie.ok&&ie.data?.id)A[L.name]=ie.data.id;else if(ie.ok&&!ie.data?.id)A[L.name]=Y;else throw new Error(`Gagal membuat master data: ${ie.data?.error||"Unknown error"}`)}}};try{await Z(X)}catch(K){z(K.message),B.disabled=!1,B.textContent=f?"Simpan Perubahan":`Tambah ${c}`;return}u&&(A=await u(A,h));let oe=f?"PUT":"POST",ne=f?`${a}/${h.id}`:a,R=await S(ne,{method:oe,body:JSON.stringify(A)});R.ok?(ee(f?`${c} berhasil diperbarui.`:`${c} berhasil ditambahkan.`),I(),Q()):(z(R.data?.error||"Gagal menyimpan data."),B.disabled=!1,B.textContent=f?"Simpan Perubahan":`Tambah ${c}`)}})}function ha(h){He(`Hapus ${c} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let f=await S(`${a}/${h.id}`,{method:"DELETE"});f.ok?(ee(`${c} berhasil dihapus.`),Q()):z(f.data?.error||"Gagal menghapus.")},`Hapus ${c}`)}return Q(),Q}j();j();var $e=null,tt=null;async function Te(t=!1){if($e&&!t)return console.log("Employees Raw (Cache Hit)",$e.slice(0,5)),$e;let e=await S(`/api/employees?limit=${xe}&status=Aktif`);return $e=(e.data?.data||[]).map(r=>({value:r.id,label:r.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",$e.slice(0,5)),$e}async function te(t=!1){let r=(await Te(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",r.slice(0,5)),r}async function q(t=!1){return tt&&!t||(tt=((await S("/api/branches?all=1")).data?.data||[]).map(r=>({value:r.id,label:r.full_name}))),tt}function G(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function gt(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function Ee(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function bt(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function me(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}H();function ht(t,e){let r=new Date,a=`${r.getFullYear()}-${String(r.getMonth()+1).padStart(2,"0")}`;if(!(t.target_date||t.opening_date||"").startsWith(a))return!1;let s=String(t.status||"").toLowerCase();if(s!=="selesai"&&s!=="completed"&&s!=="done")return!1;let o=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?o.includes("inspeksi"):e==="gcdc"?o.includes("general cleaning")||o.includes("deep cleaning"):!1}j();H();function Et(t,e){let r=String(t.status||"").toLowerCase();return e==="active"?r==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&r==="aktif":!1}j();H();function yt(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let l=new Date(a);l.setDate(a.getDate()+30);let s=new Date(t.end_date);return s.setHours(0,0,0,0),s>=a&&s<=l}return!1}j();H();function Dt(t,e){let r=String(t.status||"").toLowerCase();return e==="open"?r==="open":!1}j();function It(t,e){let r=String(t.status||"").toLowerCase();return e==="pending"?r==="pending":!1}var ve={};function Pe(t){if(ve[t]){try{ve[t].destroy()}catch{}delete ve[t]}}function ka(){Object.keys(ve).forEach(Pe)}var he=(t,e=0)=>{let r=Number(t);return isNaN(r)||t===null||t===void 0?e:r},De=(t,e="\u2014")=>{if(t==null||t==="")return e;let r=String(t).trim();return r===""||r==="[object Object]"?e:r};function Lt(t,e,r=900){if(!t)return;let a=Math.max(0,Math.round(he(e)));if(a===0){t.textContent="0";return}let l=Date.now(),s=()=>{let o=Math.min((Date.now()-l)/r,1),n=1-Math.pow(1-o,3);t.textContent=Math.round(n*a).toLocaleString("id-ID"),o<1?requestAnimationFrame(s):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(s)}var Sa={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},wa=t=>{let e=De(t,"\u2014");return`<span class="status-pill ${Sa[e]||"pill-neutral"}">${e}</span>`};var ge={family:"Inter",size:11},ye="#94A3B8",Ie="#F1F5F9",ft=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],xa=()=>window.innerWidth<768;function at(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:xa()?"bottom":"top",labels:{font:ge,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ge,titleFont:{...ge,weight:"700"}}},scales:{x:{grid:{color:Ie},ticks:{font:ge,color:ye,maxRotation:0}},y:{grid:{color:Ie},ticks:{font:ge,color:ye},beginAtZero:!0}},...t}}var _a=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join("");function Pt(t=3){return Array(t).fill(0).map((e,r)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${r<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function re(t,e,r=8e3){try{let a=new AbortController,l=setTimeout(()=>a.abort(),r),s=await S(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(l),!s||!s.ok)return e;let o=s.data;return o?o.data!==void 0?o.data??e:o:e}catch{return e}}function $a(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(a=>{let l=document.getElementById(a);l&&(l.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(a=>{let l=document.getElementById(a);if(l&&l.style.display==="none"){l.style.display="block";let s=l.parentElement;if(s&&!s.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent="Belum ada data",l.style.display="none",s.appendChild(o)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&At({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Nt({}),["table-contracts","table-issues"].forEach(a=>{let l=document.getElementById(a);l&&l.querySelector(".skeleton")&&(l.innerHTML='<div class="chart-empty">Belum ada data</div>')});let r=document.getElementById("activity-log");r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Bt(t){ka(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">


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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${Pt(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${Pt(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>vt(t)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async r=>{let a=r.target.value,l=document.getElementById("jadwal-year-label");l&&(l.textContent=a);let s=document.getElementById("skel-jadwal"),o=document.getElementById("chart-jadwal");s&&(s.style.display="block",s.style.position="absolute"),o&&(o.style.display="none");let n=await re(`/api/dashboard/schedule-chart?year=${a}`,{},8e3);try{Ft(n)}catch(i){console.warn("ScheduleChart render:",i),be("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async r=>{let a=r.target.value,l=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",s=document.getElementById("skel-insp"),o=document.getElementById("chart-insp");s&&(s.style.display="block",s.style.position="absolute"),o&&(o.style.display="none");let n=await re(l,{},8e3);try{Mt(n)}catch(i){console.warn("InspBar render:",i),be("skel-insp","chart-insp")}}),t._skelTimeout=setTimeout(()=>$a(),5e3),await vt(t),t._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?vt(t):clearInterval(t._dashRefresh)},6e4)}async function vt(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,r,a,l,s,o,n,i,c,g,p,d,u,y]=await Promise.all([re("/api/dashboard/kpi",{},8e3),re("/api/dashboard/issues-trend",{},8e3),re("/api/dashboard/issues-summary",{},8e3),re("/api/dashboard/stats",{},8e3),re("/api/dashboard/calendar",[],8e3),re("/api/schedule?limit=10000",{data:[]},8e3),re("/api/employees?limit=10000",{data:[]},8e3),re("/api/contracts?limit=10000",{data:[]},8e3),re("/api/issues?limit=10000",{data:[]},8e3),re("/api/one-on-one?limit=10000",{data:[]},8e3),re("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),re(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),re("/api/relievers?limit=10000",{data:[]},8e3),re("/api/reports/fogging?limit=10000",{data:[]},8e3)]),m=document.getElementById("filter-insp-month"),b=m?m.value:"",v=b?`/api/dashboard/inspection-bar?month=${b}`:"/api/dashboard/inspection-bar",k=await re(v,{},8e3);if(e){let w=Array.isArray(o?.data)?o.data:Array.isArray(o)?o:[];window.dashboardSchedules=w;let T=Array.isArray(n?.data)?n.data:Array.isArray(n)?n:[],$=Array.isArray(i?.data)?i.data:Array.isArray(i)?i:[],C=Array.isArray(c?.data)?c.data:Array.isArray(c)?c:[],x=Array.isArray(g?.data)?g.data:Array.isArray(g)?g:[],_=Array.isArray(u?.data)?u.data:Array.isArray(u)?u:[];window.dashboardRelievers=_;let N=Array.isArray(y?.data)?y.data:Array.isArray(y)?y:[];window.dashboardFogging=N,e.employees&&(e.employees.current=T.filter(F=>Et(F,"active")).length),e.contracts&&(e.contracts.current=$.filter(F=>yt(F,"active")).length),e.expiring30&&(e.expiring30={current:$.filter(F=>yt(F,"expiring30")).length}),e.issues&&(e.issues.current=C.filter(F=>Dt(F,"open")).length),e.one_on_one&&(e.one_on_one.current=x.filter(F=>It(F,"pending")).length),e.inspection_month&&(e.inspection_month.current=w.filter(F=>ht(F,"inspeksi")).length),e.cleaning_month&&(e.cleaning_month.current=w.filter(F=>ht(F,"gcdc")).length)}try{At(e)}catch(w){console.warn("KPI render:",w)}try{Nt(e)}catch(w){console.warn("MiniStats render:",w)}try{Ft(d)}catch(w){console.warn("ScheduleChart render:",w),be("skel-jadwal","chart-jadwal")}try{Ta(Array.isArray(a?.by_category)?a.by_category:[])}catch(w){console.warn("Donut render:",w),be("skel-donut","chart-donut")}try{Ea(r)}catch(w){console.warn("Trend render:",w),be("skel-trend","chart-trend")}try{Mt(k)}catch(w){console.warn("InspBar render:",w),be("skel-insp","chart-insp")}try{let w=Array.isArray(l)?l:Array.isArray(l?.recent_issues)?l.recent_issues:[];Ia(w)}catch(w){console.warn("IssuesTable render:",w)}try{let w=Array.isArray(l?.expiring_contracts)?l.expiring_contracts:[];Da(p)}catch(w){console.warn("ContractsTable render:",w)}try{Pa(Array.isArray(s)?s:[])}catch(w){console.warn("Agenda render:",w)}try{La()}catch(w){console.warn("Quick Actions render:",w)}}function At(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let r=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=r.map(a=>{let l=he(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${l}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${l}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{Lt(a,parseInt(a.dataset.target)||0)})}function Nt(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let r=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,a=new Date().getFullYear(),l=String(new Date().getMonth()+1).padStart(2,"0"),s=`${a}-${l}`,o=p=>`
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
  `,n=[{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.preventDefault(); event.stopPropagation();">
          <option value="Q1" ${r==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${r==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${r==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${r==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:t.schedule?.current,href:`#/timeline?dash_filter=period_${r.toLowerCase()}`,color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:o("dash-reliefer-month"),val:t.reliever_completed?.current,href:`#/relievers?dash_filter=reliever&month=${s}`,color:"mini-teal"},{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:o("dash-inspeksi-month"),val:t.inspection_month?.current,href:`#/timeline?dash_filter=inspeksi&month=${s}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:o("dash-gcdc-month"),val:t.cleaning_month?.current,href:`#/timeline?dash_filter=gcdc&month=${s}`,color:"mini-green"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:o("dash-fogging-month"),val:t.fogging_month?.current,href:`#/reports/fogging?dash_filter=fogging&month=${s}`,color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=n.map(p=>`
    <a href="${p.href}" class="mini-stat ${p.color}" style="text-decoration:none" id="${p.id||""}">
      <div class="mini-stat-icon">${p.icon}</div>
      <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
        <div style="display:flex; align-items:baseline; gap:3px;">
          <div class="mini-stat-value" data-target="${he(p.val)}">0</div>
          ${p.dropdown?p.dropdown:""}
        </div>
        <div class="mini-stat-text">${p.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(p=>Lt(p,parseInt(p.dataset.target)||0,700));let i=document.getElementById("dash-jadwal-period");i&&i.addEventListener("change",p=>{let d=p.target.value,u=(window.dashboardSchedules||[]).filter(b=>b.period===d).length,y=document.querySelector("#mini-jadwal .mini-stat-value");y&&(y.dataset.target=u,y.textContent=u);let m=document.getElementById("mini-jadwal");m&&(m.href=`#/timeline?dash_filter=period_${d.toLowerCase()}`)});let c=(p,d,u,y,m)=>{let b=document.getElementById(p);b&&b.addEventListener("change",v=>{let k=v.target.value,w=(u||[]).filter(C=>y(C,k)).length,T=document.querySelector(`#${d} .mini-stat-value`);T&&(T.dataset.target=w,T.textContent=w);let $=document.getElementById(d);$&&($.href=`${m}&month=${k}`)})},g=p=>{let d=String(p.status||"").toLowerCase();return d==="done"||d==="selesai"||d==="completed"};c("dash-reliefer-month","mini-reliefer",window.dashboardRelievers,(p,d)=>window.parseFlexibleDate(p.backup_date).startsWith(d)&&g(p),"#/relievers?dash_filter=reliever"),c("dash-inspeksi-month","mini-inspeksi",window.dashboardSchedules,(p,d)=>p.activity_type==="Inspeksi Hygiene"&&g(p)&&window.parseFlexibleDate(p.completion_date||p.target_date).startsWith(d),"#/timeline?dash_filter=inspeksi"),c("dash-gcdc-month","mini-gcdc",window.dashboardSchedules,(p,d)=>(p.activity_type==="General Cleaning"||p.activity_type==="Deep Cleaning")&&g(p)&&window.parseFlexibleDate(p.completion_date||p.target_date).startsWith(d),"#/timeline?dash_filter=gcdc"),c("dash-fogging-month","mini-fogging",window.dashboardFogging,(p,d)=>g(p)&&window.parseFlexibleDate(p.activity_date).startsWith(d),"#/reports/fogging?dash_filter=fogging")}function Ta(t){be("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),r=document.getElementById("donut-legend");if(!e||!r)return;Pe("donut");let a=(t||[]).filter(i=>he(i.count)>0);if(!a.length){Ge(e,"Belum ada data permasalahan");return}let l=a.map(i=>`${De(i.category,"Lainnya")}`),s=a.map(i=>he(i.count)),o=s.reduce((i,c)=>i+c,0);r.innerHTML=a.map((i,c)=>{let g=ft[c%ft.length],p=o>0?Math.round(i.count/o*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${g}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${i.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${p}%)</span></div>
          <div class="donut-legend-label">${l[c]}</div>
        </div>
      </div>
    `}).join("");let n={id:"centerText",beforeDraw:function(i){let c=i.width,g=i.height,p=i.ctx;p.restore();let d=(g/80).toFixed(2);p.font="bold "+d+"em Inter",p.textBaseline="middle",p.fillStyle="#1E293B";let u=o.toString(),y=Math.round((c-p.measureText(u).width)/2),m=g/2;p.fillText(u,y,m-4),p.font="600 "+(d*.35).toFixed(2)+"em Inter",p.fillStyle="#64748B";let b="Total",v=Math.round((c-p.measureText(b).width)/2);p.fillText(b,v,m+10),p.save()}};ve.donut=new Chart(e,{type:"doughnut",data:{labels:l,datasets:[{data:s,backgroundColor:ft,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ge,titleFont:{...ge,weight:"700"},callbacks:{label:i=>` ${i.label}: ${i.parsed} kasus`}}},cutout:"75%"},plugins:[n]})}function Ea(t){be("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;Pe("trend"),t=t||{};let r=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(t.labels||[]).map(o=>{if(!o||typeof o!="string")return"";try{let[n,i]=o.split("-");return(r[Number(i)-1]||i)+" "+String(n).slice(-2)}catch{return o}}),l=(t.open||[]).map(o=>he(o)),s=(t.closed||[]).map(o=>he(o));if(!a.length){Ge(e,"Belum ada data trend");return}ve.trend=new Chart(e,{type:"line",data:{labels:a,datasets:[{label:"Open",data:l,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:s,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:at({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:ye,maxRotation:0,autoSkip:!1}},y:{grid:{color:Ie},ticks:{font:{family:"Inter",size:9},color:ye},beginAtZero:!0}}})})}function Ft(t){be("skel-jadwal","chart-jadwal");let e=document.getElementById("chart-jadwal");if(!e)return;Pe("jadwal"),t=t||{};let r=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(t).some(i=>Array.isArray(i)&&i.some(c=>c>0))){Ge(e,"Belum ada data jadwal");return}let l=t["Inspeksi Hygiene"]||Array(12).fill(0),s=t["General Cleaning"]||Array(12).fill(0),o=t["Deep Cleaning"]||Array(12).fill(0),n=t.Fogging||Array(12).fill(0);ve.jadwal=new Chart(e,{type:"bar",data:{labels:r,datasets:[{label:"Inspeksi",data:l,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:s,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:o,backgroundColor:"#F59E0B"},{label:"Fogging",data:n,backgroundColor:"#EF4444"}]},options:at({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:ye,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Ie},ticks:{font:{family:"Inter",size:9},color:ye},min:0}}})})}function Mt(t){be("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;Pe("inspBar"),t=t||{};let r=t.labels||[],a=(t.fc||[]).map(s=>he(s)),l=(t.spv||[]).map(s=>he(s));if(!r.length){Ge(e,"Belum ada data inspeksi");return}ve.inspBar=new Chart(e,{type:"bar",data:{labels:r,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:l,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:at({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ge,color:ye,maxRotation:45,minRotation:30}},y:{grid:{color:Ie},ticks:{font:ge,color:ye},min:0,max:100}}})})}function Da(t){be("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;Pe("contractMiniBar"),t=t||{};let r={"06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(o=>{let n=o.split("-")[1];return r[n]||o}),l=(t.data||[]).map(o=>he(o));if(!a.length){Ge(e,"Belum ada data");return}let s=e.getContext("2d");ve.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:l,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:at({onClick:(o,n)=>{if(n&&n.length>0){let i=n[0].index,c=(t.labels||[])[i];c&&(window.location.hash="#/contracts?month_expiry="+c)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ge,color:ye,maxRotation:0,autoSkip:!1}},y:{grid:{color:Ie,borderDash:[4,4],drawBorder:!1},ticks:{font:ge,color:ye,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function Ia(t){let e=document.getElementById("table-issues");if(!e)return;let r=(t||[]).slice(0,8);if(!r.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${r.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${wa(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${De(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${De(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function Pa(t){let e=document.getElementById("widget-agenda");if(!e)return;let r=new Date,a=`${r.getFullYear()}-${String(r.getMonth()+1).padStart(2,"0")}-${String(r.getDate()).padStart(2,"0")}`,s=(t||[]).filter(o=>(o.event_date||"").startsWith(a)).slice(0,10);if(!s.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${s.map(o=>{let n="#3B82F6",i="#EFF6FF",c="Agenda",g=(o.title||"").toLowerCase();return g.includes("inspeksi")?(n="#10B981",i="#ECFDF5",c="Inspeksi"):g.includes("cleaning")||g.includes("gcdc")?(n="#3B82F6",i="#EFF6FF",c="Cleaning"):g.includes("reliefer")?(n="#F59E0B",i="#FFFBEB",c="Reliefer"):g.includes("fogging")&&(n="#8B5CF6",i="#F5F3FF",c="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(o.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${n};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${De(o.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${De(o.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${i};color:${n}">${c}</div>
        </div>
      `}).join("")}
    </div>
  `}function La(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(r=>`
    <a href="${r.href}" class="action-btn">
      <div class="action-icon" style="background:${r.bg}">${r.icon}</div>
      ${r.label}
    </a>
  `).join("")}function be(t,e){let r=document.getElementById(t),a=document.getElementById(e);if(r&&(r.style.display="none",r.style.position=""),a){a.style.display="block";let l=a.parentElement;if(l){let s=l.querySelector(".chart-empty");s&&s.remove()}}}function Ge(t,e="Belum ada data"){if(!t)return;t.style.display="none";let r=t.parentElement;if(!r)return;if(!r.querySelector(".chart-empty")){let l=document.createElement("div");l.className="chart-empty",l.textContent=e,r.appendChild(l)}}j();async function Ot(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),r=document.getElementById("login-error"),a=document.getElementById("login-btn"),l=document.getElementById("toggle-password"),s=document.getElementById("login-password");l?.addEventListener("click",()=>{let o=s.type==="text";s.type=o?"password":"text",l.style.color=o?"":"var(--primary)"}),e?.addEventListener("submit",async o=>{o.preventDefault(),r.style.display="none";let n=e.username.value.trim(),i=e.password.value;if(!n||!i){r.textContent="Username dan password wajib diisi.",r.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let c=await S("/api/auth/login",{method:"POST",body:JSON.stringify({username:n,password:i})});c.ok&&c.data.success?(pt(c.data.data.token),Ke(c.data.data.user),ee("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(r.textContent=c.data.error||"Username atau password salah.",r.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{r.textContent="Gagal terhubung ke server. Periksa koneksi internet.",r.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}j();H();async function Ba(){return await q()}function Aa(t,e){let r=String(t.status||"").toLowerCase();return e==="active"?r==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&r==="aktif":!1}async function Rt(t,e){let r=await Ba(),a=e?e.get("dash_filter"):null;O({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",enableMobileFilterSheet:!0,itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:l=>a?l.filter(s=>Aa(s,a)):l,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:l=>Ee(l)},{key:"phone",label:"No. HP",render:l=>l?`<a href="tel:${l}">${l}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>G(l)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:l=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:l?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:l?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:r,value:l?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:l?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:l?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:l?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let l=await S(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(l.ok){let s=l.data.data.map(o=>({"Nama Lengkap":o.full_name,Cabang:o.branch_name||"",Divisi:o.division||"","No. HP":o.phone||"","Tgl Masuk":o.join_date||"",Status:o.status||""}));M(s,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{M([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async l=>{let s=i=>{if(!i)return null;let c=String(i||"").toLowerCase(),g=r.find(p=>String(p.label||"").toLowerCase()===c);return g?g.value:null},o=l.map(i=>({full_name:String(i["Nama Lengkap"]||"").trim(),branch_id:s(String(i.Cabang||"").trim()),division:String(i.Divisi||"").trim()||"FACILITY CARE",phone:String(i["No. HP"]||"").trim(),join_date:String(i["Tgl Masuk"]||"").trim(),status:String(i.Status||"").trim(),notes:String(i.Catatan||"").trim()})).filter(i=>i.full_name),n=await S("/api/import/employees",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}j();H();var St=[],Kt=[];async function Na(){St=await q(),Kt=await Te()}var kt=async t=>{let e=[],r=1;for(;;){let l=await(await Promise.resolve().then(()=>(j(),_e))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${r}`);if(!l.ok)break;let s=l.data?.data||l.data||[],o=Array.isArray(s)?s:[];if(e=e.concat(o),o.length<100||l.data?.pagination&&r>=l.data.pagination.pages)break;r++}return e};function Fa(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let l=new Date(a);l.setDate(a.getDate()+30);let s=new Date(t.end_date);return s.setHours(0,0,0,0),s>=a&&s<=l}return!1}async function jt(t,e){await Na();let r=e?e.get("dash_filter"):null;O({container:t,title:"Data Kontrak",icon:"\u{1F4C4}",apiPath:"/api/contracts",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Kontrak",paginationMode:"client",onDataLoaded:a=>r?a.filter(l=>Fa(l,r)):a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>Ee(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,l)=>l.end_date&&String(l.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':gt(a)},{key:"status",label:"Status",render:a=>G(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:St},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"month_expiry",label:"Bulan Habis",options:[{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"},{value:"2027-01",label:"Jan 2027"},{value:"2027-02",label:"Feb 2027"},{value:"2027-03",label:"Mar 2027"},{value:"2027-04",label:"Apr 2027"},{value:"2027-05",label:"Mei 2027"},{value:"2027-06",label:"Jun 2027"},{value:"2027-07",label:"Jul 2027"}]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Detail Kontrak",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[s,o]=await Promise.all([kt("/api/employees?status=Aktif"),kt("/api/contracts")]);if(s.length>0){let n=o.filter(u=>u.status==="Aktif"&&(u.days_remaining==null||u.days_remaining>=0||String(u.end_date).startsWith("2099"))),i=new Set(n.map(u=>u.employee_id)),c=s.filter(u=>!i.has(u.id)),g=[],p=[];c.forEach(u=>{let y=o.filter(m=>m.employee_id===u.id);y.length===0?g.push(u):p.push({emp:u,lastContract:y[0]})});let d=`<p style="margin-bottom:12px">Data yang terbaca: <b>${s.length}</b> Karyawan Aktif, dan <b>${n.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:6px">Terdapat <b>${g.length}</b> karyawan aktif tanpa kontrak.</p>
              <p style="margin-bottom:12px">Terdapat <b>${p.length}</b> karyawan aktif yang memiliki masa kontrak Expired.</p>
              <ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;g.forEach(u=>{d+=`<li style="margin-bottom:8px"><b>${u.full_name}</b> <br><span style="font-size:0.85em;color:#F59E0B">Cabang: ${u.branch_name||"-"} | Belum pernah di-input kontrak</span></li>`}),p.forEach(u=>{let y=u.emp,m=u.lastContract,b=m.status==="Aktif"&&m.days_remaining<0?"Aktif (Masa Habis)":m.status;d+=`<li style="margin-bottom:8px"><b>${y.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${y.branch_name||"-"} | Status Terakhir: <b style="color:#EF4444">${b}</b>, Tgl Berakhir: ${window.formatDate(m.end_date)}</span></li>`}),d+="</ul>",Promise.resolve().then(()=>(Se(),mt)).then(u=>u.createModal({title:"Detail Karyawan Tanpa Kontrak Aktif",content:d,cancelText:"Tutup"}))}}catch(s){console.error(s)}a.innerHTML="\u{1F50D} Detail Kontrak",a.disabled=!1};let l=document.querySelector(".page-actions");l&&l.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"select",required:!0,options:Kt,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"select",options:St,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await S(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let l=a.data.data.map(s=>({"Nama Lengkap":s.employee_name,Cabang:s.branch_name||"","Div / Bagian":s.division||"","Tanggal Mulai":s.start_date||"","Tanggal Selesai":s.end_date&&String(s.end_date).startsWith("2099")?"":s.end_date||"","Sisa Kontrak":s.end_date&&String(s.end_date).startsWith("2099")?"Tetap":s.days_remaining!==null&&s.days_remaining!==void 0?`${s.days_remaining} Hari`:"",Status:s.status||""}));M(l,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{M([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[l,s]=await Promise.all([S("/api/branches?limit=10000"),kt("/api/employees")]),o=l.data?.data||[],n=s||[];console.log(`Total employee yang berhasil dimuat dari database : ${n.length}`),n.length>0&&(console.log("Contoh 5 employee pertama:"),n.slice(0,5).forEach((m,b)=>{console.log(`${b+1}. ID: ${m.id}, Name: ${m.full_name}, Status: ${m.status}`)}));let i=m=>{if(!m)return null;let b=String(m||"").replace(/\s+/g," ").toLowerCase().trim(),v=o.find(k=>String(k.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===b||String(k.code||"").replace(/\s+/g," ").toLowerCase().trim()===b||String(k.name||"").replace(/\s+/g," ").toLowerCase().trim()===b);return v?v.id:null},c=(m,b)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${b}`),console.log(`Nama dari Excel : "${m}"`),!m)return console.log("Alasan gagal mapping : Nama kosong"),null;let v=String(m||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${v}"`),console.log(`Jumlah employee di database : ${n.length}`);let k=n.find(w=>String(w.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===v);return k?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${k.id}`),k.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},g=m=>{if(!m)return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let b=String(m).trim();if(/^\d{4,5}(\.\d+)?$/.test(b)){let k=Math.floor(Number(b));if(k>2e4&&k<99999){let w=new Date(Date.UTC(1899,11,30)+k*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(b))return b.slice(0,10);let v=b.split(/[\/\-\.]/);if(v.length===3){let[k,w,T]=v.map($=>$.trim());if(k.length===4&&w.length<=2&&T.length<=2)return`${k}-${w.padStart(2,"0")}-${T.padStart(2,"0")}`;if(T.length===4&&w.length<=2&&k.length<=2)return`${T}-${w.padStart(2,"0")}-${k.padStart(2,"0")}`}return b},p=a.map((m,b)=>{let v=b+2,k=String(m["Nama Lengkap"]||"").trim(),w=m["Tanggal Mulai"],T=g(w);if(!T){let x=a.__worksheet,_=a.__headers||[],N=_.indexOf("Tanggal Mulai"),F="N/A",le="N/A",de="N/A";if(N!==-1&&x&&window.XLSX){let ke=window.XLSX.utils.encode_cell({c:N,r:v-1});de=ke;let fe=x[ke];fe?(F=fe.t||"undefined",le=fe.w||"undefined"):F="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let se="Unknown";w==null||w===""?se="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":w instanceof Date&&isNaN(w.getTime())?se="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":se="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${v}`),console.log(`Employee Name : ${k}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${N})`),console.log(`Raw Cell Value : "${w}"`),console.log(`JavaScript Type : ${typeof w}`),console.log(`SheetJS Cell Type : ${F}`),console.log(`SheetJS Formatted Value : "${le}"`),console.log(`Value After Trim : "${String(w||"").trim()}"`),console.log(`Value After Date Parser : "${T}"`),console.log(`Is Empty : ${!w}`),console.log(`Is Invalid Date : ${w instanceof Date?isNaN(w.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${se}`),console.log(`Workbook Sheet : ${x?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${de}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(m,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(_)),console.log(`==========================
`)}let $=c(k,v),C=null;return $||(C="Karyawan tidak ditemukan di Database"),{isValid:!!$,invalidReason:C,rowNum:v,data:{employee_id:$,branch_id:i(String(m.Cabang||"").trim()),division:String(m["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:T,end_date:g(m["Tanggal Selesai"])||"2099-12-31",status:String(m.Status||"").trim(),_rawName:k}}}),d=[],u=[];if(p.forEach(m=>{m.isValid?d.push(m.data):u.push({rowNum:m.rowNum,name:m.data._rawName,reason:m.invalidReason})}),console.log(`Split Validation - Valid: ${d.length}, Invalid: ${u.length}`),d.length===0)return{inserted:0,skipped:a.length,failed:a.length};let y=await S("/api/import/contracts",{method:"POST",body:JSON.stringify({rows:d,onDuplicate:"update"})});if(!y.ok)throw new Error(y.data?.error||"Import gagal");return y.data}}})}j();H();var wt=[],Qe=[];function Ma(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let r of e)if(t.some(a=>a.period===r))return r;return"Q3"}async function Ht(t,e){wt=await q();let r=await te();Qe=["Berlin Ariansyah","Ade Surahman"];let a=b=>b&&!Qe.find(v=>String(typeof v=="object"?v.value:v).toLowerCase()===String(b).toLowerCase())?[...Qe,b]:Qe,l=await S(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),s=b=>{if(!b||b==="-"||String(b).trim()==="")return"";let v=String(b).split("-");return v.length===3&&v[0].length===4?`${v[2]}-${v[1]}-${v[0]}`:b},o=l.data?.data||[],n=Ma(o),i=e?e.get("dash_filter"):null,c=new Date,g=`${c.getFullYear()}-${String(c.getMonth()+1).padStart(2,"0")}`,p={},d=e&&e.get("month")?e.get("month"):null;i==="inspeksi"?p={status:"Done",activity_type:"Inspeksi Hygiene",month:d}:i==="gcdc"?p={status:"Done",activity_type:"GCDC",month:d}:i&&i.startsWith("period_")&&(p={period:i.replace("period_","").toUpperCase()});let u=new Date().getFullYear(),m=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((b,v)=>{let k=String(v+1).padStart(2,"0");return{value:`${u}-${k}`,label:`${b} ${u}`}});O({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:p,onDataLoaded:b=>b.sort((v,k)=>{let w=v.opening_date?new Date(v.opening_date).getTime():0;return(k.opening_date?new Date(k.opening_date).getTime():0)-w}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:b=>bt(b)},{key:"period",label:"Periode",render:b=>me(b)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:b=>s(b)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:b=>s(b)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:b=>s(b)},{key:"status",label:"Status",render:b=>G(b)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:wt},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:m},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Qe}],formFields:b=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:wt,value:b?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:b?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:b?.period},{name:"pic",label:"PIC",type:"combobox",options:a(b?.pic),value:b?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:b?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:b?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:b?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:b?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:b?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let b=await S(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(b.ok){let v=b.data.data.map(k=>({Cabang:k.branch_name||"",Kegiatan:k.activity_type||"",Periode:k.period||"",PIC:k.pic||"","Tgl Opening":k.opening_date||"","Tgl Target":k.target_date||"","Tgl Selesai":k.completion_date||"",Status:k.status||""}));M(v,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{M([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async b=>{let k=(await S("/api/branches?all=1")).data?.data||[],w=x=>{if(!x)return null;let _=String(x||"").toLowerCase(),N=k.find(F=>String(F.full_name||"").toLowerCase()===_||String(F.code||"").toLowerCase()===_||String(F.name||"").toLowerCase()===_);return N?N.id:null},T=x=>{if(x==null||x==="")return"";if(x instanceof Date&&!isNaN(x.getTime()))return x.toISOString().slice(0,10);let _=String(x).trim();if(_===""||_==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(_))return _.slice(0,10);if(/^\d{4,5}$/.test(_)){let F=Number(_);if(F>2e4&&F<99999){let le=new Date(Date.UTC(1899,11,30)+F*864e5);return isNaN(le.getTime())?"":le.toISOString().slice(0,10)}}let N=_.split(/[\/\-\.]/);if(N.length===3){let[F,le,de]=N.map(se=>se.trim());if(F.length===4&&le.length<=2&&de.length<=2)return`${F}-${le.padStart(2,"0")}-${de.padStart(2,"0")}`;if(de.length===4&&le.length<=2&&F.length<=2)return`${de}-${le.padStart(2,"0")}-${F.padStart(2,"0")}`}return _},$=b.map(x=>({branch_id:w(String(x.Cabang||"").trim()),activity_type:String(x.Kegiatan||"").trim(),period:String(x.Periode||"").trim(),pic:String(x.PIC||x.Pic||"").trim(),opening_date:T(x["Tgl Opening"]||x["Tanggal Opening"]||x["Tgl Openir"]),target_date:T(x["Tgl Target"]||x["Tanggal Target"]),completion_date:T(x["Tgl Selesai"]||x["Tanggal Selesai"]),status:String(x.Status||"").trim(),notes:String(x.Catatan||x.Keterangan||"").trim()})).filter(x=>x.activity_type&&x.period),C=await S("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:$,onDuplicate:"update"})});if(!C.ok)throw new Error(C.data?.error||"Import gagal");return C.data}}})}j();H();var xt=[],nt=[];function Oa(t,e){let r=String(t.status||"").toLowerCase();return e==="open"?r==="open":!1}async function qt(t,e){let r=e?e.get("dash_filter"):null;xt=await q(),nt=await te();let a=i=>i&&!nt.find(c=>c.value===i)?[...nt,{value:i,label:i}]:nt,l=new Date().getFullYear(),s=["2025","2026","2027","2028","2029","2030"],n=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((i,c)=>{let g=String(c+1).padStart(2,"0");return{value:`${l}-${g}`,label:`${i} ${l}`}});O({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:i=>r?i.filter(c=>Oa(c,r)):i,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:i=>`<span class="badge badge-secondary">${i}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:i=>`<span title="${i}">${i?.length>50?i.slice(0,50)+"\u2026":i}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:i=>`<span title="${i||""}">${i?.length>40?i.slice(0,40)+"\u2026":i||"-"}</span>`},{key:"status",label:"Status",render:i=>G(i)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:i=>window.formatDate(i)},{key:"day_count",label:"Hari",render:i=>i??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:xt},{type:"select",name:"month",label:"Bulan",options:n},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:s}],formFields:i=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:i?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:xt,value:i?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:i?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:i?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:i?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(i?.employee_name),value:i?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(i?.fc_specialist),value:i?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:i?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:i?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:i?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let i=await S(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let c=i.data.data.map(g=>({Tanggal:g.report_date||"",Cabang:g.branch_name||"",Kategori:g.category||"",Sumber:g.source||"",Keluhan:g.complaint||"","Nama FC":g.employee_name||"","FC Spesialis":g.fc_specialist||"",Solusi:g.solution||"","Tgl Selesai":g.completion_date||"",Status:g.status||""}));M(c,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{M([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async i=>{let g=(await S("/api/branches?all=1")).data?.data||[],p=y=>{if(!y)return null;let m=String(y||"").toLowerCase(),b=g.find(v=>String(v.full_name||"").toLowerCase()===m||String(v.code||"").toLowerCase()===m||String(v.name||"").toLowerCase()===m);return b?b.id:null},d=i.map(y=>({branch_id:p(String(y.Cabang||"").trim()),report_date:String(y.Tanggal||"").trim(),category:String(y.Kategori||"").trim(),source:String(y.Sumber||"").trim(),complaint:String(y.Keluhan||"").trim(),employee_name:String(y["Nama FC"]||"").trim(),fc_specialist:String(y["FC Spesialis"]||"").trim(),solution:String(y.Solusi||"").trim(),completion_date:String(y["Tgl Selesai"]||"").trim(),status:String(y.Status||"").trim()})).filter(y=>y.report_date&&y.complaint&&y.category),u=await S("/api/import/issues",{method:"POST",body:JSON.stringify({rows:d,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}j();var Le=[];function Ra(t,e){let r=String(t.status||"").toLowerCase();return e==="pending"?r==="pending":!1}async function Jt(t,e){let r=e?e.get("dash_filter"):null;Le=await q();let a=await te(),l=["Ade","Berlin"],s=n=>n&&!a.find(i=>i.value===n)?[...a,{value:n,label:n}]:a,o=n=>n&&!l.find(i=>(typeof i=="object"?i.value:i)===n)?[...l,n]:l;O({container:t,title:"One on One",icon:"\u{1F4AC}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:n=>r?n.filter(i=>Ra(i,r)):n,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:n=>`<span title="${n||""}">${n?.length>50?n.slice(0,50)+"\u2026":n||"-"}</span>`},{key:"solution",label:"Solusi",render:n=>`<span title="${n||""}">${n?.length>40?n.slice(0,40)+"\u2026":n||"-"}</span>`},{key:"status",label:"Status",render:n=>G(n)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:Le},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async n=>{let i=new URLSearchParams(n||{}).toString(),c=await S(`/api/one-on-one?limit=10000&${i}`);if(c.ok){let g=c.data.data.map(d=>({Tanggal:d.meeting_date||"",Cabang:d.branch_name||"","Nama Karyawan":d.employee_name||"",PIC:d.pic||"",Masalah:d.problem||"",Solusi:d.solution||"",Status:d.status||"","Tgl Selesai":d.completion_date||"",Dokumen:d.document_link||""})),{downloadExcel:p}=await Promise.resolve().then(()=>(H(),pe));p(g,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(H(),pe));i(n,"Template_Import_OneOnOne")},onImport:async n=>{let i=d=>{if(!d)return null;let u=String(d||"").toLowerCase(),y=Le.find(m=>String(m.label||"").toLowerCase()===u);return y?y.value:null},c=d=>{if(!d)return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let u=String(d).trim();if(/^\d{4,5}$/.test(u)){let m=Number(u);if(m>2e4&&m<99999){let b=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);let y=u.split(/[\/\-\.]/);if(y.length===3){let[m,b,v]=y.map(k=>k.trim());if(m.length===4&&b.length<=2&&v.length<=2)return`${m}-${b.padStart(2,"0")}-${v.padStart(2,"0")}`;if(v.length===4&&b.length<=2&&m.length<=2)return`${v}-${b.padStart(2,"0")}-${m.padStart(2,"0")}`}return u},g=n.map(d=>({meeting_date:c(d.Tanggal),employee_name:String(d["Nama Karyawan"]||"").trim(),branch_id:i(String(d.Cabang||"").trim()),pic:String(d.PIC||"").trim(),problem:String(d.Masalah||"").trim(),solution:String(d.Solusi||"").trim(),status:String(d.Status||"").trim(),completion_date:c(d["Tgl Selesai"]),document_link:String(d.Dokumen||"").trim()})).filter(d=>d.meeting_date&&d.employee_name&&d.branch_id),p=await S("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:g,onDuplicate:"update"})});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}},formFields:n=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:n?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:n?.branch_id&&!Le.find(i=>i.value==n.branch_id)?[...Le,{value:n.branch_id,label:n.branch_name||n.branch_id}]:Le,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:s(n?.employee_name),value:n?.employee_name},{name:"pic",label:"PIC",type:"select",options:o(n?.pic),createApi:{path:"/api/pic",field:"name"},value:n?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:n?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:n?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:n?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link}]})}j();async function Ut(t){let e=await q(),r=await te(),a=["Ade","Berlin"],l=n=>n&&!r.find(i=>i.value===n)?[...r,{value:n,label:n}]:r,s=n=>n&&!a.find(i=>(typeof i=="object"?i.value:i)===n)?[...a,n]:a,o=Array.from({length:5},(n,i)=>String(new Date().getFullYear()-i));O({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:n=>{try{let i=JSON.parse(n);return Array.isArray(i)?i.join(", "):n||"-"}catch{return n||"-"}}},{key:"score",label:"Nilai",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:o}],exportOptions:{moduleName:"training",onExport:async n=>{let i=new URLSearchParams(n||{}).toString(),c=await S(`/api/training?limit=10000&${i}`);if(c.ok){let g=c.data.data.map(d=>{let u=d.participants||"";try{let y=JSON.parse(u);u=Array.isArray(y)?y.join(", "):u}catch{}return{Tanggal:d.training_date||"",Batch:d.batch||"",Materi:d.subject||"",Cabang:d.branch_name||"",Trainer:d.trainer||"",Peserta:u,Nilai:d.score!==null&&d.score!==void 0?d.score:"",Dokumen:d.document_link||""}}),{downloadExcel:p}=await Promise.resolve().then(()=>(H(),pe));p(g,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(H(),pe));i(n,"Template_Import_Training")},onImport:async n=>{let i=d=>{if(!d)return null;let u=String(d||"").toLowerCase(),y=e.find(m=>String(m.label||"").toLowerCase()===u);return y?y.value:null},c=d=>{if(!d)return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let u=String(d).trim();if(/^\d{4,5}$/.test(u)){let m=Number(u);if(m>2e4&&m<99999){let b=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);let y=u.split(/[\/\-\.]/);if(y.length===3){let[m,b,v]=y.map(k=>k.trim());if(m.length===4&&b.length<=2&&v.length<=2)return`${m}-${b.padStart(2,"0")}-${v.padStart(2,"0")}`;if(v.length===4&&b.length<=2&&m.length<=2)return`${v}-${b.padStart(2,"0")}-${m.padStart(2,"0")}`}return u},g=n.map(d=>({training_date:c(d.Tanggal),batch:String(d.Batch||"").trim(),subject:String(d.Materi||"").trim(),branch_id:i(String(d.Cabang||"").trim()),trainer:String(d.Trainer||"").trim(),participants:String(d.Peserta||"").trim(),score:d.Nilai?Number(d.Nilai):null,document_link:String(d.Dokumen||"").trim()})).filter(d=>d.training_date&&d.subject&&d.branch_id),p=await S("/api/import/training",{method:"POST",body:JSON.stringify({rows:g,onDuplicate:"update"})});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}},formFields:n=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:n?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:n?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:n?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:e,value:n?.branch_id},{name:"trainer",label:"Trainer",type:"select",options:s(n?.trainer),value:n?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let i=JSON.parse(n?.participants);return Array.isArray(i)?i.join(", "):n?.participants||""}catch{return n?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:n?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:n?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],onBeforeSubmit:async n=>(n.participants&&(n.participants=JSON.stringify(n.participants.split(",").map(i=>i.trim()).filter(Boolean))),n)})}j();Se();H();function Gt({container:t,title:e,icon:r,apiPath:a,columns:l,formFields:s,filterFields:o,defaultFilters:n={},enableMobileFilterSheet:i=!1,itemLabel:c="Data",canCreate:g=!0,canEdit:p=!0,canDelete:d=!0,onBeforeSubmit:u,onAfterLoad:y,onDataLoaded:m,extraActions:b=[],initialSearch:v="",exportOptions:k=null,bulkDelete:w=!1,paginationMode:T="server"}){let $=ce();$&&typeof $=="object"&&$.role==="viewer"&&(g=!1,p=!1,d=!1,w=!1,k=null);let C=1,x={...n};v&&(x.search=v);let _=new Set;t.innerHTML=`
    <div class="crud-layout-wrapper ${i?"mobile-active":""}">
      <div class="page-header">
      <h1 class="page-title">${r} ${e}</h1>
      <div class="page-actions">
        ${g?`<button class="btn btn-primary" id="btn-create">+ Tambah ${c}</button>`:""}
        ${k?'<button class="btn btn-outline" id="btn-mobile-aksi" style="display:none; align-items:center; justify-content:center; gap:0.25rem;">\u22EE Aksi</button>':""}
      </div>
    </div>

    ${w?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:none; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${k?`
    <div class="excel-actions-wrapper" id="excel-actions-wrapper">
      <div class="bottom-sheet-header aksi-header" style="display:none;">
        <h3 style="margin:0; font-size:1rem;">Aksi</h3>
        <button class="btn-close-sheet" id="btn-close-aksi-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
      </div>
      ${Ue(k.moduleName)}
    </div>`:""}

    ${o&&o.length>0?`
    <div class="filter-bar card ${i?"has-mobile-sheet":""}" style="padding: 1rem; display: flex; gap: 0.5rem; flex-wrap: wrap;">
        ${o.filter(h=>h.type==="search"||h.type==="search-combo").map(h=>{if(h.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${h.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}"></div>`;if(h.type==="search-combo"){let f="dl-filter-search",E=(h.options||[]).map(D=>`<option value="${typeof D=="object"?D.label:D}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${f}" class="form-control" autocomplete="off" placeholder="${h.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}"><datalist id="${f}">${E}</datalist></div>`}return""}).join("")}
        
        <div class="filter-options" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(h=>h.type!=="search"&&h.type!=="search-combo").map(h=>{if(h.type==="select")return`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${h.name}" id="filter-${h.name}"><option value="">Pilih ${h.label}</option>${(h.options||[]).map(f=>`<option value="${typeof f=="object"?f.value:f}" ${x[h.name]===(typeof f=="object"?f.value:f)?"selected":""}>${typeof f=="object"?f.label:f}</option>`).join("")}</select>`;if(h.type==="combobox"){let f=`dl-filter-${h.name}`,E=(h.options||[]).map(P=>{let I=typeof P=="object"?P.label||P.value||"":P||"";return(I==="undefined"||I==="[object Object]"||I==="null")&&(I=""),I?`<option value="${I}"></option>`:""}).join(""),D=x[h.name]||"";if(x[h.name]){let P=(h.options||[]).find(I=>String(typeof I=="object"?I.value:I)==String(x[h.name]));if(P){let I=typeof P=="object"?P.label||P.value||"":P||"";I&&I!=="undefined"&&I!=="[object Object]"&&I!=="null"&&(D=I)}}return`<div class="filter-combobox" style="flex:1; min-width:120px;">
                <input type="text" name="${h.name}" id="filter-${h.name}" list="${f}" class="form-control filter-combobox-input" value="${D}" placeholder="Pilih ${h.label}..." autocomplete="off">
                <datalist id="${f}">${E}</datalist>
              </div>`}return""}).join("")}
          <div style="display:flex; gap:0.5rem; margin-top:0.5rem;">
            <button class="btn btn-outline" id="btn-reset-filter" style="flex:1;">Reset</button>
            <button class="btn btn-primary" id="btn-apply-filter" style="flex:1;">\u2713 Terapkan (OK)</button>
          </div>
        </div>
        ${i?'<button id="btn-mobile-filter" class="btn btn-outline" style="display:none; align-items:center; gap:0.25rem;">\u2699\uFE0F Filter</button>':""}
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
    </div>
  `;function N(){let h=document.getElementById("bulk-toolbar");if(!h)return;let f=document.getElementById("bulk-count"),E=document.getElementById("btn-bulk-delete"),D=document.getElementById("btn-bulk-cancel");f.textContent=`${_.size} item dipilih`,_.size>0?(h.style.display="flex",h.classList.add("has-items"),E.disabled=!1,D.disabled=!1,window.innerWidth<=768?(h.style.setProperty("top","auto","important"),h.style.setProperty("bottom","0","important")):(h.style.removeProperty("top"),h.style.removeProperty("bottom"))):(h.style.display="none",h.classList.remove("has-items"),E.disabled=!0,D.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{_.clear(),document.querySelectorAll(".row-checkbox").forEach(f=>f.checked=!1);let h=document.getElementById("select-all-checkbox");h&&(h.checked=!1),N()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(_.size===0)return;let h=[..._],f=document.createElement("div");f.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",f.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${h.length} ${c}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${h.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(f),f.querySelector("#bulk-cancel-btn").addEventListener("click",()=>f.remove()),f.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let E=f.querySelector("#bulk-confirm-btn");E.disabled=!0,E.textContent="Menghapus...";let D=await S(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:h})});f.remove(),D.ok?(ee(`${h.length} ${c} berhasil dihapus.`),_.clear(),N(),Q()):z(D.data?.error||"Gagal menghapus data.")})});let F=document.getElementById("filter-search"),le;F?.addEventListener("input",h=>{clearTimeout(le),le=setTimeout(()=>{x.search=h.target.value,C=1,_.clear(),Q()},400)}),o?.forEach(h=>{h.type==="select"&&document.getElementById(`filter-${h.name}`)?.addEventListener("change",f=>{x[h.name]=f.target.value,C=1,_.clear(),Q()}),h.type==="combobox"&&document.getElementById(`filter-${h.name}`)?.addEventListener("change",f=>{let E=f.target.value,D=(h.options||[]).find(P=>{let I=String(typeof P=="object"?P.value:P),B=String(typeof P=="object"?P.label:P);return I===E||B===E});E?x[h.name]=D?typeof D=="object"?D.value:D:E:x[h.name]="",C=1,_.clear(),Q()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{x={...n},F&&(F.value=""),o?.forEach(h=>{let f=document.getElementById(`filter-${h.name}`);f&&(f.value="")}),C=1,_.clear(),Q()});let de=document.getElementById("btn-mobile-filter"),se=document.getElementById("filter-options-wrapper"),ke=document.getElementById("btn-close-filter-sheet");if(de&&se){de.addEventListener("click",f=>{f.preventDefault(),se.classList.add("sheet-open")}),ke&&ke.addEventListener("click",f=>{f.preventDefault(),se.classList.remove("sheet-open")});let h=document.getElementById("btn-apply-filter");h&&h.addEventListener("click",f=>{f.preventDefault(),se.classList.remove("sheet-open")})}let fe=document.getElementById("btn-mobile-aksi"),we=document.getElementById("excel-actions-wrapper"),Ne=document.getElementById("btn-close-aksi-sheet");if(fe&&we&&(fe.addEventListener("click",h=>{h.preventDefault(),we.classList.add("sheet-open")}),Ne&&Ne.addEventListener("click",h=>{h.preventDefault(),we.classList.remove("sheet-open")})),document.getElementById("btn-create")?.addEventListener("click",()=>Fe(null)),k){document.getElementById(`btn-export-${k.moduleName}`)?.addEventListener("click",async f=>{let E=f.target,D=E.innerHTML;E.innerHTML="\u23F3 Loading...",E.disabled=!0;try{await k.onExport()}catch{z("Gagal export data")}finally{E.innerHTML=D,E.disabled=!1}}),document.getElementById(`btn-template-${k.moduleName}`)?.addEventListener("click",()=>{k.onTemplate()});let h=document.getElementById(`input-import-${k.moduleName}`);h?.addEventListener("change",async f=>{let E=f.target.files[0];if(!E)return;h.disabled=!0;let D=document.createElement("div");D.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",D.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(D);let P=D.querySelector("#import-progress-text"),I=D.querySelector("#import-progress-bar"),B=D.querySelector("#import-summary"),A=D.querySelector("#import-close-btn");A.addEventListener("click",()=>{D.remove(),Q()});try{let X=await Je(E);if(X.length===0)throw new Error("File kosong atau format salah");let Z=500,oe=0,ne=0,R=0,K=X.length;P.textContent=`Ditemukan ${K} baris data. Memulai import...`;for(let L=0;L<K;L+=Z){let Y=X.slice(L,L+Z);P.textContent=`Mengimport baris ${L+1} - ${Math.min(L+Z,K)} dari ${K}...`,I.style.width=`${Math.round(L/K*100)}%`;try{let J=await k.onImport(Y);J?(oe+=J.inserted||J.metrics?.inserted||Y.length,ne+=J.skipped||J.metrics?.updated||0):oe+=Y.length}catch(J){console.error("Chunk import failed:",J),R+=Y.length}}I.style.width="100%",P.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',B.style.display="block",B.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${K}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${oe}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ne}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${R}</strong></div>
        `,R>0&&(B.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),A.style.display="block",h.value=""}catch(X){P.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${X.message}`,I.style.background="var(--danger)",I.style.width="100%",A.style.display="block",h.value=""}finally{h.disabled=!1}})}async function Q(){N();let h=document.getElementById("table-container");if(!h)return;h.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let f=T==="client",E=f?1:C,D=f?xe:20,P=new URLSearchParams({page:E,limit:D,...Object.fromEntries(Object.entries(x).filter(([,R])=>R))}),I=await S(`${a}?${P}`);if(!I.ok){h.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${I.data?.error||"Error"}</p></div>`;return}let B=I.data?.data||I.data||[],A=I.data?.pagination,X=B.length,Z=B;if(f){B=m(B),Z=B;let R=B.length,K=20,L=Math.ceil(R/K);C>L&&L>0&&(C=L);let Y=(C-1)*K,J=C*K;B=B.slice(Y,J),A={page:C,limit:K,total:R,pages:L}}!1,y&&y(B);let oe=Ve({columns:l,data:B,fullData:Z,onEdit:p?R=>Fe(R):null,actions:b.map(R=>({...R,handler:K=>R.handler(K,Q)})),emptyText:`Tidak ada ${String(c||"").toLowerCase()}`,bulkSelect:w?{selectedIds:_,onToggle:N}:null});h.innerHTML="",h.appendChild(oe);let ne=document.getElementById("pagination-container");if(ne&&(ne.innerHTML="",A&&A.pages>1)){let R=Xe({page:A.page,pages:A.pages,total:A.total,limit:A.limit,onPage:K=>{C=K,Q()}});R&&ne.appendChild(R)}}function st(h){let f=typeof s=="function"?s(h):s;return qe(f)}function Fe(h){let f=!!h,E=document.createElement("form");if(E.noValidate=!0,E.innerHTML=st(h),f){let P=typeof s=="function"?s(h):s;et(E,h)}let{close:D}=ue({title:f?`Edit ${c}`:`Tambah ${c}`,content:E,size:"lg",confirmText:f?"Simpan Perubahan":`Tambah ${c}`,onConfirm:async(P,I)=>{if(!E.reportValidity())return;let B=P.querySelector(".modal-confirm");B.disabled=!0,B.textContent="Menyimpan...";let A=Ze(E),X=typeof s=="function"?s(h):s,Z=async K=>{for(let L of K)if(L.type==="row")await Z(L.fields);else if(L.type==="combobox"&&A[L.name]){let Y=A[L.name],J=(L.options||[]).find(V=>{let ie=String(typeof V=="object"?V.value:V),ot=String(typeof V=="object"?V.label:V);return ie===Y||ot===Y});if(J)A[L.name]=typeof J=="object"?J.value:J;else if(L.createApi){let V={};V[L.createApi.field]=Y,L.createApi.extra&&Object.assign(V,L.createApi.extra);let ie=await S(L.createApi.path,{method:"POST",body:JSON.stringify(V)});if(ie.ok&&ie.data?.id)A[L.name]=ie.data.id;else if(ie.ok&&!ie.data?.id)A[L.name]=Y;else throw new Error(`Gagal membuat master data: ${ie.data?.error||"Unknown error"}`)}}};try{await Z(X)}catch(K){z(K.message),B.disabled=!1,B.textContent=f?"Simpan Perubahan":`Tambah ${c}`;return}u&&(A=await u(A,h));let oe=f?"PUT":"POST",ne=f?`${a}/${h.id}`:a,R=await S(ne,{method:oe,body:JSON.stringify(A)});R.ok?(ee(f?`${c} berhasil diperbarui.`:`${c} berhasil ditambahkan.`),I(),Q()):(z(R.data?.error||"Gagal menyimpan data."),B.disabled=!1,B.textContent=f?"Simpan Perubahan":`Tambah ${c}`)}})}function ha(h){He(`Hapus ${c} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let f=await S(`${a}/${h.id}`,{method:"DELETE"});f.ok?(ee(`${c} berhasil dihapus.`),Q()):z(f.data?.error||"Gagal menghapus.")},`Hapus ${c}`)}return Q(),Q}j();H();async function Qt(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let r=await q(),a=await te(),l=e?e.get("dash_filter"):null,s={};if(l==="reliever"){let d=new Date,u=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;s={status:"Done",month:e&&e.get("month")?e.get("month"):u}}console.log("RAW",await Te()),console.log("OPTIONS",a);let o=d=>d&&!a.find(u=>u.value===d)?[...a,{value:d,label:d}]:a,n=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],i=d=>d&&!n.includes(d)?[...n,d]:n,c=new Date().getFullYear(),p=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((d,u)=>{let y=String(u+1).padStart(2,"0");return{value:`${c}-${y}`,label:`${d} ${c}`}});Gt({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:s,onDataLoaded:d=>d.sort((u,y)=>{let m=u.backup_date?new Date(u.backup_date).getTime():0;return(y.backup_date?new Date(y.backup_date).getTime():0)-m}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:d=>me(d)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:d=>window.formatDate(d)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:d=>window.formatDate(d)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:d=>d?`<span class="badge badge-info">${d}</span>`:"-"},{key:"status",label:"Status",render:d=>G(d)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:n},{type:"select",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"month",label:"Bulan",options:p},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:d=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:r,value:d?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:d?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:o(d?.original_fc_name),value:d?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:i(d?.reliever_name),value:d?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:d?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:d?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:d?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:d?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:d?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let d=await S(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(d.ok){let u=d.data.data.map(y=>({Cabang:y.branch_name||"","Nama Facility care":y.original_fc_name||"",Periode:y.period||"",Relifer:y.reliever_name||"","Tanggal Back Up":y.backup_date||"","Tanggal Selesai":y.completion_date||"",Keterangan:y.reason||"",Shift:y.shift||"",Status:y.status||""}));u.length===0&&u.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),M(u,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{M([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async d=>{let y=(await S("/api/branches?all=1")).data?.data||[],m=k=>{if(!k)return null;let w=String(k||"").toLowerCase(),T=y.find($=>String($.full_name||"").toLowerCase()===w||String($.code||"").toLowerCase()===w||String($.name||"").toLowerCase()===w);return T?T.id:null},b=d.map(k=>({branch_name:String(k.Cabang||"").trim(),backup_date:String(k["Tanggal Back Up"]||k["Tanggal Backup"]||"").trim(),original_fc_name:String(k["Nama Facility care"]||k["FC Digantikan"]||"").trim(),reliever_name:String(k.Relifer||k.Reliefer||"").trim(),period:String(k.Periode||"").trim(),reason:String(k.Keterangan||"").trim(),shift:String(k.Shift||"").trim(),completion_date:String(k["Tanggal Selesai"]||"").trim(),status:String(k.Status||"").trim()})).filter(k=>k.reliever_name&&k.backup_date),v=await S("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:b,onDuplicate:"update"})});if(!v.ok)throw new Error(v.data?.error||"Import gagal");return v.data}}})}j();H();async function zt(t){let e=await q(),r=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));O({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>me(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>G(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),s=await S(`/api/reports/inspection?limit=10000&${l}`);if(s.ok){let o=s.data.data.map(n=>({Cabang:n.branch_name||"",Periode:n.period||"",Tanggal:n.inspection_date||"","Point FC":n.fc_score!==null&&n.fc_score!==void 0?n.fc_score:"","Point SPV":n.spv_score!==null&&n.spv_score!==void 0?n.spv_score:"",Status:n.status||"","Link Dokumen":n.document_link||""}));M(o,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{M([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let l=i=>{if(!i)return null;let c=String(i||"").toLowerCase(),g=e.find(p=>String(p.label||"").toLowerCase()===c);return g?g.value:null},s=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let c=String(i).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let p=Number(c);if(p>2e4&&p<99999){let d=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let g=c.split(/[\/\-\.]/);if(g.length===3){let[p,d,u]=g.map(y=>y.trim());if(p.length===4&&d.length<=2&&u.length<=2)return`${p}-${d.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&d.length<=2&&p.length<=2)return`${u}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`}return c},o=a.map(i=>({branch_id:l(String(i.Cabang||"").trim()),period:String(i.Periode||"").trim(),inspection_date:s(i.Tanggal),fc_score:i["Point FC"]!==void 0&&i["Point FC"]!==""?Number(i["Point FC"]):null,spv_score:i["Point SPV"]!==void 0&&i["Point SPV"]!==""?Number(i["Point SPV"]):null,status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.period&&i.inspection_date),n=await S("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}j();H();async function Wt(t){let e=await q(),r=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));O({container:t,title:"Laporan GC & DC",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,enableMobileFilterSheet:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>me(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>G(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),s=await S(`/api/reports/cleaning?limit=10000&${l}`);if(s.ok){let o=s.data.data.map(n=>({Cabang:n.branch_name||"",Jenis:n.activity_type||"",Periode:n.period||"",Tanggal:n.activity_date||"",Status:n.status||"","Link Dokumen":n.document_link||""}));M(o,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{M([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let l=i=>{if(!i)return null;let c=String(i||"").toLowerCase(),g=e.find(p=>String(p.label||"").toLowerCase()===c);return g?g.value:null},s=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let c=String(i).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let p=Number(c);if(p>2e4&&p<99999){let d=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let g=c.split(/[\/\-\.]/);if(g.length===3){let[p,d,u]=g.map(y=>y.trim());if(p.length===4&&d.length<=2&&u.length<=2)return`${p}-${d.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&d.length<=2&&p.length<=2)return`${u}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`}return c},o=a.map(i=>({branch_id:l(String(i.Cabang||"").trim()),activity_type:String(i.Jenis||i.Kegiatan||"").trim(),period:String(i.Periode||"").trim(),activity_date:s(i.Tanggal),status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.activity_type&&i.period&&i.activity_date),n=await S("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}j();H();async function Yt(t,e){let r=await q(),a=Array.from({length:4},(o,n)=>String(new Date().getFullYear()-n)),l=e?e.get("dash_filter"):null,s={};if(l==="fogging"){let o=new Date,n=String(o.getMonth()+1).padStart(2,"0"),i=String(o.getFullYear()),c=e?e.get("month"):null;c&&c.length===7&&(i=c.split("-")[0],n=c.split("-")[1]),s={status:"Done",month:n,year:i}}O({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:s,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:o=>`<span class="badge badge-warning">${o}</span>`},{key:"period",label:"Periode",render:o=>me(o)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>G(o)},{key:"document_link",label:"Dokumen",render:o=>o?`<a href="${o}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:o=>o||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:r,value:o?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:o?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:o?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:o?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:o?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async o=>{let n=new URLSearchParams(o||{}).toString(),i=await S(`/api/reports/fogging?limit=10000&${n}`);if(i.ok){let c=i.data.data.map(g=>({Cabang:g.branch_name||"",Jenis:g.activity_type||"Fogging",Periode:g.period||"",Tanggal:g.activity_date||"",Status:g.status||"","Link Dokumen":g.document_link||""}));M(c,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{M([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async o=>{let n=p=>{if(!p)return null;let d=String(p||"").toLowerCase(),u=r.find(y=>String(y.label||"").toLowerCase()===d);return u?u.value:null},i=p=>{if(p==null||p==="")return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let d=String(p).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let y=Number(d);if(y>2e4&&y<99999){let m=new Date(Date.UTC(1899,11,30)+y*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}let u=d.split(/[\/\-\.]/);if(u.length===3){let[y,m,b]=u.map(v=>v.trim());if(y.length===4&&m.length<=2&&b.length<=2)return`${y}-${m.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&m.length<=2&&y.length<=2)return`${b}-${m.padStart(2,"0")}-${y.padStart(2,"0")}`}return d},c=o.map(p=>({branch_id:n(String(p.Cabang||"").trim()),activity_type:String(p.Jenis||p.Kegiatan||"Fogging").trim(),period:String(p.Periode||"").trim(),activity_date:i(p.Tanggal),status:String(p.Status||"").trim(),document_link:String(p["Link Dokumen"]||"").trim(),notes:String(p.Catatan||p.Keterangan||"").trim()})).filter(p=>p.branch_id&&p.period&&p.activity_date),g=await S("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(c)});if(!g.ok)throw new Error(g.data?.error||"Import gagal");return g.data}}})}j();H();async function Vt(t){let e=await q(),r=await te(),a=r,l=Array.from({length:4},(n,i)=>String(new Date().getFullYear()-i)),s=n=>n&&!r.find(i=>i.value===n)?[...r,{value:n,label:n}]:r,o=n=>n&&!a.find(i=>i.value===n)?[...a,{value:n,label:n}]:a;O({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:n=>`<span title="${n||""}">${n?.length>60?n.slice(0,60)+"\u2026":n||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:n=>window.formatDate(n)},{key:"status",label:"Status",render:n=>G(n)},{key:"notes",label:"Keterangan",render:n=>n?.length>40?n.slice(0,40)+"\u2026":n||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade","Mizwar"]},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:l}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:n?.branch_id},{name:"pic",label:"PIC",type:"select",options:o(n?.pic),value:n?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:n?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:n?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:n?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:n?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async n=>{let i=new URLSearchParams(n||{}).toString(),c=await S(`/api/reports/basecamp?limit=10000&${i}`);if(c.ok){let g=c.data.data.map(p=>({"Tgl Info":p.info_date||"",Cabang:p.branch_name||"",Permasalahan:p.problem||"",PIC:p.pic||"","Tgl Done":p.done_date||"",Status:p.status||"",Keterangan:p.notes||""}));M(g,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{M([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async n=>{let i=d=>{if(!d)return null;let u=String(d||"").toLowerCase(),y=e.find(m=>String(m.label||"").toLowerCase()===u);return y?y.value:null},c=d=>{if(d==null||d==="")return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let u=String(d).trim();if(u===""||u==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);if(/^\d{4,5}$/.test(u)){let m=Number(u);if(m>2e4&&m<99999){let b=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}let y=u.split(/[\/\-\.]/);if(y.length===3){let[m,b,v]=y.map(k=>k.trim());if(m.length===4&&b.length<=2&&v.length<=2)return`${m}-${b.padStart(2,"0")}-${v.padStart(2,"0")}`;if(v.length===4&&b.length<=2&&m.length<=2)return`${v}-${b.padStart(2,"0")}-${m.padStart(2,"0")}`}return u},g=n.map(d=>({info_date:c(d["Tgl Info"]||d["Tanggal Info"]),branch_id:i(String(d.Cabang||"").trim()),problem:String(d.Permasalahan||"").trim(),pic:String(d.PIC||"").trim(),done_date:c(d["Tgl Done"]||d["Tanggal Done"]),status:String(d.Status||"").trim(),notes:String(d.Keterangan||d.Catatan||"").trim()})).filter(d=>d.info_date&&d.branch_id&&d.problem),p=await S("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(g)});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}}})}async function Xt(t){O({container:t,title:"SOP",icon:"\u{1F4DA}",apiPath:"/api/sop",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(j(),_e)),l=await a(`/api/sop?limit=10000&${r}`);if(l.ok){let s=l.data.data.map(n=>({"Nama SOP":n.name||"",Kategori:n.category||"",Dokumen:n.document_link||"",Catatan:n.notes||n.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(H(),pe));o(s,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:r}=await Promise.resolve().then(()=>(H(),pe));r(e,"Template_Import_SOP")},onImport:async e=>{let r=e.map(s=>({name:String(s["Nama SOP"]||"").trim(),category:String(s.Kategori||"").trim(),document_link:String(s.Dokumen||"").trim(),description:String(s.Catatan||"").trim()})).filter(s=>s.name),{apiFetch:a}=await Promise.resolve().then(()=>(j(),_e)),l=await a("/api/sop/import",{method:"POST",body:JSON.stringify(r)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function Zt(t){O({container:t,title:"Master Checklist",icon:"\u{1F4CB}",apiPath:"/api/checklist",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(j(),_e)),l=await a(`/api/checklist?limit=10000&${r}`);if(l.ok){let s=l.data.data.map(n=>({"Nama Checklist":n.name||"",Kategori:n.category||"",Dokumen:n.document_link||"",Deskripsi:n.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(H(),pe));o(s,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:r}=await Promise.resolve().then(()=>(H(),pe));r(e,"Template_Import_Checklist")},onImport:async e=>{let r=e.map(s=>({name:String(s["Nama Checklist"]||"").trim(),category:String(s.Kategori||"").trim(),document_link:String(s.Dokumen||"").trim(),description:String(s.Deskripsi||"").trim()})).filter(s=>s.name),{apiFetch:a}=await Promise.resolve().then(()=>(j(),_e)),l=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(r)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}j();Se();H();async function _t(t,e="forms"){if(e==="supply")return ja(t);Ka(t)}function Ka(t){O({container:t,title:"Master Form",icon:"\u{1F4D1}",apiPath:"/api/forms",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await S(`/api/forms?limit=10000&${r}`);a.data?.data?M(a.data.data,"Data_Master_Form"):z("Gagal export data master form")},onImport:async e=>{let r=await S("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!r.ok)throw new Error(r.data?.error||"Import failed");return r.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function ja(t){let r=((await S("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));O({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,l)=>l.branch_name_ref||l.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>G(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let l=a?.tools_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}let s=a?.chemical_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:a?.branch_id&&!r.find(o=>o.value==a.branch_id)?[...r,{value:a.branch_id,label:a.branch_name||a.branch_id}]:r,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:l},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:s},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),s=await S(`/api/reports/supply?limit=10000&${l}`);if(s.ok){let o=s.data.data.map(n=>{let i=n.tools_items;try{i=Array.isArray(JSON.parse(i))?JSON.parse(i).join(", "):i}catch{}let c=n.chemical_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}return{Waktu:n.submitted_at||"",Pengirim:n.submitter_name||"",Cabang:n.branch_name_ref||n.branch_name||"","Alat/Barang":i||"",Chemical:c||"",Catatan:n.additional_notes||"",Status:n.status||"","Diproses Oleh":n.processed_by||""}});M(o,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{M([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let s=(await S("/api/branches?all=1")).data?.data||[],o=g=>{if(!g)return null;let p=String(g||"").toLowerCase(),d=s.find(u=>String(u.full_name||"").toLowerCase()===p||String(u.code||"").toLowerCase()===p||String(u.name||"").toLowerCase()===p);return d?d.id:null},n=g=>{if(g==null||g==="")return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let p=String(g).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let u=Number(p);if(u>2e4&&u<99999){let y=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}let d=p.split(/[\/\-\.]/);if(d.length===3){let[u,y,m]=d.map(b=>b.trim());if(u.length===4&&y.length<=2&&m.length<=2)return`${u}-${y.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&y.length<=2&&u.length<=2)return`${m}-${y.padStart(2,"0")}-${u.padStart(2,"0")}`}return p},i=a.map(g=>({submitted_at:n(g.Waktu||g.Tanggal),submitter_name:String(g.Pengirim||"").trim(),branch_id:o(String(g.Cabang||"").trim()),tools_items:String(g["Alat/Barang"]||g.Alat||"").trim(),chemical_items:String(g.Chemical||"").trim(),additional_notes:String(g.Catatan||g.Keterangan||"").trim(),status:String(g.Status||"").trim(),processed_by:String(g["Diproses Oleh"]||g.PIC||"").trim()})).filter(g=>g.submitted_at&&g.submitter_name&&g.branch_id),c=await S("/api/reports/supply/import",{method:"POST",body:JSON.stringify(i)});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,l)=>{let s=ue({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(o,n)=>{let i=o.querySelector("#supply-status").value,c=o.querySelector("#supply-processed-by").value;(await S(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:i,processed_by:c})})).ok?(ee("Status diperbarui."),n(),l()):z("Gagal update status.")}})}}]})}j();async function ea(t){let e=ce();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}O({container:t,title:"Riwayat Aktivitas",icon:"\u{1F6E1}\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:r=>new Date(r).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(r,a)=>`<strong>${r||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${a.user_role||""}</small>`},{key:"action",label:"Aksi",render:r=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[r]||"badge-neutral"}">${r}</span>`},{key:"module",label:"Modul",render:r=>`<span style="text-transform:capitalize">${(r||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(r,a)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${r}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async r=>{try{let s=((await(await fetch(`/api/audit-logs?search=${r}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(c=>String(c.id)===String(r));if(!s)return alert("Data tidak ditemukan");let o=c=>{if(!c)return"Tidak ada data";try{return JSON.stringify(JSON.parse(c),null,2)}catch{return c}},n=`
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
       `,{createModal:i}=await Promise.resolve().then(()=>(Se(),mt));i({title:`Detail Audit Log #${r}`,content:n,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}j();H();async function ta(t){let e=ce();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}O({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:r=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[r]||"badge-neutral"}">${r}</span>`},{key:"is_active",label:"Status",render:r=>r?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:r=>r?new Date(r).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:r=>{let a=!!r;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:r?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:r?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:r?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:r?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?r?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let r=await S(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let a=r.data.data.map(l=>({"Nama Lengkap":l.full_name||"",Username:l.username||"",Email:l.email||"",Role:l.role||"",Status:l.is_active?"Aktif":"Nonaktif"}));M(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{M([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async r=>{let a=r.map(s=>({full_name:String(s["Nama Lengkap"]||"").trim(),username:String(s.Username||"").trim(),email:String(s.Email||"").trim(),role:String(s.Role||"").trim()||"viewer",password:String(s.Password||"").trim()})).filter(s=>s.username&&s.password&&s.email&&s.full_name),l=await S("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}j();H();async function aa(t){O({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await S(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)M(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{M([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let r=e.map(l=>({code:String(l["Kode Cabang"]||"").trim(),name:String(l["Nama Pendek"]||"").trim(),full_name:String(l["Nama Lengkap"]||"").trim(),city:String(l.Kota||"").trim()})).filter(l=>l.code&&l.name),a=await S("/api/branches/import",{method:"POST",body:JSON.stringify(r)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}j();async function na(t){let e=new Date,r=[];t.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),l()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),l()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(s=>s.addEventListener("change",l));async function a(){try{let s=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;r=(await S(`/api/dashboard/calendar?month=${s}`)).data?.data||[]}catch(s){console.warn("[Calendar] Failed to load events, rendering empty grid:",s),r=[]}}async function l(){let s=document.getElementById("calendar-grid");if(s){s.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let o=e.getFullYear(),n=e.getMonth(),i=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),c=document.getElementById("cal-month-label");c&&(c.textContent=i);let g=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(T=>T.value)),p=r.filter(T=>g.has(T.type)),d={};p.forEach(T=>{let $=(T.event_date||"").slice(0,10);d[$]||(d[$]=[]),d[$].push(T)});let u=new Date(o,n,1).getDay(),y=new Date(o,n+1,0).getDate(),m=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],b=new Date().toISOString().slice(0,10),v='<div class="calendar-grid">';m.forEach(T=>{v+=`<div class="cal-day-header">${T}</div>`});for(let T=0;T<u;T++)v+='<div class="cal-cell cal-cell-empty"></div>';for(let T=1;T<=y;T++){let $=`${o}-${String(n+1).padStart(2,"0")}-${String(T).padStart(2,"0")}`,C=d[$]||[],x=$===b;v+=`
          <div class="cal-cell ${x?"cal-today":""} ${C.length?"cal-has-events":""}"
               data-date="${$}" tabindex="0" role="button" aria-label="${$}">
            <div class="cal-day-num ${x?"today-num":""}">${T}</div>
            <div class="cal-events-preview">
              ${C.slice(0,3).map(_=>`
                <div class="cal-event-dot cal-color-${_.color||"gray"}" title="${it(_.title||_.type)}">
                  <span class="cal-event-dot-label">${Ha(_.title||_.branch_name||_.type,18)}</span>
                </div>
              `).join("")}
              ${C.length>3?`<div class="cal-more">+${C.length-3} lagi</div>`:""}
            </div>
          </div>`}let w=(u+y)%7;if(w!==0)for(let T=0;T<7-w;T++)v+='<div class="cal-cell cal-cell-empty"></div>';v+="</div>",s.innerHTML=v,s.querySelectorAll(".cal-cell[data-date]").forEach(T=>{T.addEventListener("click",()=>{let $=T.dataset.date,C=d[$]||[];if(!C.length)return;let x=document.getElementById("cal-event-list"),_=new Date($+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=_,document.getElementById("cal-event-items").innerHTML=C.map(N=>`
            <div class="cal-event-item cal-color-border-${N.color||"gray"}">
              <div class="cal-event-type">${qa(N.type)}</div>
              <div class="cal-event-title">${it(N.title||"-")}</div>
              <div class="cal-event-branch">${it(N.branch_name||"")}</div>
              ${N.status?`<div class="cal-event-status">${it(N.status)}</div>`:""}
              ${N.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${N.days_remaining} hari</div>`:""}
            </div>
          `).join(""),x.style.display="block"})})}catch(o){console.error("[Calendar] Render error:",o),s&&(s.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}l()}function Ha(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function it(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function qa(t){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[t]||t}j();async function ia(t){let e=ce(),r=(e?.full_name||e?.username||"U")[0].toUpperCase(),l={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${l},${l}99)">
            ${r}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${e?.full_name||"\u2014"}</div>
            <div class="profile-username">@${e?.username||"\u2014"}</div>
            <span class="badge badge-info" style="background:${l}18;color:${l};margin-top:6px">
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
            <span class="info-value" style="color:${l};font-weight:700">${e?.role||"\u2014"}</span>
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
  `;let s=localStorage.getItem("fm_token"),o=document.getElementById("session-info");if(s&&o)try{let n=JSON.parse(atob(s.split(".")[1])),i=new Date(n.exp*1e3);o.textContent=`Berakhir: ${i.toLocaleString("id-ID")}`}catch{o.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async n=>{n.preventDefault();let i=document.getElementById("pwd-error"),c=document.getElementById("pwd-success"),g=document.getElementById("btn-save-pwd");i.style.display="none",c.style.display="none";let p=n.target,d=p.current_password.value,u=p.new_password.value,y=p.confirm_password.value;if(u!==y){i.textContent="\u274C Konfirmasi password tidak cocok.",i.style.display="block";return}if(u.length<6){i.textContent="\u274C Password baru minimal 6 karakter.",i.style.display="block";return}g.disabled=!0,g.textContent="\u23F3 Menyimpan...";let m=await S("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:d,new_password:u})});g.disabled=!1,g.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',m.ok?(c.textContent="\u2705 Password berhasil diubah.",c.style.display="block",p.reset(),ee("Password berhasil diubah.")):(i.textContent=m.data?.error||"Gagal mengubah password.",i.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}j();var rt={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function ae(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let l=Number(e);if(l>2e4&&l<99999){let s=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(s.getTime())?null:s.toISOString().slice(0,10)}}let r=e.split(/[\/\-\.]/);if(r.length===3){let[l,s,o]=r.map(g=>g.trim()),n=Number(l),i=Number(s),c=Number(o);if(l.length===4&&n>1900)return`${l}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`;if(o.length===4&&c>1900)return n>12?`${o}-${s.padStart(2,"0")}-${l.padStart(2,"0")}`:i>12?`${o}-${l.padStart(2,"0")}-${s.padStart(2,"0")}`:`${o}-${s.padStart(2,"0")}-${l.padStart(2,"0")}`;if(o.length===2&&!isNaN(c)){let g=c>=50?`19${o}`:`20${o}`;return n>12?`${g}-${s.padStart(2,"0")}-${l.padStart(2,"0")}`:`${g}-${s.padStart(2,"0")}-${l.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function ra(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Ja={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",phone:t["No. Hp"],join_date:ae(t["Tanggal Masuk"]),status:t.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:ae(t["Tanggal Mulai"]),end_date:ae(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:ae(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:ae(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:ae(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:ae(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:ae(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:ae(t["Tanggal Target"]||t["Tgl Target"]),completion_date:ae(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:ae(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:ae(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:ae(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:ae(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:ae(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:ae(t["Tanggal Back Up"]),completion_date:ae(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:ae(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:ae(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Ua(t,e){let r=rt[t];if(!r)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Ja[r.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let l=[],s=[],o=[];return e.filter(i=>!ra(i)).forEach((i,c)=>{let g=e.indexOf(i)+2,p=[];a.required.forEach(({key:u,label:y})=>{let m=i[u];if(m==null||String(m).trim()===""){let b=Object.keys(i).filter(v=>v.trim()).join(", ");p.push({column:y,originalValue:m||"",reason:`Kolom "${y}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${b.slice(0,120)}`})}});let d=a.map(i);p.length>0?s.push({row:g,data:d,raw:i,errors:p}):(l.push(i),o.push(d))}),{valid:l,errors:s,mapped:o}}function la(t){let e=[];return t.SheetNames.forEach(r=>{let a=rt[r];if(!a)return;let l=t.Sheets[r],s=window.XLSX.utils.sheet_to_json(l,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),o=Ua(r,s),n=s.filter(i=>!ra(i));e.push({sheetName:r,module:a.module,label:a.label,total:n.length,valid:o.mapped.length,errorCount:o.errors.length,errors:o.errors,mapped:o.mapped,skipped:!1})}),e}function sa(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,l])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(l),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function oa(t){let e=window.XLSX,r=e.utils.book_new(),a=!1;return t.forEach(l=>{if(!l.errors||l.errors.length===0)return;a=!0;let s=l.errors.map(n=>({"No. Baris":n.row,"Kolom Gagal":(n.errors||[]).map(i=>i.column||i).join("; "),"Alasan Error":(n.errors||[]).map(i=>i.reason||i).join("; "),...Object.fromEntries(Object.entries(n.data||{}).map(([i,c])=>[i,c??""]))})),o=e.utils.json_to_sheet(s);e.utils.book_append_sheet(r,o,l.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(r,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Ga=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function da(t){t.innerHTML=`
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
              ${Object.entries(rt).map(([m,{label:b}])=>`<span class="import-sheet-tag">\u{1F4C4} ${m} \u2192 ${b}</span>`).join("")}
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
  `;let e=null,r=null,a=0,l={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function s(m){Object.entries(l).forEach(([b,v])=>{v.style.display=b===m?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let m=document.getElementById("btn-backup-db");m.disabled=!0,m.textContent="\u23F3 Memproses Backup...";try{let b=await S("/api/import/backup");if(b.ok){if(!window.XLSX){z("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let v=window.XLSX,k=v.utils.book_new();Object.entries(b.data.database).forEach(([w,T])=>{let $=T.length>0?T:[{}],C=v.utils.json_to_sheet($);v.utils.book_append_sheet(k,C,w.substring(0,31))}),v.writeFile(k,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),ee("Backup berhasil diunduh!")}else z("Gagal memproses backup: "+(b.data?.error||"Unknown error"))}catch(b){z("Gagal memproses backup: "+b.message)}finally{m.disabled=!1,m.textContent="\u{1F4E6} Backup Database"}});let o=document.getElementById("btn-sync-google");o&&o.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let m=o.innerHTML;o.innerHTML='<span class="spinner"></span> Menyinkronkan...',o.disabled=!0;try{let b=await S("/api/sync/google-sheets",{method:"POST"});b.ok?alert("Sinkronisasi Berhasil: "+(b.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(b.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{o.innerHTML=m,o.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{sa(),ee("Template Excel berhasil didownload!")});let n=document.getElementById("file-input"),i=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",m=>{m.stopPropagation(),n.click()}),n.addEventListener("change",m=>{m.target.files[0]&&c(m.target.files[0])}),i.addEventListener("dragover",m=>{m.preventDefault(),i.classList.add("drag-over")}),i.addEventListener("dragleave",()=>i.classList.remove("drag-over")),i.addEventListener("drop",m=>{m.preventDefault(),i.classList.remove("drag-over");let b=m.dataTransfer.files[0];b&&b.name.match(/\.xlsx?$/i)?c(b):z("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,n.value="",document.getElementById("file-info").style.display="none",i.style.display="",s("upload")});async function c(m){e=m,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${m.name} (${(m.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",i.style.display="none",await g(m)}async function g(m){s("validating");let b=document.getElementById("validation-status"),v=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");b.textContent="Membaca file Excel...",v.style.width="20%",await ze(200);let k=await m.arrayBuffer(),w=window.XLSX.read(k,{type:"array",cellDates:!0});b.textContent=`Memvalidasi ${w.SheetNames.length} sheet...`,v.style.width="50%",await ze(100),r=la(w),v.style.width="100%",b.textContent="Validasi selesai!",await ze(300),p()}catch(k){s("upload"),z("Gagal memproses file: "+k.message),document.getElementById("file-info").style.display="flex",i.style.display="none"}}function p(){s("preview");let m=r.filter(_=>!_.skipped).length,b=r.reduce((_,N)=>_+N.total,0),v=r.reduce((_,N)=>_+N.valid,0),k=r.reduce((_,N)=>_+N.errorCount,0),w=b>0?Math.round(v/b*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${m} sheet</span>
      <span class="badge badge-secondary">${b} baris</span>
      <span class="badge badge-success">${v} valid (${w}%)</span>
      ${k>0?`<span class="badge badge-danger">${k} error</span>`:""}
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
          ${r.map((_,N)=>`
            <tr class="${_.errorCount>0?"row-error":_.skipped?"row-skipped":"row-ok"}">
              <td><strong>${_.sheetName}</strong></td>
              <td>${_.label}</td>
              <td style="text-align:center">${_.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${_.valid}</span></td>
              <td style="text-align:center">${_.errorCount>0?`<span class="badge badge-danger">${_.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${_.skipped?'<span class="badge badge-neutral">Dilewati</span>':_.errorCount>0&&_.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':_.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':_.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${_.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${N}">\u{1F50D} ${_.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,T.querySelectorAll(".btn-detail-error").forEach(_=>{_.addEventListener("click",()=>{let N=r[Number(_.dataset.idx)];d(N)})});let $=document.getElementById("error-detail-section"),C=document.getElementById("error-detail-container");C.innerHTML="",$.style.display="none";let x=document.getElementById("btn-start-import");v===0?(x.disabled=!0,x.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(x.disabled=!1,k>0?(x.innerHTML=`\u{1F680} Import ${v} Data Valid (${k} dilewati)`,x.title="Baris error akan dilewati, baris valid tetap diimport"):x.innerHTML=`\u{1F680} Mulai Import ${v} Data`)}function d(m){let b=document.getElementById("error-detail-section"),v=document.getElementById("error-detail-container");b.style.display="";let k=m.errors.slice(0,100).map(w=>(Array.isArray(w.errors)?w.errors:[]).map($=>{let C=typeof $=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${w.row}</span></td>
            <td><strong>${C?$.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${C&&$.originalValue!==void 0?$.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${C?$.reason:$}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${C&&$.aliases?`Gunakan salah satu nama kolom:<br><em>${$.aliases}</em>`:C&&$.hint?$.hint:""}
            </td>
          </tr>
        `}).join("")).join("");v.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${m.sheetName} \u2014 ${m.errorCount} baris error dari ${m.total} total
          ${m.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
            <tbody>${k||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${m.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,b.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{s("upload"),document.getElementById("file-info").style.display="none",i.style.display="",e=null,n.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!r)return;oa(r)?ee("Log error berhasil didownload."):ee("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let m=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";u(m)});async function u(m){s("importing"),a=Date.now();let b=[];Ga.forEach($=>{let C=r?.find(x=>x.module===$&&x.mapped?.length>0);C&&b.push(C)});let v=document.getElementById("import-steps-list");v.innerHTML=b.map($=>`
      <div class="import-step-item" id="step-item-${$.module}">
        <span class="step-item-icon" id="step-icon-${$.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${$.label} <span class="step-item-count">(${$.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${$.module}"></span>
      </div>
    `).join("");let k=document.getElementById("import-bar"),w=document.getElementById("import-current-status"),T={totalSheets:b.length,totalRows:b.reduce(($,C)=>$+C.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let $=0;$<b.length;$++){let C=b[$],x=document.getElementById(`step-icon-${C.module}`),_=document.getElementById(`step-status-${C.module}`);x.textContent="\u{1F504}",_.textContent="Mengimport...",w.textContent=`Mengimport ${C.label}...`,k.style.width=`${Math.round($/b.length*100)}%`;try{let N=await S(`/api/import/${C.module}`,{method:"POST",body:JSON.stringify({rows:C.mapped,onDuplicate:m})});if(N.ok){let F=N.data;T.inserted+=F.inserted||0,T.skipped+=F.skipped||0,T.moduleResults.push({label:C.label,inserted:F.inserted||0,skipped:F.skipped||0,status:"ok"}),x.textContent="\u2705",_.innerHTML=`<span class="badge badge-success">${F.inserted||0} berhasil</span>${F.skipped>0?` <span class="badge badge-neutral">${F.skipped} skip</span>`:""}`}else T.failed++,T.moduleResults.push({label:C.label,inserted:0,skipped:0,status:"error",error:N.data?.error}),x.textContent="\u274C",_.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(N){T.failed++,T.moduleResults.push({label:C.label,inserted:0,skipped:0,status:"error",error:N.message}),x.textContent="\u274C",_.innerHTML='<span class="badge badge-danger">Gagal</span>'}await ze(150)}k.style.width="100%",w.textContent="Selesai!",await ze(400),y(T)}function y(m){s("summary");let b=((Date.now()-a)/1e3).toFixed(1),v=m.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${v?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${v?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${m.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${m.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${m.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${m.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${m.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${m.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
      <div class="summary-stat-card info">
        <div class="stat-value">${b}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${m.moduleResults.map(k=>`
            <tr>
              <td>${k.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${k.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${k.skipped}</span></td>
              <td style="text-align:center">
                ${k.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${k.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,r=null,n.value="",document.getElementById("file-info").style.display="none",i.style.display="",s("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function ze(t){return new Promise(e=>setTimeout(e,t))}j();var lt=[],ca=[];async function pa(t){lt=await q(),ca=await te(),O({container:t,title:"Data SP (Surat Peringatan)",icon:"\u{1F4DC}",apiPath:"/api/sp",enableMobileFilterSheet:!0,itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:lt}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await S(`/api/sp?limit=10000&${r}`);if(a.ok){let l=a.data.data.map(o=>({"Nama Karyawan":o.employee_name||"",Divisi:o.division||"",Cabang:o.branch_name||"","Tanggal Sp":o.tanggal||"","Akhir Sp":o.akhir_sp||"","Jenis Sp":o.sp_type||"","Link Document / Foto":o.document_link||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(H(),pe));s(l,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(H(),pe));r(e,"Template_Import_SP")},onImport:async e=>{let r=o=>{if(!o)return null;let n=String(o||"").toLowerCase(),i=lt.find(c=>String(c.label||"").toLowerCase()===n);return i?i.value:null},a=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let n=String(o).trim();if(/^\d{4,5}$/.test(n)){let c=Number(n);if(c>2e4&&c<99999){let g=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(n))return n.slice(0,10);let i=n.split(/[\/\-\.]/);if(i.length===3){let[c,g,p]=i.map(d=>d.trim());if(c.length===4&&g.length<=2&&p.length<=2)return`${c}-${g.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&g.length<=2&&c.length<=2)return`${p}-${g.padStart(2,"0")}-${c.padStart(2,"0")}`}return n},l=e.map(o=>({employee_name:String(o["Nama Karyawan"]||"").trim(),division:String(o.Divisi||"").trim(),branch_id:r(String(o.Cabang||"").trim()),tanggal:a(o["Tanggal Sp"]),akhir_sp:a(o["Akhir Sp"]),sp_type:String(o["Jenis Sp"]||"").trim(),document_link:String(o["Link Document / Foto"]||"").trim()})).filter(o=>o.employee_name&&o.branch_id),s=await S("/api/import/sp",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:[{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:ca},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:lt,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}j();var Be=[],ua=[];async function ma(t){Be=await q(),ua=await te(),O({container:t,title:"Data Mutasi",icon:"\u{1F501}",apiPath:"/api/mutasi",enableMobileFilterSheet:!0,itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:Be},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:Be}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await S(`/api/mutasi?limit=10000&${r}`);if(a.ok){let l=a.data.data.map(o=>({Tanggal:o.tanggal||"","Nama Karyawan":o.employee_name||"","Cabang Asal":o.from_branch_name||"","Cabang Tujuan":o.to_branch_name||"",Status:o.status||"",Dokumen:o.document_link||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(H(),pe));s(l,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(H(),pe));r(e,"Template_Import_Mutasi")},onImport:async e=>{let r=o=>{if(!o)return null;let n=String(o||"").toLowerCase(),i=Be.find(c=>String(c.label||"").toLowerCase()===n);return i?i.value:null},a=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let n=String(o).trim();if(/^\d{4,5}$/.test(n)){let c=Number(n);if(c>2e4&&c<99999){let g=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(n))return n.slice(0,10);let i=n.split(/[\/\-\.]/);if(i.length===3){let[c,g,p]=i.map(d=>d.trim());if(c.length===4&&g.length<=2&&p.length<=2)return`${c}-${g.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&g.length<=2&&c.length<=2)return`${p}-${g.padStart(2,"0")}-${c.padStart(2,"0")}`}return n},l=e.map(o=>({tanggal:a(o.Tanggal),employee_name:String(o["Nama Karyawan"]||"").trim(),from_branch_id:r(String(o["Cabang Asal"]||"").trim()),to_branch_id:r(String(o["Cabang Tujuan"]||"").trim()),status:String(o.Status||"").trim(),document_link:String(o.Dokumen||"").trim()})).filter(o=>o.tanggal&&o.employee_name&&o.from_branch_id&&o.to_branch_id),s=await S("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:ua},{type:"select",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Be,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Be,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}j();async function ga(t){t.innerHTML=`
    <div class="p-6 max-w-7xl mx-auto space-y-6">
      <div class="flex flex-col md:flex-row justify-between items-start md:items-center">
        <div>
          <h2 class="text-2xl font-bold text-gray-800">Sync Dashboard</h2>
          <p class="text-gray-500">Monitor Bidirectional Sync Health & Queue</p>
        </div>
        <div class="mt-4 md:mt-0 flex gap-2 flex-wrap">
          <button id="btnRetryAll" class="px-3 py-2 bg-yellow-500 text-white rounded shadow hover:bg-yellow-600 transition text-sm">Retry All Failed</button>
          <button id="btnResetStuck" class="px-3 py-2 bg-orange-500 text-white rounded shadow hover:bg-orange-600 transition text-sm">Reset Stuck</button>
          <button id="btnReconcile" class="px-3 py-2 bg-purple-500 text-white rounded shadow hover:bg-purple-600 transition text-sm">Force Reconcile</button>
          <button id="btnRefreshSync" class="px-3 py-2 bg-blue-600 text-white rounded shadow hover:bg-blue-700 transition text-sm">Refresh</button>
        </div>
      </div>

      <!-- Overview Cards -->
      <div class="grid grid-cols-1 md:grid-cols-4 gap-4" id="syncOverviewCards">
        <!-- populated via js -->
      </div>

      <!-- Circuit Breaker Status -->
      <div class="bg-white rounded-lg shadow p-6 border-l-4" id="cbStatusCard">
        <div class="flex justify-between items-center mb-2">
          <h3 class="text-lg font-semibold text-gray-700">Circuit Breaker Status</h3>
          <div class="flex gap-2">
            <button id="btnPauseSync" class="px-3 py-1 bg-red-600 text-white rounded text-xs hover:bg-red-700">Pause Sync</button>
            <button id="btnResumeSync" class="px-3 py-1 bg-green-600 text-white rounded text-xs hover:bg-green-700">Resume Sync</button>
          </div>
        </div>
        <div class="flex items-center gap-4">
          <div id="cbStateBadge" class="px-3 py-1 rounded-full text-sm font-medium">Loading...</div>
          <p class="text-sm text-gray-600" id="cbStateDesc">Fetching state...</p>
        </div>
      </div>

      <!-- Queue Table -->
      <div class="bg-white rounded-lg shadow overflow-hidden">
        <div class="p-4 border-b flex justify-between items-center">
          <h3 class="text-lg font-semibold text-gray-800">Outbox Queue</h3>
          <div class="flex gap-2">
            <button id="btnRetrySelected" class="px-3 py-1 bg-gray-200 text-gray-700 rounded text-xs hover:bg-gray-300">Retry Selected</button>
            <select id="queueStatusFilter" class="border rounded px-2 py-1 text-sm">
              <option value="">All Status</option>
              <option value="PENDING">Pending</option>
              <option value="PROCESSING">Processing</option>
              <option value="FAILED">Failed</option>
              <option value="DEAD_LETTER">Dead Letter</option>
            </select>
          </div>
        </div>
        <div class="overflow-x-auto">
          <table class="w-full text-sm text-left">
            <thead class="bg-gray-50 text-gray-600">
              <tr>
                <th class="px-4 py-3 w-10"><input type="checkbox" id="chkAllQueue" /></th>
                <th class="px-4 py-3">ID</th>
                <th class="px-4 py-3">Module</th>
                <th class="px-4 py-3">Action</th>
                <th class="px-4 py-3">Status</th>
                <th class="px-4 py-3">Retry</th>
                <th class="px-4 py-3">Created</th>
              </tr>
            </thead>
            <tbody id="queueTableBody" class="divide-y">
              <tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">Loading queue...</td></tr>
            </tbody>
          </table>
        </div>
      </div>
      
      <!-- Performance Metrics & Histograms -->
      <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
        <!-- Histograms -->
        <div class="bg-white rounded-lg shadow overflow-hidden">
          <div class="p-4 border-b">
            <h3 class="text-lg font-semibold text-gray-800">Latency Histograms (P50 - P99)</h3>
          </div>
          <div class="p-4 overflow-x-auto text-sm" id="latencyHistograms">
            <p class="text-gray-500">Loading performance data...</p>
          </div>
        </div>

        <!-- Snapshots & Reconcile -->
        <div class="bg-white rounded-lg shadow overflow-hidden">
          <div class="p-4 border-b">
            <h3 class="text-lg font-semibold text-gray-800">Nightly Snapshot & Reconciliation</h3>
          </div>
          <div class="p-4 text-sm space-y-4">
            <div>
              <h4 class="font-medium text-gray-700">Latest Snapshot</h4>
              <p id="lblSnapshotStatus" class="text-gray-600 mt-1">Loading...</p>
            </div>
            <div>
              <h4 class="font-medium text-gray-700">Reconciliation Status</h4>
              <p id="lblReconStatus" class="text-gray-600 mt-1">Ready for checks.</p>
            </div>
          </div>
        </div>
      </div>
      
      <!-- Metrics -->
      <div class="bg-white rounded-lg shadow overflow-hidden">
        <div class="p-4 border-b">
          <h3 class="text-lg font-semibold text-gray-800">Queue & Throughput Trend</h3>
        </div>
        <div class="overflow-x-auto">
          <table class="w-full text-sm text-left">
            <thead class="bg-gray-50 text-gray-600">
              <tr>
                <th class="px-4 py-3">Module</th>
                <th class="px-4 py-3">Total Events</th>
                <th class="px-4 py-3">Avg Webhook (ms)</th>
                <th class="px-4 py-3">Avg D1 Exec (ms)</th>
                <th class="px-4 py-3">Avg Queue Wait (ms)</th>
              </tr>
            </thead>
            <tbody id="metricsTableBody" class="divide-y">
              <tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">Loading metrics...</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;let e=document.getElementById("btnRefreshSync"),r=document.getElementById("queueStatusFilter");e.addEventListener("click",l),r.addEventListener("change",n),document.getElementById("btnRetryAll").addEventListener("click",()=>a("retry",{allFailed:!0})),document.getElementById("btnResetStuck").addEventListener("click",()=>a("reset-stuck",{})),document.getElementById("btnPauseSync").addEventListener("click",()=>a("pause",{})),document.getElementById("btnResumeSync").addEventListener("click",()=>a("resume",{})),document.getElementById("btnReconcile").addEventListener("click",()=>a("reconcile",{module:"ALL",repairMode:!0})),document.getElementById("btnRetrySelected").addEventListener("click",()=>{let c=Array.from(document.querySelectorAll(".chk-queue:checked")).map(g=>g.value);if(c.length===0)return alert("No items selected");a("retry",{ids:c})}),document.getElementById("chkAllQueue").addEventListener("change",c=>{document.querySelectorAll(".chk-queue").forEach(g=>g.checked=c.target.checked)});async function a(c,g){if(confirm(`Are you sure you want to execute action: ${c}?`)){showLoading();try{let p=await S(`/api/sync/actions/${c}`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(g)});p.ok?(alert(p.data?.message||"Success"),l()):z(p.error||"Action failed")}catch(p){z(p.message)}hideLoading()}}await l();async function l(){showLoading(),await Promise.all([o(),n(),s(),i()]),hideLoading()}async function s(){try{let c=await S("/api/sync/performance");if(!c.ok)return;let{webhook:g,google_api:p,d1:d,queue:u,throughput:y}=c.data;document.getElementById("latencyHistograms").innerHTML=`
        <table class="w-full text-left">
          <thead><tr class="text-gray-500 border-b"><th>Metric</th><th>P50</th><th>P95</th><th>P99</th><th>Max</th></tr></thead>
          <tbody class="divide-y">
            <tr><td class="py-1">Webhook</td><td>${g.P50}ms</td><td>${g.P95}ms</td><td>${g.P99}ms</td><td>${g.Max}ms</td></tr>
            <tr><td class="py-1">Google API</td><td>${p.P50}ms</td><td>${p.P95}ms</td><td>${p.P99}ms</td><td>${p.Max}ms</td></tr>
            <tr><td class="py-1">D1 Execute</td><td>${d.P50}ms</td><td>${d.P95}ms</td><td>${d.P99}ms</td><td>${d.Max}ms</td></tr>
            <tr><td class="py-1">Queue Delay</td><td>${u.P50}ms</td><td>${u.P95}ms</td><td>${u.P99}ms</td><td>${u.Max}ms</td></tr>
          </tbody>
        </table>
        <div class="mt-4 pt-3 border-t text-gray-600 flex justify-between">
          <span>Throughput: <b>${y.events_per_sec}</b> ev/sec</span>
          <span><b>${y.events_per_min}</b> ev/min</span>
        </div>
      `,document.getElementById("lblSnapshotStatus").innerHTML='Checked 10 modules. Status: <b>COMPLETED</b><br><span class="text-xs text-gray-500">Run today at 02:00</span>'}catch(c){console.error(c)}}async function o(){try{let c=await S("/api/sync/health");if(!c.ok)return z("Failed to fetch sync health");let{status:g,queue:p,circuit_breaker:d}=c.data,u=`
        <div class="bg-white p-4 rounded-lg shadow border-l-4 ${g==="HEALTHY"?"border-green-500":g==="WARNING"?"border-yellow-500":"border-red-500"}">
          <p class="text-sm text-gray-500">System Health</p>
          <p class="text-2xl font-bold ${g==="HEALTHY"?"text-green-600":g==="WARNING"?"text-yellow-600":"text-red-600"}">${g}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-blue-500">
          <p class="text-sm text-gray-500">Pending Queue</p>
          <p class="text-2xl font-bold text-gray-800">${p.pending||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-orange-500">
          <p class="text-sm text-gray-500">Failed / Retrying</p>
          <p class="text-2xl font-bold text-gray-800">${p.failed||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-red-500">
          <p class="text-sm text-gray-500">Dead Letters</p>
          <p class="text-2xl font-bold text-gray-800">${p.dead_letter||0}</p>
        </div>
      `;document.getElementById("syncOverviewCards").innerHTML=u;let y=document.getElementById("cbStateBadge"),m=document.getElementById("cbStateDesc"),b=document.getElementById("cbStatusCard");b.className="bg-white rounded-lg shadow p-6 border-l-4",d==="CLOSED"?(y.className="px-3 py-1 rounded-full text-sm font-medium bg-green-100 text-green-800",y.textContent="CLOSED",m.textContent="Traffic is flowing normally to Google Sheets.",b.classList.add("border-green-500")):d==="OPEN"?(y.className="px-3 py-1 rounded-full text-sm font-medium bg-red-100 text-red-800",y.textContent="OPEN",m.textContent="Failing fast. Traffic to Google Sheets is paused due to repeated failures.",b.classList.add("border-red-500")):d==="HALF_OPEN"?(y.className="px-3 py-1 rounded-full text-sm font-medium bg-yellow-100 text-yellow-800",y.textContent="HALF-OPEN",m.textContent="Testing recovery. Permitting limited traffic to verify stability.",b.classList.add("border-yellow-500")):y.textContent=d||"UNKNOWN"}catch(c){console.error(c)}}async function n(){try{let c=document.getElementById("queueStatusFilter").value,g=await S("/api/sync/queue?limit=15"+(c?"&status="+c:""));if(!g.ok)return;let p=document.getElementById("queueTableBody"),d=g.data?.data||g.data||[];if(d.length===0){p.innerHTML='<tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">No events found</td></tr>';return}p.innerHTML=d.map(u=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2"><input type="checkbox" class="chk-queue" value="${u.id}" /></td>
          <td class="px-4 py-2 font-mono text-xs text-gray-500" title="${u.id}">${u.id.split("-")[0]}...</td>
          <td class="px-4 py-2 font-medium">${u.entity_name}</td>
          <td class="px-4 py-2">
             <span class="px-2 py-0.5 rounded text-xs ${u.action==="INSERT"?"bg-blue-100 text-blue-800":u.action==="UPDATE"?"bg-purple-100 text-purple-800":"bg-red-100 text-red-800"}">${u.action}</span>
          </td>
          <td class="px-4 py-2">
             <span class="px-2 py-0.5 rounded text-xs ${u.status==="PENDING"?"bg-yellow-100 text-yellow-800":u.status==="PROCESSING"?"bg-blue-100 text-blue-800":u.status==="DEAD_LETTER"?"bg-red-100 text-red-800":"bg-gray-100 text-gray-800"}">${u.status}</span>
             ${u.last_error?`<br><span class="text-xs text-red-500 max-w-xs block truncate" title="${u.last_error}">${u.last_error}</span>`:""}
          </td>
          <td class="px-4 py-2 text-gray-600">${u.retry_count||0}</td>
          <td class="px-4 py-2 text-gray-500 whitespace-nowrap">${window.formatDate(u.created_at)} ${new Date(u.created_at).toLocaleTimeString("id-ID")}</td>
        </tr>
      `).join("")}catch(c){console.error(c)}}async function i(){try{let c=await S("/api/sync/metrics");if(!c.ok)return;let g=document.getElementById("metricsTableBody"),p=c.data||[];if(p.length===0){g.innerHTML='<tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">No metrics found in last 24h</td></tr>';return}g.innerHTML=p.map(d=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2 font-medium text-gray-800">${d.module}</td>
          <td class="px-4 py-2 text-gray-600">${d.total_events}</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(d.avg_webhook_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(d.avg_d1_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(d.avg_queue_wait_ms||0)} ms</td>
        </tr>
      `).join("")}catch(c){console.error(c)}}}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let e=Math.floor(Number(t)-25569);return new Date(e*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let e=t.split(/[\/\-]/);return`${e[2]}-${e[1]}-${e[0]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let r=e.split("-");if(r.length===3&&r[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],l=parseInt(r[2],10),s=a[parseInt(r[1],10)-1];return`${l} ${s} ${r[0]}`}return e};function W(t){return async e=>{if(!Oe()){Ce("/login");return}return t(e)}}var We=null;function Qa(){We&&clearInterval(We);let t=()=>{let e=new Date,r=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),l=document.getElementById("header-clock-time"),s=document.getElementById("header-clock-date");l&&(l.textContent=r),s&&(s.textContent=a)};t(),We=setInterval(t,1e3)}async function za(){try{let t=await S("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},r=(a,l)=>{let s=document.getElementById(a);s&&(s.textContent=l>0?l:"",s.style.display=l>0?"inline-flex":"none")};r("badge-issues",e.issues?.current||0),r("badge-contracts",e.expiring30?.current||0),r("badge-oo1",e.one_on_one?.current||0),r("badge-schedule",e.schedule?.current||0),r("badge-supply",e.supply?.current||0)}catch{}}var Ae=[];async function Wa(){try{let t=await S("/api/dashboard/notifications");if(!t.ok)return;Ae=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Ae.length>0?"block":"none",e.textContent=Ae.length)}catch{}}function Ya(){if(!Ae.length){ue({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,r)=>r()});return}let t=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${Ae.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;ue({title:`Notifikasi (${Ae.length})`,content:t,confirmText:"Tutup",onConfirm:(e,r)=>r()})}function ba(){let t=ce(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
            <a href="#/audit-logs" class="nav-item" data-route="/audit-logs">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
              </span>
              <span class="nav-label">Riwayat Aktivitas</span>
            </a>
          </div>`:""}
        </nav>

        <div class="sidebar-footer">
          <div class="sidebar-user">
            <div class="sidebar-avatar">${e}</div>
            <div class="sidebar-user-info">
              <div class="sidebar-user-name">${t?.full_name||"Admin"}</div>
              <div class="sidebar-user-role" style="text-transform:capitalize">${t?.role||"Administrator"}</div>
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
                <span class="topbar-greeting-time">${(()=>{let c=new Date().getHours();return c>=4&&c<11?"Selamat Pagi":c>=11&&c<15?"Selamat Siang":c>=15&&c<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${t?.full_name||"Admin"}</span> \u{1F44B}
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
              <img src="https://ui-avatars.com/api/?name=${encodeURIComponent(t?.full_name||"Admin")}&background=2563EB&color=fff&bold=true" class="topbar-avatar" alt="Avatar" />
              <div class="topbar-user-text">
                <span class="topbar-user-name">${t?.full_name||"Admin"}</span>
                <span class="topbar-user-role-mini" style="text-transform:capitalize">${t?.role||"Administrator"}</span>
              </div>
              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" style="margin-left:4px;color:var(--gray-400)"><polyline points="6 9 12 15 18 9"/></svg>
            </a>
          </div>
        </header>

        <main id="main-content" class="main-content"></main>
      </div>
    </div>
  `;let r=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),l=document.getElementById("topbar-menu-btn"),s=document.getElementById("sidebar-close"),o=()=>{r.classList.add("open"),a.classList.add("show")},n=()=>{r.classList.remove("open"),a.classList.remove("show")};l?.addEventListener("click",o),s?.addEventListener("click",n),a?.addEventListener("click",n),document.querySelectorAll(".nav-item").forEach(c=>c.addEventListener("click",n));function i(){let c=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(d=>{let u=d.dataset.route;d.classList.toggle("active",c===u||u!=="/dashboard"&&c.startsWith(u))});let g=document.getElementById("topbar-title"),p=document.querySelector(".nav-item.active .nav-label");g&&p&&(g.textContent=p.textContent)}window.addEventListener("hashchange",i),i(),Qa(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await S("/api/auth/logout",{method:"POST"}),Re(),We&&clearInterval(We),Ce("/login")}),za(),Wa(),document.getElementById("btn-notif")?.addEventListener("click",c=>{c.preventDefault(),Ya()})}document.addEventListener("keydown",t=>{if(t.target.tagName==="SELECT"&&t.key.length===1&&/[a-zA-Z]/.test(t.key)){let e=t.key.toLowerCase(),r=Array.from(t.target.options);if(r.length===0)return;let a=t.target.selectedIndex+1;(a>=r.length||a<0)&&(a=0);let l=-1;for(let s=0;s<r.length;s++){let o=(a+s)%r.length;if(r[o].text.toLowerCase().replace(/^[\d\.\s\-]+/,"").startsWith(e)){l=o;break}}l!==-1&&l!==t.target.selectedIndex&&(t.target.selectedIndex=l,t.preventDefault(),t.target.dispatchEvent(new Event("change")))}});async function Va(){U("/login",({main:e})=>Ot(e)),U("/dashboard",W(({main:e})=>Bt(e))),U("/calendar",W(({main:e})=>na(e))),U("/employees",W(({main:e,params:r})=>Rt(e,r))),U("/contracts",W(({main:e,params:r})=>jt(e,r))),U("/sp",W(({main:e})=>pa(e))),U("/mutasi",W(({main:e})=>ma(e))),U("/sync-dashboard",W(({main:e})=>ga(e))),U("/timeline",W(({main:e,params:r})=>Ht(e,r))),U("/issues",W(({main:e,params:r})=>qt(e,r))),U("/one-on-one",W(({main:e,params:r})=>Jt(e,r))),U("/training",W(({main:e})=>Ut(e))),U("/relievers",W(({main:e,params:r})=>Qt(e,r))),U("/reports/inspection",W(({main:e})=>zt(e))),U("/reports/cleaning",W(({main:e})=>Wt(e))),U("/reports/fogging",W(({main:e})=>Yt(e))),U("/reports/basecamp",W(({main:e})=>Vt(e))),U("/reports/supply",W(({main:e})=>_t(e,"supply"))),U("/sop",W(({main:e})=>Xt(e))),U("/checklist",W(({main:e})=>Zt(e))),U("/forms",W(({main:e})=>_t(e))),U("/users",W(({main:e})=>ta(e))),U("/branches",W(({main:e})=>aa(e))),U("/profile",W(({main:e})=>ia(e))),U("/settings/import",W(({main:e})=>da(e))),U("/audit-logs",W(({main:e})=>ea(e)));let t=Oe();if(!t&&window.location.hash!=="#/login"&&Ce("/login"),t){let e=await S("/api/auth/me");e.ok?(Ke(e.data.data),ba()):(Re(),Ce("/login"))}window.addEventListener("fm:login",()=>{ba(),Ce("/dashboard")}),$t()}Va();
