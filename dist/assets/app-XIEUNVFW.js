var fa=Object.defineProperty;var dt=(t,e)=>()=>(t&&(e=t(t=0)),e);var ct=(t,e)=>{for(var i in e)fa(t,i,{get:e[i],enumerable:!0})};var Le={};ct(Le,{API:()=>Et,CLIENT_SIDE_MAX_ROWS:()=>Se,IS_DEVELOPMENT:()=>$e,apiFetch:()=>k,clearToken:()=>We,getToken:()=>Qe,getUser:()=>oe,setToken:()=>pt,setUser:()=>Ve});function Qe(){return localStorage.getItem("fm_token")}function pt(t){localStorage.setItem("fm_token",t)}function We(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function oe(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Ve(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function k(t,e={}){let i=Qe(),a={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...e.headers||{}};try{let o=`cb=${Date.now()}`,d=t.includes("?")?"&":"?",r=`${Et}${t}${d}${o}`,n=await fetch(r,{...e,headers:a}),s;try{let l=await n.text();try{s=JSON.parse(l)}catch{s={error:`Server Error (${n.status}): ${l.substring(0,80)}...`}}}catch{s={error:"Gagal membaca respon dari server"}}return n.status===401&&(We(),window.location.hash="#/login"),{ok:n.ok,status:n.status,data:s}}catch(o){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${o.message})`}}}}var $e,va,Et,Se,j=dt(()=>{$e=!1,va="https://fm-operations-api.facilitycare-audydental.workers.dev",Et=va,Se=1e4});var mt={};ct(mt,{confirmDialog:()=>Ee,createModal:()=>de});function de({title:t,content:e,onConfirm:i,onCancel:a,confirmText:o="Simpan",cancelText:d="Batal",size:r="md",confirmClass:n="btn-primary"}){let s={sm:"400px",md:"560px",lg:"720px",xl:"900px"},l=document.createElement("div");l.className="modal-overlay",l.innerHTML=`
    <div class="modal" style="max-width:${s[r]||s.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${d}</button>
        ${i?`<button class="btn ${n} modal-confirm">${o}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&l.querySelector(".modal-body").appendChild(e);let b=()=>{l.classList.remove("show"),setTimeout(()=>l.remove(),250)};return l.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),b()}),l.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),b()}),i&&l.querySelector(".modal-confirm").addEventListener("click",()=>i(l,b)),l.addEventListener("click",c=>{c.target===l&&(a&&a(),b())}),document.body.appendChild(l),requestAnimationFrame(()=>l.classList.add("show")),{overlay:l,close:b}}function Ee(t,e,i="Konfirmasi"){return de({title:i,content:`<p>${t}</p>`,onConfirm:(a,o)=>{e(),o()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var we=dt(()=>{});var ce={};ct(ce,{downloadExcel:()=>R,parseExcel:()=>De,renderExcelButtons:()=>Ie});function De(t){return new Promise((e,i)=>{let a=new FileReader;a.onload=o=>{try{let d=new Uint8Array(o.target.result),r=XLSX.read(d,{type:"array"}),n=r.SheetNames[0],s=r.Sheets[n];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${r.SheetNames.join(", ")}`),console.log(`Sheet Used: ${n}`);let l=XLSX.utils.decode_range(s["!ref"]||"A1:A1"),b=l.e.r-l.s.r+1,c=l.e.c-l.s.c+1;console.log(`Total Rows (including empty): ${b}`),console.log(`Total Columns: ${c}`);let u=[];for(let y=l.s.c;y<=l.e.c;++y){let g=s[XLSX.utils.encode_cell({c:y,r:l.s.r})];g&&g.v&&u.push(g.v)}console.log(`Headers Found: ${u.join(", ")}`),console.log("---------------------------");let m=XLSX.utils.sheet_to_json(s,{defval:""});Object.defineProperty(m,"__worksheet",{value:s,enumerable:!1}),Object.defineProperty(m,"__headers",{value:u,enumerable:!1}),e(m)}catch(d){i(d)}},a.onerror=o=>i(o),a.readAsArrayBuffer(t)})}function R(t,e){try{let i=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,i,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function Ie(t){return`
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
  `}var G=dt(()=>{});j();var ut={},at=null;function Y(t,e){ut[t]=e}function Be(t){window.location.hash=t}function Tt(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[i,...a]=e.split("?"),o=ut[i];if(!o){for(let[r,n]of Object.entries(ut))if(r.endsWith("/*")&&i.startsWith(r.slice(0,-2))){o=n;break}}at&&(at(),at=null);let d=document.getElementById("main-content");if(d&&(d.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),o){let r=new URLSearchParams(a.join("?")),n=i.split("/").filter(Boolean),s=await o({path:i,params:r,segments:n,main:d});s&&(at=s)}else{let r=d||document.getElementById("app");r&&(r.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var Ye;function ka(){return Ye||(Ye=document.createElement("div"),Ye.id="toast-container",document.body.appendChild(Ye)),Ye}function Dt(t,e="info",i=3500){let a=ka(),o=document.createElement("div");o.className=`toast toast-${e}`;let d={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};o.innerHTML=`<span class="toast-icon">${d[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(o),requestAnimationFrame(()=>o.classList.add("show")),setTimeout(()=>{o.classList.remove("show"),setTimeout(()=>o.remove(),350)},i)}var ne=t=>Dt(t,"success"),z=t=>Dt(t,"error");we();j();j();we();function Pe({columns:t,data:e,onEdit:i,onDelete:a,onView:o,actions:d=[],emptyText:r="Tidak ada data",bulkSelect:n=null}){let s=document.createElement("div");if(s.className="table-wrapper",!e||e.length===0)return s.innerHTML=`<div class="empty-state"><p>${r}</p></div>`,s;let l=document.createElement("table");l.className="data-table";let b=document.createElement("thead"),c=document.createElement("tr");if(n){let m=document.createElement("th");m.style.width="40px",m.style.textAlign="center";let y=document.createElement("input");y.type="checkbox",y.id="select-all-checkbox",y.title="Pilih semua",y.addEventListener("change",()=>{e.forEach(g=>{y.checked?n.selectedIds.add(g.id):n.selectedIds.delete(g.id)}),s.querySelectorAll(".row-checkbox").forEach(g=>g.checked=y.checked),n.onToggle()}),m.appendChild(y),c.appendChild(m)}if(t.forEach(m=>{let y=document.createElement("th");y.textContent=m.label,m.width&&(y.style.width=m.width),c.appendChild(y)}),i||a||o||d.length>0){let m=document.createElement("th");m.textContent="Aksi",m.style.width="120px",c.appendChild(m)}b.appendChild(c),l.appendChild(b);let u=document.createElement("tbody");return e.forEach(m=>{let y=document.createElement("tr");if(n){let g=document.createElement("td");g.style.textAlign="center",g.style.width="40px";let f=document.createElement("input");f.type="checkbox",f.className="row-checkbox",f.checked=n.selectedIds.has(m.id),f.addEventListener("change",()=>{if(f.checked)n.selectedIds.add(m.id);else{n.selectedIds.delete(m.id);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1)}n.onToggle()}),g.appendChild(f),y.appendChild(g)}if(t.forEach(g=>{let f=document.createElement("td");if(g.render){let v=g.render(m[g.key],m);v instanceof HTMLElement?f.appendChild(v):f.innerHTML=v||""}else f.textContent=m[g.key]!==null&&m[g.key]!==void 0&&m[g.key]!==""?m[g.key]:"";g.nowrap&&(f.style.whiteSpace="nowrap"),y.appendChild(f)}),i||a||o||d.length>0){let g=document.createElement("td");g.className="actions-cell";let f=document.createElement("div");if(f.className="btn-group",o){let v=document.createElement("button");v.className="btn btn-xs btn-ghost",v.innerHTML="\u{1F441}",v.title="Lihat",v.addEventListener("click",()=>o(m)),f.appendChild(v)}if(i){let v=document.createElement("button");v.className="btn btn-xs btn-secondary",v.innerHTML="\u270F\uFE0F",v.title="Edit",v.addEventListener("click",()=>i(m)),f.appendChild(v)}d.forEach(v=>{let S=document.createElement("button");S.className=`btn btn-xs ${v.class||"btn-ghost"}`,S.innerHTML=v.icon||v.label,S.title=v.label,S.addEventListener("click",()=>v.handler(m)),f.appendChild(S)}),g.appendChild(f),y.appendChild(g)}u.appendChild(y)}),l.appendChild(u),s.appendChild(l),s}function Ae({page:t,pages:e,total:i,limit:a,onPage:o}){if(e<=1)return null;let d=document.createElement("div");d.className="pagination";let r=document.createElement("span");r.className="pagination-info",r.textContent=`Total: ${i} data`,d.appendChild(r);let n=document.createElement("div");n.className="pagination-btns";let s=(c,u,m=!1,y=!1)=>{let g=document.createElement("button");g.className=`btn btn-sm ${y?"btn-primary":"btn-ghost"} pagination-btn`,g.textContent=c,g.disabled=m,g.addEventListener("click",()=>o(u)),n.appendChild(g)};s("\xAB",1,t===1),s("\u2039",t-1,t===1);let l=Math.max(1,t-2),b=Math.min(e,t+2);for(let c=l;c<=b;c++)s(c,c,!1,c===t);return s("\u203A",t+1,t===e),s("\xBB",e,t===e),d.appendChild(n),d}we();function Te(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${Te(e.fields)}</div>`;let i=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",o="";switch(e.type){case"textarea":o=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${i} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let r=(e.options||[]).map(c=>{let u=typeof c=="object"?c.value:c,m=typeof c=="object"?c.label:c,y=e.value==u?"selected":"";return`<option value="${u}" ${y}>${m}</option>`}).join("");o=`<select name="${e.name}" class="form-control" ${i}><option value="">-- Pilih ${e.label||""} --</option>${r}</select>`;break;case"combobox":let n=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,s=(e.options||[]).map(c=>{let u=typeof c=="object"?c.value:c,m=typeof c=="object"?c.label||c.value||"":c||"";return(m==="undefined"||m==="[object Object]"||m==="null")&&(m=""),m?`<option value="${m}"></option>`:""}).join(""),l=e.value||"";if(e.value){let c=(e.options||[]).find(u=>(typeof u=="object"?u.value:u)==e.value);if(c){let u=typeof c=="object"?c.label||c.value||"":c||"";u&&u!=="undefined"&&u!=="[object Object]"&&u!=="null"&&(l=u)}}o=`
          <input type="text" name="${e.name}" list="${n}" class="form-control" value="${l}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${n}">${s}</datalist>
        `;break;case"checkbox":o=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let b=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";o=`<input type="date" name="${e.name}" class="form-control" value="${b}" ${i}>`;break;case"number":o=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${i}>`;break;case"email":o=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i}>`;break;case"url":o=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${i}>`;break;default:o=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i} autocomplete="off">`}let d=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${o}${d}</div>`}).join("")}function Ne(t){let e={},i=new FormData(t);for(let[a,o]of i.entries())e[a]=o===""?null:o;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function Me(t,e){e&&Object.entries(e).forEach(([i,a])=>{let o=t.querySelector(`[name="${i}"]`);o&&(o.hasAttribute("list")||(o.type==="checkbox"?o.checked=!!a:o.type==="date"&&a&&window.parseFlexibleDate?o.value=window.parseFlexibleDate(a):o.value=a??""))})}G();function gt({container:t,title:e,icon:i,apiPath:a,columns:o,formFields:d,filterFields:r,defaultFilters:n={},enableMobileFilterSheet:s=!1,itemLabel:l="Data",canCreate:b=!0,canEdit:c=!0,canDelete:u=!0,onBeforeSubmit:m,onAfterLoad:y,onDataLoaded:g,extraActions:f=[],initialSearch:v="",exportOptions:S=null,bulkDelete:D=!1,paginationMode:C="server"}){let P=oe();P&&typeof P=="object"&&P.role==="viewer"&&(b=!1,c=!1,u=!1,D=!1,S=null);let $=1,L={...n};v&&(L.search=v);let E=new Set;t.innerHTML=`
    <div class="crud-layout-wrapper ${s?"mobile-active":""}">
      <div class="page-header">
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions">
        ${b?`<button class="btn btn-primary" id="btn-create">+ Tambah ${l}</button>`:""}
        ${S?'<button class="btn btn-outline" id="btn-mobile-aksi" style="display:none; align-items:center; justify-content:center; gap:0.25rem;">\u22EE Aksi</button>':""}
      </div>
    </div>

    ${D?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${S?`
    <div class="excel-actions-wrapper" id="excel-actions-wrapper">
      <div class="bottom-sheet-header aksi-header" style="display:none;">
        <h3 style="margin:0; font-size:1rem;">Aksi</h3>
        <button class="btn-close-sheet" id="btn-close-aksi-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
      </div>
      ${Ie(S.moduleName)}
    </div>`:""}

    ${r&&r.length>0?`
    <div class="filter-bar card ${s?"has-mobile-sheet":""}" style="padding: 1rem; display: flex; gap: 0.5rem; flex-wrap: wrap;">
        ${r.filter(p=>p.type==="search"||p.type==="search-combo").map(p=>{if(p.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${p.placeholder||"Cari..."}" id="filter-search" value="${L.search||""}"></div>`;if(p.type==="search-combo"){let h="dl-filter-search",w=(p.options||[]).map(x=>`<option value="${typeof x=="object"?x.label:x}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${h}" class="form-control" autocomplete="off" placeholder="${p.placeholder||"Cari..."}" id="filter-search" value="${L.search||""}"><datalist id="${h}">${w}</datalist></div>`}return""}).join("")}
        
        <div class="filter-options" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${r.filter(p=>p.type!=="search"&&p.type!=="search-combo").map(p=>{if(p.type==="select")return`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${p.name}" id="filter-${p.name}"><option value="">Pilih ${p.label}</option>${(p.options||[]).map(h=>`<option value="${typeof h=="object"?h.value:h}" ${L[p.name]===(typeof h=="object"?h.value:h)?"selected":""}>${typeof h=="object"?h.label:h}</option>`).join("")}</select>`;if(p.type==="combobox"){let h=`dl-filter-${p.name}`,w=(p.options||[]).map(T=>{let _=typeof T=="object"?T.label||T.value||"":T||"";return(_==="undefined"||_==="[object Object]"||_==="null")&&(_=""),_?`<option value="${_}"></option>`:""}).join(""),x=L[p.name]||"";if(L[p.name]){let T=(p.options||[]).find(_=>String(typeof _=="object"?_.value:_)==String(L[p.name]));if(T){let _=typeof T=="object"?T.label||T.value||"":T||"";_&&_!=="undefined"&&_!=="[object Object]"&&_!=="null"&&(x=_)}}return`<div class="filter-combobox" style="flex:1; min-width:120px;">
                <input type="text" name="${p.name}" id="filter-${p.name}" list="${h}" class="form-control filter-combobox-input" value="${x}" placeholder="Pilih ${p.label}..." autocomplete="off">
                <datalist id="${h}">${w}</datalist>
              </div>`}return""}).join("")}
          <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
        </div>
        ${s?'<button id="btn-mobile-filter" class="btn btn-outline" style="display:none; align-items:center; gap:0.25rem;">\u2699\uFE0F Filter</button>':""}
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
    </div>
  `;function F(){let p=document.getElementById("bulk-toolbar");if(!p)return;let h=document.getElementById("bulk-count"),w=document.getElementById("btn-bulk-delete"),x=document.getElementById("btn-bulk-cancel");h.textContent=`${E.size} item dipilih`,E.size>0?(p.classList.add("has-items"),w.disabled=!1,x.disabled=!1):(p.classList.remove("has-items"),w.disabled=!0,x.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{E.clear(),document.querySelectorAll(".row-checkbox").forEach(h=>h.checked=!1);let p=document.getElementById("select-all-checkbox");p&&(p.checked=!1),F()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(E.size===0)return;let p=[...E],h=document.createElement("div");h.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",h.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${p.length} ${l}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${p.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(h),h.querySelector("#bulk-cancel-btn").addEventListener("click",()=>h.remove()),h.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let w=h.querySelector("#bulk-confirm-btn");w.disabled=!0,w.textContent="Menghapus...";let x=await k(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:p})});h.remove(),x.ok?(ne(`${p.length} ${l} berhasil dihapus.`),E.clear(),F(),q()):z(x.data?.error||"Gagal menghapus data.")})});let K=document.getElementById("filter-search"),ue;K?.addEventListener("input",p=>{clearTimeout(ue),ue=setTimeout(()=>{L.search=p.target.value,$=1,E.clear(),q()},400)}),r?.forEach(p=>{p.type==="select"&&document.getElementById(`filter-${p.name}`)?.addEventListener("change",h=>{L[p.name]=h.target.value,$=1,E.clear(),q()}),p.type==="combobox"&&document.getElementById(`filter-${p.name}`)?.addEventListener("change",h=>{let w=h.target.value,x=(p.options||[]).find(T=>{let _=String(typeof T=="object"?T.value:T),B=String(typeof T=="object"?T.label:T);return _===w||B===w});w?L[p.name]=x?typeof x=="object"?x.value:x:w:L[p.name]="",$=1,E.clear(),q()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{L={...n},K&&(K.value=""),r?.forEach(p=>{let h=document.getElementById(`filter-${p.name}`);h&&(h.value="")}),$=1,E.clear(),q()});let me=document.getElementById("btn-mobile-filter"),se=document.getElementById("filter-options-wrapper"),ge=document.getElementById("btn-close-filter-sheet");me&&se&&(me.addEventListener("click",p=>{p.preventDefault(),se.classList.add("sheet-open")}),ge&&ge.addEventListener("click",p=>{p.preventDefault(),se.classList.remove("sheet-open")}));let pe=document.getElementById("btn-mobile-aksi"),be=document.getElementById("excel-actions-wrapper"),_e=document.getElementById("btn-close-aksi-sheet");if(pe&&be&&(pe.addEventListener("click",p=>{p.preventDefault(),be.classList.add("sheet-open")}),_e&&_e.addEventListener("click",p=>{p.preventDefault(),be.classList.remove("sheet-open")})),document.getElementById("btn-create")?.addEventListener("click",()=>Ce(null)),S){document.getElementById(`btn-export-${S.moduleName}`)?.addEventListener("click",async h=>{let w=h.target,x=w.innerHTML;w.innerHTML="\u23F3 Loading...",w.disabled=!0;try{await S.onExport()}catch{z("Gagal export data")}finally{w.innerHTML=x,w.disabled=!1}}),document.getElementById(`btn-template-${S.moduleName}`)?.addEventListener("click",()=>{S.onTemplate()});let p=document.getElementById(`input-import-${S.moduleName}`);p?.addEventListener("change",async h=>{let w=h.target.files[0];if(!w)return;p.disabled=!0;let x=document.createElement("div");x.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",x.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(x);let T=x.querySelector("#import-progress-text"),_=x.querySelector("#import-progress-bar"),B=x.querySelector("#import-summary"),A=x.querySelector("#import-close-btn");A.addEventListener("click",()=>{x.remove(),q()});try{let W=await De(w);if(W.length===0)throw new Error("File kosong atau format salah");let V=500,ae=0,Z=0,N=0,M=W.length;T.textContent=`Ditemukan ${M} baris data. Memulai import...`;for(let I=0;I<M;I+=V){let J=W.slice(I,I+V);T.textContent=`Mengimport baris ${I+1} - ${Math.min(I+V,M)} dari ${M}...`,_.style.width=`${Math.round(I/M*100)}%`;try{let O=await S.onImport(J);O?(ae+=O.inserted||O.metrics?.inserted||J.length,Z+=O.skipped||O.metrics?.updated||0):ae+=J.length}catch(O){console.error("Chunk import failed:",O),N+=J.length}}_.style.width="100%",T.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',B.style.display="block",B.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${M}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${Z}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${N}</strong></div>
        `,N>0&&(B.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),A.style.display="block",p.value=""}catch(W){T.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${W.message}`,_.style.background="var(--danger)",_.style.width="100%",A.style.display="block",p.value=""}finally{p.disabled=!1}})}async function q(){F();let p=document.getElementById("table-container");if(!p)return;p.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let h=C==="client",w=h?1:$,x=h?Se:20,T=new URLSearchParams({page:w,limit:x,...Object.fromEntries(Object.entries(L).filter(([,N])=>N))}),_=await k(`${a}?${T}`);if(!_.ok){p.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${_.data?.error||"Error"}</p></div>`;return}let B=_.data?.data||_.data||[],A=_.data?.pagination,W=B.length,V=B;if(h){B=g(B),V=B;let N=B.length,M=20,I=Math.ceil(N/M);$>I&&I>0&&($=I);let J=($-1)*M,O=$*M;B=B.slice(J,O),A={page:$,limit:M,total:N,pages:I}}!1,y&&y(B);let ae=Pe({columns:o,data:B,fullData:V,onEdit:c?N=>Ce(N):null,actions:f.map(N=>({...N,handler:M=>N.handler(M,q)})),emptyText:`Tidak ada ${String(l||"").toLowerCase()}`,bulkSelect:D?{selectedIds:E,onToggle:F}:null});p.innerHTML="",p.appendChild(ae);let Z=document.getElementById("pagination-container");if(Z&&(Z.innerHTML="",A&&A.pages>1)){let N=Ae({page:A.page,pages:A.pages,total:A.total,limit:A.limit,onPage:M=>{$=M,q()}});N&&Z.appendChild(N)}}function Ge(p){let h=typeof d=="function"?d(p):d;return Te(h)}function Ce(p){let h=!!p,w=document.createElement("form");if(w.noValidate=!0,w.innerHTML=Ge(p),h){let T=typeof d=="function"?d(p):d;Me(w,p)}let{close:x}=de({title:h?`Edit ${l}`:`Tambah ${l}`,content:w,size:"lg",confirmText:h?"Simpan Perubahan":`Tambah ${l}`,onConfirm:async(T,_)=>{if(!w.reportValidity())return;let B=T.querySelector(".modal-confirm");B.disabled=!0,B.textContent="Menyimpan...";let A=Ne(w),W=typeof d=="function"?d(p):d,V=async M=>{for(let I of M)if(I.type==="row")await V(I.fields);else if(I.type==="combobox"&&A[I.name]){let J=A[I.name],O=(I.options||[]).find(U=>{let ee=String(typeof U=="object"?U.value:U),ze=String(typeof U=="object"?U.label:U);return ee===J||ze===J});if(O)A[I.name]=typeof O=="object"?O.value:O;else if(I.createApi){let U={};U[I.createApi.field]=J,I.createApi.extra&&Object.assign(U,I.createApi.extra);let ee=await k(I.createApi.path,{method:"POST",body:JSON.stringify(U)});if(ee.ok&&ee.data?.id)A[I.name]=ee.data.id;else if(ee.ok&&!ee.data?.id)A[I.name]=J;else throw new Error(`Gagal membuat master data: ${ee.data?.error||"Unknown error"}`)}}};try{await V(W)}catch(M){z(M.message),B.disabled=!1,B.textContent=h?"Simpan Perubahan":`Tambah ${l}`;return}m&&(A=await m(A,p));let ae=h?"PUT":"POST",Z=h?`${a}/${p.id}`:a,N=await k(Z,{method:ae,body:JSON.stringify(A)});N.ok?(ne(h?`${l} berhasil diperbarui.`:`${l} berhasil ditambahkan.`),_(),q()):(z(N.data?.error||"Gagal menyimpan data."),B.disabled=!1,B.textContent=h?"Simpan Perubahan":`Tambah ${l}`)}})}function $t(p){Ee(`Hapus ${l} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let h=await k(`${a}/${p.id}`,{method:"DELETE"});h.ok?(ne(`${l} berhasil dihapus.`),q()):z(h.data?.error||"Gagal menghapus.")},`Hapus ${l}`)}return q(),q}j();j();var Fe=null,nt=null;async function Oe(t=!1){if(Fe&&!t)return console.log("Employees Raw (Cache Hit)",Fe.slice(0,5)),Fe;let e=await k(`/api/employees?limit=${Se}&status=Aktif`);return Fe=(e.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",Fe.slice(0,5)),Fe}async function ie(t=!1){let i=(await Oe(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",i.slice(0,5)),i}async function Q(t=!1){return nt&&!t||(nt=((await k("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),nt}function X(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function bt(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function Re(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function ht(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function he(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}G();function yt(t,e){let i=new Date,a=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}`;if(!(t.target_date||t.opening_date||"").startsWith(a))return!1;let d=String(t.status||"").toLowerCase();if(d!=="selesai"&&d!=="completed"&&d!=="done")return!1;let r=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?r.includes("inspeksi"):e==="gcdc"?r.includes("general cleaning")||r.includes("deep cleaning"):!1}j();we();G();function H({container:t,title:e,icon:i,apiPath:a,columns:o,formFields:d,filterFields:r,defaultFilters:n={},enableMobileFilterSheet:s=!1,itemLabel:l="Data",canCreate:b=!0,canEdit:c=!0,canDelete:u=!0,onBeforeSubmit:m,onAfterLoad:y,onDataLoaded:g,extraActions:f=[],initialSearch:v="",exportOptions:S=null,bulkDelete:D=!1,paginationMode:C="server"}){let P=oe();P&&typeof P=="object"&&P.role==="viewer"&&(b=!1,c=!1,u=!1,D=!1,S=null);let $=1,L={...n};v&&(L.search=v);let E=new Set;t.innerHTML=`
    <div class="crud-layout-wrapper ${s?"mobile-active":""}">
      <div class="page-header">
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions">
        ${b?`<button class="btn btn-primary" id="btn-create">+ Tambah ${l}</button>`:""}
        ${S?'<button class="btn btn-outline" id="btn-mobile-aksi" style="display:none; align-items:center; justify-content:center; gap:0.25rem;">\u22EE Aksi</button>':""}
      </div>
    </div>

    ${D?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${S?`
    <div class="excel-actions-wrapper" id="excel-actions-wrapper">
      <div class="bottom-sheet-header aksi-header" style="display:none;">
        <h3 style="margin:0; font-size:1rem;">Aksi</h3>
        <button class="btn-close-sheet" id="btn-close-aksi-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
      </div>
      ${Ie(S.moduleName)}
    </div>`:""}

    ${r&&r.length>0?`
    <div class="filter-bar card ${s?"has-mobile-sheet":""}" style="padding: 1rem; display: flex; gap: 0.5rem; flex-wrap: wrap;">
        ${r.filter(p=>p.type==="search"||p.type==="search-combo").map(p=>{if(p.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${p.placeholder||"Cari..."}" id="filter-search" value="${L.search||""}"></div>`;if(p.type==="search-combo"){let h="dl-filter-search",w=(p.options||[]).map(x=>`<option value="${typeof x=="object"?x.label:x}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${h}" class="form-control" autocomplete="off" placeholder="${p.placeholder||"Cari..."}" id="filter-search" value="${L.search||""}"><datalist id="${h}">${w}</datalist></div>`}return""}).join("")}
        
        <div class="filter-options" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${r.filter(p=>p.type!=="search"&&p.type!=="search-combo").map(p=>{if(p.type==="select")return`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${p.name}" id="filter-${p.name}"><option value="">Pilih ${p.label}</option>${(p.options||[]).map(h=>`<option value="${typeof h=="object"?h.value:h}" ${L[p.name]===(typeof h=="object"?h.value:h)?"selected":""}>${typeof h=="object"?h.label:h}</option>`).join("")}</select>`;if(p.type==="combobox"){let h=`dl-filter-${p.name}`,w=(p.options||[]).map(T=>{let _=typeof T=="object"?T.label||T.value||"":T||"";return(_==="undefined"||_==="[object Object]"||_==="null")&&(_=""),_?`<option value="${_}"></option>`:""}).join(""),x=L[p.name]||"";if(L[p.name]){let T=(p.options||[]).find(_=>String(typeof _=="object"?_.value:_)==String(L[p.name]));if(T){let _=typeof T=="object"?T.label||T.value||"":T||"";_&&_!=="undefined"&&_!=="[object Object]"&&_!=="null"&&(x=_)}}return`<div class="filter-combobox" style="flex:1; min-width:120px;">
                <input type="text" name="${p.name}" id="filter-${p.name}" list="${h}" class="form-control filter-combobox-input" value="${x}" placeholder="Pilih ${p.label}..." autocomplete="off">
                <datalist id="${h}">${w}</datalist>
              </div>`}return""}).join("")}
          <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
        </div>
        ${s?'<button id="btn-mobile-filter" class="btn btn-outline" style="display:none; align-items:center; gap:0.25rem;">\u2699\uFE0F Filter</button>':""}
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
    </div>
  `;function F(){let p=document.getElementById("bulk-toolbar");if(!p)return;let h=document.getElementById("bulk-count"),w=document.getElementById("btn-bulk-delete"),x=document.getElementById("btn-bulk-cancel");h.textContent=`${E.size} item dipilih`,E.size>0?(p.classList.add("has-items"),w.disabled=!1,x.disabled=!1):(p.classList.remove("has-items"),w.disabled=!0,x.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{E.clear(),document.querySelectorAll(".row-checkbox").forEach(h=>h.checked=!1);let p=document.getElementById("select-all-checkbox");p&&(p.checked=!1),F()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(E.size===0)return;let p=[...E],h=document.createElement("div");h.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",h.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${p.length} ${l}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${p.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(h),h.querySelector("#bulk-cancel-btn").addEventListener("click",()=>h.remove()),h.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let w=h.querySelector("#bulk-confirm-btn");w.disabled=!0,w.textContent="Menghapus...";let x=await k(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:p})});h.remove(),x.ok?(ne(`${p.length} ${l} berhasil dihapus.`),E.clear(),F(),q()):z(x.data?.error||"Gagal menghapus data.")})});let K=document.getElementById("filter-search"),ue;K?.addEventListener("input",p=>{clearTimeout(ue),ue=setTimeout(()=>{L.search=p.target.value,$=1,E.clear(),q()},400)}),r?.forEach(p=>{p.type==="select"&&document.getElementById(`filter-${p.name}`)?.addEventListener("change",h=>{L[p.name]=h.target.value,$=1,E.clear(),q()}),p.type==="combobox"&&document.getElementById(`filter-${p.name}`)?.addEventListener("change",h=>{let w=h.target.value,x=(p.options||[]).find(T=>{let _=String(typeof T=="object"?T.value:T),B=String(typeof T=="object"?T.label:T);return _===w||B===w});w?L[p.name]=x?typeof x=="object"?x.value:x:w:L[p.name]="",$=1,E.clear(),q()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{L={...n},K&&(K.value=""),r?.forEach(p=>{let h=document.getElementById(`filter-${p.name}`);h&&(h.value="")}),$=1,E.clear(),q()});let me=document.getElementById("btn-mobile-filter"),se=document.getElementById("filter-options-wrapper"),ge=document.getElementById("btn-close-filter-sheet");me&&se&&(me.addEventListener("click",p=>{p.preventDefault(),se.classList.add("sheet-open")}),ge&&ge.addEventListener("click",p=>{p.preventDefault(),se.classList.remove("sheet-open")}));let pe=document.getElementById("btn-mobile-aksi"),be=document.getElementById("excel-actions-wrapper"),_e=document.getElementById("btn-close-aksi-sheet");if(pe&&be&&(pe.addEventListener("click",p=>{p.preventDefault(),be.classList.add("sheet-open")}),_e&&_e.addEventListener("click",p=>{p.preventDefault(),be.classList.remove("sheet-open")})),document.getElementById("btn-create")?.addEventListener("click",()=>Ce(null)),S){document.getElementById(`btn-export-${S.moduleName}`)?.addEventListener("click",async h=>{let w=h.target,x=w.innerHTML;w.innerHTML="\u23F3 Loading...",w.disabled=!0;try{await S.onExport()}catch{z("Gagal export data")}finally{w.innerHTML=x,w.disabled=!1}}),document.getElementById(`btn-template-${S.moduleName}`)?.addEventListener("click",()=>{S.onTemplate()});let p=document.getElementById(`input-import-${S.moduleName}`);p?.addEventListener("change",async h=>{let w=h.target.files[0];if(!w)return;p.disabled=!0;let x=document.createElement("div");x.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",x.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(x);let T=x.querySelector("#import-progress-text"),_=x.querySelector("#import-progress-bar"),B=x.querySelector("#import-summary"),A=x.querySelector("#import-close-btn");A.addEventListener("click",()=>{x.remove(),q()});try{let W=await De(w);if(W.length===0)throw new Error("File kosong atau format salah");let V=500,ae=0,Z=0,N=0,M=W.length;T.textContent=`Ditemukan ${M} baris data. Memulai import...`;for(let I=0;I<M;I+=V){let J=W.slice(I,I+V);T.textContent=`Mengimport baris ${I+1} - ${Math.min(I+V,M)} dari ${M}...`,_.style.width=`${Math.round(I/M*100)}%`;try{let O=await S.onImport(J);O?(ae+=O.inserted||O.metrics?.inserted||J.length,Z+=O.skipped||O.metrics?.updated||0):ae+=J.length}catch(O){console.error("Chunk import failed:",O),N+=J.length}}_.style.width="100%",T.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',B.style.display="block",B.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${M}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${Z}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${N}</strong></div>
        `,N>0&&(B.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),A.style.display="block",p.value=""}catch(W){T.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${W.message}`,_.style.background="var(--danger)",_.style.width="100%",A.style.display="block",p.value=""}finally{p.disabled=!1}})}async function q(){F();let p=document.getElementById("table-container");if(!p)return;p.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let h=C==="client",w=h?1:$,x=h?Se:20,T=new URLSearchParams({page:w,limit:x,...Object.fromEntries(Object.entries(L).filter(([,N])=>N))}),_=await k(`${a}?${T}`);if(!_.ok){p.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${_.data?.error||"Error"}</p></div>`;return}let B=_.data?.data||_.data||[],A=_.data?.pagination,W=B.length,V=B;if(h){B=g(B),V=B;let N=B.length,M=20,I=Math.ceil(N/M);$>I&&I>0&&($=I);let J=($-1)*M,O=$*M;B=B.slice(J,O),A={page:$,limit:M,total:N,pages:I}}!1,y&&y(B);let ae=Pe({columns:o,data:B,fullData:V,onEdit:c?N=>Ce(N):null,actions:f.map(N=>({...N,handler:M=>N.handler(M,q)})),emptyText:`Tidak ada ${String(l||"").toLowerCase()}`,bulkSelect:D?{selectedIds:E,onToggle:F}:null});p.innerHTML="",p.appendChild(ae);let Z=document.getElementById("pagination-container");if(Z&&(Z.innerHTML="",A&&A.pages>1)){let N=Ae({page:A.page,pages:A.pages,total:A.total,limit:A.limit,onPage:M=>{$=M,q()}});N&&Z.appendChild(N)}}function Ge(p){let h=typeof d=="function"?d(p):d;return Te(h)}function Ce(p){let h=!!p,w=document.createElement("form");if(w.noValidate=!0,w.innerHTML=Ge(p),h){let T=typeof d=="function"?d(p):d;Me(w,p)}let{close:x}=de({title:h?`Edit ${l}`:`Tambah ${l}`,content:w,size:"lg",confirmText:h?"Simpan Perubahan":`Tambah ${l}`,onConfirm:async(T,_)=>{if(!w.reportValidity())return;let B=T.querySelector(".modal-confirm");B.disabled=!0,B.textContent="Menyimpan...";let A=Ne(w),W=typeof d=="function"?d(p):d,V=async M=>{for(let I of M)if(I.type==="row")await V(I.fields);else if(I.type==="combobox"&&A[I.name]){let J=A[I.name],O=(I.options||[]).find(U=>{let ee=String(typeof U=="object"?U.value:U),ze=String(typeof U=="object"?U.label:U);return ee===J||ze===J});if(O)A[I.name]=typeof O=="object"?O.value:O;else if(I.createApi){let U={};U[I.createApi.field]=J,I.createApi.extra&&Object.assign(U,I.createApi.extra);let ee=await k(I.createApi.path,{method:"POST",body:JSON.stringify(U)});if(ee.ok&&ee.data?.id)A[I.name]=ee.data.id;else if(ee.ok&&!ee.data?.id)A[I.name]=J;else throw new Error(`Gagal membuat master data: ${ee.data?.error||"Unknown error"}`)}}};try{await V(W)}catch(M){z(M.message),B.disabled=!1,B.textContent=h?"Simpan Perubahan":`Tambah ${l}`;return}m&&(A=await m(A,p));let ae=h?"PUT":"POST",Z=h?`${a}/${p.id}`:a,N=await k(Z,{method:ae,body:JSON.stringify(A)});N.ok?(ne(h?`${l} berhasil diperbarui.`:`${l} berhasil ditambahkan.`),_(),q()):(z(N.data?.error||"Gagal menyimpan data."),B.disabled=!1,B.textContent=h?"Simpan Perubahan":`Tambah ${l}`)}})}function $t(p){Ee(`Hapus ${l} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let h=await k(`${a}/${p.id}`,{method:"DELETE"});h.ok?(ne(`${l} berhasil dihapus.`),q()):z(h.data?.error||"Gagal menghapus.")},`Hapus ${l}`)}return q(),q}j();G();function It(t,e){let i=String(t.status||"").toLowerCase();return e==="active"?i==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&i==="aktif":!1}j();G();function ft(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let o=new Date(a);o.setDate(a.getDate()+30);let d=new Date(t.end_date);return d.setHours(0,0,0,0),d>=a&&d<=o}return!1}j();G();function Lt(t,e){let i=String(t.status||"").toLowerCase();return e==="open"?i==="open":!1}j();function Bt(t,e){let i=String(t.status||"").toLowerCase();return e==="pending"?i==="pending":!1}var xe={};function He(t){if(xe[t]){try{xe[t].destroy()}catch{}delete xe[t]}}function Sa(){Object.keys(xe).forEach(He)}var ve=(t,e=0)=>{let i=Number(t);return isNaN(i)||t===null||t===void 0?e:i},Ke=(t,e="\u2014")=>{if(t==null||t==="")return e;let i=String(t).trim();return i===""||i==="[object Object]"?e:i};function At(t,e,i=900){if(!t)return;let a=Math.max(0,Math.round(ve(e)));if(a===0){t.textContent="0";return}let o=Date.now(),d=()=>{let r=Math.min((Date.now()-o)/i,1),n=1-Math.pow(1-r,3);t.textContent=Math.round(n*a).toLocaleString("id-ID"),r<1?requestAnimationFrame(d):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(d)}var wa={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},xa=t=>{let e=Ke(t,"\u2014");return`<span class="status-pill ${wa[e]||"pill-neutral"}">${e}</span>`};var ye={family:"Inter",size:11},ke="#94A3B8",je="#F1F5F9",vt=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],_a=()=>window.innerWidth<768;function it(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:_a()?"bottom":"top",labels:{font:ye,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ye,titleFont:{...ye,weight:"700"}}},scales:{x:{grid:{color:je},ticks:{font:ye,color:ke,maxRotation:0}},y:{grid:{color:je},ticks:{font:ye,color:ke},beginAtZero:!0}},...t}}var Ca=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),$a=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Pt(t=3){return Array(t).fill(0).map((e,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function le(t,e,i=8e3){try{let a=new AbortController,o=setTimeout(()=>a.abort(),i),d=await k(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(o),!d||!d.ok)return e;let r=d.data;return r?r.data!==void 0?r.data??e:r:e}catch{return e}}function Ea(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(a=>{let o=document.getElementById(a);o&&(o.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(a=>{let o=document.getElementById(a);if(o&&o.style.display==="none"){o.style.display="block";let d=o.parentElement;if(d&&!d.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent="Belum ada data",o.style.display="none",d.appendChild(r)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Mt({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Ft({}),["table-contracts","table-issues"].forEach(a=>{let o=document.getElementById(a);o&&o.querySelector(".skeleton")&&(o.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Nt(t){Sa(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">


      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${Ca()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${$a()}</div>

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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>kt(t)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async i=>{let a=i.target.value,o=document.getElementById("jadwal-year-label");o&&(o.textContent=a);let d=document.getElementById("skel-jadwal"),r=document.getElementById("chart-jadwal");d&&(d.style.display="block",d.style.position="absolute"),r&&(r.style.display="none");let n=await le(`/api/dashboard/schedule-chart?year=${a}`,{},8e3);try{Ot(n)}catch(s){console.warn("ScheduleChart render:",s),fe("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let a=i.target.value,o=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",d=document.getElementById("skel-insp"),r=document.getElementById("chart-insp");d&&(d.style.display="block",d.style.position="absolute"),r&&(r.style.display="none");let n=await le(o,{},8e3);try{Rt(n)}catch(s){console.warn("InspBar render:",s),fe("skel-insp","chart-insp")}}),t._skelTimeout=setTimeout(()=>Ea(),5e3),await kt(t),t._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?kt(t):clearInterval(t._dashRefresh)},6e4)}async function kt(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,i,a,o,d,r,n,s,l,b,c,u,m,y]=await Promise.all([le("/api/dashboard/kpi",{},8e3),le("/api/dashboard/issues-trend",{},8e3),le("/api/dashboard/issues-summary",{},8e3),le("/api/dashboard/stats",{},8e3),le("/api/dashboard/calendar",[],8e3),le("/api/schedule?limit=10000",{data:[]},8e3),le("/api/employees?limit=10000",{data:[]},8e3),le("/api/contracts?limit=10000",{data:[]},8e3),le("/api/issues?limit=10000",{data:[]},8e3),le("/api/one-on-one?limit=10000",{data:[]},8e3),le("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),le(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),le("/api/relievers?limit=10000",{data:[]},8e3),le("/api/reports/fogging?limit=10000",{data:[]},8e3)]),g=document.getElementById("filter-insp-month"),f=g?g.value:"",v=f?`/api/dashboard/inspection-bar?month=${f}`:"/api/dashboard/inspection-bar",S=await le(v,{},8e3);if(e){let D=Array.isArray(r?.data)?r.data:Array.isArray(r)?r:[];window.dashboardSchedules=D;let C=Array.isArray(n?.data)?n.data:Array.isArray(n)?n:[],P=Array.isArray(s?.data)?s.data:Array.isArray(s)?s:[],$=Array.isArray(l?.data)?l.data:Array.isArray(l)?l:[],L=Array.isArray(b?.data)?b.data:Array.isArray(b)?b:[],E=Array.isArray(m?.data)?m.data:Array.isArray(m)?m:[];window.dashboardRelievers=E;let F=Array.isArray(y?.data)?y.data:Array.isArray(y)?y:[];window.dashboardFogging=F,e.employees&&(e.employees.current=C.filter(K=>It(K,"active")).length),e.contracts&&(e.contracts.current=P.filter(K=>ft(K,"active")).length),e.expiring30&&(e.expiring30={current:P.filter(K=>ft(K,"expiring30")).length}),e.issues&&(e.issues.current=$.filter(K=>Lt(K,"open")).length),e.one_on_one&&(e.one_on_one.current=L.filter(K=>Bt(K,"pending")).length),e.inspection_month&&(e.inspection_month.current=D.filter(K=>yt(K,"inspeksi")).length),e.cleaning_month&&(e.cleaning_month.current=D.filter(K=>yt(K,"gcdc")).length)}try{Mt(e)}catch(D){console.warn("KPI render:",D)}try{Ft(e)}catch(D){console.warn("MiniStats render:",D)}try{Ot(u)}catch(D){console.warn("ScheduleChart render:",D),fe("skel-jadwal","chart-jadwal")}try{Ta(Array.isArray(a?.by_category)?a.by_category:[])}catch(D){console.warn("Donut render:",D),fe("skel-donut","chart-donut")}try{Da(i)}catch(D){console.warn("Trend render:",D),fe("skel-trend","chart-trend")}try{Rt(S)}catch(D){console.warn("InspBar render:",D),fe("skel-insp","chart-insp")}try{let D=Array.isArray(o)?o:Array.isArray(o?.recent_issues)?o.recent_issues:[];La(D)}catch(D){console.warn("IssuesTable render:",D)}try{let D=Array.isArray(o?.expiring_contracts)?o.expiring_contracts:[];Ia(c)}catch(D){console.warn("ContractsTable render:",D)}try{Ba(Array.isArray(d)?d:[])}catch(D){console.warn("Agenda render:",D)}try{Pa()}catch(D){console.warn("Quick Actions render:",D)}}function Mt(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let i=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=i.map(a=>{let o=ve(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${o}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${o}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{At(a,parseInt(a.dataset.target)||0)})}function Ft(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let i=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,a=new Date().getFullYear(),o=String(new Date().getMonth()+1).padStart(2,"0"),d=`${a}-${o}`,r=c=>`
    <select id="${c}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.preventDefault(); event.stopPropagation();">
      <option value="${a}-01" ${d===`${a}-01`?"selected":""}>Jan</option>
      <option value="${a}-02" ${d===`${a}-02`?"selected":""}>Feb</option>
      <option value="${a}-03" ${d===`${a}-03`?"selected":""}>Mar</option>
      <option value="${a}-04" ${d===`${a}-04`?"selected":""}>Apr</option>
      <option value="${a}-05" ${d===`${a}-05`?"selected":""}>Mei</option>
      <option value="${a}-06" ${d===`${a}-06`?"selected":""}>Jun</option>
      <option value="${a}-07" ${d===`${a}-07`?"selected":""}>Jul</option>
      <option value="${a}-08" ${d===`${a}-08`?"selected":""}>Agu</option>
      <option value="${a}-09" ${d===`${a}-09`?"selected":""}>Sep</option>
      <option value="${a}-10" ${d===`${a}-10`?"selected":""}>Okt</option>
      <option value="${a}-11" ${d===`${a}-11`?"selected":""}>Nov</option>
      <option value="${a}-12" ${d===`${a}-12`?"selected":""}>Des</option>
    </select>
  `,n=[{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.preventDefault(); event.stopPropagation();">
          <option value="Q1" ${i==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${i==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${i==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${i==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:t.schedule?.current,href:`#/timeline?dash_filter=period_${i.toLowerCase()}`,color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:r("dash-reliefer-month"),val:t.reliever_completed?.current,href:`#/relievers?dash_filter=reliever&month=${d}`,color:"mini-teal"},{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:r("dash-inspeksi-month"),val:t.inspection_month?.current,href:`#/timeline?dash_filter=inspeksi&month=${d}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:r("dash-gcdc-month"),val:t.cleaning_month?.current,href:`#/timeline?dash_filter=gcdc&month=${d}`,color:"mini-green"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:r("dash-fogging-month"),val:t.fogging_month?.current,href:`#/reports/fogging?dash_filter=fogging&month=${d}`,color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=n.map(c=>`
    <a href="${c.href}" class="mini-stat ${c.color}" style="text-decoration:none" id="${c.id||""}">
      <div class="mini-stat-icon">${c.icon}</div>
      <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
        <div style="display:flex; align-items:baseline; gap:3px;">
          <div class="mini-stat-value" data-target="${ve(c.val)}">0</div>
          ${c.dropdown?c.dropdown:""}
        </div>
        <div class="mini-stat-text">${c.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(c=>At(c,parseInt(c.dataset.target)||0,700));let s=document.getElementById("dash-jadwal-period");s&&s.addEventListener("change",c=>{let u=c.target.value,m=(window.dashboardSchedules||[]).filter(f=>f.period===u).length,y=document.querySelector("#mini-jadwal .mini-stat-value");y&&(y.dataset.target=m,y.textContent=m);let g=document.getElementById("mini-jadwal");g&&(g.href=`#/timeline?dash_filter=period_${u.toLowerCase()}`)});let l=(c,u,m,y,g)=>{let f=document.getElementById(c);f&&f.addEventListener("change",v=>{let S=v.target.value,D=(m||[]).filter($=>y($,S)).length,C=document.querySelector(`#${u} .mini-stat-value`);C&&(C.dataset.target=D,C.textContent=D);let P=document.getElementById(u);P&&(P.href=`${g}&month=${S}`)})},b=c=>{let u=String(c.status||"").toLowerCase();return u==="done"||u==="selesai"||u==="completed"};l("dash-reliefer-month","mini-reliefer",window.dashboardRelievers,(c,u)=>window.parseFlexibleDate(c.backup_date).startsWith(u)&&b(c),"#/relievers?dash_filter=reliever"),l("dash-inspeksi-month","mini-inspeksi",window.dashboardSchedules,(c,u)=>c.activity_type==="Inspeksi Hygiene"&&b(c)&&window.parseFlexibleDate(c.opening_date||c.target_date).startsWith(u),"#/timeline?dash_filter=inspeksi"),l("dash-gcdc-month","mini-gcdc",window.dashboardSchedules,(c,u)=>(c.activity_type==="General Cleaning"||c.activity_type==="Deep Cleaning")&&b(c)&&window.parseFlexibleDate(c.opening_date||c.target_date).startsWith(u),"#/timeline?dash_filter=gcdc"),l("dash-fogging-month","mini-fogging",window.dashboardFogging,(c,u)=>b(c)&&window.parseFlexibleDate(c.activity_date).startsWith(u),"#/reports/fogging?dash_filter=fogging")}function Ta(t){fe("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!e||!i)return;He("donut");let a=(t||[]).filter(s=>ve(s.count)>0);if(!a.length){Xe(e,"Belum ada data permasalahan");return}let o=a.map(s=>`${Ke(s.category,"Lainnya")}`),d=a.map(s=>ve(s.count)),r=d.reduce((s,l)=>s+l,0);i.innerHTML=a.map((s,l)=>{let b=vt[l%vt.length],c=r>0?Math.round(s.count/r*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${b}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${s.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${c}%)</span></div>
          <div class="donut-legend-label">${o[l]}</div>
        </div>
      </div>
    `}).join("");let n={id:"centerText",beforeDraw:function(s){let l=s.width,b=s.height,c=s.ctx;c.restore();let u=(b/80).toFixed(2);c.font="bold "+u+"em Inter",c.textBaseline="middle",c.fillStyle="#1E293B";let m=r.toString(),y=Math.round((l-c.measureText(m).width)/2),g=b/2;c.fillText(m,y,g-4),c.font="600 "+(u*.35).toFixed(2)+"em Inter",c.fillStyle="#64748B";let f="Total",v=Math.round((l-c.measureText(f).width)/2);c.fillText(f,v,g+10),c.save()}};xe.donut=new Chart(e,{type:"doughnut",data:{labels:o,datasets:[{data:d,backgroundColor:vt,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ye,titleFont:{...ye,weight:"700"},callbacks:{label:s=>` ${s.label}: ${s.parsed} kasus`}}},cutout:"75%"},plugins:[n]})}function Da(t){fe("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;He("trend"),t=t||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(t.labels||[]).map(r=>{if(!r||typeof r!="string")return"";try{let[n,s]=r.split("-");return(i[Number(s)-1]||s)+" "+String(n).slice(-2)}catch{return r}}),o=(t.open||[]).map(r=>ve(r)),d=(t.closed||[]).map(r=>ve(r));if(!a.length){Xe(e,"Belum ada data trend");return}xe.trend=new Chart(e,{type:"line",data:{labels:a,datasets:[{label:"Open",data:o,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:d,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:it({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:ke,maxRotation:0,autoSkip:!1}},y:{grid:{color:je},ticks:{font:{family:"Inter",size:9},color:ke},beginAtZero:!0}}})})}function Ot(t){fe("skel-jadwal","chart-jadwal");let e=document.getElementById("chart-jadwal");if(!e)return;He("jadwal"),t=t||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(t).some(s=>Array.isArray(s)&&s.some(l=>l>0))){Xe(e,"Belum ada data jadwal");return}let o=t["Inspeksi Hygiene"]||Array(12).fill(0),d=t["General Cleaning"]||Array(12).fill(0),r=t["Deep Cleaning"]||Array(12).fill(0),n=t.Fogging||Array(12).fill(0);xe.jadwal=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Inspeksi",data:o,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:d,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:r,backgroundColor:"#F59E0B"},{label:"Fogging",data:n,backgroundColor:"#EF4444"}]},options:it({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:ke,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:je},ticks:{font:{family:"Inter",size:9},color:ke},min:0}}})})}function Rt(t){fe("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;He("inspBar"),t=t||{};let i=t.labels||[],a=(t.fc||[]).map(d=>ve(d)),o=(t.spv||[]).map(d=>ve(d));if(!i.length){Xe(e,"Belum ada data inspeksi");return}xe.inspBar=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:o,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:it({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ye,color:ke,maxRotation:45,minRotation:30}},y:{grid:{color:je},ticks:{font:ye,color:ke},min:0,max:100}}})})}function Ia(t){fe("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;He("contractMiniBar"),t=t||{};let i={"06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(r=>{let n=r.split("-")[1];return i[n]||r}),o=(t.data||[]).map(r=>ve(r));if(!a.length){Xe(e,"Belum ada data");return}let d=e.getContext("2d");xe.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:o,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:it({onClick:(r,n)=>{if(n&&n.length>0){let s=n[0].index,l=(t.labels||[])[s];l&&(window.location.hash="#/contracts?month_expiry="+l)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ye,color:ke,maxRotation:0,autoSkip:!1}},y:{grid:{color:je,borderDash:[4,4],drawBorder:!1},ticks:{font:ye,color:ke,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function La(t){let e=document.getElementById("table-issues");if(!e)return;let i=(t||[]).slice(0,8);if(!i.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${i.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${xa(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Ke(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Ke(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function Ba(t){let e=document.getElementById("widget-agenda");if(!e)return;let i=new Date,a=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,d=(t||[]).filter(r=>(r.event_date||"").startsWith(a)).slice(0,10);if(!d.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${d.map(r=>{let n="#3B82F6",s="#EFF6FF",l="Agenda",b=(r.title||"").toLowerCase();return b.includes("inspeksi")?(n="#10B981",s="#ECFDF5",l="Inspeksi"):b.includes("cleaning")||b.includes("gcdc")?(n="#3B82F6",s="#EFF6FF",l="Cleaning"):b.includes("reliefer")?(n="#F59E0B",s="#FFFBEB",l="Reliefer"):b.includes("fogging")&&(n="#8B5CF6",s="#F5F3FF",l="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(r.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${n};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Ke(r.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Ke(r.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${s};color:${n}">${l}</div>
        </div>
      `}).join("")}
    </div>
  `}function Pa(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function fe(t,e){let i=document.getElementById(t),a=document.getElementById(e);if(i&&(i.style.display="none",i.style.position=""),a){a.style.display="block";let o=a.parentElement;if(o){let d=o.querySelector(".chart-empty");d&&d.remove()}}}function Xe(t,e="Belum ada data"){if(!t)return;t.style.display="none";let i=t.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent=e,i.appendChild(o)}}j();async function Kt(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),i=document.getElementById("login-error"),a=document.getElementById("login-btn"),o=document.getElementById("toggle-password"),d=document.getElementById("login-password");o?.addEventListener("click",()=>{let r=d.type==="text";d.type=r?"password":"text",o.style.color=r?"":"var(--primary)"}),e?.addEventListener("submit",async r=>{r.preventDefault(),i.style.display="none";let n=e.username.value.trim(),s=e.password.value;if(!n||!s){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let l=await k("/api/auth/login",{method:"POST",body:JSON.stringify({username:n,password:s})});l.ok&&l.data.success?(pt(l.data.data.token),Ve(l.data.data.user),ne("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=l.data.error||"Username atau password salah.",i.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}j();G();async function Aa(){return await Q()}function Na(t,e){let i=String(t.status||"").toLowerCase();return e==="active"?i==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&i==="aktif":!1}async function jt(t,e){let i=await Aa(),a=e?e.get("dash_filter"):null;H({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:o=>a?o.filter(d=>Na(d,a)):o,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:o=>Re(o)},{key:"phone",label:"No. HP",render:o=>o?`<a href="tel:${o}">${o}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>X(o)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:o=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:o?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:o?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:i,value:o?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:o?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:o?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:o?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let o=await k(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(o.ok){let d=o.data.data.map(r=>({"Nama Lengkap":r.full_name,Cabang:r.branch_name||"",Divisi:r.division||"","No. HP":r.phone||"","Tgl Masuk":r.join_date||"",Status:r.status||""}));R(d,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async o=>{let d=s=>{if(!s)return null;let l=String(s||"").toLowerCase(),b=i.find(c=>String(c.label||"").toLowerCase()===l);return b?b.value:null},r=o.map(s=>({full_name:String(s["Nama Lengkap"]||"").trim(),branch_id:d(String(s.Cabang||"").trim()),division:String(s.Divisi||"").trim()||"FACILITY CARE",phone:String(s["No. HP"]||"").trim(),join_date:String(s["Tgl Masuk"]||"").trim(),status:String(s.Status||"").trim(),notes:String(s.Catatan||"").trim()})).filter(s=>s.full_name),n=await k("/api/import/employees",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}j();G();var wt=[],Ht=[];async function Ma(){wt=await Q(),Ht=await Oe()}var St=async t=>{let e=[],i=1;for(;;){let o=await(await Promise.resolve().then(()=>(j(),Le))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${i}`);if(!o.ok)break;let d=o.data?.data||o.data||[],r=Array.isArray(d)?d:[];if(e=e.concat(r),r.length<100||o.data?.pagination&&i>=o.data.pagination.pages)break;i++}return e};function Fa(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let o=new Date(a);o.setDate(a.getDate()+30);let d=new Date(t.end_date);return d.setHours(0,0,0,0),d>=a&&d<=o}return!1}async function qt(t,e){await Ma();let i=e?e.get("dash_filter"):null;H({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",onDataLoaded:a=>i?a.filter(o=>Fa(o,i)):a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>Re(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,o)=>o.end_date&&String(o.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':bt(a)},{key:"status",label:"Status",render:a=>X(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:wt},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"month_expiry",label:"Bulan Habis",options:[{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"},{value:"2027-01",label:"Jan 2027"},{value:"2027-02",label:"Feb 2027"},{value:"2027-03",label:"Mar 2027"},{value:"2027-04",label:"Apr 2027"},{value:"2027-05",label:"Mei 2027"},{value:"2027-06",label:"Jun 2027"},{value:"2027-07",label:"Jul 2027"}]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Detail Kontrak",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[d,r]=await Promise.all([St("/api/employees?status=Aktif"),St("/api/contracts")]);if(d.length>0){let n=r.filter(m=>m.status==="Aktif"&&(m.days_remaining==null||m.days_remaining>=0||String(m.end_date).startsWith("2099"))),s=new Set(n.map(m=>m.employee_id)),l=d.filter(m=>!s.has(m.id)),b=[],c=[];l.forEach(m=>{let y=r.filter(g=>g.employee_id===m.id);y.length===0?b.push(m):c.push({emp:m,lastContract:y[0]})});let u=`<p style="margin-bottom:12px">Data yang terbaca: <b>${d.length}</b> Karyawan Aktif, dan <b>${n.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:6px">Terdapat <b>${b.length}</b> karyawan aktif tanpa kontrak.</p>
              <p style="margin-bottom:12px">Terdapat <b>${c.length}</b> karyawan aktif yang memiliki masa kontrak Expired.</p>
              <ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;b.forEach(m=>{u+=`<li style="margin-bottom:8px"><b>${m.full_name}</b> <br><span style="font-size:0.85em;color:#F59E0B">Cabang: ${m.branch_name||"-"} | Belum pernah di-input kontrak</span></li>`}),c.forEach(m=>{let y=m.emp,g=m.lastContract,f=g.status==="Aktif"&&g.days_remaining<0?"Aktif (Masa Habis)":g.status;u+=`<li style="margin-bottom:8px"><b>${y.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${y.branch_name||"-"} | Status Terakhir: <b style="color:#EF4444">${f}</b>, Tgl Berakhir: ${window.formatDate(g.end_date)}</span></li>`}),u+="</ul>",Promise.resolve().then(()=>(we(),mt)).then(m=>m.createModal({title:"Detail Karyawan Tanpa Kontrak Aktif",content:u,cancelText:"Tutup"}))}}catch(d){console.error(d)}a.innerHTML="\u{1F50D} Detail Kontrak",a.disabled=!1};let o=document.querySelector(".page-actions");o&&o.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"select",required:!0,options:Ht,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"select",options:wt,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await k(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let o=a.data.data.map(d=>({"Nama Lengkap":d.employee_name,Cabang:d.branch_name||"","Div / Bagian":d.division||"","Tanggal Mulai":d.start_date||"","Tanggal Selesai":d.end_date&&String(d.end_date).startsWith("2099")?"":d.end_date||"","Sisa Kontrak":d.end_date&&String(d.end_date).startsWith("2099")?"Tetap":d.days_remaining!==null&&d.days_remaining!==void 0?`${d.days_remaining} Hari`:"",Status:d.status||""}));R(o,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[o,d]=await Promise.all([k("/api/branches?limit=10000"),St("/api/employees")]),r=o.data?.data||[],n=d||[];console.log(`Total employee yang berhasil dimuat dari database : ${n.length}`),n.length>0&&(console.log("Contoh 5 employee pertama:"),n.slice(0,5).forEach((g,f)=>{console.log(`${f+1}. ID: ${g.id}, Name: ${g.full_name}, Status: ${g.status}`)}));let s=g=>{if(!g)return null;let f=String(g||"").replace(/\s+/g," ").toLowerCase().trim(),v=r.find(S=>String(S.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===f||String(S.code||"").replace(/\s+/g," ").toLowerCase().trim()===f||String(S.name||"").replace(/\s+/g," ").toLowerCase().trim()===f);return v?v.id:null},l=(g,f)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${f}`),console.log(`Nama dari Excel : "${g}"`),!g)return console.log("Alasan gagal mapping : Nama kosong"),null;let v=String(g||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${v}"`),console.log(`Jumlah employee di database : ${n.length}`);let S=n.find(D=>String(D.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===v);return S?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${S.id}`),S.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},b=g=>{if(!g)return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let f=String(g).trim();if(/^\d{4,5}(\.\d+)?$/.test(f)){let S=Math.floor(Number(f));if(S>2e4&&S<99999){let D=new Date(Date.UTC(1899,11,30)+S*864e5);return isNaN(D.getTime())?"":D.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(f))return f.slice(0,10);let v=f.split(/[\/\-\.]/);if(v.length===3){let[S,D,C]=v.map(P=>P.trim());if(S.length===4&&D.length<=2&&C.length<=2)return`${S}-${D.padStart(2,"0")}-${C.padStart(2,"0")}`;if(C.length===4&&D.length<=2&&S.length<=2)return`${C}-${D.padStart(2,"0")}-${S.padStart(2,"0")}`}return f},c=a.map((g,f)=>{let v=f+2,S=String(g["Nama Lengkap"]||"").trim(),D=g["Tanggal Mulai"],C=b(D);if(!C){let L=a.__worksheet,E=a.__headers||[],F=E.indexOf("Tanggal Mulai"),K="N/A",ue="N/A",me="N/A";if(F!==-1&&L&&window.XLSX){let ge=window.XLSX.utils.encode_cell({c:F,r:v-1});me=ge;let pe=L[ge];pe?(K=pe.t||"undefined",ue=pe.w||"undefined"):K="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let se="Unknown";D==null||D===""?se="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":D instanceof Date&&isNaN(D.getTime())?se="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":se="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${v}`),console.log(`Employee Name : ${S}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${F})`),console.log(`Raw Cell Value : "${D}"`),console.log(`JavaScript Type : ${typeof D}`),console.log(`SheetJS Cell Type : ${K}`),console.log(`SheetJS Formatted Value : "${ue}"`),console.log(`Value After Trim : "${String(D||"").trim()}"`),console.log(`Value After Date Parser : "${C}"`),console.log(`Is Empty : ${!D}`),console.log(`Is Invalid Date : ${D instanceof Date?isNaN(D.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${se}`),console.log(`Workbook Sheet : ${L?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${me}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(g,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(E)),console.log(`==========================
`)}let P=l(S,v),$=null;return P||($="Karyawan tidak ditemukan di Database"),{isValid:!!P,invalidReason:$,rowNum:v,data:{employee_id:P,branch_id:s(String(g.Cabang||"").trim()),division:String(g["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:C,end_date:b(g["Tanggal Selesai"])||"2099-12-31",status:String(g.Status||"").trim(),_rawName:S}}}),u=[],m=[];if(c.forEach(g=>{g.isValid?u.push(g.data):m.push({rowNum:g.rowNum,name:g.data._rawName,reason:g.invalidReason})}),console.log(`Split Validation - Valid: ${u.length}, Invalid: ${m.length}`),u.length===0)return{inserted:0,skipped:a.length,failed:a.length};let y=await k("/api/import/contracts",{method:"POST",body:JSON.stringify({rows:u,onDuplicate:"update"})});if(!y.ok)throw new Error(y.data?.error||"Import gagal");return y.data}}})}j();G();var xt=[],Ze=[];function Oa(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let i of e)if(t.some(a=>a.period===i))return i;return"Q3"}async function Jt(t,e){xt=await Q();let i=await ie();Ze=["Berlin Ariansyah","Ade Surahman"];let a=m=>m&&!Ze.find(y=>String(typeof y=="object"?y.value:y).toLowerCase()===String(m).toLowerCase())?[...Ze,m]:Ze,o=await k(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),d=m=>{if(!m||m==="-"||String(m).trim()==="")return"";let y=String(m).split("-");return y.length===3&&y[0].length===4?`${y[2]}-${y[1]}-${y[0]}`:m},r=o.data?.data||[],n=Oa(r),s=e?e.get("dash_filter"):null,l=new Date,b=`${l.getFullYear()}-${String(l.getMonth()+1).padStart(2,"0")}`,c={},u=e&&e.get("month")?e.get("month"):null;s==="inspeksi"?c={status:"Done",activity_type:"Inspeksi Hygiene",month:u}:s==="gcdc"?c={status:"Done",activity_type:"GCDC",month:u}:s&&s.startsWith("period_")&&(c={period:s.replace("period_","").toUpperCase()}),gt({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:c,onDataLoaded:m=>m.sort((y,g)=>{let f=y.opening_date?new Date(y.opening_date).getTime():0;return(g.opening_date?new Date(g.opening_date).getTime():0)-f}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:m=>ht(m)},{key:"period",label:"Periode",render:m=>he(m)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:m=>d(m)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:m=>d(m)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:m=>d(m)},{key:"status",label:"Status",render:m=>X(m)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:xt},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Ze}],formFields:m=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:xt,value:m?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:m?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:m?.period},{name:"pic",label:"PIC",type:"select",options:a(m?.pic),value:m?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:m?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:m?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:m?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:m?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:m?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let m=await k(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(m.ok){let y=m.data.data.map(g=>({Cabang:g.branch_name||"",Kegiatan:g.activity_type||"",Periode:g.period||"",PIC:g.pic||"","Tgl Opening":g.opening_date||"","Tgl Target":g.target_date||"","Tgl Selesai":g.completion_date||"",Status:g.status||""}));R(y,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async m=>{let g=(await k("/api/branches?all=1")).data?.data||[],f=C=>{if(!C)return null;let P=String(C||"").toLowerCase(),$=g.find(L=>String(L.full_name||"").toLowerCase()===P||String(L.code||"").toLowerCase()===P||String(L.name||"").toLowerCase()===P);return $?$.id:null},v=C=>{if(C==null||C==="")return"";if(C instanceof Date&&!isNaN(C.getTime()))return C.toISOString().slice(0,10);let P=String(C).trim();if(P===""||P==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(P))return P.slice(0,10);if(/^\d{4,5}$/.test(P)){let L=Number(P);if(L>2e4&&L<99999){let E=new Date(Date.UTC(1899,11,30)+L*864e5);return isNaN(E.getTime())?"":E.toISOString().slice(0,10)}}let $=P.split(/[\/\-\.]/);if($.length===3){let[L,E,F]=$.map(K=>K.trim());if(L.length===4&&E.length<=2&&F.length<=2)return`${L}-${E.padStart(2,"0")}-${F.padStart(2,"0")}`;if(F.length===4&&E.length<=2&&L.length<=2)return`${F}-${E.padStart(2,"0")}-${L.padStart(2,"0")}`}return P},S=m.map(C=>({branch_id:f(String(C.Cabang||"").trim()),activity_type:String(C.Kegiatan||"").trim(),period:String(C.Periode||"").trim(),pic:String(C.PIC||C.Pic||"").trim(),opening_date:v(C["Tgl Opening"]||C["Tanggal Opening"]||C["Tgl Openir"]),target_date:v(C["Tgl Target"]||C["Tanggal Target"]),completion_date:v(C["Tgl Selesai"]||C["Tanggal Selesai"]),status:String(C.Status||"").trim(),notes:String(C.Catatan||C.Keterangan||"").trim()})).filter(C=>C.activity_type&&C.period),D=await k("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:S,onDuplicate:"update"})});if(!D.ok)throw new Error(D.data?.error||"Import gagal");return D.data}}})}j();G();var _t=[],rt=[];function Ra(t,e){let i=String(t.status||"").toLowerCase();return e==="open"?i==="open":!1}async function Ut(t,e){let i=e?e.get("dash_filter"):null;_t=await Q(),rt=await ie();let a=r=>r&&!rt.find(n=>n.value===r)?[...rt,{value:r,label:r}]:rt,o=new Date().getFullYear();H({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:r=>i?r.filter(n=>Ra(n,i)):r,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:r=>`<span class="badge badge-secondary">${r}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:r=>`<span title="${r}">${r?.length>50?r.slice(0,50)+"\u2026":r}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:r=>`<span title="${r||""}">${r?.length>40?r.slice(0,40)+"\u2026":r||"-"}</span>`},{key:"status",label:"Status",render:r=>X(r)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:r=>window.formatDate(r)},{key:"day_count",label:"Hari",render:r=>r??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:_t},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:["2025","2026","2027","2028","2029","2030"]}],formFields:r=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:r?.report_date},{name:"branch_id",label:"Cabang",type:"select",required:!0,options:_t,value:r?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:r?.category},{name:"source",label:"Sumber Laporan",type:"select",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:r?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:r?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"select",options:a(r?.employee_name),value:r?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"select",options:a(r?.fc_specialist),value:r?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:r?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:r?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:r?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let r=await k(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let n=r.data.data.map(s=>({Tanggal:s.report_date||"",Cabang:s.branch_name||"",Kategori:s.category||"",Sumber:s.source||"",Keluhan:s.complaint||"","Nama FC":s.employee_name||"","FC Spesialis":s.fc_specialist||"",Solusi:s.solution||"","Tgl Selesai":s.completion_date||"",Status:s.status||""}));R(n,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async r=>{let s=(await k("/api/branches?all=1")).data?.data||[],l=u=>{if(!u)return null;let m=String(u||"").toLowerCase(),y=s.find(g=>String(g.full_name||"").toLowerCase()===m||String(g.code||"").toLowerCase()===m||String(g.name||"").toLowerCase()===m);return y?y.id:null},b=r.map(u=>({branch_id:l(String(u.Cabang||"").trim()),report_date:String(u.Tanggal||"").trim(),category:String(u.Kategori||"").trim(),source:String(u.Sumber||"").trim(),complaint:String(u.Keluhan||"").trim(),employee_name:String(u["Nama FC"]||"").trim(),fc_specialist:String(u["FC Spesialis"]||"").trim(),solution:String(u.Solusi||"").trim(),completion_date:String(u["Tgl Selesai"]||"").trim(),status:String(u.Status||"").trim()})).filter(u=>u.report_date&&u.complaint&&u.category),c=await k("/api/import/issues",{method:"POST",body:JSON.stringify({rows:b,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}}})}j();var qe=[];function Ka(t,e){let i=String(t.status||"").toLowerCase();return e==="pending"?i==="pending":!1}async function Gt(t,e){let i=e?e.get("dash_filter"):null;qe=await Q();let a=await ie(),o=["Ade","Berlin"],d=n=>n&&!a.find(s=>s.value===n)?[...a,{value:n,label:n}]:a,r=n=>n&&!o.find(s=>(typeof s=="object"?s.value:s)===n)?[...o,n]:o;H({container:t,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:n=>i?n.filter(s=>Ka(s,i)):n,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:n=>`<span title="${n||""}">${n?.length>50?n.slice(0,50)+"\u2026":n||"-"}</span>`},{key:"solution",label:"Solusi",render:n=>`<span title="${n||""}">${n?.length>40?n.slice(0,40)+"\u2026":n||"-"}</span>`},{key:"status",label:"Status",render:n=>X(n)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:qe},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async n=>{let s=new URLSearchParams(n||{}).toString(),l=await k(`/api/one-on-one?limit=10000&${s}`);if(l.ok){let b=l.data.data.map(u=>({Tanggal:u.meeting_date||"",Cabang:u.branch_name||"","Nama Karyawan":u.employee_name||"",PIC:u.pic||"",Masalah:u.problem||"",Solusi:u.solution||"",Status:u.status||"","Tgl Selesai":u.completion_date||"",Dokumen:u.document_link||""})),{downloadExcel:c}=await Promise.resolve().then(()=>(G(),ce));c(b,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:s}=await Promise.resolve().then(()=>(G(),ce));s(n,"Template_Import_OneOnOne")},onImport:async n=>{let s=u=>{if(!u)return null;let m=String(u||"").toLowerCase(),y=qe.find(g=>String(g.label||"").toLowerCase()===m);return y?y.value:null},l=u=>{if(!u)return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let m=String(u).trim();if(/^\d{4,5}$/.test(m)){let g=Number(m);if(g>2e4&&g<99999){let f=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);let y=m.split(/[\/\-\.]/);if(y.length===3){let[g,f,v]=y.map(S=>S.trim());if(g.length===4&&f.length<=2&&v.length<=2)return`${g}-${f.padStart(2,"0")}-${v.padStart(2,"0")}`;if(v.length===4&&f.length<=2&&g.length<=2)return`${v}-${f.padStart(2,"0")}-${g.padStart(2,"0")}`}return m},b=n.map(u=>({meeting_date:l(u.Tanggal),employee_name:String(u["Nama Karyawan"]||"").trim(),branch_id:s(String(u.Cabang||"").trim()),pic:String(u.PIC||"").trim(),problem:String(u.Masalah||"").trim(),solution:String(u.Solusi||"").trim(),status:String(u.Status||"").trim(),completion_date:l(u["Tgl Selesai"]),document_link:String(u.Dokumen||"").trim()})).filter(u=>u.meeting_date&&u.employee_name&&u.branch_id),c=await k("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:b,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},formFields:n=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:n?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:n?.branch_id&&!qe.find(s=>s.value==n.branch_id)?[...qe,{value:n.branch_id,label:n.branch_name||n.branch_id}]:qe,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:d(n?.employee_name),value:n?.employee_name},{name:"pic",label:"PIC",type:"select",options:r(n?.pic),createApi:{path:"/api/pic",field:"name"},value:n?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:n?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:n?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:n?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link}]})}j();async function zt(t){let e=await Q(),i=await ie(),a=["Ade","Berlin"],o=n=>n&&!i.find(s=>s.value===n)?[...i,{value:n,label:n}]:i,d=n=>n&&!a.find(s=>(typeof s=="object"?s.value:s)===n)?[...a,n]:a,r=Array.from({length:5},(n,s)=>String(new Date().getFullYear()-s));H({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:n=>{try{let s=JSON.parse(n);return Array.isArray(s)?s.join(", "):n||"-"}catch{return n||"-"}}},{key:"score",label:"Nilai",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:r}],exportOptions:{moduleName:"training",onExport:async n=>{let s=new URLSearchParams(n||{}).toString(),l=await k(`/api/training?limit=10000&${s}`);if(l.ok){let b=l.data.data.map(u=>{let m=u.participants||"";try{let y=JSON.parse(m);m=Array.isArray(y)?y.join(", "):m}catch{}return{Tanggal:u.training_date||"",Batch:u.batch||"",Materi:u.subject||"",Cabang:u.branch_name||"",Trainer:u.trainer||"",Peserta:m,Nilai:u.score!==null&&u.score!==void 0?u.score:"",Dokumen:u.document_link||""}}),{downloadExcel:c}=await Promise.resolve().then(()=>(G(),ce));c(b,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:s}=await Promise.resolve().then(()=>(G(),ce));s(n,"Template_Import_Training")},onImport:async n=>{let s=u=>{if(!u)return null;let m=String(u||"").toLowerCase(),y=e.find(g=>String(g.label||"").toLowerCase()===m);return y?y.value:null},l=u=>{if(!u)return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let m=String(u).trim();if(/^\d{4,5}$/.test(m)){let g=Number(m);if(g>2e4&&g<99999){let f=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);let y=m.split(/[\/\-\.]/);if(y.length===3){let[g,f,v]=y.map(S=>S.trim());if(g.length===4&&f.length<=2&&v.length<=2)return`${g}-${f.padStart(2,"0")}-${v.padStart(2,"0")}`;if(v.length===4&&f.length<=2&&g.length<=2)return`${v}-${f.padStart(2,"0")}-${g.padStart(2,"0")}`}return m},b=n.map(u=>({training_date:l(u.Tanggal),batch:String(u.Batch||"").trim(),subject:String(u.Materi||"").trim(),branch_id:s(String(u.Cabang||"").trim()),trainer:String(u.Trainer||"").trim(),participants:String(u.Peserta||"").trim(),score:u.Nilai?Number(u.Nilai):null,document_link:String(u.Dokumen||"").trim()})).filter(u=>u.training_date&&u.subject&&u.branch_id),c=await k("/api/import/training",{method:"POST",body:JSON.stringify({rows:b,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},formFields:n=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:n?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:n?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:n?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:e,value:n?.branch_id},{name:"trainer",label:"Trainer",type:"select",options:d(n?.trainer),value:n?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let s=JSON.parse(n?.participants);return Array.isArray(s)?s.join(", "):n?.participants||""}catch{return n?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:n?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:n?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],onBeforeSubmit:async n=>(n.participants&&(n.participants=JSON.stringify(n.participants.split(",").map(s=>s.trim()).filter(Boolean))),n)})}j();we();G();function Qt({container:t,title:e,icon:i,apiPath:a,columns:o,formFields:d,filterFields:r,defaultFilters:n={},enableMobileFilterSheet:s=!1,itemLabel:l="Data",canCreate:b=!0,canEdit:c=!0,canDelete:u=!0,onBeforeSubmit:m,onAfterLoad:y,onDataLoaded:g,extraActions:f=[],initialSearch:v="",exportOptions:S=null,bulkDelete:D=!1,paginationMode:C="server"}){let P=oe();P&&typeof P=="object"&&P.role==="viewer"&&(b=!1,c=!1,u=!1,D=!1,S=null);let $=1,L={...n};v&&(L.search=v);let E=new Set;t.innerHTML=`
    <div class="crud-layout-wrapper ${s?"mobile-active":""}">
      <div class="page-header">
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions">
        ${b?`<button class="btn btn-primary" id="btn-create">+ Tambah ${l}</button>`:""}
        ${S?'<button class="btn btn-outline" id="btn-mobile-aksi" style="display:none; align-items:center; justify-content:center; gap:0.25rem;">\u22EE Aksi</button>':""}
      </div>
    </div>

    ${D?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${S?`
    <div class="excel-actions-wrapper" id="excel-actions-wrapper">
      <div class="bottom-sheet-header aksi-header" style="display:none;">
        <h3 style="margin:0; font-size:1rem;">Aksi</h3>
        <button class="btn-close-sheet" id="btn-close-aksi-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
      </div>
      ${Ie(S.moduleName)}
    </div>`:""}

    ${r&&r.length>0?`
    <div class="filter-bar card ${s?"has-mobile-sheet":""}" style="padding: 1rem; display: flex; gap: 0.5rem; flex-wrap: wrap;">
        ${r.filter(p=>p.type==="search"||p.type==="search-combo").map(p=>{if(p.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${p.placeholder||"Cari..."}" id="filter-search" value="${L.search||""}"></div>`;if(p.type==="search-combo"){let h="dl-filter-search",w=(p.options||[]).map(x=>`<option value="${typeof x=="object"?x.label:x}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${h}" class="form-control" autocomplete="off" placeholder="${p.placeholder||"Cari..."}" id="filter-search" value="${L.search||""}"><datalist id="${h}">${w}</datalist></div>`}return""}).join("")}
        
        <div class="filter-options" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${r.filter(p=>p.type!=="search"&&p.type!=="search-combo").map(p=>{if(p.type==="select")return`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${p.name}" id="filter-${p.name}"><option value="">Pilih ${p.label}</option>${(p.options||[]).map(h=>`<option value="${typeof h=="object"?h.value:h}" ${L[p.name]===(typeof h=="object"?h.value:h)?"selected":""}>${typeof h=="object"?h.label:h}</option>`).join("")}</select>`;if(p.type==="combobox"){let h=`dl-filter-${p.name}`,w=(p.options||[]).map(T=>{let _=typeof T=="object"?T.label||T.value||"":T||"";return(_==="undefined"||_==="[object Object]"||_==="null")&&(_=""),_?`<option value="${_}"></option>`:""}).join(""),x=L[p.name]||"";if(L[p.name]){let T=(p.options||[]).find(_=>String(typeof _=="object"?_.value:_)==String(L[p.name]));if(T){let _=typeof T=="object"?T.label||T.value||"":T||"";_&&_!=="undefined"&&_!=="[object Object]"&&_!=="null"&&(x=_)}}return`<div class="filter-combobox" style="flex:1; min-width:120px;">
                <input type="text" name="${p.name}" id="filter-${p.name}" list="${h}" class="form-control filter-combobox-input" value="${x}" placeholder="Pilih ${p.label}..." autocomplete="off">
                <datalist id="${h}">${w}</datalist>
              </div>`}return""}).join("")}
          <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
        </div>
        ${s?'<button id="btn-mobile-filter" class="btn btn-outline" style="display:none; align-items:center; gap:0.25rem;">\u2699\uFE0F Filter</button>':""}
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
    </div>
  `;function F(){let p=document.getElementById("bulk-toolbar");if(!p)return;let h=document.getElementById("bulk-count"),w=document.getElementById("btn-bulk-delete"),x=document.getElementById("btn-bulk-cancel");h.textContent=`${E.size} item dipilih`,E.size>0?(p.classList.add("has-items"),w.disabled=!1,x.disabled=!1):(p.classList.remove("has-items"),w.disabled=!0,x.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{E.clear(),document.querySelectorAll(".row-checkbox").forEach(h=>h.checked=!1);let p=document.getElementById("select-all-checkbox");p&&(p.checked=!1),F()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(E.size===0)return;let p=[...E],h=document.createElement("div");h.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",h.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${p.length} ${l}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${p.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(h),h.querySelector("#bulk-cancel-btn").addEventListener("click",()=>h.remove()),h.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let w=h.querySelector("#bulk-confirm-btn");w.disabled=!0,w.textContent="Menghapus...";let x=await k(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:p})});h.remove(),x.ok?(ne(`${p.length} ${l} berhasil dihapus.`),E.clear(),F(),q()):z(x.data?.error||"Gagal menghapus data.")})});let K=document.getElementById("filter-search"),ue;K?.addEventListener("input",p=>{clearTimeout(ue),ue=setTimeout(()=>{L.search=p.target.value,$=1,E.clear(),q()},400)}),r?.forEach(p=>{p.type==="select"&&document.getElementById(`filter-${p.name}`)?.addEventListener("change",h=>{L[p.name]=h.target.value,$=1,E.clear(),q()}),p.type==="combobox"&&document.getElementById(`filter-${p.name}`)?.addEventListener("change",h=>{let w=h.target.value,x=(p.options||[]).find(T=>{let _=String(typeof T=="object"?T.value:T),B=String(typeof T=="object"?T.label:T);return _===w||B===w});w?L[p.name]=x?typeof x=="object"?x.value:x:w:L[p.name]="",$=1,E.clear(),q()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{L={...n},K&&(K.value=""),r?.forEach(p=>{let h=document.getElementById(`filter-${p.name}`);h&&(h.value="")}),$=1,E.clear(),q()});let me=document.getElementById("btn-mobile-filter"),se=document.getElementById("filter-options-wrapper"),ge=document.getElementById("btn-close-filter-sheet");me&&se&&(me.addEventListener("click",p=>{p.preventDefault(),se.classList.add("sheet-open")}),ge&&ge.addEventListener("click",p=>{p.preventDefault(),se.classList.remove("sheet-open")}));let pe=document.getElementById("btn-mobile-aksi"),be=document.getElementById("excel-actions-wrapper"),_e=document.getElementById("btn-close-aksi-sheet");if(pe&&be&&(pe.addEventListener("click",p=>{p.preventDefault(),be.classList.add("sheet-open")}),_e&&_e.addEventListener("click",p=>{p.preventDefault(),be.classList.remove("sheet-open")})),document.getElementById("btn-create")?.addEventListener("click",()=>Ce(null)),S){document.getElementById(`btn-export-${S.moduleName}`)?.addEventListener("click",async h=>{let w=h.target,x=w.innerHTML;w.innerHTML="\u23F3 Loading...",w.disabled=!0;try{await S.onExport()}catch{z("Gagal export data")}finally{w.innerHTML=x,w.disabled=!1}}),document.getElementById(`btn-template-${S.moduleName}`)?.addEventListener("click",()=>{S.onTemplate()});let p=document.getElementById(`input-import-${S.moduleName}`);p?.addEventListener("change",async h=>{let w=h.target.files[0];if(!w)return;p.disabled=!0;let x=document.createElement("div");x.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",x.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(x);let T=x.querySelector("#import-progress-text"),_=x.querySelector("#import-progress-bar"),B=x.querySelector("#import-summary"),A=x.querySelector("#import-close-btn");A.addEventListener("click",()=>{x.remove(),q()});try{let W=await De(w);if(W.length===0)throw new Error("File kosong atau format salah");let V=500,ae=0,Z=0,N=0,M=W.length;T.textContent=`Ditemukan ${M} baris data. Memulai import...`;for(let I=0;I<M;I+=V){let J=W.slice(I,I+V);T.textContent=`Mengimport baris ${I+1} - ${Math.min(I+V,M)} dari ${M}...`,_.style.width=`${Math.round(I/M*100)}%`;try{let O=await S.onImport(J);O?(ae+=O.inserted||O.metrics?.inserted||J.length,Z+=O.skipped||O.metrics?.updated||0):ae+=J.length}catch(O){console.error("Chunk import failed:",O),N+=J.length}}_.style.width="100%",T.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',B.style.display="block",B.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${M}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${Z}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${N}</strong></div>
        `,N>0&&(B.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),A.style.display="block",p.value=""}catch(W){T.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${W.message}`,_.style.background="var(--danger)",_.style.width="100%",A.style.display="block",p.value=""}finally{p.disabled=!1}})}async function q(){F();let p=document.getElementById("table-container");if(!p)return;p.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let h=C==="client",w=h?1:$,x=h?Se:20,T=new URLSearchParams({page:w,limit:x,...Object.fromEntries(Object.entries(L).filter(([,N])=>N))}),_=await k(`${a}?${T}`);if(!_.ok){p.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${_.data?.error||"Error"}</p></div>`;return}let B=_.data?.data||_.data||[],A=_.data?.pagination,W=B.length,V=B;if(h){B=g(B),V=B;let N=B.length,M=20,I=Math.ceil(N/M);$>I&&I>0&&($=I);let J=($-1)*M,O=$*M;B=B.slice(J,O),A={page:$,limit:M,total:N,pages:I}}!1,y&&y(B);let ae=Pe({columns:o,data:B,fullData:V,onEdit:c?N=>Ce(N):null,actions:f.map(N=>({...N,handler:M=>N.handler(M,q)})),emptyText:`Tidak ada ${String(l||"").toLowerCase()}`,bulkSelect:D?{selectedIds:E,onToggle:F}:null});p.innerHTML="",p.appendChild(ae);let Z=document.getElementById("pagination-container");if(Z&&(Z.innerHTML="",A&&A.pages>1)){let N=Ae({page:A.page,pages:A.pages,total:A.total,limit:A.limit,onPage:M=>{$=M,q()}});N&&Z.appendChild(N)}}function Ge(p){let h=typeof d=="function"?d(p):d;return Te(h)}function Ce(p){let h=!!p,w=document.createElement("form");if(w.noValidate=!0,w.innerHTML=Ge(p),h){let T=typeof d=="function"?d(p):d;Me(w,p)}let{close:x}=de({title:h?`Edit ${l}`:`Tambah ${l}`,content:w,size:"lg",confirmText:h?"Simpan Perubahan":`Tambah ${l}`,onConfirm:async(T,_)=>{if(!w.reportValidity())return;let B=T.querySelector(".modal-confirm");B.disabled=!0,B.textContent="Menyimpan...";let A=Ne(w),W=typeof d=="function"?d(p):d,V=async M=>{for(let I of M)if(I.type==="row")await V(I.fields);else if(I.type==="combobox"&&A[I.name]){let J=A[I.name],O=(I.options||[]).find(U=>{let ee=String(typeof U=="object"?U.value:U),ze=String(typeof U=="object"?U.label:U);return ee===J||ze===J});if(O)A[I.name]=typeof O=="object"?O.value:O;else if(I.createApi){let U={};U[I.createApi.field]=J,I.createApi.extra&&Object.assign(U,I.createApi.extra);let ee=await k(I.createApi.path,{method:"POST",body:JSON.stringify(U)});if(ee.ok&&ee.data?.id)A[I.name]=ee.data.id;else if(ee.ok&&!ee.data?.id)A[I.name]=J;else throw new Error(`Gagal membuat master data: ${ee.data?.error||"Unknown error"}`)}}};try{await V(W)}catch(M){z(M.message),B.disabled=!1,B.textContent=h?"Simpan Perubahan":`Tambah ${l}`;return}m&&(A=await m(A,p));let ae=h?"PUT":"POST",Z=h?`${a}/${p.id}`:a,N=await k(Z,{method:ae,body:JSON.stringify(A)});N.ok?(ne(h?`${l} berhasil diperbarui.`:`${l} berhasil ditambahkan.`),_(),q()):(z(N.data?.error||"Gagal menyimpan data."),B.disabled=!1,B.textContent=h?"Simpan Perubahan":`Tambah ${l}`)}})}function $t(p){Ee(`Hapus ${l} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let h=await k(`${a}/${p.id}`,{method:"DELETE"});h.ok?(ne(`${l} berhasil dihapus.`),q()):z(h.data?.error||"Gagal menghapus.")},`Hapus ${l}`)}return q(),q}j();G();async function Wt(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await Q(),a=await ie(),o=e?e.get("dash_filter"):null,d={};if(o==="reliever"){let l=new Date,b=`${l.getFullYear()}-${String(l.getMonth()+1).padStart(2,"0")}`;d={status:"Done",month:e&&e.get("month")?e.get("month"):b}}console.log("RAW",await Oe()),console.log("OPTIONS",a);let r=l=>l&&!a.find(b=>b.value===l)?[...a,{value:l,label:l}]:a,n=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],s=l=>l&&!n.includes(l)?[...n,l]:n;Qt({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",defaultFilters:d,onDataLoaded:l=>l.sort((b,c)=>{let u=b.backup_date?new Date(b.backup_date).getTime():0;return(c.backup_date?new Date(c.backup_date).getTime():0)-u}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:l=>he(l)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:l=>window.formatDate(l)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:l=>window.formatDate(l)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:l=>l?`<span class="badge badge-info">${l}</span>`:"-"},{key:"status",label:"Status",render:l=>X(l)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:n},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:l=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:i,value:l?.branch_id},{name:"period",label:"Periode",type:"select",options:["Q1","Q2","Q3","Q4"],value:l?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"select",options:r(l?.original_fc_name),value:l?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"select",required:!0,options:s(l?.reliever_name),value:l?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:l?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:l?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"select",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:l?.reason},{name:"shift",label:"Shift",type:"select",options:["Pagi","Siang","Full Shift","Middle"],value:l?.shift}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done","Tidak Datang"],value:l?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let l=await k(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(l.ok){let b=l.data.data.map(c=>({Cabang:c.branch_name||"","Nama Facility care":c.original_fc_name||"",Periode:c.period||"",Relifer:c.reliever_name||"","Tanggal Back Up":c.backup_date||"","Tanggal Selesai":c.completion_date||"",Keterangan:c.reason||"",Shift:c.shift||"",Status:c.status||""}));b.length===0&&b.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),R(b,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async l=>{let c=(await k("/api/branches?all=1")).data?.data||[],u=g=>{if(!g)return null;let f=String(g||"").toLowerCase(),v=c.find(S=>String(S.full_name||"").toLowerCase()===f||String(S.code||"").toLowerCase()===f||String(S.name||"").toLowerCase()===f);return v?v.id:null},m=l.map(g=>({branch_name:String(g.Cabang||"").trim(),backup_date:String(g["Tanggal Back Up"]||g["Tanggal Backup"]||"").trim(),original_fc_name:String(g["Nama Facility care"]||g["FC Digantikan"]||"").trim(),reliever_name:String(g.Relifer||g.Reliefer||"").trim(),period:String(g.Periode||"").trim(),reason:String(g.Keterangan||"").trim(),shift:String(g.Shift||"").trim(),completion_date:String(g["Tanggal Selesai"]||"").trim(),status:String(g.Status||"").trim()})).filter(g=>g.reliever_name&&g.backup_date),y=await k("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!y.ok)throw new Error(y.data?.error||"Import gagal");return y.data}}})}j();G();async function Vt(t){let e=await Q(),i=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));H({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>he(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>X(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),d=await k(`/api/reports/inspection?limit=10000&${o}`);if(d.ok){let r=d.data.data.map(n=>({Cabang:n.branch_name||"",Periode:n.period||"",Tanggal:n.inspection_date||"","Point FC":n.fc_score!==null&&n.fc_score!==void 0?n.fc_score:"","Point SPV":n.spv_score!==null&&n.spv_score!==void 0?n.spv_score:"",Status:n.status||"","Link Dokumen":n.document_link||""}));R(r,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let o=s=>{if(!s)return null;let l=String(s||"").toLowerCase(),b=e.find(c=>String(c.label||"").toLowerCase()===l);return b?b.value:null},d=s=>{if(s==null||s==="")return"";if(s instanceof Date&&!isNaN(s.getTime()))return s.toISOString().slice(0,10);let l=String(s).trim();if(l===""||l==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);if(/^\d{4,5}$/.test(l)){let c=Number(l);if(c>2e4&&c<99999){let u=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}let b=l.split(/[\/\-\.]/);if(b.length===3){let[c,u,m]=b.map(y=>y.trim());if(c.length===4&&u.length<=2&&m.length<=2)return`${c}-${u.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&u.length<=2&&c.length<=2)return`${m}-${u.padStart(2,"0")}-${c.padStart(2,"0")}`}return l},r=a.map(s=>({branch_id:o(String(s.Cabang||"").trim()),period:String(s.Periode||"").trim(),inspection_date:d(s.Tanggal),fc_score:s["Point FC"]!==void 0&&s["Point FC"]!==""?Number(s["Point FC"]):null,spv_score:s["Point SPV"]!==void 0&&s["Point SPV"]!==""?Number(s["Point SPV"]):null,status:String(s.Status||"").trim(),document_link:String(s["Link Dokumen"]||"").trim(),notes:String(s.Catatan||s.Keterangan||"").trim()})).filter(s=>s.branch_id&&s.period&&s.inspection_date),n=await k("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}j();G();async function Yt(t){let e=await Q(),i=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));H({container:t,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>he(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>X(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),d=await k(`/api/reports/cleaning?limit=10000&${o}`);if(d.ok){let r=d.data.data.map(n=>({Cabang:n.branch_name||"",Jenis:n.activity_type||"",Periode:n.period||"",Tanggal:n.activity_date||"",Status:n.status||"","Link Dokumen":n.document_link||""}));R(r,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let o=s=>{if(!s)return null;let l=String(s||"").toLowerCase(),b=e.find(c=>String(c.label||"").toLowerCase()===l);return b?b.value:null},d=s=>{if(s==null||s==="")return"";if(s instanceof Date&&!isNaN(s.getTime()))return s.toISOString().slice(0,10);let l=String(s).trim();if(l===""||l==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);if(/^\d{4,5}$/.test(l)){let c=Number(l);if(c>2e4&&c<99999){let u=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}let b=l.split(/[\/\-\.]/);if(b.length===3){let[c,u,m]=b.map(y=>y.trim());if(c.length===4&&u.length<=2&&m.length<=2)return`${c}-${u.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&u.length<=2&&c.length<=2)return`${m}-${u.padStart(2,"0")}-${c.padStart(2,"0")}`}return l},r=a.map(s=>({branch_id:o(String(s.Cabang||"").trim()),activity_type:String(s.Jenis||s.Kegiatan||"").trim(),period:String(s.Periode||"").trim(),activity_date:d(s.Tanggal),status:String(s.Status||"").trim(),document_link:String(s["Link Dokumen"]||"").trim(),notes:String(s.Catatan||s.Keterangan||"").trim()})).filter(s=>s.branch_id&&s.activity_type&&s.period&&s.activity_date),n=await k("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}j();G();async function Xt(t,e){let i=await Q(),a=Array.from({length:4},(r,n)=>String(new Date().getFullYear()-n)),o=e?e.get("dash_filter"):null,d={};if(o==="fogging"){let r=new Date,n=String(r.getMonth()+1).padStart(2,"0"),s=String(r.getFullYear()),l=e?e.get("month"):null;l&&l.length===7&&(s=l.split("-")[0],n=l.split("-")[1]),d={status:"Done",month:n,year:s}}H({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,defaultFilters:d,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:r=>`<span class="badge badge-warning">${r}</span>`},{key:"period",label:"Periode",render:r=>he(r)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"status",label:"Status",render:r=>X(r)},{key:"document_link",label:"Dokumen",render:r=>r?`<a href="${r}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:r=>r||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:r=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:i,value:r?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:r?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:r?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:r?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:r?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async r=>{let n=new URLSearchParams(r||{}).toString(),s=await k(`/api/reports/fogging?limit=10000&${n}`);if(s.ok){let l=s.data.data.map(b=>({Cabang:b.branch_name||"",Jenis:b.activity_type||"Fogging",Periode:b.period||"",Tanggal:b.activity_date||"",Status:b.status||"","Link Dokumen":b.document_link||""}));R(l,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async r=>{let n=c=>{if(!c)return null;let u=String(c||"").toLowerCase(),m=i.find(y=>String(y.label||"").toLowerCase()===u);return m?m.value:null},s=c=>{if(c==null||c==="")return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let u=String(c).trim();if(u===""||u==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);if(/^\d{4,5}$/.test(u)){let y=Number(u);if(y>2e4&&y<99999){let g=new Date(Date.UTC(1899,11,30)+y*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}let m=u.split(/[\/\-\.]/);if(m.length===3){let[y,g,f]=m.map(v=>v.trim());if(y.length===4&&g.length<=2&&f.length<=2)return`${y}-${g.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&g.length<=2&&y.length<=2)return`${f}-${g.padStart(2,"0")}-${y.padStart(2,"0")}`}return u},l=r.map(c=>({branch_id:n(String(c.Cabang||"").trim()),activity_type:String(c.Jenis||c.Kegiatan||"Fogging").trim(),period:String(c.Periode||"").trim(),activity_date:s(c.Tanggal),status:String(c.Status||"").trim(),document_link:String(c["Link Dokumen"]||"").trim(),notes:String(c.Catatan||c.Keterangan||"").trim()})).filter(c=>c.branch_id&&c.period&&c.activity_date),b=await k("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(l)});if(!b.ok)throw new Error(b.data?.error||"Import gagal");return b.data}}})}j();G();async function Zt(t){let e=await Q(),i=await ie(),a=i,o=Array.from({length:4},(n,s)=>String(new Date().getFullYear()-s)),d=n=>n&&!i.find(s=>s.value===n)?[...i,{value:n,label:n}]:i,r=n=>n&&!a.find(s=>s.value===n)?[...a,{value:n,label:n}]:a;H({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:n=>`<span title="${n||""}">${n?.length>60?n.slice(0,60)+"\u2026":n||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:n=>window.formatDate(n)},{key:"status",label:"Status",render:n=>X(n)},{key:"notes",label:"Keterangan",render:n=>n?.length>40?n.slice(0,40)+"\u2026":n||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade","Mizwar"]},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:o}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:n?.branch_id},{name:"pic",label:"PIC",type:"select",options:r(n?.pic),value:n?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:n?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:n?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:n?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:n?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async n=>{let s=new URLSearchParams(n||{}).toString(),l=await k(`/api/reports/basecamp?limit=10000&${s}`);if(l.ok){let b=l.data.data.map(c=>({"Tgl Info":c.info_date||"",Cabang:c.branch_name||"",Permasalahan:c.problem||"",PIC:c.pic||"","Tgl Done":c.done_date||"",Status:c.status||"",Keterangan:c.notes||""}));R(b,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async n=>{let s=u=>{if(!u)return null;let m=String(u||"").toLowerCase(),y=e.find(g=>String(g.label||"").toLowerCase()===m);return y?y.value:null},l=u=>{if(u==null||u==="")return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let m=String(u).trim();if(m===""||m==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);if(/^\d{4,5}$/.test(m)){let g=Number(m);if(g>2e4&&g<99999){let f=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}let y=m.split(/[\/\-\.]/);if(y.length===3){let[g,f,v]=y.map(S=>S.trim());if(g.length===4&&f.length<=2&&v.length<=2)return`${g}-${f.padStart(2,"0")}-${v.padStart(2,"0")}`;if(v.length===4&&f.length<=2&&g.length<=2)return`${v}-${f.padStart(2,"0")}-${g.padStart(2,"0")}`}return m},b=n.map(u=>({info_date:l(u["Tgl Info"]||u["Tanggal Info"]),branch_id:s(String(u.Cabang||"").trim()),problem:String(u.Permasalahan||"").trim(),pic:String(u.PIC||"").trim(),done_date:l(u["Tgl Done"]||u["Tanggal Done"]),status:String(u.Status||"").trim(),notes:String(u.Keterangan||u.Catatan||"").trim()})).filter(u=>u.info_date&&u.branch_id&&u.problem),c=await k("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(b)});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}}})}async function ea(t){H({container:t,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(j(),Le)),o=await a(`/api/sop?limit=10000&${i}`);if(o.ok){let d=o.data.data.map(n=>({"Nama SOP":n.name||"",Kategori:n.category||"",Dokumen:n.document_link||"",Catatan:n.notes||n.description||""})),{downloadExcel:r}=await Promise.resolve().then(()=>(G(),ce));r(d,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(G(),ce));i(e,"Template_Import_SOP")},onImport:async e=>{let i=e.map(d=>({name:String(d["Nama SOP"]||"").trim(),category:String(d.Kategori||"").trim(),document_link:String(d.Dokumen||"").trim(),description:String(d.Catatan||"").trim()})).filter(d=>d.name),{apiFetch:a}=await Promise.resolve().then(()=>(j(),Le)),o=await a("/api/sop/import",{method:"POST",body:JSON.stringify(i)});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function ta(t){H({container:t,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(j(),Le)),o=await a(`/api/checklist?limit=10000&${i}`);if(o.ok){let d=o.data.data.map(n=>({"Nama Checklist":n.name||"",Kategori:n.category||"",Dokumen:n.document_link||"",Deskripsi:n.description||""})),{downloadExcel:r}=await Promise.resolve().then(()=>(G(),ce));r(d,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(G(),ce));i(e,"Template_Import_Checklist")},onImport:async e=>{let i=e.map(d=>({name:String(d["Nama Checklist"]||"").trim(),category:String(d.Kategori||"").trim(),document_link:String(d.Dokumen||"").trim(),description:String(d.Deskripsi||"").trim()})).filter(d=>d.name),{apiFetch:a}=await Promise.resolve().then(()=>(j(),Le)),o=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(i)});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}j();we();G();async function Ct(t,e="forms"){if(e==="supply")return Ha(t);ja(t)}function ja(t){H({container:t,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await k(`/api/forms?limit=10000&${i}`);a.data?.data?R(a.data.data,"Data_Master_Form"):z("Gagal export data master form")},onImport:async e=>{let i=await k("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!i.ok)throw new Error(i.data?.error||"Import failed");return i.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function Ha(t){let i=((await k("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));H({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,o)=>o.branch_name_ref||o.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let o=JSON.parse(a);return Array.isArray(o)?o.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let o=JSON.parse(a);return Array.isArray(o)?o.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>X(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let o=a?.tools_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}let d=a?.chemical_items;try{d=Array.isArray(JSON.parse(d))?JSON.parse(d).join(", "):d}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:a?.branch_id&&!i.find(r=>r.value==a.branch_id)?[...i,{value:a.branch_id,label:a.branch_name||a.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:o},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:d},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),d=await k(`/api/reports/supply?limit=10000&${o}`);if(d.ok){let r=d.data.data.map(n=>{let s=n.tools_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}let l=n.chemical_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}return{Waktu:n.submitted_at||"",Pengirim:n.submitter_name||"",Cabang:n.branch_name_ref||n.branch_name||"","Alat/Barang":s||"",Chemical:l||"",Catatan:n.additional_notes||"",Status:n.status||"","Diproses Oleh":n.processed_by||""}});R(r,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let d=(await k("/api/branches?all=1")).data?.data||[],r=b=>{if(!b)return null;let c=String(b||"").toLowerCase(),u=d.find(m=>String(m.full_name||"").toLowerCase()===c||String(m.code||"").toLowerCase()===c||String(m.name||"").toLowerCase()===c);return u?u.id:null},n=b=>{if(b==null||b==="")return"";if(b instanceof Date&&!isNaN(b.getTime()))return b.toISOString().slice(0,10);let c=String(b).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let m=Number(c);if(m>2e4&&m<99999){let y=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}let u=c.split(/[\/\-\.]/);if(u.length===3){let[m,y,g]=u.map(f=>f.trim());if(m.length===4&&y.length<=2&&g.length<=2)return`${m}-${y.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&y.length<=2&&m.length<=2)return`${g}-${y.padStart(2,"0")}-${m.padStart(2,"0")}`}return c},s=a.map(b=>({submitted_at:n(b.Waktu||b.Tanggal),submitter_name:String(b.Pengirim||"").trim(),branch_id:r(String(b.Cabang||"").trim()),tools_items:String(b["Alat/Barang"]||b.Alat||"").trim(),chemical_items:String(b.Chemical||"").trim(),additional_notes:String(b.Catatan||b.Keterangan||"").trim(),status:String(b.Status||"").trim(),processed_by:String(b["Diproses Oleh"]||b.PIC||"").trim()})).filter(b=>b.submitted_at&&b.submitter_name&&b.branch_id),l=await k("/api/reports/supply/import",{method:"POST",body:JSON.stringify(s)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,o)=>{let d=de({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(r,n)=>{let s=r.querySelector("#supply-status").value,l=r.querySelector("#supply-processed-by").value;(await k(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:s,processed_by:l})})).ok?(ne("Status diperbarui."),n(),o()):z("Gagal update status.")}})}}]})}j();async function aa(t){let e=oe();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}H({container:t,title:"Riwayat Aktivitas",icon:"\u{1F575}\uFE0F\u200D\u2642\uFE0F",apiPath:"/api/audit-logs",itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:i=>new Date(i).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(i,a)=>`<strong>${i||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${a.user_role||""}</small>`},{key:"action",label:"Aksi",render:i=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[i]||"badge-neutral"}">${i}</span>`},{key:"module",label:"Modul",render:i=>`<span style="text-transform:capitalize">${(i||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(i,a)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${i}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async i=>{try{let d=((await(await fetch(`/api/audit-logs?search=${i}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(l=>String(l.id)===String(i));if(!d)return alert("Data tidak ditemukan");let r=l=>{if(!l)return"Tidak ada data";try{return JSON.stringify(JSON.parse(l),null,2)}catch{return l}},n=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${r(d.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${r(d.new_data)}</pre>
           </div>
         </div>
       `,{createModal:s}=await Promise.resolve().then(()=>(we(),mt));s({title:`Detail Audit Log #${i}`,content:n,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}j();G();async function na(t){let e=oe();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}H({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let a=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await k(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let a=i.data.data.map(o=>({"Nama Lengkap":o.full_name||"",Username:o.username||"",Email:o.email||"",Role:o.role||"",Status:o.is_active?"Aktif":"Nonaktif"}));R(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let a=i.map(d=>({full_name:String(d["Nama Lengkap"]||"").trim(),username:String(d.Username||"").trim(),email:String(d.Email||"").trim(),role:String(d.Role||"").trim()||"viewer",password:String(d.Password||"").trim()})).filter(d=>d.username&&d.password&&d.email&&d.full_name),o=await k("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}}})}j();G();async function ia(t){H({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await k(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)R(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{R([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let i=e.map(o=>({code:String(o["Kode Cabang"]||"").trim(),name:String(o["Nama Pendek"]||"").trim(),full_name:String(o["Nama Lengkap"]||"").trim(),city:String(o.Kota||"").trim()})).filter(o=>o.code&&o.name),a=await k("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}j();async function ra(t){let e=new Date,i=[];t.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),o()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),o()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(d=>d.addEventListener("change",o));async function a(){try{let d=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;i=(await k(`/api/dashboard/calendar?month=${d}`)).data?.data||[]}catch(d){console.warn("[Calendar] Failed to load events, rendering empty grid:",d),i=[]}}async function o(){let d=document.getElementById("calendar-grid");if(d){d.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let r=e.getFullYear(),n=e.getMonth(),s=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),l=document.getElementById("cal-month-label");l&&(l.textContent=s);let b=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(C=>C.value)),c=i.filter(C=>b.has(C.type)),u={};c.forEach(C=>{let P=(C.event_date||"").slice(0,10);u[P]||(u[P]=[]),u[P].push(C)});let m=new Date(r,n,1).getDay(),y=new Date(r,n+1,0).getDate(),g=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],f=new Date().toISOString().slice(0,10),v='<div class="calendar-grid">';g.forEach(C=>{v+=`<div class="cal-day-header">${C}</div>`});for(let C=0;C<m;C++)v+='<div class="cal-cell cal-cell-empty"></div>';for(let C=1;C<=y;C++){let P=`${r}-${String(n+1).padStart(2,"0")}-${String(C).padStart(2,"0")}`,$=u[P]||[],L=P===f;v+=`
          <div class="cal-cell ${L?"cal-today":""} ${$.length?"cal-has-events":""}"
               data-date="${P}" tabindex="0" role="button" aria-label="${P}">
            <div class="cal-day-num ${L?"today-num":""}">${C}</div>
            <div class="cal-events-preview">
              ${$.slice(0,3).map(E=>`
                <div class="cal-event-dot cal-color-${E.color||"gray"}" title="${lt(E.title||E.type)}">
                  <span class="cal-event-dot-label">${qa(E.title||E.branch_name||E.type,18)}</span>
                </div>
              `).join("")}
              ${$.length>3?`<div class="cal-more">+${$.length-3} lagi</div>`:""}
            </div>
          </div>`}let D=(m+y)%7;if(D!==0)for(let C=0;C<7-D;C++)v+='<div class="cal-cell cal-cell-empty"></div>';v+="</div>",d.innerHTML=v,d.querySelectorAll(".cal-cell[data-date]").forEach(C=>{C.addEventListener("click",()=>{let P=C.dataset.date,$=u[P]||[];if(!$.length)return;let L=document.getElementById("cal-event-list"),E=new Date(P+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=E,document.getElementById("cal-event-items").innerHTML=$.map(F=>`
            <div class="cal-event-item cal-color-border-${F.color||"gray"}">
              <div class="cal-event-type">${Ja(F.type)}</div>
              <div class="cal-event-title">${lt(F.title||"-")}</div>
              <div class="cal-event-branch">${lt(F.branch_name||"")}</div>
              ${F.status?`<div class="cal-event-status">${lt(F.status)}</div>`:""}
              ${F.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${F.days_remaining} hari</div>`:""}
            </div>
          `).join(""),L.style.display="block"})})}catch(r){console.error("[Calendar] Render error:",r),d&&(d.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}o()}function qa(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function lt(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Ja(t){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[t]||t}j();async function la(t){let e=oe(),i=(e?.full_name||e?.username||"U")[0].toUpperCase(),o={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
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
            <div class="profile-fullname">${e?.full_name||"\u2014"}</div>
            <div class="profile-username">@${e?.username||"\u2014"}</div>
            <span class="badge badge-info" style="background:${o}18;color:${o};margin-top:6px">
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
            <span class="info-value" style="color:${o};font-weight:700">${e?.role||"\u2014"}</span>
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
  `;let d=localStorage.getItem("fm_token"),r=document.getElementById("session-info");if(d&&r)try{let n=JSON.parse(atob(d.split(".")[1])),s=new Date(n.exp*1e3);r.textContent=`Berakhir: ${s.toLocaleString("id-ID")}`}catch{r.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async n=>{n.preventDefault();let s=document.getElementById("pwd-error"),l=document.getElementById("pwd-success"),b=document.getElementById("btn-save-pwd");s.style.display="none",l.style.display="none";let c=n.target,u=c.current_password.value,m=c.new_password.value,y=c.confirm_password.value;if(m!==y){s.textContent="\u274C Konfirmasi password tidak cocok.",s.style.display="block";return}if(m.length<6){s.textContent="\u274C Password baru minimal 6 karakter.",s.style.display="block";return}b.disabled=!0,b.textContent="\u23F3 Menyimpan...";let g=await k("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:u,new_password:m})});b.disabled=!1,b.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',g.ok?(l.textContent="\u2705 Password berhasil diubah.",l.style.display="block",c.reset(),ne("Password berhasil diubah.")):(s.textContent=g.data?.error||"Gagal mengubah password.",s.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}j();var st={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function re(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let o=Number(e);if(o>2e4&&o<99999){let d=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(d.getTime())?null:d.toISOString().slice(0,10)}}let i=e.split(/[\/\-\.]/);if(i.length===3){let[o,d,r]=i.map(b=>b.trim()),n=Number(o),s=Number(d),l=Number(r);if(o.length===4&&n>1900)return`${o}-${d.padStart(2,"0")}-${r.padStart(2,"0")}`;if(r.length===4&&l>1900)return n>12?`${r}-${d.padStart(2,"0")}-${o.padStart(2,"0")}`:s>12?`${r}-${o.padStart(2,"0")}-${d.padStart(2,"0")}`:`${r}-${d.padStart(2,"0")}-${o.padStart(2,"0")}`;if(r.length===2&&!isNaN(l)){let b=l>=50?`19${r}`:`20${r}`;return n>12?`${b}-${d.padStart(2,"0")}-${o.padStart(2,"0")}`:`${b}-${d.padStart(2,"0")}-${o.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function sa(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Ua={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",phone:t["No. Hp"],join_date:re(t["Tanggal Masuk"]),status:t.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:re(t["Tanggal Mulai"]),end_date:re(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:re(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:re(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:re(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:re(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:re(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:re(t["Tanggal Target"]||t["Tgl Target"]),completion_date:re(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:re(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:re(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:re(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:re(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:re(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:re(t["Tanggal Back Up"]),completion_date:re(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:re(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:re(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Ga(t,e){let i=st[t];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Ua[i.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let o=[],d=[],r=[];return e.filter(s=>!sa(s)).forEach((s,l)=>{let b=e.indexOf(s)+2,c=[];a.required.forEach(({key:m,label:y})=>{let g=s[m];if(g==null||String(g).trim()===""){let f=Object.keys(s).filter(v=>v.trim()).join(", ");c.push({column:y,originalValue:g||"",reason:`Kolom "${y}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${f.slice(0,120)}`})}});let u=a.map(s);c.length>0?d.push({row:b,data:u,raw:s,errors:c}):(o.push(s),r.push(u))}),{valid:o,errors:d,mapped:r}}function oa(t){let e=[];return t.SheetNames.forEach(i=>{let a=st[i];if(!a)return;let o=t.Sheets[i],d=window.XLSX.utils.sheet_to_json(o,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),r=Ga(i,d),n=d.filter(s=>!sa(s));e.push({sheetName:i,module:a.module,label:a.label,total:n.length,valid:r.mapped.length,errorCount:r.errors.length,errors:r.errors,mapped:r.mapped,skipped:!1})}),e}function da(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,o])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(o),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function ca(t){let e=window.XLSX,i=e.utils.book_new(),a=!1;return t.forEach(o=>{if(!o.errors||o.errors.length===0)return;a=!0;let d=o.errors.map(n=>({"No. Baris":n.row,"Kolom Gagal":(n.errors||[]).map(s=>s.column||s).join("; "),"Alasan Error":(n.errors||[]).map(s=>s.reason||s).join("; "),...Object.fromEntries(Object.entries(n.data||{}).map(([s,l])=>[s,l??""]))})),r=e.utils.json_to_sheet(d);e.utils.book_append_sheet(i,r,o.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var za=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function pa(t){t.innerHTML=`
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
              ${Object.entries(st).map(([g,{label:f}])=>`<span class="import-sheet-tag">\u{1F4C4} ${g} \u2192 ${f}</span>`).join("")}
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
  `;let e=null,i=null,a=0,o={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function d(g){Object.entries(o).forEach(([f,v])=>{v.style.display=f===g?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let g=document.getElementById("btn-backup-db");g.disabled=!0,g.textContent="\u23F3 Memproses Backup...";try{let f=await k("/api/import/backup");if(f.ok){if(!window.XLSX){z("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let v=window.XLSX,S=v.utils.book_new();Object.entries(f.data.database).forEach(([D,C])=>{let P=C.length>0?C:[{}],$=v.utils.json_to_sheet(P);v.utils.book_append_sheet(S,$,D.substring(0,31))}),v.writeFile(S,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),ne("Backup berhasil diunduh!")}else z("Gagal memproses backup: "+(f.data?.error||"Unknown error"))}catch(f){z("Gagal memproses backup: "+f.message)}finally{g.disabled=!1,g.textContent="\u{1F4E6} Backup Database"}});let r=document.getElementById("btn-sync-google");r&&r.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let g=r.innerHTML;r.innerHTML='<span class="spinner"></span> Menyinkronkan...',r.disabled=!0;try{let f=await k("/api/sync/google-sheets",{method:"POST"});f.ok?alert("Sinkronisasi Berhasil: "+(f.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(f.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{r.innerHTML=g,r.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{da(),ne("Template Excel berhasil didownload!")});let n=document.getElementById("file-input"),s=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",g=>{g.stopPropagation(),n.click()}),n.addEventListener("change",g=>{g.target.files[0]&&l(g.target.files[0])}),s.addEventListener("dragover",g=>{g.preventDefault(),s.classList.add("drag-over")}),s.addEventListener("dragleave",()=>s.classList.remove("drag-over")),s.addEventListener("drop",g=>{g.preventDefault(),s.classList.remove("drag-over");let f=g.dataTransfer.files[0];f&&f.name.match(/\.xlsx?$/i)?l(f):z("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,n.value="",document.getElementById("file-info").style.display="none",s.style.display="",d("upload")});async function l(g){e=g,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${g.name} (${(g.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",s.style.display="none",await b(g)}async function b(g){d("validating");let f=document.getElementById("validation-status"),v=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");f.textContent="Membaca file Excel...",v.style.width="20%",await et(200);let S=await g.arrayBuffer(),D=window.XLSX.read(S,{type:"array",cellDates:!0});f.textContent=`Memvalidasi ${D.SheetNames.length} sheet...`,v.style.width="50%",await et(100),i=oa(D),v.style.width="100%",f.textContent="Validasi selesai!",await et(300),c()}catch(S){d("upload"),z("Gagal memproses file: "+S.message),document.getElementById("file-info").style.display="flex",s.style.display="none"}}function c(){d("preview");let g=i.filter(E=>!E.skipped).length,f=i.reduce((E,F)=>E+F.total,0),v=i.reduce((E,F)=>E+F.valid,0),S=i.reduce((E,F)=>E+F.errorCount,0),D=f>0?Math.round(v/f*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${g} sheet</span>
      <span class="badge badge-secondary">${f} baris</span>
      <span class="badge badge-success">${v} valid (${D}%)</span>
      ${S>0?`<span class="badge badge-danger">${S} error</span>`:""}
    `;let C=document.getElementById("preview-table-container");C.innerHTML=`
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
          ${i.map((E,F)=>`
            <tr class="${E.errorCount>0?"row-error":E.skipped?"row-skipped":"row-ok"}">
              <td><strong>${E.sheetName}</strong></td>
              <td>${E.label}</td>
              <td style="text-align:center">${E.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${E.valid}</span></td>
              <td style="text-align:center">${E.errorCount>0?`<span class="badge badge-danger">${E.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${E.skipped?'<span class="badge badge-neutral">Dilewati</span>':E.errorCount>0&&E.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':E.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':E.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${E.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${F}">\u{1F50D} ${E.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,C.querySelectorAll(".btn-detail-error").forEach(E=>{E.addEventListener("click",()=>{let F=i[Number(E.dataset.idx)];u(F)})});let P=document.getElementById("error-detail-section"),$=document.getElementById("error-detail-container");$.innerHTML="",P.style.display="none";let L=document.getElementById("btn-start-import");v===0?(L.disabled=!0,L.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(L.disabled=!1,S>0?(L.innerHTML=`\u{1F680} Import ${v} Data Valid (${S} dilewati)`,L.title="Baris error akan dilewati, baris valid tetap diimport"):L.innerHTML=`\u{1F680} Mulai Import ${v} Data`)}function u(g){let f=document.getElementById("error-detail-section"),v=document.getElementById("error-detail-container");f.style.display="";let S=g.errors.slice(0,100).map(D=>(Array.isArray(D.errors)?D.errors:[]).map(P=>{let $=typeof P=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${D.row}</span></td>
            <td><strong>${$?P.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${$&&P.originalValue!==void 0?P.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${$?P.reason:P}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${$&&P.aliases?`Gunakan salah satu nama kolom:<br><em>${P.aliases}</em>`:$&&P.hint?P.hint:""}
            </td>
          </tr>
        `}).join("")).join("");v.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${g.sheetName} \u2014 ${g.errorCount} baris error dari ${g.total} total
          ${g.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
        ${g.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,f.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{d("upload"),document.getElementById("file-info").style.display="none",s.style.display="",e=null,n.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;ca(i)?ne("Log error berhasil didownload."):ne("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let g=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";m(g)});async function m(g){d("importing"),a=Date.now();let f=[];za.forEach(P=>{let $=i?.find(L=>L.module===P&&L.mapped?.length>0);$&&f.push($)});let v=document.getElementById("import-steps-list");v.innerHTML=f.map(P=>`
      <div class="import-step-item" id="step-item-${P.module}">
        <span class="step-item-icon" id="step-icon-${P.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${P.label} <span class="step-item-count">(${P.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${P.module}"></span>
      </div>
    `).join("");let S=document.getElementById("import-bar"),D=document.getElementById("import-current-status"),C={totalSheets:f.length,totalRows:f.reduce((P,$)=>P+$.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let P=0;P<f.length;P++){let $=f[P],L=document.getElementById(`step-icon-${$.module}`),E=document.getElementById(`step-status-${$.module}`);L.textContent="\u{1F504}",E.textContent="Mengimport...",D.textContent=`Mengimport ${$.label}...`,S.style.width=`${Math.round(P/f.length*100)}%`;try{let F=await k(`/api/import/${$.module}`,{method:"POST",body:JSON.stringify({rows:$.mapped,onDuplicate:g})});if(F.ok){let K=F.data;C.inserted+=K.inserted||0,C.skipped+=K.skipped||0,C.moduleResults.push({label:$.label,inserted:K.inserted||0,skipped:K.skipped||0,status:"ok"}),L.textContent="\u2705",E.innerHTML=`<span class="badge badge-success">${K.inserted||0} berhasil</span>${K.skipped>0?` <span class="badge badge-neutral">${K.skipped} skip</span>`:""}`}else C.failed++,C.moduleResults.push({label:$.label,inserted:0,skipped:0,status:"error",error:F.data?.error}),L.textContent="\u274C",E.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(F){C.failed++,C.moduleResults.push({label:$.label,inserted:0,skipped:0,status:"error",error:F.message}),L.textContent="\u274C",E.innerHTML='<span class="badge badge-danger">Gagal</span>'}await et(150)}S.style.width="100%",D.textContent="Selesai!",await et(400),y(C)}function y(g){d("summary");let f=((Date.now()-a)/1e3).toFixed(1),v=g.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${v?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${v?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${g.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${g.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${g.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${g.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${g.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${g.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
      <div class="summary-stat-card info">
        <div class="stat-value">${f}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${g.moduleResults.map(S=>`
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,i=null,n.value="",document.getElementById("file-info").style.display="none",s.style.display="",d("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function et(t){return new Promise(e=>setTimeout(e,t))}j();var ot=[],ua=[];async function ma(t){ot=await Q(),ua=await ie(),H({container:t,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:ot}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await k(`/api/sp?limit=10000&${i}`);if(a.ok){let o=a.data.data.map(r=>({"Nama Karyawan":r.employee_name||"",Divisi:r.division||"",Cabang:r.branch_name||"","Tanggal Sp":r.tanggal||"","Akhir Sp":r.akhir_sp||"","Jenis Sp":r.sp_type||"","Link Document / Foto":r.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(G(),ce));d(o,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(G(),ce));i(e,"Template_Import_SP")},onImport:async e=>{let i=r=>{if(!r)return null;let n=String(r||"").toLowerCase(),s=ot.find(l=>String(l.label||"").toLowerCase()===n);return s?s.value:null},a=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let n=String(r).trim();if(/^\d{4,5}$/.test(n)){let l=Number(n);if(l>2e4&&l<99999){let b=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(n))return n.slice(0,10);let s=n.split(/[\/\-\.]/);if(s.length===3){let[l,b,c]=s.map(u=>u.trim());if(l.length===4&&b.length<=2&&c.length<=2)return`${l}-${b.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&b.length<=2&&l.length<=2)return`${c}-${b.padStart(2,"0")}-${l.padStart(2,"0")}`}return n},o=e.map(r=>({employee_name:String(r["Nama Karyawan"]||"").trim(),division:String(r.Divisi||"").trim(),branch_id:i(String(r.Cabang||"").trim()),tanggal:a(r["Tanggal Sp"]),akhir_sp:a(r["Akhir Sp"]),sp_type:String(r["Jenis Sp"]||"").trim(),document_link:String(r["Link Document / Foto"]||"").trim()})).filter(r=>r.employee_name&&r.branch_id),d=await k("/api/import/sp",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:[{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:ua},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:ot,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}j();var Je=[],ga=[];async function ba(t){Je=await Q(),ga=await ie(),H({container:t,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:Je},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:Je}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await k(`/api/mutasi?limit=10000&${i}`);if(a.ok){let o=a.data.data.map(r=>({Tanggal:r.tanggal||"","Nama Karyawan":r.employee_name||"","Cabang Asal":r.from_branch_name||"","Cabang Tujuan":r.to_branch_name||"",Status:r.status||"",Dokumen:r.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(G(),ce));d(o,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(G(),ce));i(e,"Template_Import_Mutasi")},onImport:async e=>{let i=r=>{if(!r)return null;let n=String(r||"").toLowerCase(),s=Je.find(l=>String(l.label||"").toLowerCase()===n);return s?s.value:null},a=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let n=String(r).trim();if(/^\d{4,5}$/.test(n)){let l=Number(n);if(l>2e4&&l<99999){let b=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(n))return n.slice(0,10);let s=n.split(/[\/\-\.]/);if(s.length===3){let[l,b,c]=s.map(u=>u.trim());if(l.length===4&&b.length<=2&&c.length<=2)return`${l}-${b.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&b.length<=2&&l.length<=2)return`${c}-${b.padStart(2,"0")}-${l.padStart(2,"0")}`}return n},o=e.map(r=>({tanggal:a(r.Tanggal),employee_name:String(r["Nama Karyawan"]||"").trim(),from_branch_id:i(String(r["Cabang Asal"]||"").trim()),to_branch_id:i(String(r["Cabang Tujuan"]||"").trim()),status:String(r.Status||"").trim(),document_link:String(r.Dokumen||"").trim()})).filter(r=>r.tanggal&&r.employee_name&&r.from_branch_id&&r.to_branch_id),d=await k("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:ga},{type:"select",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Je,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Je,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}j();async function ha(t){t.innerHTML=`
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
  `;let e=document.getElementById("btnRefreshSync"),i=document.getElementById("queueStatusFilter");e.addEventListener("click",o),i.addEventListener("change",n),document.getElementById("btnRetryAll").addEventListener("click",()=>a("retry",{allFailed:!0})),document.getElementById("btnResetStuck").addEventListener("click",()=>a("reset-stuck",{})),document.getElementById("btnPauseSync").addEventListener("click",()=>a("pause",{})),document.getElementById("btnResumeSync").addEventListener("click",()=>a("resume",{})),document.getElementById("btnReconcile").addEventListener("click",()=>a("reconcile",{module:"ALL",repairMode:!0})),document.getElementById("btnRetrySelected").addEventListener("click",()=>{let l=Array.from(document.querySelectorAll(".chk-queue:checked")).map(b=>b.value);if(l.length===0)return alert("No items selected");a("retry",{ids:l})}),document.getElementById("chkAllQueue").addEventListener("change",l=>{document.querySelectorAll(".chk-queue").forEach(b=>b.checked=l.target.checked)});async function a(l,b){if(confirm(`Are you sure you want to execute action: ${l}?`)){showLoading();try{let c=await k(`/api/sync/actions/${l}`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(b)});c.ok?(alert(c.data?.message||"Success"),o()):z(c.error||"Action failed")}catch(c){z(c.message)}hideLoading()}}await o();async function o(){showLoading(),await Promise.all([r(),n(),d(),s()]),hideLoading()}async function d(){try{let l=await k("/api/sync/performance");if(!l.ok)return;let{webhook:b,google_api:c,d1:u,queue:m,throughput:y}=l.data;document.getElementById("latencyHistograms").innerHTML=`
        <table class="w-full text-left">
          <thead><tr class="text-gray-500 border-b"><th>Metric</th><th>P50</th><th>P95</th><th>P99</th><th>Max</th></tr></thead>
          <tbody class="divide-y">
            <tr><td class="py-1">Webhook</td><td>${b.P50}ms</td><td>${b.P95}ms</td><td>${b.P99}ms</td><td>${b.Max}ms</td></tr>
            <tr><td class="py-1">Google API</td><td>${c.P50}ms</td><td>${c.P95}ms</td><td>${c.P99}ms</td><td>${c.Max}ms</td></tr>
            <tr><td class="py-1">D1 Execute</td><td>${u.P50}ms</td><td>${u.P95}ms</td><td>${u.P99}ms</td><td>${u.Max}ms</td></tr>
            <tr><td class="py-1">Queue Delay</td><td>${m.P50}ms</td><td>${m.P95}ms</td><td>${m.P99}ms</td><td>${m.Max}ms</td></tr>
          </tbody>
        </table>
        <div class="mt-4 pt-3 border-t text-gray-600 flex justify-between">
          <span>Throughput: <b>${y.events_per_sec}</b> ev/sec</span>
          <span><b>${y.events_per_min}</b> ev/min</span>
        </div>
      `,document.getElementById("lblSnapshotStatus").innerHTML='Checked 10 modules. Status: <b>COMPLETED</b><br><span class="text-xs text-gray-500">Run today at 02:00</span>'}catch(l){console.error(l)}}async function r(){try{let l=await k("/api/sync/health");if(!l.ok)return z("Failed to fetch sync health");let{status:b,queue:c,circuit_breaker:u}=l.data,m=`
        <div class="bg-white p-4 rounded-lg shadow border-l-4 ${b==="HEALTHY"?"border-green-500":b==="WARNING"?"border-yellow-500":"border-red-500"}">
          <p class="text-sm text-gray-500">System Health</p>
          <p class="text-2xl font-bold ${b==="HEALTHY"?"text-green-600":b==="WARNING"?"text-yellow-600":"text-red-600"}">${b}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-blue-500">
          <p class="text-sm text-gray-500">Pending Queue</p>
          <p class="text-2xl font-bold text-gray-800">${c.pending||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-orange-500">
          <p class="text-sm text-gray-500">Failed / Retrying</p>
          <p class="text-2xl font-bold text-gray-800">${c.failed||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-red-500">
          <p class="text-sm text-gray-500">Dead Letters</p>
          <p class="text-2xl font-bold text-gray-800">${c.dead_letter||0}</p>
        </div>
      `;document.getElementById("syncOverviewCards").innerHTML=m;let y=document.getElementById("cbStateBadge"),g=document.getElementById("cbStateDesc"),f=document.getElementById("cbStatusCard");f.className="bg-white rounded-lg shadow p-6 border-l-4",u==="CLOSED"?(y.className="px-3 py-1 rounded-full text-sm font-medium bg-green-100 text-green-800",y.textContent="CLOSED",g.textContent="Traffic is flowing normally to Google Sheets.",f.classList.add("border-green-500")):u==="OPEN"?(y.className="px-3 py-1 rounded-full text-sm font-medium bg-red-100 text-red-800",y.textContent="OPEN",g.textContent="Failing fast. Traffic to Google Sheets is paused due to repeated failures.",f.classList.add("border-red-500")):u==="HALF_OPEN"?(y.className="px-3 py-1 rounded-full text-sm font-medium bg-yellow-100 text-yellow-800",y.textContent="HALF-OPEN",g.textContent="Testing recovery. Permitting limited traffic to verify stability.",f.classList.add("border-yellow-500")):y.textContent=u||"UNKNOWN"}catch(l){console.error(l)}}async function n(){try{let l=document.getElementById("queueStatusFilter").value,b=await k("/api/sync/queue?limit=15"+(l?"&status="+l:""));if(!b.ok)return;let c=document.getElementById("queueTableBody"),u=b.data?.data||b.data||[];if(u.length===0){c.innerHTML='<tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">No events found</td></tr>';return}c.innerHTML=u.map(m=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2"><input type="checkbox" class="chk-queue" value="${m.id}" /></td>
          <td class="px-4 py-2 font-mono text-xs text-gray-500" title="${m.id}">${m.id.split("-")[0]}...</td>
          <td class="px-4 py-2 font-medium">${m.entity_name}</td>
          <td class="px-4 py-2">
             <span class="px-2 py-0.5 rounded text-xs ${m.action==="INSERT"?"bg-blue-100 text-blue-800":m.action==="UPDATE"?"bg-purple-100 text-purple-800":"bg-red-100 text-red-800"}">${m.action}</span>
          </td>
          <td class="px-4 py-2">
             <span class="px-2 py-0.5 rounded text-xs ${m.status==="PENDING"?"bg-yellow-100 text-yellow-800":m.status==="PROCESSING"?"bg-blue-100 text-blue-800":m.status==="DEAD_LETTER"?"bg-red-100 text-red-800":"bg-gray-100 text-gray-800"}">${m.status}</span>
             ${m.last_error?`<br><span class="text-xs text-red-500 max-w-xs block truncate" title="${m.last_error}">${m.last_error}</span>`:""}
          </td>
          <td class="px-4 py-2 text-gray-600">${m.retry_count||0}</td>
          <td class="px-4 py-2 text-gray-500 whitespace-nowrap">${window.formatDate(m.created_at)} ${new Date(m.created_at).toLocaleTimeString("id-ID")}</td>
        </tr>
      `).join("")}catch(l){console.error(l)}}async function s(){try{let l=await k("/api/sync/metrics");if(!l.ok)return;let b=document.getElementById("metricsTableBody"),c=l.data||[];if(c.length===0){b.innerHTML='<tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">No metrics found in last 24h</td></tr>';return}b.innerHTML=c.map(u=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2 font-medium text-gray-800">${u.module}</td>
          <td class="px-4 py-2 text-gray-600">${u.total_events}</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(u.avg_webhook_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(u.avg_d1_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(u.avg_queue_wait_ms||0)} ms</td>
        </tr>
      `).join("")}catch(l){console.error(l)}}}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let e=Math.floor(Number(t)-25569);return new Date(e*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let e=t.split(/[\/\-]/);return`${e[2]}-${e[1]}-${e[0]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let i=e.split("-");if(i.length===3&&i[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],o=parseInt(i[2],10),d=a[parseInt(i[1],10)-1];return`${o} ${d} ${i[0]}`}return e};function te(t){return async e=>{if(!Qe()){Be("/login");return}return t(e)}}var tt=null;function Qa(){tt&&clearInterval(tt);let t=()=>{let e=new Date,i=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),o=document.getElementById("header-clock-time"),d=document.getElementById("header-clock-date");o&&(o.textContent=i),d&&(d.textContent=a)};t(),tt=setInterval(t,1e3)}async function Wa(){try{let t=await k("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},i=(a,o)=>{let d=document.getElementById(a);d&&(d.textContent=o>0?o:"",d.style.display=o>0?"inline-flex":"none")};i("badge-issues",e.issues?.current||0),i("badge-contracts",e.expiring30?.current||0),i("badge-oo1",e.one_on_one?.current||0),i("badge-schedule",e.schedule?.current||0),i("badge-supply",e.supply?.current||0)}catch{}}var Ue=[];async function Va(){try{let t=await k("/api/dashboard/notifications");if(!t.ok)return;Ue=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Ue.length>0?"block":"none",e.textContent=Ue.length)}catch{}}function Ya(){if(!Ue.length){de({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,i)=>i()});return}let t=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${Ue.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;de({title:`Notifikasi (${Ue.length})`,content:t,confirmText:"Tutup",onConfirm:(e,i)=>i()})}function ya(){let t=oe(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
                <span class="topbar-greeting-time">${(()=>{let l=new Date().getHours();return l>=4&&l<11?"Selamat Pagi":l>=11&&l<15?"Selamat Siang":l>=15&&l<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${t?.full_name||"Admin"}</span> \u{1F44B}
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
  `;let i=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),o=document.getElementById("topbar-menu-btn"),d=document.getElementById("sidebar-close"),r=()=>{i.classList.add("open"),a.classList.add("show")},n=()=>{i.classList.remove("open"),a.classList.remove("show")};o?.addEventListener("click",r),d?.addEventListener("click",n),a?.addEventListener("click",n),document.querySelectorAll(".nav-item").forEach(l=>l.addEventListener("click",n));function s(){let l=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(u=>{let m=u.dataset.route;u.classList.toggle("active",l===m||m!=="/dashboard"&&l.startsWith(m))});let b=document.getElementById("topbar-title"),c=document.querySelector(".nav-item.active .nav-label");b&&c&&(b.textContent=c.textContent)}window.addEventListener("hashchange",s),s(),Qa(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await k("/api/auth/logout",{method:"POST"}),We(),tt&&clearInterval(tt),Be("/login")}),Wa(),Va(),document.getElementById("btn-notif")?.addEventListener("click",l=>{l.preventDefault(),Ya()})}document.addEventListener("keydown",t=>{if(t.target.tagName==="SELECT"&&t.key.length===1&&/[a-zA-Z]/.test(t.key)){let e=t.key.toLowerCase(),i=Array.from(t.target.options);if(i.length===0)return;let a=t.target.selectedIndex+1;(a>=i.length||a<0)&&(a=0);let o=-1;for(let d=0;d<i.length;d++){let r=(a+d)%i.length;if(i[r].text.toLowerCase().replace(/^[\d\.\s\-]+/,"").startsWith(e)){o=r;break}}o!==-1&&o!==t.target.selectedIndex&&(t.target.selectedIndex=o,t.preventDefault(),t.target.dispatchEvent(new Event("change")))}});async function Xa(){Y("/login",({main:e})=>Kt(e)),Y("/dashboard",te(({main:e})=>Nt(e))),Y("/calendar",te(({main:e})=>ra(e))),Y("/employees",te(({main:e,params:i})=>jt(e,i))),Y("/contracts",te(({main:e,params:i})=>qt(e,i))),Y("/sp",te(({main:e})=>ma(e))),Y("/mutasi",te(({main:e})=>ba(e))),Y("/sync-dashboard",te(({main:e})=>ha(e))),Y("/timeline",te(({main:e,params:i})=>Jt(e,i))),Y("/issues",te(({main:e,params:i})=>Ut(e,i))),Y("/one-on-one",te(({main:e,params:i})=>Gt(e,i))),Y("/training",te(({main:e})=>zt(e))),Y("/relievers",te(({main:e,params:i})=>Wt(e,i))),Y("/reports/inspection",te(({main:e})=>Vt(e))),Y("/reports/cleaning",te(({main:e})=>Yt(e))),Y("/reports/fogging",te(({main:e})=>Xt(e))),Y("/reports/basecamp",te(({main:e})=>Zt(e))),Y("/reports/supply",te(({main:e})=>Ct(e,"supply"))),Y("/sop",te(({main:e})=>ea(e))),Y("/checklist",te(({main:e})=>ta(e))),Y("/forms",te(({main:e})=>Ct(e))),Y("/users",te(({main:e})=>na(e))),Y("/branches",te(({main:e})=>ia(e))),Y("/profile",te(({main:e})=>la(e))),Y("/settings/import",te(({main:e})=>pa(e))),Y("/audit-logs",te(({main:e})=>aa(e)));let t=Qe();if(!t&&window.location.hash!=="#/login"&&Be("/login"),t){let e=await k("/api/auth/me");e.ok?(Ve(e.data.data),ya()):(We(),Be("/login"))}window.addEventListener("fm:login",()=>{ya(),Be("/dashboard")}),Tt()}Xa();
