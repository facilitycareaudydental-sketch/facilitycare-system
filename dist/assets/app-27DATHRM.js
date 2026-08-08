var fa=Object.defineProperty;var dt=(t,e)=>()=>(t&&(e=t(t=0)),e);var ct=(t,e)=>{for(var r in e)fa(t,r,{get:e[r],enumerable:!0})};var Le={};ct(Le,{API:()=>Et,CLIENT_SIDE_MAX_ROWS:()=>Se,IS_DEVELOPMENT:()=>$e,apiFetch:()=>S,clearToken:()=>We,getToken:()=>Qe,getUser:()=>ce,setToken:()=>pt,setUser:()=>Ve});function Qe(){return localStorage.getItem("fm_token")}function pt(t){localStorage.setItem("fm_token",t)}function We(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ce(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Ve(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function S(t,e={}){let r=Qe(),a={"Content-Type":"application/json",...r?{Authorization:`Bearer ${r}`}:{},...e.headers||{}};try{let l=`cb=${Date.now()}`,s=t.includes("?")?"&":"?",o=`${Et}${t}${s}${l}`,n=await fetch(o,{...e,headers:a}),i;try{let d=await n.text();try{i=JSON.parse(d)}catch{i={error:`Server Error (${n.status}): ${d.substring(0,80)}...`}}}catch{i={error:"Gagal membaca respon dari server"}}return n.status===401&&(We(),window.location.hash="#/login"),{ok:n.ok,status:n.status,data:i}}catch(l){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${l.message})`}}}}var $e,va,Et,Se,j=dt(()=>{$e=!1,va="https://fm-operations-api.facilitycare-audydental.workers.dev",Et=va,Se=1e4});var mt={};ct(mt,{confirmDialog:()=>Ee,createModal:()=>pe});function pe({title:t,content:e,onConfirm:r,onCancel:a,confirmText:l="Simpan",cancelText:s="Batal",size:o="md",confirmClass:n="btn-primary"}){let i={sm:"400px",md:"560px",lg:"720px",xl:"900px"},d=document.createElement("div");d.className="modal-overlay",d.innerHTML=`
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
  `,e instanceof HTMLElement&&d.querySelector(".modal-body").appendChild(e);let h=()=>{d.classList.remove("show"),setTimeout(()=>d.remove(),250)};return d.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),h()}),d.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),h()}),r&&d.querySelector(".modal-confirm").addEventListener("click",()=>r(d,h)),d.addEventListener("click",u=>{u.target===d&&(a&&a(),h())}),document.body.appendChild(d),requestAnimationFrame(()=>d.classList.add("show")),{overlay:d,close:h}}function Ee(t,e,r="Konfirmasi"){return pe({title:r,content:`<p>${t}</p>`,onConfirm:(a,l)=>{e(),l()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var we=dt(()=>{});var ue={};ct(ue,{downloadExcel:()=>K,parseExcel:()=>De,renderExcelButtons:()=>Ie});function De(t){return new Promise((e,r)=>{let a=new FileReader;a.onload=l=>{try{let s=new Uint8Array(l.target.result),o=XLSX.read(s,{type:"array"}),n=o.SheetNames[0],i=o.Sheets[n];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${o.SheetNames.join(", ")}`),console.log(`Sheet Used: ${n}`);let d=XLSX.utils.decode_range(i["!ref"]||"A1:A1"),h=d.e.r-d.s.r+1,u=d.e.c-d.s.c+1;console.log(`Total Rows (including empty): ${h}`),console.log(`Total Columns: ${u}`);let c=[];for(let f=d.s.c;f<=d.e.c;++f){let g=i[XLSX.utils.encode_cell({c:f,r:d.s.r})];g&&g.v&&c.push(g.v)}console.log(`Headers Found: ${c.join(", ")}`),console.log("---------------------------");let m=XLSX.utils.sheet_to_json(i,{defval:""});Object.defineProperty(m,"__worksheet",{value:i,enumerable:!1}),Object.defineProperty(m,"__headers",{value:c,enumerable:!1}),e(m)}catch(s){r(s)}},a.onerror=l=>r(l),a.readAsArrayBuffer(t)})}function K(t,e){try{let r=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,r,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(r){throw console.error("Error generating Excel file:",r),r}}function Ie(t){return`
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
  `}var G=dt(()=>{});j();var ut={},at=null;function Y(t,e){ut[t]=e}function Pe(t){window.location.hash=t}function Tt(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[r,...a]=e.split("?"),l=ut[r];if(!l){for(let[o,n]of Object.entries(ut))if(o.endsWith("/*")&&r.startsWith(o.slice(0,-2))){l=n;break}}at&&(at(),at=null);let s=document.getElementById("main-content");if(s&&(s.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),l){let o=new URLSearchParams(a.join("?")),n=r.split("/").filter(Boolean),i=await l({path:r,params:o,segments:n,main:s});i&&(at=i)}else{let o=s||document.getElementById("app");o&&(o.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var Ye;function ka(){return Ye||(Ye=document.createElement("div"),Ye.id="toast-container",document.body.appendChild(Ye)),Ye}function Dt(t,e="info",r=3500){let a=ka(),l=document.createElement("div");l.className=`toast toast-${e}`;let s={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};l.innerHTML=`<span class="toast-icon">${s[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(l),requestAnimationFrame(()=>l.classList.add("show")),setTimeout(()=>{l.classList.remove("show"),setTimeout(()=>l.remove(),350)},r)}var ne=t=>Dt(t,"success"),z=t=>Dt(t,"error");we();j();j();we();function Be({columns:t,data:e,onEdit:r,onDelete:a,onView:l,actions:s=[],emptyText:o="Tidak ada data",bulkSelect:n=null}){let i=document.createElement("div");if(i.className="table-wrapper",!e||e.length===0)return i.innerHTML=`<div class="empty-state"><p>${o}</p></div>`,i;let d=document.createElement("table");d.className="data-table";let h=document.createElement("thead"),u=document.createElement("tr");if(n){let m=document.createElement("th");m.style.width="40px",m.style.textAlign="center";let f=document.createElement("input");f.type="checkbox",f.id="select-all-checkbox",f.title="Pilih semua",f.addEventListener("change",()=>{e.forEach(g=>{f.checked?n.selectedIds.add(g.id):n.selectedIds.delete(g.id)}),i.querySelectorAll(".row-checkbox").forEach(g=>g.checked=f.checked),n.onToggle()}),m.appendChild(f),u.appendChild(m)}if(t.forEach(m=>{let f=document.createElement("th");f.textContent=m.label,m.width&&(f.style.width=m.width),u.appendChild(f)}),r||a||l||s.length>0){let m=document.createElement("th");m.textContent="Aksi",m.style.width="120px",u.appendChild(m)}h.appendChild(u),d.appendChild(h);let c=document.createElement("tbody");return e.forEach(m=>{let f=document.createElement("tr");if(n){let g=document.createElement("td");g.style.textAlign="center",g.style.width="40px";let y=document.createElement("input");y.type="checkbox",y.className="row-checkbox",y.checked=n.selectedIds.has(m.id),y.addEventListener("change",()=>{if(y.checked)n.selectedIds.add(m.id);else{n.selectedIds.delete(m.id);let k=document.getElementById("select-all-checkbox");k&&(k.checked=!1)}n.onToggle()}),g.appendChild(y),f.appendChild(g)}if(t.forEach(g=>{let y=document.createElement("td");if(g.render){let k=g.render(m[g.key],m);k instanceof HTMLElement?y.appendChild(k):y.innerHTML=k||""}else y.textContent=m[g.key]!==null&&m[g.key]!==void 0&&m[g.key]!==""?m[g.key]:"";g.nowrap&&(y.style.whiteSpace="nowrap"),f.appendChild(y)}),r||a||l||s.length>0){let g=document.createElement("td");g.className="actions-cell";let y=document.createElement("div");if(y.className="btn-group",l){let k=document.createElement("button");k.className="btn btn-xs btn-ghost",k.innerHTML="\u{1F441}",k.title="Lihat",k.addEventListener("click",()=>l(m)),y.appendChild(k)}if(r){let k=document.createElement("button");k.className="btn btn-xs btn-secondary",k.innerHTML="\u270F\uFE0F",k.title="Edit",k.addEventListener("click",()=>r(m)),y.appendChild(k)}s.forEach(k=>{let v=document.createElement("button");v.className=`btn btn-xs ${k.class||"btn-ghost"}`,v.innerHTML=k.icon||k.label,v.title=k.label,v.addEventListener("click",()=>k.handler(m)),y.appendChild(v)}),g.appendChild(y),f.appendChild(g)}c.appendChild(f)}),d.appendChild(c),i.appendChild(d),i}function Ae({page:t,pages:e,total:r,limit:a,onPage:l}){if(e<=1)return null;let s=document.createElement("div");s.className="pagination";let o=document.createElement("span");o.className="pagination-info",o.textContent=`Total: ${r} data`,s.appendChild(o);let n=document.createElement("div");n.className="pagination-btns";let i=(u,c,m=!1,f=!1)=>{let g=document.createElement("button");g.className=`btn btn-sm ${f?"btn-primary":"btn-ghost"} pagination-btn`,g.textContent=u,g.disabled=m,g.addEventListener("click",()=>l(c)),n.appendChild(g)};i("\xAB",1,t===1),i("\u2039",t-1,t===1);let d=Math.max(1,t-2),h=Math.min(e,t+2);for(let u=d;u<=h;u++)i(u,u,!1,u===t);return i("\u203A",t+1,t===e),i("\xBB",e,t===e),s.appendChild(n),s}we();function Te(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${Te(e.fields)}</div>`;let r=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",l="";switch(e.type){case"textarea":l=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${r} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let o=(e.options||[]).map(u=>{let c=typeof u=="object"?u.value:u,m=typeof u=="object"?u.label:u,f=e.value==c?"selected":"";return`<option value="${c}" ${f}>${m}</option>`}).join("");l=`<select name="${e.name}" class="form-control" ${r}><option value="">-- Pilih ${e.label||""} --</option>${o}</select>`;break;case"combobox":let n=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,i=(e.options||[]).map(u=>{let c=typeof u=="object"?u.value:u,m=typeof u=="object"?u.label||u.value||"":u||"";return(m==="undefined"||m==="[object Object]"||m==="null")&&(m=""),m?`<option value="${m}"></option>`:""}).join(""),d=e.value||"";if(e.value){let u=(e.options||[]).find(c=>(typeof c=="object"?c.value:c)==e.value);if(u){let c=typeof u=="object"?u.label||u.value||"":u||"";c&&c!=="undefined"&&c!=="[object Object]"&&c!=="null"&&(d=c)}}l=`
          <input type="text" name="${e.name}" list="${n}" class="form-control" value="${d}" placeholder="Pilih atau ketik baru..." ${r} autocomplete="off">
          <datalist id="${n}">${i}</datalist>
        `;break;case"checkbox":l=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let h=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";l=`<input type="date" name="${e.name}" class="form-control" value="${h}" ${r}>`;break;case"number":l=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${r}>`;break;case"email":l=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${r}>`;break;case"url":l=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${r}>`;break;default:l=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${r} autocomplete="off">`}let s=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${l}${s}</div>`}).join("")}function Ne(t){let e={},r=new FormData(t);for(let[a,l]of r.entries())e[a]=l===""?null:l;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function Me(t,e){e&&Object.entries(e).forEach(([r,a])=>{let l=t.querySelector(`[name="${r}"]`);l&&(l.hasAttribute("list")||(l.type==="checkbox"?l.checked=!!a:l.type==="date"&&a&&window.parseFlexibleDate?l.value=window.parseFlexibleDate(a):l.value=a??""))})}G();function gt({container:t,title:e,icon:r,apiPath:a,columns:l,formFields:s,filterFields:o,defaultFilters:n={},enableMobileFilterSheet:i=!1,itemLabel:d="Data",canCreate:h=!0,canEdit:u=!0,canDelete:c=!0,onBeforeSubmit:m,onAfterLoad:f,onDataLoaded:g,extraActions:y=[],initialSearch:k="",exportOptions:v=null,bulkDelete:E=!1,paginationMode:A="server"}){let B=ce();B&&typeof B=="object"&&B.role==="viewer"&&(h=!1,u=!1,c=!1,E=!1,v=null);let T=1,x={...n};k&&(x.search=k);let $=new Set;t.innerHTML=`
    <div class="crud-layout-wrapper ${i?"mobile-active":""}">
      <div class="page-header">
      <h1 class="page-title">${r} ${e}</h1>
      <div class="page-actions">
        ${h?`<button class="btn btn-primary" id="btn-create">+ Tambah ${d}</button>`:""}
        ${v?'<button class="btn btn-outline" id="btn-mobile-aksi" style="display:none; align-items:center; justify-content:center; gap:0.25rem;">\u22EE Aksi</button>':""}
      </div>
    </div>

    ${E?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:none; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${v?`
    <div class="excel-actions-wrapper" id="excel-actions-wrapper">
      <div class="bottom-sheet-header aksi-header" style="display:none;">
        <h3 style="margin:0; font-size:1rem;">Aksi</h3>
        <button class="btn-close-sheet" id="btn-close-aksi-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
      </div>
      ${Ie(v.moduleName)}
    </div>`:""}

    ${o&&o.length>0?`
    <div class="filter-bar card ${i?"has-mobile-sheet":""}" style="padding: 1rem; display: flex; gap: 0.5rem; flex-wrap: wrap;">
        ${o.filter(p=>p.type==="search"||p.type==="search-combo").map(p=>{if(p.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${p.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}"></div>`;if(p.type==="search-combo"){let b="dl-filter-search",w=(p.options||[]).map(_=>`<option value="${typeof _=="object"?_.label:_}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${b}" class="form-control" autocomplete="off" placeholder="${p.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}"><datalist id="${b}">${w}</datalist></div>`}return""}).join("")}
        
        <div class="filter-options" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(p=>p.type!=="search"&&p.type!=="search-combo").map(p=>{if(p.type==="select")return`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${p.name}" id="filter-${p.name}"><option value="">Pilih ${p.label}</option>${(p.options||[]).map(b=>`<option value="${typeof b=="object"?b.value:b}" ${x[p.name]===(typeof b=="object"?b.value:b)?"selected":""}>${typeof b=="object"?b.label:b}</option>`).join("")}</select>`;if(p.type==="combobox"){let b=`dl-filter-${p.name}`,w=(p.options||[]).map(D=>{let C=typeof D=="object"?D.label||D.value||"":D||"";return(C==="undefined"||C==="[object Object]"||C==="null")&&(C=""),C?`<option value="${C}"></option>`:""}).join(""),_=x[p.name]||"";if(x[p.name]){let D=(p.options||[]).find(C=>String(typeof C=="object"?C.value:C)==String(x[p.name]));if(D){let C=typeof D=="object"?D.label||D.value||"":D||"";C&&C!=="undefined"&&C!=="[object Object]"&&C!=="null"&&(_=C)}}return`<div class="filter-combobox" style="flex:1; min-width:120px;">
                <input type="text" name="${p.name}" id="filter-${p.name}" list="${b}" class="form-control filter-combobox-input" value="${_}" placeholder="Pilih ${p.label}..." autocomplete="off">
                <datalist id="${b}">${w}</datalist>
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
  `;function F(){let p=document.getElementById("bulk-toolbar");if(!p)return;let b=document.getElementById("bulk-count"),w=document.getElementById("btn-bulk-delete"),_=document.getElementById("btn-bulk-cancel");b.textContent=`${$.size} item dipilih`,$.size>0?(p.style.display="flex",p.classList.add("has-items"),w.disabled=!1,_.disabled=!1,window.innerWidth<=768?(p.style.setProperty("top","auto","important"),p.style.setProperty("bottom","0","important")):(p.style.removeProperty("top"),p.style.removeProperty("bottom"))):(p.style.display="none",p.classList.remove("has-items"),w.disabled=!0,_.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{$.clear(),document.querySelectorAll(".row-checkbox").forEach(b=>b.checked=!1);let p=document.getElementById("select-all-checkbox");p&&(p.checked=!1),F()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if($.size===0)return;let p=[...$],b=document.createElement("div");b.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",b.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${p.length} ${d}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${p.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(b),b.querySelector("#bulk-cancel-btn").addEventListener("click",()=>b.remove()),b.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let w=b.querySelector("#bulk-confirm-btn");w.disabled=!0,w.textContent="Menghapus...";let _=await S(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:p})});b.remove(),_.ok?(ne(`${p.length} ${d} berhasil dihapus.`),$.clear(),F(),q()):z(_.data?.error||"Gagal menghapus data.")})});let O=document.getElementById("filter-search"),re;O?.addEventListener("input",p=>{clearTimeout(re),re=setTimeout(()=>{x.search=p.target.value,T=1,$.clear(),q()},400)}),o?.forEach(p=>{p.type==="select"&&document.getElementById(`filter-${p.name}`)?.addEventListener("change",b=>{x[p.name]=b.target.value,T=1,$.clear(),q()}),p.type==="combobox"&&document.getElementById(`filter-${p.name}`)?.addEventListener("change",b=>{let w=b.target.value,_=(p.options||[]).find(D=>{let C=String(typeof D=="object"?D.value:D),L=String(typeof D=="object"?D.label:D);return C===w||L===w});w?x[p.name]=_?typeof _=="object"?_.value:_:w:x[p.name]="",T=1,$.clear(),q()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{x={...n},O&&(O.value=""),o?.forEach(p=>{let b=document.getElementById(`filter-${p.name}`);b&&(b.value="")}),T=1,$.clear(),q()});let de=document.getElementById("btn-mobile-filter"),ie=document.getElementById("filter-options-wrapper"),ge=document.getElementById("btn-close-filter-sheet");if(de&&ie){de.addEventListener("click",b=>{b.preventDefault(),ie.classList.add("sheet-open")}),ge&&ge.addEventListener("click",b=>{b.preventDefault(),ie.classList.remove("sheet-open")});let p=document.getElementById("btn-apply-filter");p&&p.addEventListener("click",b=>{b.preventDefault(),ie.classList.remove("sheet-open")})}let me=document.getElementById("btn-mobile-aksi"),be=document.getElementById("excel-actions-wrapper"),_e=document.getElementById("btn-close-aksi-sheet");if(me&&be&&(me.addEventListener("click",p=>{p.preventDefault(),be.classList.add("sheet-open")}),_e&&_e.addEventListener("click",p=>{p.preventDefault(),be.classList.remove("sheet-open")})),document.getElementById("btn-create")?.addEventListener("click",()=>Ce(null)),v){document.getElementById(`btn-export-${v.moduleName}`)?.addEventListener("click",async b=>{let w=b.target,_=w.innerHTML;w.innerHTML="\u23F3 Loading...",w.disabled=!0;try{await v.onExport()}catch{z("Gagal export data")}finally{w.innerHTML=_,w.disabled=!1}}),document.getElementById(`btn-template-${v.moduleName}`)?.addEventListener("click",()=>{v.onTemplate()});let p=document.getElementById(`input-import-${v.moduleName}`);p?.addEventListener("change",async b=>{let w=b.target.files[0];if(!w)return;p.disabled=!0;let _=document.createElement("div");_.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",_.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(_);let D=_.querySelector("#import-progress-text"),C=_.querySelector("#import-progress-bar"),L=_.querySelector("#import-summary"),P=_.querySelector("#import-close-btn");P.addEventListener("click",()=>{_.remove(),q()});try{let W=await De(w);if(W.length===0)throw new Error("File kosong atau format salah");let V=500,ae=0,Z=0,N=0,M=W.length;D.textContent=`Ditemukan ${M} baris data. Memulai import...`;for(let I=0;I<M;I+=V){let J=W.slice(I,I+V);D.textContent=`Mengimport baris ${I+1} - ${Math.min(I+V,M)} dari ${M}...`,C.style.width=`${Math.round(I/M*100)}%`;try{let R=await v.onImport(J);R?(ae+=R.inserted||R.metrics?.inserted||J.length,Z+=R.skipped||R.metrics?.updated||0):ae+=J.length}catch(R){console.error("Chunk import failed:",R),N+=J.length}}C.style.width="100%",D.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',L.style.display="block",L.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${M}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${Z}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${N}</strong></div>
        `,N>0&&(L.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),P.style.display="block",p.value=""}catch(W){D.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${W.message}`,C.style.background="var(--danger)",C.style.width="100%",P.style.display="block",p.value=""}finally{p.disabled=!1}})}async function q(){F();let p=document.getElementById("table-container");if(!p)return;p.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let b=A==="client",w=b?1:T,_=b?Se:20,D=new URLSearchParams({page:w,limit:_,...Object.fromEntries(Object.entries(x).filter(([,N])=>N))}),C=await S(`${a}?${D}`);if(!C.ok){p.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${C.data?.error||"Error"}</p></div>`;return}let L=C.data?.data||C.data||[],P=C.data?.pagination,W=L.length,V=L;if(b){L=g(L),V=L;let N=L.length,M=20,I=Math.ceil(N/M);T>I&&I>0&&(T=I);let J=(T-1)*M,R=T*M;L=L.slice(J,R),P={page:T,limit:M,total:N,pages:I}}!1,f&&f(L);let ae=Be({columns:l,data:L,fullData:V,onEdit:u?N=>Ce(N):null,actions:y.map(N=>({...N,handler:M=>N.handler(M,q)})),emptyText:`Tidak ada ${String(d||"").toLowerCase()}`,bulkSelect:E?{selectedIds:$,onToggle:F}:null});p.innerHTML="",p.appendChild(ae);let Z=document.getElementById("pagination-container");if(Z&&(Z.innerHTML="",P&&P.pages>1)){let N=Ae({page:P.page,pages:P.pages,total:P.total,limit:P.limit,onPage:M=>{T=M,q()}});N&&Z.appendChild(N)}}function Ge(p){let b=typeof s=="function"?s(p):s;return Te(b)}function Ce(p){let b=!!p,w=document.createElement("form");if(w.noValidate=!0,w.innerHTML=Ge(p),b){let D=typeof s=="function"?s(p):s;Me(w,p)}let{close:_}=pe({title:b?`Edit ${d}`:`Tambah ${d}`,content:w,size:"lg",confirmText:b?"Simpan Perubahan":`Tambah ${d}`,onConfirm:async(D,C)=>{if(!w.reportValidity())return;let L=D.querySelector(".modal-confirm");L.disabled=!0,L.textContent="Menyimpan...";let P=Ne(w),W=typeof s=="function"?s(p):s,V=async M=>{for(let I of M)if(I.type==="row")await V(I.fields);else if(I.type==="combobox"&&P[I.name]){let J=P[I.name],R=(I.options||[]).find(U=>{let ee=String(typeof U=="object"?U.value:U),ze=String(typeof U=="object"?U.label:U);return ee===J||ze===J});if(R)P[I.name]=typeof R=="object"?R.value:R;else if(I.createApi){let U={};U[I.createApi.field]=J,I.createApi.extra&&Object.assign(U,I.createApi.extra);let ee=await S(I.createApi.path,{method:"POST",body:JSON.stringify(U)});if(ee.ok&&ee.data?.id)P[I.name]=ee.data.id;else if(ee.ok&&!ee.data?.id)P[I.name]=J;else throw new Error(`Gagal membuat master data: ${ee.data?.error||"Unknown error"}`)}}};try{await V(W)}catch(M){z(M.message),L.disabled=!1,L.textContent=b?"Simpan Perubahan":`Tambah ${d}`;return}m&&(P=await m(P,p));let ae=b?"PUT":"POST",Z=b?`${a}/${p.id}`:a,N=await S(Z,{method:ae,body:JSON.stringify(P)});N.ok?(ne(b?`${d} berhasil diperbarui.`:`${d} berhasil ditambahkan.`),C(),q()):(z(N.data?.error||"Gagal menyimpan data."),L.disabled=!1,L.textContent=b?"Simpan Perubahan":`Tambah ${d}`)}})}function $t(p){Ee(`Hapus ${d} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let b=await S(`${a}/${p.id}`,{method:"DELETE"});b.ok?(ne(`${d} berhasil dihapus.`),q()):z(b.data?.error||"Gagal menghapus.")},`Hapus ${d}`)}return q(),q}j();j();var Fe=null,nt=null;async function Oe(t=!1){if(Fe&&!t)return console.log("Employees Raw (Cache Hit)",Fe.slice(0,5)),Fe;let e=await S(`/api/employees?limit=${Se}&status=Aktif`);return Fe=(e.data?.data||[]).map(r=>({value:r.id,label:r.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",Fe.slice(0,5)),Fe}async function le(t=!1){let r=(await Oe(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",r.slice(0,5)),r}async function Q(t=!1){return nt&&!t||(nt=((await S("/api/branches?all=1")).data?.data||[]).map(r=>({value:r.id,label:r.full_name}))),nt}function X(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function bt(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function Re(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function ht(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function he(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}G();function yt(t,e){let r=new Date,a=`${r.getFullYear()}-${String(r.getMonth()+1).padStart(2,"0")}`;if(!(t.target_date||t.opening_date||"").startsWith(a))return!1;let s=String(t.status||"").toLowerCase();if(s!=="selesai"&&s!=="completed"&&s!=="done")return!1;let o=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?o.includes("inspeksi"):e==="gcdc"?o.includes("general cleaning")||o.includes("deep cleaning"):!1}j();we();G();function H({container:t,title:e,icon:r,apiPath:a,columns:l,formFields:s,filterFields:o,defaultFilters:n={},enableMobileFilterSheet:i=!1,itemLabel:d="Data",canCreate:h=!0,canEdit:u=!0,canDelete:c=!0,onBeforeSubmit:m,onAfterLoad:f,onDataLoaded:g,extraActions:y=[],initialSearch:k="",exportOptions:v=null,bulkDelete:E=!1,paginationMode:A="server"}){let B=ce();B&&typeof B=="object"&&B.role==="viewer"&&(h=!1,u=!1,c=!1,E=!1,v=null);let T=1,x={...n};k&&(x.search=k);let $=new Set;t.innerHTML=`
    <div class="crud-layout-wrapper ${i?"mobile-active":""}">
      <div class="page-header">
      <h1 class="page-title">${r} ${e}</h1>
      <div class="page-actions">
        ${h?`<button class="btn btn-primary" id="btn-create">+ Tambah ${d}</button>`:""}
        ${v?'<button class="btn btn-outline" id="btn-mobile-aksi" style="display:none; align-items:center; justify-content:center; gap:0.25rem;">\u22EE Aksi</button>':""}
      </div>
    </div>

    ${E?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:none; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${v?`
    <div class="excel-actions-wrapper" id="excel-actions-wrapper">
      <div class="bottom-sheet-header aksi-header" style="display:none;">
        <h3 style="margin:0; font-size:1rem;">Aksi</h3>
        <button class="btn-close-sheet" id="btn-close-aksi-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
      </div>
      ${Ie(v.moduleName)}
    </div>`:""}

    ${o&&o.length>0?`
    <div class="filter-bar card ${i?"has-mobile-sheet":""}" style="padding: 1rem; display: flex; gap: 0.5rem; flex-wrap: wrap;">
        ${o.filter(p=>p.type==="search"||p.type==="search-combo").map(p=>{if(p.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${p.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}"></div>`;if(p.type==="search-combo"){let b="dl-filter-search",w=(p.options||[]).map(_=>`<option value="${typeof _=="object"?_.label:_}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${b}" class="form-control" autocomplete="off" placeholder="${p.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}"><datalist id="${b}">${w}</datalist></div>`}return""}).join("")}
        
        <div class="filter-options" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(p=>p.type!=="search"&&p.type!=="search-combo").map(p=>{if(p.type==="select")return`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${p.name}" id="filter-${p.name}"><option value="">Pilih ${p.label}</option>${(p.options||[]).map(b=>`<option value="${typeof b=="object"?b.value:b}" ${x[p.name]===(typeof b=="object"?b.value:b)?"selected":""}>${typeof b=="object"?b.label:b}</option>`).join("")}</select>`;if(p.type==="combobox"){let b=`dl-filter-${p.name}`,w=(p.options||[]).map(D=>{let C=typeof D=="object"?D.label||D.value||"":D||"";return(C==="undefined"||C==="[object Object]"||C==="null")&&(C=""),C?`<option value="${C}"></option>`:""}).join(""),_=x[p.name]||"";if(x[p.name]){let D=(p.options||[]).find(C=>String(typeof C=="object"?C.value:C)==String(x[p.name]));if(D){let C=typeof D=="object"?D.label||D.value||"":D||"";C&&C!=="undefined"&&C!=="[object Object]"&&C!=="null"&&(_=C)}}return`<div class="filter-combobox" style="flex:1; min-width:120px;">
                <input type="text" name="${p.name}" id="filter-${p.name}" list="${b}" class="form-control filter-combobox-input" value="${_}" placeholder="Pilih ${p.label}..." autocomplete="off">
                <datalist id="${b}">${w}</datalist>
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
  `;function F(){let p=document.getElementById("bulk-toolbar");if(!p)return;let b=document.getElementById("bulk-count"),w=document.getElementById("btn-bulk-delete"),_=document.getElementById("btn-bulk-cancel");b.textContent=`${$.size} item dipilih`,$.size>0?(p.style.display="flex",p.classList.add("has-items"),w.disabled=!1,_.disabled=!1,window.innerWidth<=768?(p.style.setProperty("top","auto","important"),p.style.setProperty("bottom","0","important")):(p.style.removeProperty("top"),p.style.removeProperty("bottom"))):(p.style.display="none",p.classList.remove("has-items"),w.disabled=!0,_.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{$.clear(),document.querySelectorAll(".row-checkbox").forEach(b=>b.checked=!1);let p=document.getElementById("select-all-checkbox");p&&(p.checked=!1),F()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if($.size===0)return;let p=[...$],b=document.createElement("div");b.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",b.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${p.length} ${d}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${p.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(b),b.querySelector("#bulk-cancel-btn").addEventListener("click",()=>b.remove()),b.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let w=b.querySelector("#bulk-confirm-btn");w.disabled=!0,w.textContent="Menghapus...";let _=await S(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:p})});b.remove(),_.ok?(ne(`${p.length} ${d} berhasil dihapus.`),$.clear(),F(),q()):z(_.data?.error||"Gagal menghapus data.")})});let O=document.getElementById("filter-search"),re;O?.addEventListener("input",p=>{clearTimeout(re),re=setTimeout(()=>{x.search=p.target.value,T=1,$.clear(),q()},400)}),o?.forEach(p=>{p.type==="select"&&document.getElementById(`filter-${p.name}`)?.addEventListener("change",b=>{x[p.name]=b.target.value,T=1,$.clear(),q()}),p.type==="combobox"&&document.getElementById(`filter-${p.name}`)?.addEventListener("change",b=>{let w=b.target.value,_=(p.options||[]).find(D=>{let C=String(typeof D=="object"?D.value:D),L=String(typeof D=="object"?D.label:D);return C===w||L===w});w?x[p.name]=_?typeof _=="object"?_.value:_:w:x[p.name]="",T=1,$.clear(),q()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{x={...n},O&&(O.value=""),o?.forEach(p=>{let b=document.getElementById(`filter-${p.name}`);b&&(b.value="")}),T=1,$.clear(),q()});let de=document.getElementById("btn-mobile-filter"),ie=document.getElementById("filter-options-wrapper"),ge=document.getElementById("btn-close-filter-sheet");if(de&&ie){de.addEventListener("click",b=>{b.preventDefault(),ie.classList.add("sheet-open")}),ge&&ge.addEventListener("click",b=>{b.preventDefault(),ie.classList.remove("sheet-open")});let p=document.getElementById("btn-apply-filter");p&&p.addEventListener("click",b=>{b.preventDefault(),ie.classList.remove("sheet-open")})}let me=document.getElementById("btn-mobile-aksi"),be=document.getElementById("excel-actions-wrapper"),_e=document.getElementById("btn-close-aksi-sheet");if(me&&be&&(me.addEventListener("click",p=>{p.preventDefault(),be.classList.add("sheet-open")}),_e&&_e.addEventListener("click",p=>{p.preventDefault(),be.classList.remove("sheet-open")})),document.getElementById("btn-create")?.addEventListener("click",()=>Ce(null)),v){document.getElementById(`btn-export-${v.moduleName}`)?.addEventListener("click",async b=>{let w=b.target,_=w.innerHTML;w.innerHTML="\u23F3 Loading...",w.disabled=!0;try{await v.onExport()}catch{z("Gagal export data")}finally{w.innerHTML=_,w.disabled=!1}}),document.getElementById(`btn-template-${v.moduleName}`)?.addEventListener("click",()=>{v.onTemplate()});let p=document.getElementById(`input-import-${v.moduleName}`);p?.addEventListener("change",async b=>{let w=b.target.files[0];if(!w)return;p.disabled=!0;let _=document.createElement("div");_.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",_.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(_);let D=_.querySelector("#import-progress-text"),C=_.querySelector("#import-progress-bar"),L=_.querySelector("#import-summary"),P=_.querySelector("#import-close-btn");P.addEventListener("click",()=>{_.remove(),q()});try{let W=await De(w);if(W.length===0)throw new Error("File kosong atau format salah");let V=500,ae=0,Z=0,N=0,M=W.length;D.textContent=`Ditemukan ${M} baris data. Memulai import...`;for(let I=0;I<M;I+=V){let J=W.slice(I,I+V);D.textContent=`Mengimport baris ${I+1} - ${Math.min(I+V,M)} dari ${M}...`,C.style.width=`${Math.round(I/M*100)}%`;try{let R=await v.onImport(J);R?(ae+=R.inserted||R.metrics?.inserted||J.length,Z+=R.skipped||R.metrics?.updated||0):ae+=J.length}catch(R){console.error("Chunk import failed:",R),N+=J.length}}C.style.width="100%",D.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',L.style.display="block",L.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${M}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${Z}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${N}</strong></div>
        `,N>0&&(L.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),P.style.display="block",p.value=""}catch(W){D.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${W.message}`,C.style.background="var(--danger)",C.style.width="100%",P.style.display="block",p.value=""}finally{p.disabled=!1}})}async function q(){F();let p=document.getElementById("table-container");if(!p)return;p.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let b=A==="client",w=b?1:T,_=b?Se:20,D=new URLSearchParams({page:w,limit:_,...Object.fromEntries(Object.entries(x).filter(([,N])=>N))}),C=await S(`${a}?${D}`);if(!C.ok){p.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${C.data?.error||"Error"}</p></div>`;return}let L=C.data?.data||C.data||[],P=C.data?.pagination,W=L.length,V=L;if(b){L=g(L),V=L;let N=L.length,M=20,I=Math.ceil(N/M);T>I&&I>0&&(T=I);let J=(T-1)*M,R=T*M;L=L.slice(J,R),P={page:T,limit:M,total:N,pages:I}}!1,f&&f(L);let ae=Be({columns:l,data:L,fullData:V,onEdit:u?N=>Ce(N):null,actions:y.map(N=>({...N,handler:M=>N.handler(M,q)})),emptyText:`Tidak ada ${String(d||"").toLowerCase()}`,bulkSelect:E?{selectedIds:$,onToggle:F}:null});p.innerHTML="",p.appendChild(ae);let Z=document.getElementById("pagination-container");if(Z&&(Z.innerHTML="",P&&P.pages>1)){let N=Ae({page:P.page,pages:P.pages,total:P.total,limit:P.limit,onPage:M=>{T=M,q()}});N&&Z.appendChild(N)}}function Ge(p){let b=typeof s=="function"?s(p):s;return Te(b)}function Ce(p){let b=!!p,w=document.createElement("form");if(w.noValidate=!0,w.innerHTML=Ge(p),b){let D=typeof s=="function"?s(p):s;Me(w,p)}let{close:_}=pe({title:b?`Edit ${d}`:`Tambah ${d}`,content:w,size:"lg",confirmText:b?"Simpan Perubahan":`Tambah ${d}`,onConfirm:async(D,C)=>{if(!w.reportValidity())return;let L=D.querySelector(".modal-confirm");L.disabled=!0,L.textContent="Menyimpan...";let P=Ne(w),W=typeof s=="function"?s(p):s,V=async M=>{for(let I of M)if(I.type==="row")await V(I.fields);else if(I.type==="combobox"&&P[I.name]){let J=P[I.name],R=(I.options||[]).find(U=>{let ee=String(typeof U=="object"?U.value:U),ze=String(typeof U=="object"?U.label:U);return ee===J||ze===J});if(R)P[I.name]=typeof R=="object"?R.value:R;else if(I.createApi){let U={};U[I.createApi.field]=J,I.createApi.extra&&Object.assign(U,I.createApi.extra);let ee=await S(I.createApi.path,{method:"POST",body:JSON.stringify(U)});if(ee.ok&&ee.data?.id)P[I.name]=ee.data.id;else if(ee.ok&&!ee.data?.id)P[I.name]=J;else throw new Error(`Gagal membuat master data: ${ee.data?.error||"Unknown error"}`)}}};try{await V(W)}catch(M){z(M.message),L.disabled=!1,L.textContent=b?"Simpan Perubahan":`Tambah ${d}`;return}m&&(P=await m(P,p));let ae=b?"PUT":"POST",Z=b?`${a}/${p.id}`:a,N=await S(Z,{method:ae,body:JSON.stringify(P)});N.ok?(ne(b?`${d} berhasil diperbarui.`:`${d} berhasil ditambahkan.`),C(),q()):(z(N.data?.error||"Gagal menyimpan data."),L.disabled=!1,L.textContent=b?"Simpan Perubahan":`Tambah ${d}`)}})}function $t(p){Ee(`Hapus ${d} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let b=await S(`${a}/${p.id}`,{method:"DELETE"});b.ok?(ne(`${d} berhasil dihapus.`),q()):z(b.data?.error||"Gagal menghapus.")},`Hapus ${d}`)}return q(),q}j();G();function It(t,e){let r=String(t.status||"").toLowerCase();return e==="active"?r==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&r==="aktif":!1}j();G();function ft(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let l=new Date(a);l.setDate(a.getDate()+30);let s=new Date(t.end_date);return s.setHours(0,0,0,0),s>=a&&s<=l}return!1}j();G();function Lt(t,e){let r=String(t.status||"").toLowerCase();return e==="open"?r==="open":!1}j();function Pt(t,e){let r=String(t.status||"").toLowerCase();return e==="pending"?r==="pending":!1}var xe={};function He(t){if(xe[t]){try{xe[t].destroy()}catch{}delete xe[t]}}function Sa(){Object.keys(xe).forEach(He)}var ve=(t,e=0)=>{let r=Number(t);return isNaN(r)||t===null||t===void 0?e:r},Ke=(t,e="\u2014")=>{if(t==null||t==="")return e;let r=String(t).trim();return r===""||r==="[object Object]"?e:r};function At(t,e,r=900){if(!t)return;let a=Math.max(0,Math.round(ve(e)));if(a===0){t.textContent="0";return}let l=Date.now(),s=()=>{let o=Math.min((Date.now()-l)/r,1),n=1-Math.pow(1-o,3);t.textContent=Math.round(n*a).toLocaleString("id-ID"),o<1?requestAnimationFrame(s):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(s)}var wa={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},xa=t=>{let e=Ke(t,"\u2014");return`<span class="status-pill ${wa[e]||"pill-neutral"}">${e}</span>`};var ye={family:"Inter",size:11},ke="#94A3B8",je="#F1F5F9",vt=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],_a=()=>window.innerWidth<768;function it(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:_a()?"bottom":"top",labels:{font:ye,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ye,titleFont:{...ye,weight:"700"}}},scales:{x:{grid:{color:je},ticks:{font:ye,color:ke,maxRotation:0}},y:{grid:{color:je},ticks:{font:ye,color:ke},beginAtZero:!0}},...t}}var Ca=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join("");function Bt(t=3){return Array(t).fill(0).map((e,r)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${r<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function oe(t,e,r=8e3){try{let a=new AbortController,l=setTimeout(()=>a.abort(),r),s=await S(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(l),!s||!s.ok)return e;let o=s.data;return o?o.data!==void 0?o.data??e:o:e}catch{return e}}function Ea(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(a=>{let l=document.getElementById(a);l&&(l.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(a=>{let l=document.getElementById(a);if(l&&l.style.display==="none"){l.style.display="block";let s=l.parentElement;if(s&&!s.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent="Belum ada data",l.style.display="none",s.appendChild(o)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Mt({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Ft({}),["table-contracts","table-issues"].forEach(a=>{let l=document.getElementById(a);l&&l.querySelector(".skeleton")&&(l.innerHTML='<div class="chart-empty">Belum ada data</div>')});let r=document.getElementById("activity-log");r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Nt(t){Sa(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${Bt(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${Bt(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>kt(t)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async r=>{let a=r.target.value,l=document.getElementById("jadwal-year-label");l&&(l.textContent=a);let s=document.getElementById("skel-jadwal"),o=document.getElementById("chart-jadwal");s&&(s.style.display="block",s.style.position="absolute"),o&&(o.style.display="none");let n=await oe(`/api/dashboard/schedule-chart?year=${a}`,{},8e3);try{Ot(n)}catch(i){console.warn("ScheduleChart render:",i),fe("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async r=>{let a=r.target.value,l=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",s=document.getElementById("skel-insp"),o=document.getElementById("chart-insp");s&&(s.style.display="block",s.style.position="absolute"),o&&(o.style.display="none");let n=await oe(l,{},8e3);try{Rt(n)}catch(i){console.warn("InspBar render:",i),fe("skel-insp","chart-insp")}}),t._skelTimeout=setTimeout(()=>Ea(),5e3),await kt(t),t._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?kt(t):clearInterval(t._dashRefresh)},6e4)}async function kt(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,r,a,l,s,o,n,i,d,h,u,c,m,f]=await Promise.all([oe("/api/dashboard/kpi",{},8e3),oe("/api/dashboard/issues-trend",{},8e3),oe("/api/dashboard/issues-summary",{},8e3),oe("/api/dashboard/stats",{},8e3),oe("/api/dashboard/calendar",[],8e3),oe("/api/schedule?limit=10000",{data:[]},8e3),oe("/api/employees?limit=10000",{data:[]},8e3),oe("/api/contracts?limit=10000",{data:[]},8e3),oe("/api/issues?limit=10000",{data:[]},8e3),oe("/api/one-on-one?limit=10000",{data:[]},8e3),oe("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),oe(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),oe("/api/relievers?limit=10000",{data:[]},8e3),oe("/api/reports/fogging?limit=10000",{data:[]},8e3)]),g=document.getElementById("filter-insp-month"),y=g?g.value:"",k=y?`/api/dashboard/inspection-bar?month=${y}`:"/api/dashboard/inspection-bar",v=await oe(k,{},8e3);if(e){let E=Array.isArray(o?.data)?o.data:Array.isArray(o)?o:[];window.dashboardSchedules=E;let A=Array.isArray(n?.data)?n.data:Array.isArray(n)?n:[],B=Array.isArray(i?.data)?i.data:Array.isArray(i)?i:[],T=Array.isArray(d?.data)?d.data:Array.isArray(d)?d:[],x=Array.isArray(h?.data)?h.data:Array.isArray(h)?h:[],$=Array.isArray(m?.data)?m.data:Array.isArray(m)?m:[];window.dashboardRelievers=$;let F=Array.isArray(f?.data)?f.data:Array.isArray(f)?f:[];window.dashboardFogging=F,e.employees&&(e.employees.current=A.filter(O=>It(O,"active")).length),e.contracts&&(e.contracts.current=B.filter(O=>ft(O,"active")).length),e.expiring30&&(e.expiring30={current:B.filter(O=>ft(O,"expiring30")).length}),e.issues&&(e.issues.current=T.filter(O=>Lt(O,"open")).length),e.one_on_one&&(e.one_on_one.current=x.filter(O=>Pt(O,"pending")).length),e.inspection_month&&(e.inspection_month.current=E.filter(O=>yt(O,"inspeksi")).length),e.cleaning_month&&(e.cleaning_month.current=E.filter(O=>yt(O,"gcdc")).length)}try{Mt(e)}catch(E){console.warn("KPI render:",E)}try{Ft(e)}catch(E){console.warn("MiniStats render:",E)}try{Ot(c)}catch(E){console.warn("ScheduleChart render:",E),fe("skel-jadwal","chart-jadwal")}try{Ta(Array.isArray(a?.by_category)?a.by_category:[])}catch(E){console.warn("Donut render:",E),fe("skel-donut","chart-donut")}try{Da(r)}catch(E){console.warn("Trend render:",E),fe("skel-trend","chart-trend")}try{Rt(v)}catch(E){console.warn("InspBar render:",E),fe("skel-insp","chart-insp")}try{let E=Array.isArray(l)?l:Array.isArray(l?.recent_issues)?l.recent_issues:[];La(E)}catch(E){console.warn("IssuesTable render:",E)}try{let E=Array.isArray(l?.expiring_contracts)?l.expiring_contracts:[];Ia(u)}catch(E){console.warn("ContractsTable render:",E)}try{Pa(Array.isArray(s)?s:[])}catch(E){console.warn("Agenda render:",E)}try{Ba()}catch(E){console.warn("Quick Actions render:",E)}}function Mt(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let r=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=r.map(a=>{let l=ve(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${l}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${l}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{At(a,parseInt(a.dataset.target)||0)})}function Ft(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let r=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,a=new Date().getFullYear(),l=String(new Date().getMonth()+1).padStart(2,"0"),s=`${a}-${l}`,o=u=>`
    <select id="${u}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.preventDefault(); event.stopPropagation();">
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
      `,val:t.schedule?.current,href:`#/timeline?dash_filter=period_${r.toLowerCase()}`,color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:o("dash-reliefer-month"),val:t.reliever_completed?.current,href:`#/relievers?dash_filter=reliever&month=${s}`,color:"mini-teal"},{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:o("dash-inspeksi-month"),val:t.inspection_month?.current,href:`#/timeline?dash_filter=inspeksi&month=${s}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:o("dash-gcdc-month"),val:t.cleaning_month?.current,href:`#/timeline?dash_filter=gcdc&month=${s}`,color:"mini-green"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:o("dash-fogging-month"),val:t.fogging_month?.current,href:`#/reports/fogging?dash_filter=fogging&month=${s}`,color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=n.map(u=>`
    <a href="${u.href}" class="mini-stat ${u.color}" style="text-decoration:none" id="${u.id||""}">
      <div class="mini-stat-icon">${u.icon}</div>
      <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
        <div style="display:flex; align-items:baseline; gap:3px;">
          <div class="mini-stat-value" data-target="${ve(u.val)}">0</div>
          ${u.dropdown?u.dropdown:""}
        </div>
        <div class="mini-stat-text">${u.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(u=>At(u,parseInt(u.dataset.target)||0,700));let i=document.getElementById("dash-jadwal-period");i&&i.addEventListener("change",u=>{let c=u.target.value,m=(window.dashboardSchedules||[]).filter(y=>y.period===c).length,f=document.querySelector("#mini-jadwal .mini-stat-value");f&&(f.dataset.target=m,f.textContent=m);let g=document.getElementById("mini-jadwal");g&&(g.href=`#/timeline?dash_filter=period_${c.toLowerCase()}`)});let d=(u,c,m,f,g)=>{let y=document.getElementById(u);y&&y.addEventListener("change",k=>{let v=k.target.value,E=(m||[]).filter(T=>f(T,v)).length,A=document.querySelector(`#${c} .mini-stat-value`);A&&(A.dataset.target=E,A.textContent=E);let B=document.getElementById(c);B&&(B.href=`${g}&month=${v}`)})},h=u=>{let c=String(u.status||"").toLowerCase();return c==="done"||c==="selesai"||c==="completed"};d("dash-reliefer-month","mini-reliefer",window.dashboardRelievers,(u,c)=>window.parseFlexibleDate(u.backup_date).startsWith(c)&&h(u),"#/relievers?dash_filter=reliever"),d("dash-inspeksi-month","mini-inspeksi",window.dashboardSchedules,(u,c)=>u.activity_type==="Inspeksi Hygiene"&&h(u)&&window.parseFlexibleDate(u.opening_date||u.target_date).startsWith(c),"#/timeline?dash_filter=inspeksi"),d("dash-gcdc-month","mini-gcdc",window.dashboardSchedules,(u,c)=>(u.activity_type==="General Cleaning"||u.activity_type==="Deep Cleaning")&&h(u)&&window.parseFlexibleDate(u.opening_date||u.target_date).startsWith(c),"#/timeline?dash_filter=gcdc"),d("dash-fogging-month","mini-fogging",window.dashboardFogging,(u,c)=>h(u)&&window.parseFlexibleDate(u.activity_date).startsWith(c),"#/reports/fogging?dash_filter=fogging")}function Ta(t){fe("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),r=document.getElementById("donut-legend");if(!e||!r)return;He("donut");let a=(t||[]).filter(i=>ve(i.count)>0);if(!a.length){Xe(e,"Belum ada data permasalahan");return}let l=a.map(i=>`${Ke(i.category,"Lainnya")}`),s=a.map(i=>ve(i.count)),o=s.reduce((i,d)=>i+d,0);r.innerHTML=a.map((i,d)=>{let h=vt[d%vt.length],u=o>0?Math.round(i.count/o*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${h}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${i.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${u}%)</span></div>
          <div class="donut-legend-label">${l[d]}</div>
        </div>
      </div>
    `}).join("");let n={id:"centerText",beforeDraw:function(i){let d=i.width,h=i.height,u=i.ctx;u.restore();let c=(h/80).toFixed(2);u.font="bold "+c+"em Inter",u.textBaseline="middle",u.fillStyle="#1E293B";let m=o.toString(),f=Math.round((d-u.measureText(m).width)/2),g=h/2;u.fillText(m,f,g-4),u.font="600 "+(c*.35).toFixed(2)+"em Inter",u.fillStyle="#64748B";let y="Total",k=Math.round((d-u.measureText(y).width)/2);u.fillText(y,k,g+10),u.save()}};xe.donut=new Chart(e,{type:"doughnut",data:{labels:l,datasets:[{data:s,backgroundColor:vt,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ye,titleFont:{...ye,weight:"700"},callbacks:{label:i=>` ${i.label}: ${i.parsed} kasus`}}},cutout:"75%"},plugins:[n]})}function Da(t){fe("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;He("trend"),t=t||{};let r=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(t.labels||[]).map(o=>{if(!o||typeof o!="string")return"";try{let[n,i]=o.split("-");return(r[Number(i)-1]||i)+" "+String(n).slice(-2)}catch{return o}}),l=(t.open||[]).map(o=>ve(o)),s=(t.closed||[]).map(o=>ve(o));if(!a.length){Xe(e,"Belum ada data trend");return}xe.trend=new Chart(e,{type:"line",data:{labels:a,datasets:[{label:"Open",data:l,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:s,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:it({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:ke,maxRotation:0,autoSkip:!1}},y:{grid:{color:je},ticks:{font:{family:"Inter",size:9},color:ke},beginAtZero:!0}}})})}function Ot(t){fe("skel-jadwal","chart-jadwal");let e=document.getElementById("chart-jadwal");if(!e)return;He("jadwal"),t=t||{};let r=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(t).some(i=>Array.isArray(i)&&i.some(d=>d>0))){Xe(e,"Belum ada data jadwal");return}let l=t["Inspeksi Hygiene"]||Array(12).fill(0),s=t["General Cleaning"]||Array(12).fill(0),o=t["Deep Cleaning"]||Array(12).fill(0),n=t.Fogging||Array(12).fill(0);xe.jadwal=new Chart(e,{type:"bar",data:{labels:r,datasets:[{label:"Inspeksi",data:l,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:s,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:o,backgroundColor:"#F59E0B"},{label:"Fogging",data:n,backgroundColor:"#EF4444"}]},options:it({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:ke,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:je},ticks:{font:{family:"Inter",size:9},color:ke},min:0}}})})}function Rt(t){fe("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;He("inspBar"),t=t||{};let r=t.labels||[],a=(t.fc||[]).map(s=>ve(s)),l=(t.spv||[]).map(s=>ve(s));if(!r.length){Xe(e,"Belum ada data inspeksi");return}xe.inspBar=new Chart(e,{type:"bar",data:{labels:r,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:l,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:it({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ye,color:ke,maxRotation:45,minRotation:30}},y:{grid:{color:je},ticks:{font:ye,color:ke},min:0,max:100}}})})}function Ia(t){fe("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;He("contractMiniBar"),t=t||{};let r={"06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(o=>{let n=o.split("-")[1];return r[n]||o}),l=(t.data||[]).map(o=>ve(o));if(!a.length){Xe(e,"Belum ada data");return}let s=e.getContext("2d");xe.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:l,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:it({onClick:(o,n)=>{if(n&&n.length>0){let i=n[0].index,d=(t.labels||[])[i];d&&(window.location.hash="#/contracts?month_expiry="+d)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ye,color:ke,maxRotation:0,autoSkip:!1}},y:{grid:{color:je,borderDash:[4,4],drawBorder:!1},ticks:{font:ye,color:ke,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function La(t){let e=document.getElementById("table-issues");if(!e)return;let r=(t||[]).slice(0,8);if(!r.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${r.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${xa(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Ke(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Ke(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function Pa(t){let e=document.getElementById("widget-agenda");if(!e)return;let r=new Date,a=`${r.getFullYear()}-${String(r.getMonth()+1).padStart(2,"0")}-${String(r.getDate()).padStart(2,"0")}`,s=(t||[]).filter(o=>(o.event_date||"").startsWith(a)).slice(0,10);if(!s.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${s.map(o=>{let n="#3B82F6",i="#EFF6FF",d="Agenda",h=(o.title||"").toLowerCase();return h.includes("inspeksi")?(n="#10B981",i="#ECFDF5",d="Inspeksi"):h.includes("cleaning")||h.includes("gcdc")?(n="#3B82F6",i="#EFF6FF",d="Cleaning"):h.includes("reliefer")?(n="#F59E0B",i="#FFFBEB",d="Reliefer"):h.includes("fogging")&&(n="#8B5CF6",i="#F5F3FF",d="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(o.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${n};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Ke(o.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Ke(o.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${i};color:${n}">${d}</div>
        </div>
      `}).join("")}
    </div>
  `}function Ba(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(r=>`
    <a href="${r.href}" class="action-btn">
      <div class="action-icon" style="background:${r.bg}">${r.icon}</div>
      ${r.label}
    </a>
  `).join("")}function fe(t,e){let r=document.getElementById(t),a=document.getElementById(e);if(r&&(r.style.display="none",r.style.position=""),a){a.style.display="block";let l=a.parentElement;if(l){let s=l.querySelector(".chart-empty");s&&s.remove()}}}function Xe(t,e="Belum ada data"){if(!t)return;t.style.display="none";let r=t.parentElement;if(!r)return;if(!r.querySelector(".chart-empty")){let l=document.createElement("div");l.className="chart-empty",l.textContent=e,r.appendChild(l)}}j();async function Kt(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),r=document.getElementById("login-error"),a=document.getElementById("login-btn"),l=document.getElementById("toggle-password"),s=document.getElementById("login-password");l?.addEventListener("click",()=>{let o=s.type==="text";s.type=o?"password":"text",l.style.color=o?"":"var(--primary)"}),e?.addEventListener("submit",async o=>{o.preventDefault(),r.style.display="none";let n=e.username.value.trim(),i=e.password.value;if(!n||!i){r.textContent="Username dan password wajib diisi.",r.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let d=await S("/api/auth/login",{method:"POST",body:JSON.stringify({username:n,password:i})});d.ok&&d.data.success?(pt(d.data.data.token),Ve(d.data.data.user),ne("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(r.textContent=d.data.error||"Username atau password salah.",r.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{r.textContent="Gagal terhubung ke server. Periksa koneksi internet.",r.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}j();G();async function Aa(){return await Q()}function Na(t,e){let r=String(t.status||"").toLowerCase();return e==="active"?r==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&r==="aktif":!1}async function jt(t,e){let r=await Aa(),a=e?e.get("dash_filter"):null;H({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",enableMobileFilterSheet:!0,itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:l=>a?l.filter(s=>Na(s,a)):l,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:l=>Re(l)},{key:"phone",label:"No. HP",render:l=>l?`<a href="tel:${l}">${l}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>X(l)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:l=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:l?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:l?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:r,value:l?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:l?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:l?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:l?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let l=await S(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(l.ok){let s=l.data.data.map(o=>({"Nama Lengkap":o.full_name,Cabang:o.branch_name||"",Divisi:o.division||"","No. HP":o.phone||"","Tgl Masuk":o.join_date||"",Status:o.status||""}));K(s,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async l=>{let s=i=>{if(!i)return null;let d=String(i||"").toLowerCase(),h=r.find(u=>String(u.label||"").toLowerCase()===d);return h?h.value:null},o=l.map(i=>({full_name:String(i["Nama Lengkap"]||"").trim(),branch_id:s(String(i.Cabang||"").trim()),division:String(i.Divisi||"").trim()||"FACILITY CARE",phone:String(i["No. HP"]||"").trim(),join_date:String(i["Tgl Masuk"]||"").trim(),status:String(i.Status||"").trim(),notes:String(i.Catatan||"").trim()})).filter(i=>i.full_name),n=await S("/api/import/employees",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}j();G();var wt=[],Ht=[];async function Ma(){wt=await Q(),Ht=await Oe()}var St=async t=>{let e=[],r=1;for(;;){let l=await(await Promise.resolve().then(()=>(j(),Le))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${r}`);if(!l.ok)break;let s=l.data?.data||l.data||[],o=Array.isArray(s)?s:[];if(e=e.concat(o),o.length<100||l.data?.pagination&&r>=l.data.pagination.pages)break;r++}return e};function Fa(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let l=new Date(a);l.setDate(a.getDate()+30);let s=new Date(t.end_date);return s.setHours(0,0,0,0),s>=a&&s<=l}return!1}async function qt(t,e){await Ma();let r=e?e.get("dash_filter"):null;H({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Kontrak",paginationMode:"client",onDataLoaded:a=>r?a.filter(l=>Fa(l,r)):a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>Re(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,l)=>l.end_date&&String(l.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':bt(a)},{key:"status",label:"Status",render:a=>X(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:wt},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"month_expiry",label:"Bulan Habis",options:[{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"},{value:"2027-01",label:"Jan 2027"},{value:"2027-02",label:"Feb 2027"},{value:"2027-03",label:"Mar 2027"},{value:"2027-04",label:"Apr 2027"},{value:"2027-05",label:"Mei 2027"},{value:"2027-06",label:"Jun 2027"},{value:"2027-07",label:"Jul 2027"}]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Detail Kontrak",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[s,o]=await Promise.all([St("/api/employees?status=Aktif"),St("/api/contracts")]);if(s.length>0){let n=o.filter(m=>m.status==="Aktif"&&(m.days_remaining==null||m.days_remaining>=0||String(m.end_date).startsWith("2099"))),i=new Set(n.map(m=>m.employee_id)),d=s.filter(m=>!i.has(m.id)),h=[],u=[];d.forEach(m=>{let f=o.filter(g=>g.employee_id===m.id);f.length===0?h.push(m):u.push({emp:m,lastContract:f[0]})});let c=`<p style="margin-bottom:12px">Data yang terbaca: <b>${s.length}</b> Karyawan Aktif, dan <b>${n.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:6px">Terdapat <b>${h.length}</b> karyawan aktif tanpa kontrak.</p>
              <p style="margin-bottom:12px">Terdapat <b>${u.length}</b> karyawan aktif yang memiliki masa kontrak Expired.</p>
              <ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;h.forEach(m=>{c+=`<li style="margin-bottom:8px"><b>${m.full_name}</b> <br><span style="font-size:0.85em;color:#F59E0B">Cabang: ${m.branch_name||"-"} | Belum pernah di-input kontrak</span></li>`}),u.forEach(m=>{let f=m.emp,g=m.lastContract,y=g.status==="Aktif"&&g.days_remaining<0?"Aktif (Masa Habis)":g.status;c+=`<li style="margin-bottom:8px"><b>${f.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${f.branch_name||"-"} | Status Terakhir: <b style="color:#EF4444">${y}</b>, Tgl Berakhir: ${window.formatDate(g.end_date)}</span></li>`}),c+="</ul>",Promise.resolve().then(()=>(we(),mt)).then(m=>m.createModal({title:"Detail Karyawan Tanpa Kontrak Aktif",content:c,cancelText:"Tutup"}))}}catch(s){console.error(s)}a.innerHTML="\u{1F50D} Detail Kontrak",a.disabled=!1};let l=document.querySelector(".page-actions");l&&l.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"select",required:!0,options:Ht,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"select",options:wt,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await S(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let l=a.data.data.map(s=>({"Nama Lengkap":s.employee_name,Cabang:s.branch_name||"","Div / Bagian":s.division||"","Tanggal Mulai":s.start_date||"","Tanggal Selesai":s.end_date&&String(s.end_date).startsWith("2099")?"":s.end_date||"","Sisa Kontrak":s.end_date&&String(s.end_date).startsWith("2099")?"Tetap":s.days_remaining!==null&&s.days_remaining!==void 0?`${s.days_remaining} Hari`:"",Status:s.status||""}));K(l,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[l,s]=await Promise.all([S("/api/branches?limit=10000"),St("/api/employees")]),o=l.data?.data||[],n=s||[];console.log(`Total employee yang berhasil dimuat dari database : ${n.length}`),n.length>0&&(console.log("Contoh 5 employee pertama:"),n.slice(0,5).forEach((g,y)=>{console.log(`${y+1}. ID: ${g.id}, Name: ${g.full_name}, Status: ${g.status}`)}));let i=g=>{if(!g)return null;let y=String(g||"").replace(/\s+/g," ").toLowerCase().trim(),k=o.find(v=>String(v.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===y||String(v.code||"").replace(/\s+/g," ").toLowerCase().trim()===y||String(v.name||"").replace(/\s+/g," ").toLowerCase().trim()===y);return k?k.id:null},d=(g,y)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${y}`),console.log(`Nama dari Excel : "${g}"`),!g)return console.log("Alasan gagal mapping : Nama kosong"),null;let k=String(g||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${k}"`),console.log(`Jumlah employee di database : ${n.length}`);let v=n.find(E=>String(E.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===k);return v?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${v.id}`),v.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},h=g=>{if(!g)return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let y=String(g).trim();if(/^\d{4,5}(\.\d+)?$/.test(y)){let v=Math.floor(Number(y));if(v>2e4&&v<99999){let E=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(E.getTime())?"":E.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);let k=y.split(/[\/\-\.]/);if(k.length===3){let[v,E,A]=k.map(B=>B.trim());if(v.length===4&&E.length<=2&&A.length<=2)return`${v}-${E.padStart(2,"0")}-${A.padStart(2,"0")}`;if(A.length===4&&E.length<=2&&v.length<=2)return`${A}-${E.padStart(2,"0")}-${v.padStart(2,"0")}`}return y},u=a.map((g,y)=>{let k=y+2,v=String(g["Nama Lengkap"]||"").trim(),E=g["Tanggal Mulai"],A=h(E);if(!A){let x=a.__worksheet,$=a.__headers||[],F=$.indexOf("Tanggal Mulai"),O="N/A",re="N/A",de="N/A";if(F!==-1&&x&&window.XLSX){let ge=window.XLSX.utils.encode_cell({c:F,r:k-1});de=ge;let me=x[ge];me?(O=me.t||"undefined",re=me.w||"undefined"):O="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let ie="Unknown";E==null||E===""?ie="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":E instanceof Date&&isNaN(E.getTime())?ie="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":ie="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${k}`),console.log(`Employee Name : ${v}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${F})`),console.log(`Raw Cell Value : "${E}"`),console.log(`JavaScript Type : ${typeof E}`),console.log(`SheetJS Cell Type : ${O}`),console.log(`SheetJS Formatted Value : "${re}"`),console.log(`Value After Trim : "${String(E||"").trim()}"`),console.log(`Value After Date Parser : "${A}"`),console.log(`Is Empty : ${!E}`),console.log(`Is Invalid Date : ${E instanceof Date?isNaN(E.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${ie}`),console.log(`Workbook Sheet : ${x?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${de}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(g,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify($)),console.log(`==========================
`)}let B=d(v,k),T=null;return B||(T="Karyawan tidak ditemukan di Database"),{isValid:!!B,invalidReason:T,rowNum:k,data:{employee_id:B,branch_id:i(String(g.Cabang||"").trim()),division:String(g["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:A,end_date:h(g["Tanggal Selesai"])||"2099-12-31",status:String(g.Status||"").trim(),_rawName:v}}}),c=[],m=[];if(u.forEach(g=>{g.isValid?c.push(g.data):m.push({rowNum:g.rowNum,name:g.data._rawName,reason:g.invalidReason})}),console.log(`Split Validation - Valid: ${c.length}, Invalid: ${m.length}`),c.length===0)return{inserted:0,skipped:a.length,failed:a.length};let f=await S("/api/import/contracts",{method:"POST",body:JSON.stringify({rows:c,onDuplicate:"update"})});if(!f.ok)throw new Error(f.data?.error||"Import gagal");return f.data}}})}j();G();var xt=[],Ze=[];function Oa(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let r of e)if(t.some(a=>a.period===r))return r;return"Q3"}async function Jt(t,e){xt=await Q();let r=await le();Ze=["Berlin Ariansyah","Ade Surahman"];let a=y=>y&&!Ze.find(k=>String(typeof k=="object"?k.value:k).toLowerCase()===String(y).toLowerCase())?[...Ze,y]:Ze,l=await S(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),s=y=>{if(!y||y==="-"||String(y).trim()==="")return"";let k=String(y).split("-");return k.length===3&&k[0].length===4?`${k[2]}-${k[1]}-${k[0]}`:y},o=l.data?.data||[],n=Oa(o),i=e?e.get("dash_filter"):null,d=new Date,h=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`,u={},c=e&&e.get("month")?e.get("month"):null;i==="inspeksi"?u={status:"Done",activity_type:"Inspeksi Hygiene",month:c}:i==="gcdc"?u={status:"Done",activity_type:"GCDC",month:c}:i&&i.startsWith("period_")&&(u={period:i.replace("period_","").toUpperCase()});let m=new Date().getFullYear(),g=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((y,k)=>{let v=String(k+1).padStart(2,"0");return{value:"-",label:" "}});gt({container:t,title:"Jadwal Kegiatan",icon:"\xF0\u0178\u201C\u2026",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",enableMobileFilterSheet:!0,defaultFilters:u,onDataLoaded:y=>y.sort((k,v)=>{let E=k.opening_date?new Date(k.opening_date).getTime():0;return(v.opening_date?new Date(v.opening_date).getTime():0)-E}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:y=>ht(y)},{key:"period",label:"Periode",render:y=>he(y)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:y=>s(y)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:y=>s(y)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:y=>s(y)},{key:"status",label:"Status",render:y=>X(y)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:xt},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Ze}],formFields:y=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:xt,value:y?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:y?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:y?.period},{name:"pic",label:"PIC",type:"select",options:a(y?.pic),value:y?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:y?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:y?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:y?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:y?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:y?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let y=await S(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(y.ok){let k=y.data.data.map(v=>({Cabang:v.branch_name||"",Kegiatan:v.activity_type||"",Periode:v.period||"",PIC:v.pic||"","Tgl Opening":v.opening_date||"","Tgl Target":v.target_date||"","Tgl Selesai":v.completion_date||"",Status:v.status||""}));K(k,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async y=>{let v=(await S("/api/branches?all=1")).data?.data||[],E=x=>{if(!x)return null;let $=String(x||"").toLowerCase(),F=v.find(O=>String(O.full_name||"").toLowerCase()===$||String(O.code||"").toLowerCase()===$||String(O.name||"").toLowerCase()===$);return F?F.id:null},A=x=>{if(x==null||x==="")return"";if(x instanceof Date&&!isNaN(x.getTime()))return x.toISOString().slice(0,10);let $=String(x).trim();if($===""||$==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test($))return $.slice(0,10);if(/^\d{4,5}$/.test($)){let O=Number($);if(O>2e4&&O<99999){let re=new Date(Date.UTC(1899,11,30)+O*864e5);return isNaN(re.getTime())?"":re.toISOString().slice(0,10)}}let F=$.split(/[\/\-\.]/);if(F.length===3){let[O,re,de]=F.map(ie=>ie.trim());if(O.length===4&&re.length<=2&&de.length<=2)return`${O}-${re.padStart(2,"0")}-${de.padStart(2,"0")}`;if(de.length===4&&re.length<=2&&O.length<=2)return`${de}-${re.padStart(2,"0")}-${O.padStart(2,"0")}`}return $},B=y.map(x=>({branch_id:E(String(x.Cabang||"").trim()),activity_type:String(x.Kegiatan||"").trim(),period:String(x.Periode||"").trim(),pic:String(x.PIC||x.Pic||"").trim(),opening_date:A(x["Tgl Opening"]||x["Tanggal Opening"]||x["Tgl Openir"]),target_date:A(x["Tgl Target"]||x["Tanggal Target"]),completion_date:A(x["Tgl Selesai"]||x["Tanggal Selesai"]),status:String(x.Status||"").trim(),notes:String(x.Catatan||x.Keterangan||"").trim()})).filter(x=>x.activity_type&&x.period),T=await S("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:B,onDuplicate:"update"})});if(!T.ok)throw new Error(T.data?.error||"Import gagal");return T.data}}})}j();G();var _t=[],rt=[];function Ra(t,e){let r=String(t.status||"").toLowerCase();return e==="open"?r==="open":!1}async function Ut(t,e){let r=e?e.get("dash_filter"):null;_t=await Q(),rt=await le();let a=i=>i&&!rt.find(d=>d.value===i)?[...rt,{value:i,label:i}]:rt,l=new Date().getFullYear(),s=["2025","2026","2027","2028","2029","2030"],n=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((i,d)=>{let h=String(d+1).padStart(2,"0");return{value:`${l}-${h}`,label:`${i} ${l}`}});H({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:i=>r?i.filter(d=>Ra(d,r)):i,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:i=>`<span class="badge badge-secondary">${i}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:i=>`<span title="${i}">${i?.length>50?i.slice(0,50)+"\u2026":i}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:i=>`<span title="${i||""}">${i?.length>40?i.slice(0,40)+"\u2026":i||"-"}</span>`},{key:"status",label:"Status",render:i=>X(i)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:i=>window.formatDate(i)},{key:"day_count",label:"Hari",render:i=>i??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:_t},{type:"select",name:"month",label:"Bulan",options:n},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:s}],formFields:i=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:i?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:_t,value:i?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:i?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:i?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:i?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(i?.employee_name),value:i?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(i?.fc_specialist),value:i?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:i?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:i?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:i?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let i=await S(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let d=i.data.data.map(h=>({Tanggal:h.report_date||"",Cabang:h.branch_name||"",Kategori:h.category||"",Sumber:h.source||"",Keluhan:h.complaint||"","Nama FC":h.employee_name||"","FC Spesialis":h.fc_specialist||"",Solusi:h.solution||"","Tgl Selesai":h.completion_date||"",Status:h.status||""}));K(d,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async i=>{let h=(await S("/api/branches?all=1")).data?.data||[],u=f=>{if(!f)return null;let g=String(f||"").toLowerCase(),y=h.find(k=>String(k.full_name||"").toLowerCase()===g||String(k.code||"").toLowerCase()===g||String(k.name||"").toLowerCase()===g);return y?y.id:null},c=i.map(f=>({branch_id:u(String(f.Cabang||"").trim()),report_date:String(f.Tanggal||"").trim(),category:String(f.Kategori||"").trim(),source:String(f.Sumber||"").trim(),complaint:String(f.Keluhan||"").trim(),employee_name:String(f["Nama FC"]||"").trim(),fc_specialist:String(f["FC Spesialis"]||"").trim(),solution:String(f.Solusi||"").trim(),completion_date:String(f["Tgl Selesai"]||"").trim(),status:String(f.Status||"").trim()})).filter(f=>f.report_date&&f.complaint&&f.category),m=await S("/api/import/issues",{method:"POST",body:JSON.stringify({rows:c,onDuplicate:"update"})});if(!m.ok)throw new Error(m.data?.error||"Import gagal");return m.data}}})}j();var qe=[];function Ka(t,e){let r=String(t.status||"").toLowerCase();return e==="pending"?r==="pending":!1}async function Gt(t,e){let r=e?e.get("dash_filter"):null;qe=await Q();let a=await le(),l=["Ade","Berlin"],s=n=>n&&!a.find(i=>i.value===n)?[...a,{value:n,label:n}]:a,o=n=>n&&!l.find(i=>(typeof i=="object"?i.value:i)===n)?[...l,n]:l;H({container:t,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:n=>r?n.filter(i=>Ka(i,r)):n,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:n=>`<span title="${n||""}">${n?.length>50?n.slice(0,50)+"\u2026":n||"-"}</span>`},{key:"solution",label:"Solusi",render:n=>`<span title="${n||""}">${n?.length>40?n.slice(0,40)+"\u2026":n||"-"}</span>`},{key:"status",label:"Status",render:n=>X(n)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:qe},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async n=>{let i=new URLSearchParams(n||{}).toString(),d=await S(`/api/one-on-one?limit=10000&${i}`);if(d.ok){let h=d.data.data.map(c=>({Tanggal:c.meeting_date||"",Cabang:c.branch_name||"","Nama Karyawan":c.employee_name||"",PIC:c.pic||"",Masalah:c.problem||"",Solusi:c.solution||"",Status:c.status||"","Tgl Selesai":c.completion_date||"",Dokumen:c.document_link||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(G(),ue));u(h,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(G(),ue));i(n,"Template_Import_OneOnOne")},onImport:async n=>{let i=c=>{if(!c)return null;let m=String(c||"").toLowerCase(),f=qe.find(g=>String(g.label||"").toLowerCase()===m);return f?f.value:null},d=c=>{if(!c)return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let m=String(c).trim();if(/^\d{4,5}$/.test(m)){let g=Number(m);if(g>2e4&&g<99999){let y=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);let f=m.split(/[\/\-\.]/);if(f.length===3){let[g,y,k]=f.map(v=>v.trim());if(g.length===4&&y.length<=2&&k.length<=2)return`${g}-${y.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&y.length<=2&&g.length<=2)return`${k}-${y.padStart(2,"0")}-${g.padStart(2,"0")}`}return m},h=n.map(c=>({meeting_date:d(c.Tanggal),employee_name:String(c["Nama Karyawan"]||"").trim(),branch_id:i(String(c.Cabang||"").trim()),pic:String(c.PIC||"").trim(),problem:String(c.Masalah||"").trim(),solution:String(c.Solusi||"").trim(),status:String(c.Status||"").trim(),completion_date:d(c["Tgl Selesai"]),document_link:String(c.Dokumen||"").trim()})).filter(c=>c.meeting_date&&c.employee_name&&c.branch_id),u=await S("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:h,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}},formFields:n=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:n?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:n?.branch_id&&!qe.find(i=>i.value==n.branch_id)?[...qe,{value:n.branch_id,label:n.branch_name||n.branch_id}]:qe,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:s(n?.employee_name),value:n?.employee_name},{name:"pic",label:"PIC",type:"select",options:o(n?.pic),createApi:{path:"/api/pic",field:"name"},value:n?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:n?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:n?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:n?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link}]})}j();async function zt(t){let e=await Q(),r=await le(),a=["Ade","Berlin"],l=n=>n&&!r.find(i=>i.value===n)?[...r,{value:n,label:n}]:r,s=n=>n&&!a.find(i=>(typeof i=="object"?i.value:i)===n)?[...a,n]:a,o=Array.from({length:5},(n,i)=>String(new Date().getFullYear()-i));H({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:n=>{try{let i=JSON.parse(n);return Array.isArray(i)?i.join(", "):n||"-"}catch{return n||"-"}}},{key:"score",label:"Nilai",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:o}],exportOptions:{moduleName:"training",onExport:async n=>{let i=new URLSearchParams(n||{}).toString(),d=await S(`/api/training?limit=10000&${i}`);if(d.ok){let h=d.data.data.map(c=>{let m=c.participants||"";try{let f=JSON.parse(m);m=Array.isArray(f)?f.join(", "):m}catch{}return{Tanggal:c.training_date||"",Batch:c.batch||"",Materi:c.subject||"",Cabang:c.branch_name||"",Trainer:c.trainer||"",Peserta:m,Nilai:c.score!==null&&c.score!==void 0?c.score:"",Dokumen:c.document_link||""}}),{downloadExcel:u}=await Promise.resolve().then(()=>(G(),ue));u(h,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(G(),ue));i(n,"Template_Import_Training")},onImport:async n=>{let i=c=>{if(!c)return null;let m=String(c||"").toLowerCase(),f=e.find(g=>String(g.label||"").toLowerCase()===m);return f?f.value:null},d=c=>{if(!c)return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let m=String(c).trim();if(/^\d{4,5}$/.test(m)){let g=Number(m);if(g>2e4&&g<99999){let y=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);let f=m.split(/[\/\-\.]/);if(f.length===3){let[g,y,k]=f.map(v=>v.trim());if(g.length===4&&y.length<=2&&k.length<=2)return`${g}-${y.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&y.length<=2&&g.length<=2)return`${k}-${y.padStart(2,"0")}-${g.padStart(2,"0")}`}return m},h=n.map(c=>({training_date:d(c.Tanggal),batch:String(c.Batch||"").trim(),subject:String(c.Materi||"").trim(),branch_id:i(String(c.Cabang||"").trim()),trainer:String(c.Trainer||"").trim(),participants:String(c.Peserta||"").trim(),score:c.Nilai?Number(c.Nilai):null,document_link:String(c.Dokumen||"").trim()})).filter(c=>c.training_date&&c.subject&&c.branch_id),u=await S("/api/import/training",{method:"POST",body:JSON.stringify({rows:h,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}},formFields:n=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:n?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:n?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:n?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:e,value:n?.branch_id},{name:"trainer",label:"Trainer",type:"select",options:s(n?.trainer),value:n?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let i=JSON.parse(n?.participants);return Array.isArray(i)?i.join(", "):n?.participants||""}catch{return n?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:n?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:n?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],onBeforeSubmit:async n=>(n.participants&&(n.participants=JSON.stringify(n.participants.split(",").map(i=>i.trim()).filter(Boolean))),n)})}j();we();G();function Qt({container:t,title:e,icon:r,apiPath:a,columns:l,formFields:s,filterFields:o,defaultFilters:n={},enableMobileFilterSheet:i=!1,itemLabel:d="Data",canCreate:h=!0,canEdit:u=!0,canDelete:c=!0,onBeforeSubmit:m,onAfterLoad:f,onDataLoaded:g,extraActions:y=[],initialSearch:k="",exportOptions:v=null,bulkDelete:E=!1,paginationMode:A="server"}){let B=ce();B&&typeof B=="object"&&B.role==="viewer"&&(h=!1,u=!1,c=!1,E=!1,v=null);let T=1,x={...n};k&&(x.search=k);let $=new Set;t.innerHTML=`
    <div class="crud-layout-wrapper ${i?"mobile-active":""}">
      <div class="page-header">
      <h1 class="page-title">${r} ${e}</h1>
      <div class="page-actions">
        ${h?`<button class="btn btn-primary" id="btn-create">+ Tambah ${d}</button>`:""}
        ${v?'<button class="btn btn-outline" id="btn-mobile-aksi" style="display:none; align-items:center; justify-content:center; gap:0.25rem;">\u22EE Aksi</button>':""}
      </div>
    </div>

    ${E?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:none; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${v?`
    <div class="excel-actions-wrapper" id="excel-actions-wrapper">
      <div class="bottom-sheet-header aksi-header" style="display:none;">
        <h3 style="margin:0; font-size:1rem;">Aksi</h3>
        <button class="btn-close-sheet" id="btn-close-aksi-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
      </div>
      ${Ie(v.moduleName)}
    </div>`:""}

    ${o&&o.length>0?`
    <div class="filter-bar card ${i?"has-mobile-sheet":""}" style="padding: 1rem; display: flex; gap: 0.5rem; flex-wrap: wrap;">
        ${o.filter(p=>p.type==="search"||p.type==="search-combo").map(p=>{if(p.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${p.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}"></div>`;if(p.type==="search-combo"){let b="dl-filter-search",w=(p.options||[]).map(_=>`<option value="${typeof _=="object"?_.label:_}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${b}" class="form-control" autocomplete="off" placeholder="${p.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}"><datalist id="${b}">${w}</datalist></div>`}return""}).join("")}
        
        <div class="filter-options" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${o.filter(p=>p.type!=="search"&&p.type!=="search-combo").map(p=>{if(p.type==="select")return`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${p.name}" id="filter-${p.name}"><option value="">Pilih ${p.label}</option>${(p.options||[]).map(b=>`<option value="${typeof b=="object"?b.value:b}" ${x[p.name]===(typeof b=="object"?b.value:b)?"selected":""}>${typeof b=="object"?b.label:b}</option>`).join("")}</select>`;if(p.type==="combobox"){let b=`dl-filter-${p.name}`,w=(p.options||[]).map(D=>{let C=typeof D=="object"?D.label||D.value||"":D||"";return(C==="undefined"||C==="[object Object]"||C==="null")&&(C=""),C?`<option value="${C}"></option>`:""}).join(""),_=x[p.name]||"";if(x[p.name]){let D=(p.options||[]).find(C=>String(typeof C=="object"?C.value:C)==String(x[p.name]));if(D){let C=typeof D=="object"?D.label||D.value||"":D||"";C&&C!=="undefined"&&C!=="[object Object]"&&C!=="null"&&(_=C)}}return`<div class="filter-combobox" style="flex:1; min-width:120px;">
                <input type="text" name="${p.name}" id="filter-${p.name}" list="${b}" class="form-control filter-combobox-input" value="${_}" placeholder="Pilih ${p.label}..." autocomplete="off">
                <datalist id="${b}">${w}</datalist>
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
  `;function F(){let p=document.getElementById("bulk-toolbar");if(!p)return;let b=document.getElementById("bulk-count"),w=document.getElementById("btn-bulk-delete"),_=document.getElementById("btn-bulk-cancel");b.textContent=`${$.size} item dipilih`,$.size>0?(p.style.display="flex",p.classList.add("has-items"),w.disabled=!1,_.disabled=!1,window.innerWidth<=768?(p.style.setProperty("top","auto","important"),p.style.setProperty("bottom","0","important")):(p.style.removeProperty("top"),p.style.removeProperty("bottom"))):(p.style.display="none",p.classList.remove("has-items"),w.disabled=!0,_.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{$.clear(),document.querySelectorAll(".row-checkbox").forEach(b=>b.checked=!1);let p=document.getElementById("select-all-checkbox");p&&(p.checked=!1),F()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if($.size===0)return;let p=[...$],b=document.createElement("div");b.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",b.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${p.length} ${d}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${p.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(b),b.querySelector("#bulk-cancel-btn").addEventListener("click",()=>b.remove()),b.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let w=b.querySelector("#bulk-confirm-btn");w.disabled=!0,w.textContent="Menghapus...";let _=await S(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:p})});b.remove(),_.ok?(ne(`${p.length} ${d} berhasil dihapus.`),$.clear(),F(),q()):z(_.data?.error||"Gagal menghapus data.")})});let O=document.getElementById("filter-search"),re;O?.addEventListener("input",p=>{clearTimeout(re),re=setTimeout(()=>{x.search=p.target.value,T=1,$.clear(),q()},400)}),o?.forEach(p=>{p.type==="select"&&document.getElementById(`filter-${p.name}`)?.addEventListener("change",b=>{x[p.name]=b.target.value,T=1,$.clear(),q()}),p.type==="combobox"&&document.getElementById(`filter-${p.name}`)?.addEventListener("change",b=>{let w=b.target.value,_=(p.options||[]).find(D=>{let C=String(typeof D=="object"?D.value:D),L=String(typeof D=="object"?D.label:D);return C===w||L===w});w?x[p.name]=_?typeof _=="object"?_.value:_:w:x[p.name]="",T=1,$.clear(),q()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{x={...n},O&&(O.value=""),o?.forEach(p=>{let b=document.getElementById(`filter-${p.name}`);b&&(b.value="")}),T=1,$.clear(),q()});let de=document.getElementById("btn-mobile-filter"),ie=document.getElementById("filter-options-wrapper"),ge=document.getElementById("btn-close-filter-sheet");if(de&&ie){de.addEventListener("click",b=>{b.preventDefault(),ie.classList.add("sheet-open")}),ge&&ge.addEventListener("click",b=>{b.preventDefault(),ie.classList.remove("sheet-open")});let p=document.getElementById("btn-apply-filter");p&&p.addEventListener("click",b=>{b.preventDefault(),ie.classList.remove("sheet-open")})}let me=document.getElementById("btn-mobile-aksi"),be=document.getElementById("excel-actions-wrapper"),_e=document.getElementById("btn-close-aksi-sheet");if(me&&be&&(me.addEventListener("click",p=>{p.preventDefault(),be.classList.add("sheet-open")}),_e&&_e.addEventListener("click",p=>{p.preventDefault(),be.classList.remove("sheet-open")})),document.getElementById("btn-create")?.addEventListener("click",()=>Ce(null)),v){document.getElementById(`btn-export-${v.moduleName}`)?.addEventListener("click",async b=>{let w=b.target,_=w.innerHTML;w.innerHTML="\u23F3 Loading...",w.disabled=!0;try{await v.onExport()}catch{z("Gagal export data")}finally{w.innerHTML=_,w.disabled=!1}}),document.getElementById(`btn-template-${v.moduleName}`)?.addEventListener("click",()=>{v.onTemplate()});let p=document.getElementById(`input-import-${v.moduleName}`);p?.addEventListener("change",async b=>{let w=b.target.files[0];if(!w)return;p.disabled=!0;let _=document.createElement("div");_.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",_.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(_);let D=_.querySelector("#import-progress-text"),C=_.querySelector("#import-progress-bar"),L=_.querySelector("#import-summary"),P=_.querySelector("#import-close-btn");P.addEventListener("click",()=>{_.remove(),q()});try{let W=await De(w);if(W.length===0)throw new Error("File kosong atau format salah");let V=500,ae=0,Z=0,N=0,M=W.length;D.textContent=`Ditemukan ${M} baris data. Memulai import...`;for(let I=0;I<M;I+=V){let J=W.slice(I,I+V);D.textContent=`Mengimport baris ${I+1} - ${Math.min(I+V,M)} dari ${M}...`,C.style.width=`${Math.round(I/M*100)}%`;try{let R=await v.onImport(J);R?(ae+=R.inserted||R.metrics?.inserted||J.length,Z+=R.skipped||R.metrics?.updated||0):ae+=J.length}catch(R){console.error("Chunk import failed:",R),N+=J.length}}C.style.width="100%",D.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',L.style.display="block",L.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${M}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${Z}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${N}</strong></div>
        `,N>0&&(L.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),P.style.display="block",p.value=""}catch(W){D.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${W.message}`,C.style.background="var(--danger)",C.style.width="100%",P.style.display="block",p.value=""}finally{p.disabled=!1}})}async function q(){F();let p=document.getElementById("table-container");if(!p)return;p.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let b=A==="client",w=b?1:T,_=b?Se:20,D=new URLSearchParams({page:w,limit:_,...Object.fromEntries(Object.entries(x).filter(([,N])=>N))}),C=await S(`${a}?${D}`);if(!C.ok){p.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${C.data?.error||"Error"}</p></div>`;return}let L=C.data?.data||C.data||[],P=C.data?.pagination,W=L.length,V=L;if(b){L=g(L),V=L;let N=L.length,M=20,I=Math.ceil(N/M);T>I&&I>0&&(T=I);let J=(T-1)*M,R=T*M;L=L.slice(J,R),P={page:T,limit:M,total:N,pages:I}}!1,f&&f(L);let ae=Be({columns:l,data:L,fullData:V,onEdit:u?N=>Ce(N):null,actions:y.map(N=>({...N,handler:M=>N.handler(M,q)})),emptyText:`Tidak ada ${String(d||"").toLowerCase()}`,bulkSelect:E?{selectedIds:$,onToggle:F}:null});p.innerHTML="",p.appendChild(ae);let Z=document.getElementById("pagination-container");if(Z&&(Z.innerHTML="",P&&P.pages>1)){let N=Ae({page:P.page,pages:P.pages,total:P.total,limit:P.limit,onPage:M=>{T=M,q()}});N&&Z.appendChild(N)}}function Ge(p){let b=typeof s=="function"?s(p):s;return Te(b)}function Ce(p){let b=!!p,w=document.createElement("form");if(w.noValidate=!0,w.innerHTML=Ge(p),b){let D=typeof s=="function"?s(p):s;Me(w,p)}let{close:_}=pe({title:b?`Edit ${d}`:`Tambah ${d}`,content:w,size:"lg",confirmText:b?"Simpan Perubahan":`Tambah ${d}`,onConfirm:async(D,C)=>{if(!w.reportValidity())return;let L=D.querySelector(".modal-confirm");L.disabled=!0,L.textContent="Menyimpan...";let P=Ne(w),W=typeof s=="function"?s(p):s,V=async M=>{for(let I of M)if(I.type==="row")await V(I.fields);else if(I.type==="combobox"&&P[I.name]){let J=P[I.name],R=(I.options||[]).find(U=>{let ee=String(typeof U=="object"?U.value:U),ze=String(typeof U=="object"?U.label:U);return ee===J||ze===J});if(R)P[I.name]=typeof R=="object"?R.value:R;else if(I.createApi){let U={};U[I.createApi.field]=J,I.createApi.extra&&Object.assign(U,I.createApi.extra);let ee=await S(I.createApi.path,{method:"POST",body:JSON.stringify(U)});if(ee.ok&&ee.data?.id)P[I.name]=ee.data.id;else if(ee.ok&&!ee.data?.id)P[I.name]=J;else throw new Error(`Gagal membuat master data: ${ee.data?.error||"Unknown error"}`)}}};try{await V(W)}catch(M){z(M.message),L.disabled=!1,L.textContent=b?"Simpan Perubahan":`Tambah ${d}`;return}m&&(P=await m(P,p));let ae=b?"PUT":"POST",Z=b?`${a}/${p.id}`:a,N=await S(Z,{method:ae,body:JSON.stringify(P)});N.ok?(ne(b?`${d} berhasil diperbarui.`:`${d} berhasil ditambahkan.`),C(),q()):(z(N.data?.error||"Gagal menyimpan data."),L.disabled=!1,L.textContent=b?"Simpan Perubahan":`Tambah ${d}`)}})}function $t(p){Ee(`Hapus ${d} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let b=await S(`${a}/${p.id}`,{method:"DELETE"});b.ok?(ne(`${d} berhasil dihapus.`),q()):z(b.data?.error||"Gagal menghapus.")},`Hapus ${d}`)}return q(),q}j();G();async function Wt(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let r=await Q(),a=await le(),l=e?e.get("dash_filter"):null,s={};if(l==="reliever"){let c=new Date,m=`${c.getFullYear()}-${String(c.getMonth()+1).padStart(2,"0")}`;s={status:"Done",month:e&&e.get("month")?e.get("month"):m}}console.log("RAW",await Oe()),console.log("OPTIONS",a);let o=c=>c&&!a.find(m=>m.value===c)?[...a,{value:c,label:c}]:a,n=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],i=c=>c&&!n.includes(c)?[...n,c]:n,d=new Date().getFullYear(),u=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"].map((c,m)=>{let f=String(m+1).padStart(2,"0");return{value:`${d}-${f}`,label:`${c} ${d}`}});Qt({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",defaultFilters:s,onDataLoaded:c=>c.sort((m,f)=>{let g=m.backup_date?new Date(m.backup_date).getTime():0;return(f.backup_date?new Date(f.backup_date).getTime():0)-g}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:c=>he(c)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:c=>window.formatDate(c)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:c=>window.formatDate(c)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:c=>c?`<span class="badge badge-info">${c}</span>`:"-"},{key:"status",label:"Status",render:c=>X(c)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:n},{type:"select",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"month",label:"Bulan",options:u},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:c=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:r,value:c?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:c?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:o(c?.original_fc_name),value:c?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:i(c?.reliever_name),value:c?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:c?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:c?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:c?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:c?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:c?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let c=await S(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(c.ok){let m=c.data.data.map(f=>({Cabang:f.branch_name||"","Nama Facility care":f.original_fc_name||"",Periode:f.period||"",Relifer:f.reliever_name||"","Tanggal Back Up":f.backup_date||"","Tanggal Selesai":f.completion_date||"",Keterangan:f.reason||"",Shift:f.shift||"",Status:f.status||""}));m.length===0&&m.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),K(m,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async c=>{let f=(await S("/api/branches?all=1")).data?.data||[],g=v=>{if(!v)return null;let E=String(v||"").toLowerCase(),A=f.find(B=>String(B.full_name||"").toLowerCase()===E||String(B.code||"").toLowerCase()===E||String(B.name||"").toLowerCase()===E);return A?A.id:null},y=c.map(v=>({branch_name:String(v.Cabang||"").trim(),backup_date:String(v["Tanggal Back Up"]||v["Tanggal Backup"]||"").trim(),original_fc_name:String(v["Nama Facility care"]||v["FC Digantikan"]||"").trim(),reliever_name:String(v.Relifer||v.Reliefer||"").trim(),period:String(v.Periode||"").trim(),reason:String(v.Keterangan||"").trim(),shift:String(v.Shift||"").trim(),completion_date:String(v["Tanggal Selesai"]||"").trim(),status:String(v.Status||"").trim()})).filter(v=>v.reliever_name&&v.backup_date),k=await S("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:y,onDuplicate:"update"})});if(!k.ok)throw new Error(k.data?.error||"Import gagal");return k.data}}})}j();G();async function Vt(t){let e=await Q(),r=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));H({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",enableMobileFilterSheet:!0,itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>he(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>X(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang..."},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),s=await S(`/api/reports/inspection?limit=10000&${l}`);if(s.ok){let o=s.data.data.map(n=>({Cabang:n.branch_name||"",Periode:n.period||"",Tanggal:n.inspection_date||"","Point FC":n.fc_score!==null&&n.fc_score!==void 0?n.fc_score:"","Point SPV":n.spv_score!==null&&n.spv_score!==void 0?n.spv_score:"",Status:n.status||"","Link Dokumen":n.document_link||""}));K(o,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let l=i=>{if(!i)return null;let d=String(i||"").toLowerCase(),h=e.find(u=>String(u.label||"").toLowerCase()===d);return h?h.value:null},s=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let d=String(i).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let u=Number(d);if(u>2e4&&u<99999){let c=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(c.getTime())?"":c.toISOString().slice(0,10)}}let h=d.split(/[\/\-\.]/);if(h.length===3){let[u,c,m]=h.map(f=>f.trim());if(u.length===4&&c.length<=2&&m.length<=2)return`${u}-${c.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&c.length<=2&&u.length<=2)return`${m}-${c.padStart(2,"0")}-${u.padStart(2,"0")}`}return d},o=a.map(i=>({branch_id:l(String(i.Cabang||"").trim()),period:String(i.Periode||"").trim(),inspection_date:s(i.Tanggal),fc_score:i["Point FC"]!==void 0&&i["Point FC"]!==""?Number(i["Point FC"]):null,spv_score:i["Point SPV"]!==void 0&&i["Point SPV"]!==""?Number(i["Point SPV"]):null,status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.period&&i.inspection_date),n=await S("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}j();G();async function Yt(t){let e=await Q(),r=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));H({container:t,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",enableMobileFilterSheet:!0,itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>he(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>X(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),s=await S(`/api/reports/cleaning?limit=10000&${l}`);if(s.ok){let o=s.data.data.map(n=>({Cabang:n.branch_name||"",Jenis:n.activity_type||"",Periode:n.period||"",Tanggal:n.activity_date||"",Status:n.status||"","Link Dokumen":n.document_link||""}));K(o,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let l=i=>{if(!i)return null;let d=String(i||"").toLowerCase(),h=e.find(u=>String(u.label||"").toLowerCase()===d);return h?h.value:null},s=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let d=String(i).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let u=Number(d);if(u>2e4&&u<99999){let c=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(c.getTime())?"":c.toISOString().slice(0,10)}}let h=d.split(/[\/\-\.]/);if(h.length===3){let[u,c,m]=h.map(f=>f.trim());if(u.length===4&&c.length<=2&&m.length<=2)return`${u}-${c.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&c.length<=2&&u.length<=2)return`${m}-${c.padStart(2,"0")}-${u.padStart(2,"0")}`}return d},o=a.map(i=>({branch_id:l(String(i.Cabang||"").trim()),activity_type:String(i.Jenis||i.Kegiatan||"").trim(),period:String(i.Periode||"").trim(),activity_date:s(i.Tanggal),status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.activity_type&&i.period&&i.activity_date),n=await S("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}j();G();async function Xt(t,e){let r=await Q(),a=Array.from({length:4},(o,n)=>String(new Date().getFullYear()-n)),l=e?e.get("dash_filter"):null,s={};if(l==="fogging"){let o=new Date,n=String(o.getMonth()+1).padStart(2,"0"),i=String(o.getFullYear()),d=e?e.get("month"):null;d&&d.length===7&&(i=d.split("-")[0],n=d.split("-")[1]),s={status:"Done",month:n,year:i}}H({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,enableMobileFilterSheet:!0,defaultFilters:s,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:o=>`<span class="badge badge-warning">${o}</span>`},{key:"period",label:"Periode",render:o=>he(o)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>X(o)},{key:"document_link",label:"Dokumen",render:o=>o?`<a href="${o}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:o=>o||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:r,value:o?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:o?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:o?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:o?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:o?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async o=>{let n=new URLSearchParams(o||{}).toString(),i=await S(`/api/reports/fogging?limit=10000&${n}`);if(i.ok){let d=i.data.data.map(h=>({Cabang:h.branch_name||"",Jenis:h.activity_type||"Fogging",Periode:h.period||"",Tanggal:h.activity_date||"",Status:h.status||"","Link Dokumen":h.document_link||""}));K(d,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async o=>{let n=u=>{if(!u)return null;let c=String(u||"").toLowerCase(),m=r.find(f=>String(f.label||"").toLowerCase()===c);return m?m.value:null},i=u=>{if(u==null||u==="")return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let c=String(u).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let f=Number(c);if(f>2e4&&f<99999){let g=new Date(Date.UTC(1899,11,30)+f*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}let m=c.split(/[\/\-\.]/);if(m.length===3){let[f,g,y]=m.map(k=>k.trim());if(f.length===4&&g.length<=2&&y.length<=2)return`${f}-${g.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&g.length<=2&&f.length<=2)return`${y}-${g.padStart(2,"0")}-${f.padStart(2,"0")}`}return c},d=o.map(u=>({branch_id:n(String(u.Cabang||"").trim()),activity_type:String(u.Jenis||u.Kegiatan||"Fogging").trim(),period:String(u.Periode||"").trim(),activity_date:i(u.Tanggal),status:String(u.Status||"").trim(),document_link:String(u["Link Dokumen"]||"").trim(),notes:String(u.Catatan||u.Keterangan||"").trim()})).filter(u=>u.branch_id&&u.period&&u.activity_date),h=await S("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(d)});if(!h.ok)throw new Error(h.data?.error||"Import gagal");return h.data}}})}j();G();async function Zt(t){let e=await Q(),r=await le(),a=r,l=Array.from({length:4},(n,i)=>String(new Date().getFullYear()-i)),s=n=>n&&!r.find(i=>i.value===n)?[...r,{value:n,label:n}]:r,o=n=>n&&!a.find(i=>i.value===n)?[...a,{value:n,label:n}]:a;H({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:n=>`<span title="${n||""}">${n?.length>60?n.slice(0,60)+"\u2026":n||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:n=>window.formatDate(n)},{key:"status",label:"Status",render:n=>X(n)},{key:"notes",label:"Keterangan",render:n=>n?.length>40?n.slice(0,40)+"\u2026":n||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade","Mizwar"]},{type:"select",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:l}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:e,value:n?.branch_id},{name:"pic",label:"PIC",type:"select",options:o(n?.pic),value:n?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:n?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:n?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:n?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:n?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async n=>{let i=new URLSearchParams(n||{}).toString(),d=await S(`/api/reports/basecamp?limit=10000&${i}`);if(d.ok){let h=d.data.data.map(u=>({"Tgl Info":u.info_date||"",Cabang:u.branch_name||"",Permasalahan:u.problem||"",PIC:u.pic||"","Tgl Done":u.done_date||"",Status:u.status||"",Keterangan:u.notes||""}));K(h,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async n=>{let i=c=>{if(!c)return null;let m=String(c||"").toLowerCase(),f=e.find(g=>String(g.label||"").toLowerCase()===m);return f?f.value:null},d=c=>{if(c==null||c==="")return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let m=String(c).trim();if(m===""||m==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);if(/^\d{4,5}$/.test(m)){let g=Number(m);if(g>2e4&&g<99999){let y=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}let f=m.split(/[\/\-\.]/);if(f.length===3){let[g,y,k]=f.map(v=>v.trim());if(g.length===4&&y.length<=2&&k.length<=2)return`${g}-${y.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&y.length<=2&&g.length<=2)return`${k}-${y.padStart(2,"0")}-${g.padStart(2,"0")}`}return m},h=n.map(c=>({info_date:d(c["Tgl Info"]||c["Tanggal Info"]),branch_id:i(String(c.Cabang||"").trim()),problem:String(c.Permasalahan||"").trim(),pic:String(c.PIC||"").trim(),done_date:d(c["Tgl Done"]||c["Tanggal Done"]),status:String(c.Status||"").trim(),notes:String(c.Keterangan||c.Catatan||"").trim()})).filter(c=>c.info_date&&c.branch_id&&c.problem),u=await S("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(h)});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}async function ea(t){H({container:t,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(j(),Le)),l=await a(`/api/sop?limit=10000&${r}`);if(l.ok){let s=l.data.data.map(n=>({"Nama SOP":n.name||"",Kategori:n.category||"",Dokumen:n.document_link||"",Catatan:n.notes||n.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(G(),ue));o(s,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:r}=await Promise.resolve().then(()=>(G(),ue));r(e,"Template_Import_SOP")},onImport:async e=>{let r=e.map(s=>({name:String(s["Nama SOP"]||"").trim(),category:String(s.Kategori||"").trim(),document_link:String(s.Dokumen||"").trim(),description:String(s.Catatan||"").trim()})).filter(s=>s.name),{apiFetch:a}=await Promise.resolve().then(()=>(j(),Le)),l=await a("/api/sop/import",{method:"POST",body:JSON.stringify(r)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function ta(t){H({container:t,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(j(),Le)),l=await a(`/api/checklist?limit=10000&${r}`);if(l.ok){let s=l.data.data.map(n=>({"Nama Checklist":n.name||"",Kategori:n.category||"",Dokumen:n.document_link||"",Deskripsi:n.description||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(G(),ue));o(s,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:r}=await Promise.resolve().then(()=>(G(),ue));r(e,"Template_Import_Checklist")},onImport:async e=>{let r=e.map(s=>({name:String(s["Nama Checklist"]||"").trim(),category:String(s.Kategori||"").trim(),document_link:String(s.Dokumen||"").trim(),description:String(s.Deskripsi||"").trim()})).filter(s=>s.name),{apiFetch:a}=await Promise.resolve().then(()=>(j(),Le)),l=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(r)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}j();we();G();async function Ct(t,e="forms"){if(e==="supply")return Ha(t);ja(t)}function ja(t){H({container:t,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await S(`/api/forms?limit=10000&${r}`);a.data?.data?K(a.data.data,"Data_Master_Form"):z("Gagal export data master form")},onImport:async e=>{let r=await S("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!r.ok)throw new Error(r.data?.error||"Import failed");return r.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function Ha(t){let r=((await S("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));H({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,l)=>l.branch_name_ref||l.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>X(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let l=a?.tools_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}let s=a?.chemical_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:a?.branch_id&&!r.find(o=>o.value==a.branch_id)?[...r,{value:a.branch_id,label:a.branch_name||a.branch_id}]:r,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:l},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:s},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),s=await S(`/api/reports/supply?limit=10000&${l}`);if(s.ok){let o=s.data.data.map(n=>{let i=n.tools_items;try{i=Array.isArray(JSON.parse(i))?JSON.parse(i).join(", "):i}catch{}let d=n.chemical_items;try{d=Array.isArray(JSON.parse(d))?JSON.parse(d).join(", "):d}catch{}return{Waktu:n.submitted_at||"",Pengirim:n.submitter_name||"",Cabang:n.branch_name_ref||n.branch_name||"","Alat/Barang":i||"",Chemical:d||"",Catatan:n.additional_notes||"",Status:n.status||"","Diproses Oleh":n.processed_by||""}});K(o,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let s=(await S("/api/branches?all=1")).data?.data||[],o=h=>{if(!h)return null;let u=String(h||"").toLowerCase(),c=s.find(m=>String(m.full_name||"").toLowerCase()===u||String(m.code||"").toLowerCase()===u||String(m.name||"").toLowerCase()===u);return c?c.id:null},n=h=>{if(h==null||h==="")return"";if(h instanceof Date&&!isNaN(h.getTime()))return h.toISOString().slice(0,10);let u=String(h).trim();if(u===""||u==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);if(/^\d{4,5}$/.test(u)){let m=Number(u);if(m>2e4&&m<99999){let f=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(f.getTime())?"":f.toISOString().slice(0,10)}}let c=u.split(/[\/\-\.]/);if(c.length===3){let[m,f,g]=c.map(y=>y.trim());if(m.length===4&&f.length<=2&&g.length<=2)return`${m}-${f.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&f.length<=2&&m.length<=2)return`${g}-${f.padStart(2,"0")}-${m.padStart(2,"0")}`}return u},i=a.map(h=>({submitted_at:n(h.Waktu||h.Tanggal),submitter_name:String(h.Pengirim||"").trim(),branch_id:o(String(h.Cabang||"").trim()),tools_items:String(h["Alat/Barang"]||h.Alat||"").trim(),chemical_items:String(h.Chemical||"").trim(),additional_notes:String(h.Catatan||h.Keterangan||"").trim(),status:String(h.Status||"").trim(),processed_by:String(h["Diproses Oleh"]||h.PIC||"").trim()})).filter(h=>h.submitted_at&&h.submitter_name&&h.branch_id),d=await S("/api/reports/supply/import",{method:"POST",body:JSON.stringify(i)});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,l)=>{let s=pe({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(o,n)=>{let i=o.querySelector("#supply-status").value,d=o.querySelector("#supply-processed-by").value;(await S(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:i,processed_by:d})})).ok?(ne("Status diperbarui."),n(),l()):z("Gagal update status.")}})}}]})}j();async function aa(t){let e=ce();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}H({container:t,title:"Riwayat Aktivitas",icon:"\u{1F575}\uFE0F\u200D\u2642\uFE0F",apiPath:"/api/audit-logs",enableMobileFilterSheet:!0,itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:r=>new Date(r).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(r,a)=>`<strong>${r||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${a.user_role||""}</small>`},{key:"action",label:"Aksi",render:r=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[r]||"badge-neutral"}">${r}</span>`},{key:"module",label:"Modul",render:r=>`<span style="text-transform:capitalize">${(r||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(r,a)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${r}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async r=>{try{let s=((await(await fetch(`/api/audit-logs?search=${r}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(d=>String(d.id)===String(r));if(!s)return alert("Data tidak ditemukan");let o=d=>{if(!d)return"Tidak ada data";try{return JSON.stringify(JSON.parse(d),null,2)}catch{return d}},n=`
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
       `,{createModal:i}=await Promise.resolve().then(()=>(we(),mt));i({title:`Detail Audit Log #${r}`,content:n,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}j();G();async function na(t){let e=ce();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}H({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,enableMobileFilterSheet:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:r=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[r]||"badge-neutral"}">${r}</span>`},{key:"is_active",label:"Status",render:r=>r?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:r=>r?new Date(r).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:r=>{let a=!!r;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:r?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:r?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:r?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:r?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?r?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let r=await S(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let a=r.data.data.map(l=>({"Nama Lengkap":l.full_name||"",Username:l.username||"",Email:l.email||"",Role:l.role||"",Status:l.is_active?"Aktif":"Nonaktif"}));K(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async r=>{let a=r.map(s=>({full_name:String(s["Nama Lengkap"]||"").trim(),username:String(s.Username||"").trim(),email:String(s.Email||"").trim(),role:String(s.Role||"").trim()||"viewer",password:String(s.Password||"").trim()})).filter(s=>s.username&&s.password&&s.email&&s.full_name),l=await S("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}j();G();async function ia(t){H({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",enableMobileFilterSheet:!0,itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await S(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)K(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{K([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let r=e.map(l=>({code:String(l["Kode Cabang"]||"").trim(),name:String(l["Nama Pendek"]||"").trim(),full_name:String(l["Nama Lengkap"]||"").trim(),city:String(l.Kota||"").trim()})).filter(l=>l.code&&l.name),a=await S("/api/branches/import",{method:"POST",body:JSON.stringify(r)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}j();async function ra(t){let e=new Date,r=[];t.innerHTML=`
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
    </div>`,await a();try{let o=e.getFullYear(),n=e.getMonth(),i=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),d=document.getElementById("cal-month-label");d&&(d.textContent=i);let h=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(A=>A.value)),u=r.filter(A=>h.has(A.type)),c={};u.forEach(A=>{let B=(A.event_date||"").slice(0,10);c[B]||(c[B]=[]),c[B].push(A)});let m=new Date(o,n,1).getDay(),f=new Date(o,n+1,0).getDate(),g=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],y=new Date().toISOString().slice(0,10),k='<div class="calendar-grid">';g.forEach(A=>{k+=`<div class="cal-day-header">${A}</div>`});for(let A=0;A<m;A++)k+='<div class="cal-cell cal-cell-empty"></div>';for(let A=1;A<=f;A++){let B=`${o}-${String(n+1).padStart(2,"0")}-${String(A).padStart(2,"0")}`,T=c[B]||[],x=B===y;k+=`
          <div class="cal-cell ${x?"cal-today":""} ${T.length?"cal-has-events":""}"
               data-date="${B}" tabindex="0" role="button" aria-label="${B}">
            <div class="cal-day-num ${x?"today-num":""}">${A}</div>
            <div class="cal-events-preview">
              ${T.slice(0,3).map($=>`
                <div class="cal-event-dot cal-color-${$.color||"gray"}" title="${lt($.title||$.type)}">
                  <span class="cal-event-dot-label">${qa($.title||$.branch_name||$.type,18)}</span>
                </div>
              `).join("")}
              ${T.length>3?`<div class="cal-more">+${T.length-3} lagi</div>`:""}
            </div>
          </div>`}let E=(m+f)%7;if(E!==0)for(let A=0;A<7-E;A++)k+='<div class="cal-cell cal-cell-empty"></div>';k+="</div>",s.innerHTML=k,s.querySelectorAll(".cal-cell[data-date]").forEach(A=>{A.addEventListener("click",()=>{let B=A.dataset.date,T=c[B]||[];if(!T.length)return;let x=document.getElementById("cal-event-list"),$=new Date(B+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=$,document.getElementById("cal-event-items").innerHTML=T.map(F=>`
            <div class="cal-event-item cal-color-border-${F.color||"gray"}">
              <div class="cal-event-type">${Ja(F.type)}</div>
              <div class="cal-event-title">${lt(F.title||"-")}</div>
              <div class="cal-event-branch">${lt(F.branch_name||"")}</div>
              ${F.status?`<div class="cal-event-status">${lt(F.status)}</div>`:""}
              ${F.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${F.days_remaining} hari</div>`:""}
            </div>
          `).join(""),x.style.display="block"})})}catch(o){console.error("[Calendar] Render error:",o),s&&(s.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}l()}function qa(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function lt(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Ja(t){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[t]||t}j();async function la(t){let e=ce(),r=(e?.full_name||e?.username||"U")[0].toUpperCase(),l={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
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
  `;let s=localStorage.getItem("fm_token"),o=document.getElementById("session-info");if(s&&o)try{let n=JSON.parse(atob(s.split(".")[1])),i=new Date(n.exp*1e3);o.textContent=`Berakhir: ${i.toLocaleString("id-ID")}`}catch{o.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async n=>{n.preventDefault();let i=document.getElementById("pwd-error"),d=document.getElementById("pwd-success"),h=document.getElementById("btn-save-pwd");i.style.display="none",d.style.display="none";let u=n.target,c=u.current_password.value,m=u.new_password.value,f=u.confirm_password.value;if(m!==f){i.textContent="\u274C Konfirmasi password tidak cocok.",i.style.display="block";return}if(m.length<6){i.textContent="\u274C Password baru minimal 6 karakter.",i.style.display="block";return}h.disabled=!0,h.textContent="\u23F3 Menyimpan...";let g=await S("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:c,new_password:m})});h.disabled=!1,h.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',g.ok?(d.textContent="\u2705 Password berhasil diubah.",d.style.display="block",u.reset(),ne("Password berhasil diubah.")):(i.textContent=g.data?.error||"Gagal mengubah password.",i.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}j();var st={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function se(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let l=Number(e);if(l>2e4&&l<99999){let s=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(s.getTime())?null:s.toISOString().slice(0,10)}}let r=e.split(/[\/\-\.]/);if(r.length===3){let[l,s,o]=r.map(h=>h.trim()),n=Number(l),i=Number(s),d=Number(o);if(l.length===4&&n>1900)return`${l}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`;if(o.length===4&&d>1900)return n>12?`${o}-${s.padStart(2,"0")}-${l.padStart(2,"0")}`:i>12?`${o}-${l.padStart(2,"0")}-${s.padStart(2,"0")}`:`${o}-${s.padStart(2,"0")}-${l.padStart(2,"0")}`;if(o.length===2&&!isNaN(d)){let h=d>=50?`19${o}`:`20${o}`;return n>12?`${h}-${s.padStart(2,"0")}-${l.padStart(2,"0")}`:`${h}-${s.padStart(2,"0")}-${l.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function sa(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Ua={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",phone:t["No. Hp"],join_date:se(t["Tanggal Masuk"]),status:t.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:se(t["Tanggal Mulai"]),end_date:se(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:se(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:se(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:se(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:se(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:se(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:se(t["Tanggal Target"]||t["Tgl Target"]),completion_date:se(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:se(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:se(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:se(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:se(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:se(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:se(t["Tanggal Back Up"]),completion_date:se(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:se(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:se(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Ga(t,e){let r=st[t];if(!r)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Ua[r.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let l=[],s=[],o=[];return e.filter(i=>!sa(i)).forEach((i,d)=>{let h=e.indexOf(i)+2,u=[];a.required.forEach(({key:m,label:f})=>{let g=i[m];if(g==null||String(g).trim()===""){let y=Object.keys(i).filter(k=>k.trim()).join(", ");u.push({column:f,originalValue:g||"",reason:`Kolom "${f}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${y.slice(0,120)}`})}});let c=a.map(i);u.length>0?s.push({row:h,data:c,raw:i,errors:u}):(l.push(i),o.push(c))}),{valid:l,errors:s,mapped:o}}function oa(t){let e=[];return t.SheetNames.forEach(r=>{let a=st[r];if(!a)return;let l=t.Sheets[r],s=window.XLSX.utils.sheet_to_json(l,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),o=Ga(r,s),n=s.filter(i=>!sa(i));e.push({sheetName:r,module:a.module,label:a.label,total:n.length,valid:o.mapped.length,errorCount:o.errors.length,errors:o.errors,mapped:o.mapped,skipped:!1})}),e}function da(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,l])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(l),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function ca(t){let e=window.XLSX,r=e.utils.book_new(),a=!1;return t.forEach(l=>{if(!l.errors||l.errors.length===0)return;a=!0;let s=l.errors.map(n=>({"No. Baris":n.row,"Kolom Gagal":(n.errors||[]).map(i=>i.column||i).join("; "),"Alasan Error":(n.errors||[]).map(i=>i.reason||i).join("; "),...Object.fromEntries(Object.entries(n.data||{}).map(([i,d])=>[i,d??""]))})),o=e.utils.json_to_sheet(s);e.utils.book_append_sheet(r,o,l.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(r,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var za=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function pa(t){t.innerHTML=`
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
              ${Object.entries(st).map(([g,{label:y}])=>`<span class="import-sheet-tag">\u{1F4C4} ${g} \u2192 ${y}</span>`).join("")}
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
  `;let e=null,r=null,a=0,l={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function s(g){Object.entries(l).forEach(([y,k])=>{k.style.display=y===g?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let g=document.getElementById("btn-backup-db");g.disabled=!0,g.textContent="\u23F3 Memproses Backup...";try{let y=await S("/api/import/backup");if(y.ok){if(!window.XLSX){z("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let k=window.XLSX,v=k.utils.book_new();Object.entries(y.data.database).forEach(([E,A])=>{let B=A.length>0?A:[{}],T=k.utils.json_to_sheet(B);k.utils.book_append_sheet(v,T,E.substring(0,31))}),k.writeFile(v,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),ne("Backup berhasil diunduh!")}else z("Gagal memproses backup: "+(y.data?.error||"Unknown error"))}catch(y){z("Gagal memproses backup: "+y.message)}finally{g.disabled=!1,g.textContent="\u{1F4E6} Backup Database"}});let o=document.getElementById("btn-sync-google");o&&o.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let g=o.innerHTML;o.innerHTML='<span class="spinner"></span> Menyinkronkan...',o.disabled=!0;try{let y=await S("/api/sync/google-sheets",{method:"POST"});y.ok?alert("Sinkronisasi Berhasil: "+(y.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(y.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{o.innerHTML=g,o.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{da(),ne("Template Excel berhasil didownload!")});let n=document.getElementById("file-input"),i=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",g=>{g.stopPropagation(),n.click()}),n.addEventListener("change",g=>{g.target.files[0]&&d(g.target.files[0])}),i.addEventListener("dragover",g=>{g.preventDefault(),i.classList.add("drag-over")}),i.addEventListener("dragleave",()=>i.classList.remove("drag-over")),i.addEventListener("drop",g=>{g.preventDefault(),i.classList.remove("drag-over");let y=g.dataTransfer.files[0];y&&y.name.match(/\.xlsx?$/i)?d(y):z("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,n.value="",document.getElementById("file-info").style.display="none",i.style.display="",s("upload")});async function d(g){e=g,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${g.name} (${(g.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",i.style.display="none",await h(g)}async function h(g){s("validating");let y=document.getElementById("validation-status"),k=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");y.textContent="Membaca file Excel...",k.style.width="20%",await et(200);let v=await g.arrayBuffer(),E=window.XLSX.read(v,{type:"array",cellDates:!0});y.textContent=`Memvalidasi ${E.SheetNames.length} sheet...`,k.style.width="50%",await et(100),r=oa(E),k.style.width="100%",y.textContent="Validasi selesai!",await et(300),u()}catch(v){s("upload"),z("Gagal memproses file: "+v.message),document.getElementById("file-info").style.display="flex",i.style.display="none"}}function u(){s("preview");let g=r.filter($=>!$.skipped).length,y=r.reduce(($,F)=>$+F.total,0),k=r.reduce(($,F)=>$+F.valid,0),v=r.reduce(($,F)=>$+F.errorCount,0),E=y>0?Math.round(k/y*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${g} sheet</span>
      <span class="badge badge-secondary">${y} baris</span>
      <span class="badge badge-success">${k} valid (${E}%)</span>
      ${v>0?`<span class="badge badge-danger">${v} error</span>`:""}
    `;let A=document.getElementById("preview-table-container");A.innerHTML=`
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
          ${r.map(($,F)=>`
            <tr class="${$.errorCount>0?"row-error":$.skipped?"row-skipped":"row-ok"}">
              <td><strong>${$.sheetName}</strong></td>
              <td>${$.label}</td>
              <td style="text-align:center">${$.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${$.valid}</span></td>
              <td style="text-align:center">${$.errorCount>0?`<span class="badge badge-danger">${$.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${$.skipped?'<span class="badge badge-neutral">Dilewati</span>':$.errorCount>0&&$.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':$.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':$.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${$.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${F}">\u{1F50D} ${$.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,A.querySelectorAll(".btn-detail-error").forEach($=>{$.addEventListener("click",()=>{let F=r[Number($.dataset.idx)];c(F)})});let B=document.getElementById("error-detail-section"),T=document.getElementById("error-detail-container");T.innerHTML="",B.style.display="none";let x=document.getElementById("btn-start-import");k===0?(x.disabled=!0,x.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(x.disabled=!1,v>0?(x.innerHTML=`\u{1F680} Import ${k} Data Valid (${v} dilewati)`,x.title="Baris error akan dilewati, baris valid tetap diimport"):x.innerHTML=`\u{1F680} Mulai Import ${k} Data`)}function c(g){let y=document.getElementById("error-detail-section"),k=document.getElementById("error-detail-container");y.style.display="";let v=g.errors.slice(0,100).map(E=>(Array.isArray(E.errors)?E.errors:[]).map(B=>{let T=typeof B=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${E.row}</span></td>
            <td><strong>${T?B.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${T&&B.originalValue!==void 0?B.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${T?B.reason:B}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${T&&B.aliases?`Gunakan salah satu nama kolom:<br><em>${B.aliases}</em>`:T&&B.hint?B.hint:""}
            </td>
          </tr>
        `}).join("")).join("");k.innerHTML=`
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
            <tbody>${v||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${g.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,y.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{s("upload"),document.getElementById("file-info").style.display="none",i.style.display="",e=null,n.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!r)return;ca(r)?ne("Log error berhasil didownload."):ne("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let g=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";m(g)});async function m(g){s("importing"),a=Date.now();let y=[];za.forEach(B=>{let T=r?.find(x=>x.module===B&&x.mapped?.length>0);T&&y.push(T)});let k=document.getElementById("import-steps-list");k.innerHTML=y.map(B=>`
      <div class="import-step-item" id="step-item-${B.module}">
        <span class="step-item-icon" id="step-icon-${B.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${B.label} <span class="step-item-count">(${B.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${B.module}"></span>
      </div>
    `).join("");let v=document.getElementById("import-bar"),E=document.getElementById("import-current-status"),A={totalSheets:y.length,totalRows:y.reduce((B,T)=>B+T.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let B=0;B<y.length;B++){let T=y[B],x=document.getElementById(`step-icon-${T.module}`),$=document.getElementById(`step-status-${T.module}`);x.textContent="\u{1F504}",$.textContent="Mengimport...",E.textContent=`Mengimport ${T.label}...`,v.style.width=`${Math.round(B/y.length*100)}%`;try{let F=await S(`/api/import/${T.module}`,{method:"POST",body:JSON.stringify({rows:T.mapped,onDuplicate:g})});if(F.ok){let O=F.data;A.inserted+=O.inserted||0,A.skipped+=O.skipped||0,A.moduleResults.push({label:T.label,inserted:O.inserted||0,skipped:O.skipped||0,status:"ok"}),x.textContent="\u2705",$.innerHTML=`<span class="badge badge-success">${O.inserted||0} berhasil</span>${O.skipped>0?` <span class="badge badge-neutral">${O.skipped} skip</span>`:""}`}else A.failed++,A.moduleResults.push({label:T.label,inserted:0,skipped:0,status:"error",error:F.data?.error}),x.textContent="\u274C",$.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(F){A.failed++,A.moduleResults.push({label:T.label,inserted:0,skipped:0,status:"error",error:F.message}),x.textContent="\u274C",$.innerHTML='<span class="badge badge-danger">Gagal</span>'}await et(150)}v.style.width="100%",E.textContent="Selesai!",await et(400),f(A)}function f(g){s("summary");let y=((Date.now()-a)/1e3).toFixed(1),k=g.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${k?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${k?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
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
        <div class="stat-value">${y}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${g.moduleResults.map(v=>`
            <tr>
              <td>${v.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${v.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${v.skipped}</span></td>
              <td style="text-align:center">
                ${v.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${v.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,r=null,n.value="",document.getElementById("file-info").style.display="none",i.style.display="",s("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function et(t){return new Promise(e=>setTimeout(e,t))}j();var ot=[],ua=[];async function ma(t){ot=await Q(),ua=await le(),H({container:t,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",enableMobileFilterSheet:!0,itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:ot}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await S(`/api/sp?limit=10000&${r}`);if(a.ok){let l=a.data.data.map(o=>({"Nama Karyawan":o.employee_name||"",Divisi:o.division||"",Cabang:o.branch_name||"","Tanggal Sp":o.tanggal||"","Akhir Sp":o.akhir_sp||"","Jenis Sp":o.sp_type||"","Link Document / Foto":o.document_link||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(G(),ue));s(l,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(G(),ue));r(e,"Template_Import_SP")},onImport:async e=>{let r=o=>{if(!o)return null;let n=String(o||"").toLowerCase(),i=ot.find(d=>String(d.label||"").toLowerCase()===n);return i?i.value:null},a=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let n=String(o).trim();if(/^\d{4,5}$/.test(n)){let d=Number(n);if(d>2e4&&d<99999){let h=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(n))return n.slice(0,10);let i=n.split(/[\/\-\.]/);if(i.length===3){let[d,h,u]=i.map(c=>c.trim());if(d.length===4&&h.length<=2&&u.length<=2)return`${d}-${h.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&h.length<=2&&d.length<=2)return`${u}-${h.padStart(2,"0")}-${d.padStart(2,"0")}`}return n},l=e.map(o=>({employee_name:String(o["Nama Karyawan"]||"").trim(),division:String(o.Divisi||"").trim(),branch_id:r(String(o.Cabang||"").trim()),tanggal:a(o["Tanggal Sp"]),akhir_sp:a(o["Akhir Sp"]),sp_type:String(o["Jenis Sp"]||"").trim(),document_link:String(o["Link Document / Foto"]||"").trim()})).filter(o=>o.employee_name&&o.branch_id),s=await S("/api/import/sp",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:[{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:ua},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:ot,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}j();var Je=[],ga=[];async function ba(t){Je=await Q(),ga=await le(),H({container:t,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",enableMobileFilterSheet:!0,itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:Je},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:Je}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await S(`/api/mutasi?limit=10000&${r}`);if(a.ok){let l=a.data.data.map(o=>({Tanggal:o.tanggal||"","Nama Karyawan":o.employee_name||"","Cabang Asal":o.from_branch_name||"","Cabang Tujuan":o.to_branch_name||"",Status:o.status||"",Dokumen:o.document_link||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(G(),ue));s(l,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(G(),ue));r(e,"Template_Import_Mutasi")},onImport:async e=>{let r=o=>{if(!o)return null;let n=String(o||"").toLowerCase(),i=Je.find(d=>String(d.label||"").toLowerCase()===n);return i?i.value:null},a=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let n=String(o).trim();if(/^\d{4,5}$/.test(n)){let d=Number(n);if(d>2e4&&d<99999){let h=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(n))return n.slice(0,10);let i=n.split(/[\/\-\.]/);if(i.length===3){let[d,h,u]=i.map(c=>c.trim());if(d.length===4&&h.length<=2&&u.length<=2)return`${d}-${h.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&h.length<=2&&d.length<=2)return`${u}-${h.padStart(2,"0")}-${d.padStart(2,"0")}`}return n},l=e.map(o=>({tanggal:a(o.Tanggal),employee_name:String(o["Nama Karyawan"]||"").trim(),from_branch_id:r(String(o["Cabang Asal"]||"").trim()),to_branch_id:r(String(o["Cabang Tujuan"]||"").trim()),status:String(o.Status||"").trim(),document_link:String(o.Dokumen||"").trim()})).filter(o=>o.tanggal&&o.employee_name&&o.from_branch_id&&o.to_branch_id),s=await S("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"select",name:"employee_name",label:"Nama Karyawan",required:!0,options:ga},{type:"select",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Je,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Je,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}j();async function ha(t){t.innerHTML=`
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
  `;let e=document.getElementById("btnRefreshSync"),r=document.getElementById("queueStatusFilter");e.addEventListener("click",l),r.addEventListener("change",n),document.getElementById("btnRetryAll").addEventListener("click",()=>a("retry",{allFailed:!0})),document.getElementById("btnResetStuck").addEventListener("click",()=>a("reset-stuck",{})),document.getElementById("btnPauseSync").addEventListener("click",()=>a("pause",{})),document.getElementById("btnResumeSync").addEventListener("click",()=>a("resume",{})),document.getElementById("btnReconcile").addEventListener("click",()=>a("reconcile",{module:"ALL",repairMode:!0})),document.getElementById("btnRetrySelected").addEventListener("click",()=>{let d=Array.from(document.querySelectorAll(".chk-queue:checked")).map(h=>h.value);if(d.length===0)return alert("No items selected");a("retry",{ids:d})}),document.getElementById("chkAllQueue").addEventListener("change",d=>{document.querySelectorAll(".chk-queue").forEach(h=>h.checked=d.target.checked)});async function a(d,h){if(confirm(`Are you sure you want to execute action: ${d}?`)){showLoading();try{let u=await S(`/api/sync/actions/${d}`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(h)});u.ok?(alert(u.data?.message||"Success"),l()):z(u.error||"Action failed")}catch(u){z(u.message)}hideLoading()}}await l();async function l(){showLoading(),await Promise.all([o(),n(),s(),i()]),hideLoading()}async function s(){try{let d=await S("/api/sync/performance");if(!d.ok)return;let{webhook:h,google_api:u,d1:c,queue:m,throughput:f}=d.data;document.getElementById("latencyHistograms").innerHTML=`
        <table class="w-full text-left">
          <thead><tr class="text-gray-500 border-b"><th>Metric</th><th>P50</th><th>P95</th><th>P99</th><th>Max</th></tr></thead>
          <tbody class="divide-y">
            <tr><td class="py-1">Webhook</td><td>${h.P50}ms</td><td>${h.P95}ms</td><td>${h.P99}ms</td><td>${h.Max}ms</td></tr>
            <tr><td class="py-1">Google API</td><td>${u.P50}ms</td><td>${u.P95}ms</td><td>${u.P99}ms</td><td>${u.Max}ms</td></tr>
            <tr><td class="py-1">D1 Execute</td><td>${c.P50}ms</td><td>${c.P95}ms</td><td>${c.P99}ms</td><td>${c.Max}ms</td></tr>
            <tr><td class="py-1">Queue Delay</td><td>${m.P50}ms</td><td>${m.P95}ms</td><td>${m.P99}ms</td><td>${m.Max}ms</td></tr>
          </tbody>
        </table>
        <div class="mt-4 pt-3 border-t text-gray-600 flex justify-between">
          <span>Throughput: <b>${f.events_per_sec}</b> ev/sec</span>
          <span><b>${f.events_per_min}</b> ev/min</span>
        </div>
      `,document.getElementById("lblSnapshotStatus").innerHTML='Checked 10 modules. Status: <b>COMPLETED</b><br><span class="text-xs text-gray-500">Run today at 02:00</span>'}catch(d){console.error(d)}}async function o(){try{let d=await S("/api/sync/health");if(!d.ok)return z("Failed to fetch sync health");let{status:h,queue:u,circuit_breaker:c}=d.data,m=`
        <div class="bg-white p-4 rounded-lg shadow border-l-4 ${h==="HEALTHY"?"border-green-500":h==="WARNING"?"border-yellow-500":"border-red-500"}">
          <p class="text-sm text-gray-500">System Health</p>
          <p class="text-2xl font-bold ${h==="HEALTHY"?"text-green-600":h==="WARNING"?"text-yellow-600":"text-red-600"}">${h}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-blue-500">
          <p class="text-sm text-gray-500">Pending Queue</p>
          <p class="text-2xl font-bold text-gray-800">${u.pending||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-orange-500">
          <p class="text-sm text-gray-500">Failed / Retrying</p>
          <p class="text-2xl font-bold text-gray-800">${u.failed||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-red-500">
          <p class="text-sm text-gray-500">Dead Letters</p>
          <p class="text-2xl font-bold text-gray-800">${u.dead_letter||0}</p>
        </div>
      `;document.getElementById("syncOverviewCards").innerHTML=m;let f=document.getElementById("cbStateBadge"),g=document.getElementById("cbStateDesc"),y=document.getElementById("cbStatusCard");y.className="bg-white rounded-lg shadow p-6 border-l-4",c==="CLOSED"?(f.className="px-3 py-1 rounded-full text-sm font-medium bg-green-100 text-green-800",f.textContent="CLOSED",g.textContent="Traffic is flowing normally to Google Sheets.",y.classList.add("border-green-500")):c==="OPEN"?(f.className="px-3 py-1 rounded-full text-sm font-medium bg-red-100 text-red-800",f.textContent="OPEN",g.textContent="Failing fast. Traffic to Google Sheets is paused due to repeated failures.",y.classList.add("border-red-500")):c==="HALF_OPEN"?(f.className="px-3 py-1 rounded-full text-sm font-medium bg-yellow-100 text-yellow-800",f.textContent="HALF-OPEN",g.textContent="Testing recovery. Permitting limited traffic to verify stability.",y.classList.add("border-yellow-500")):f.textContent=c||"UNKNOWN"}catch(d){console.error(d)}}async function n(){try{let d=document.getElementById("queueStatusFilter").value,h=await S("/api/sync/queue?limit=15"+(d?"&status="+d:""));if(!h.ok)return;let u=document.getElementById("queueTableBody"),c=h.data?.data||h.data||[];if(c.length===0){u.innerHTML='<tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">No events found</td></tr>';return}u.innerHTML=c.map(m=>`
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
      `).join("")}catch(d){console.error(d)}}async function i(){try{let d=await S("/api/sync/metrics");if(!d.ok)return;let h=document.getElementById("metricsTableBody"),u=d.data||[];if(u.length===0){h.innerHTML='<tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">No metrics found in last 24h</td></tr>';return}h.innerHTML=u.map(c=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2 font-medium text-gray-800">${c.module}</td>
          <td class="px-4 py-2 text-gray-600">${c.total_events}</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(c.avg_webhook_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(c.avg_d1_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(c.avg_queue_wait_ms||0)} ms</td>
        </tr>
      `).join("")}catch(d){console.error(d)}}}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let e=Math.floor(Number(t)-25569);return new Date(e*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let e=t.split(/[\/\-]/);return`${e[2]}-${e[1]}-${e[0]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let r=e.split("-");if(r.length===3&&r[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],l=parseInt(r[2],10),s=a[parseInt(r[1],10)-1];return`${l} ${s} ${r[0]}`}return e};function te(t){return async e=>{if(!Qe()){Pe("/login");return}return t(e)}}var tt=null;function Qa(){tt&&clearInterval(tt);let t=()=>{let e=new Date,r=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),l=document.getElementById("header-clock-time"),s=document.getElementById("header-clock-date");l&&(l.textContent=r),s&&(s.textContent=a)};t(),tt=setInterval(t,1e3)}async function Wa(){try{let t=await S("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},r=(a,l)=>{let s=document.getElementById(a);s&&(s.textContent=l>0?l:"",s.style.display=l>0?"inline-flex":"none")};r("badge-issues",e.issues?.current||0),r("badge-contracts",e.expiring30?.current||0),r("badge-oo1",e.one_on_one?.current||0),r("badge-schedule",e.schedule?.current||0),r("badge-supply",e.supply?.current||0)}catch{}}var Ue=[];async function Va(){try{let t=await S("/api/dashboard/notifications");if(!t.ok)return;Ue=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Ue.length>0?"block":"none",e.textContent=Ue.length)}catch{}}function Ya(){if(!Ue.length){pe({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,r)=>r()});return}let t=`
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
  `;pe({title:`Notifikasi (${Ue.length})`,content:t,confirmText:"Tutup",onConfirm:(e,r)=>r()})}function ya(){let t=ce(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
                <span class="topbar-greeting-time">${(()=>{let d=new Date().getHours();return d>=4&&d<11?"Selamat Pagi":d>=11&&d<15?"Selamat Siang":d>=15&&d<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">${t?.full_name||"Admin"}</span> \u{1F44B}
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
  `;let r=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),l=document.getElementById("topbar-menu-btn"),s=document.getElementById("sidebar-close"),o=()=>{r.classList.add("open"),a.classList.add("show")},n=()=>{r.classList.remove("open"),a.classList.remove("show")};l?.addEventListener("click",o),s?.addEventListener("click",n),a?.addEventListener("click",n),document.querySelectorAll(".nav-item").forEach(d=>d.addEventListener("click",n));function i(){let d=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(c=>{let m=c.dataset.route;c.classList.toggle("active",d===m||m!=="/dashboard"&&d.startsWith(m))});let h=document.getElementById("topbar-title"),u=document.querySelector(".nav-item.active .nav-label");h&&u&&(h.textContent=u.textContent)}window.addEventListener("hashchange",i),i(),Qa(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await S("/api/auth/logout",{method:"POST"}),We(),tt&&clearInterval(tt),Pe("/login")}),Wa(),Va(),document.getElementById("btn-notif")?.addEventListener("click",d=>{d.preventDefault(),Ya()})}document.addEventListener("keydown",t=>{if(t.target.tagName==="SELECT"&&t.key.length===1&&/[a-zA-Z]/.test(t.key)){let e=t.key.toLowerCase(),r=Array.from(t.target.options);if(r.length===0)return;let a=t.target.selectedIndex+1;(a>=r.length||a<0)&&(a=0);let l=-1;for(let s=0;s<r.length;s++){let o=(a+s)%r.length;if(r[o].text.toLowerCase().replace(/^[\d\.\s\-]+/,"").startsWith(e)){l=o;break}}l!==-1&&l!==t.target.selectedIndex&&(t.target.selectedIndex=l,t.preventDefault(),t.target.dispatchEvent(new Event("change")))}});async function Xa(){Y("/login",({main:e})=>Kt(e)),Y("/dashboard",te(({main:e})=>Nt(e))),Y("/calendar",te(({main:e})=>ra(e))),Y("/employees",te(({main:e,params:r})=>jt(e,r))),Y("/contracts",te(({main:e,params:r})=>qt(e,r))),Y("/sp",te(({main:e})=>ma(e))),Y("/mutasi",te(({main:e})=>ba(e))),Y("/sync-dashboard",te(({main:e})=>ha(e))),Y("/timeline",te(({main:e,params:r})=>Jt(e,r))),Y("/issues",te(({main:e,params:r})=>Ut(e,r))),Y("/one-on-one",te(({main:e,params:r})=>Gt(e,r))),Y("/training",te(({main:e})=>zt(e))),Y("/relievers",te(({main:e,params:r})=>Wt(e,r))),Y("/reports/inspection",te(({main:e})=>Vt(e))),Y("/reports/cleaning",te(({main:e})=>Yt(e))),Y("/reports/fogging",te(({main:e})=>Xt(e))),Y("/reports/basecamp",te(({main:e})=>Zt(e))),Y("/reports/supply",te(({main:e})=>Ct(e,"supply"))),Y("/sop",te(({main:e})=>ea(e))),Y("/checklist",te(({main:e})=>ta(e))),Y("/forms",te(({main:e})=>Ct(e))),Y("/users",te(({main:e})=>na(e))),Y("/branches",te(({main:e})=>ia(e))),Y("/profile",te(({main:e})=>la(e))),Y("/settings/import",te(({main:e})=>pa(e))),Y("/audit-logs",te(({main:e})=>aa(e)));let t=Qe();if(!t&&window.location.hash!=="#/login"&&Pe("/login"),t){let e=await S("/api/auth/me");e.ok?(Ve(e.data.data),ya()):(We(),Pe("/login"))}window.addEventListener("fm:login",()=>{ya(),Pe("/dashboard")}),Tt()}Xa();
