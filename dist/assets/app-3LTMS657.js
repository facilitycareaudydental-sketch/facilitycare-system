var oa=Object.defineProperty;var at=(t,e)=>()=>(t&&(e=t(t=0)),e);var nt=(t,e)=>{for(var l in e)oa(t,l,{get:e[l],enumerable:!0})};var ke={};nt(ke,{API:()=>yt,CLIENT_SIDE_MAX_ROWS:()=>be,IS_DEVELOPMENT:()=>Pe,apiFetch:()=>k,clearToken:()=>Le,getToken:()=>Be,getUser:()=>he,setToken:()=>it,setUser:()=>Ne});function Be(){return localStorage.getItem("fm_token")}function it(t){localStorage.setItem("fm_token",t)}function Le(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function he(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Ne(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function k(t,e={}){let l=Be(),a={"Content-Type":"application/json",...l?{Authorization:`Bearer ${l}`}:{},...e.headers||{}};try{let s=`cb=${Date.now()}`,d=t.includes("?")?"&":"?",r=`${yt}${t}${d}${s}`,n=await fetch(r,{...e,headers:a}),i;try{let c=await n.text();try{i=JSON.parse(c)}catch{i={error:`Server Error (${n.status}): ${c.substring(0,80)}...`}}}catch{i={error:"Gagal membaca respon dari server"}}return n.status===401&&(Le(),window.location.hash="#/login"),{ok:n.ok,status:n.status,data:i}}catch(s){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${s.message})`}}}}var Pe,da,yt,be,M=at(()=>{Pe=!1,da="https://fm-operations-api.facilitycare-audydental.workers.dev",yt=da,be=1e4});var kt={};nt(kt,{confirmDialog:()=>Fe,createModal:()=>se});function se({title:t,content:e,onConfirm:l,onCancel:a,confirmText:s="Simpan",cancelText:d="Batal",size:r="md",confirmClass:n="btn-primary"}){let i={sm:"400px",md:"560px",lg:"720px",xl:"900px"},c=document.createElement("div");c.className="modal-overlay",c.innerHTML=`
    <div class="modal" style="max-width:${i[r]||i.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${d}</button>
        ${l?`<button class="btn ${n} modal-confirm">${s}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&c.querySelector(".modal-body").appendChild(e);let m=()=>{c.classList.remove("show"),setTimeout(()=>c.remove(),250)};return c.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),m()}),c.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),m()}),l&&c.querySelector(".modal-confirm").addEventListener("click",()=>l(c,m)),c.addEventListener("click",p=>{p.target===c&&(a&&a(),m())}),document.body.appendChild(c),requestAnimationFrame(()=>c.classList.add("show")),{overlay:c,close:m}}function Fe(t,e,l="Konfirmasi"){return se({title:l,content:`<p>${t}</p>`,onConfirm:(a,s)=>{e(),s()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var ye=at(()=>{});var le={};nt(le,{downloadExcel:()=>B,parseExcel:()=>Oe,renderExcelButtons:()=>Re});function Oe(t){return new Promise((e,l)=>{let a=new FileReader;a.onload=s=>{try{let d=new Uint8Array(s.target.result),r=XLSX.read(d,{type:"array"}),n=r.SheetNames[0],i=r.Sheets[n];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${r.SheetNames.join(", ")}`),console.log(`Sheet Used: ${n}`);let c=XLSX.utils.decode_range(i["!ref"]||"A1:A1"),m=c.e.r-c.s.r+1,p=c.e.c-c.s.c+1;console.log(`Total Rows (including empty): ${m}`),console.log(`Total Columns: ${p}`);let o=[];for(let y=c.s.c;y<=c.e.c;++y){let u=i[XLSX.utils.encode_cell({c:y,r:c.s.r})];u&&u.v&&o.push(u.v)}console.log(`Headers Found: ${o.join(", ")}`),console.log("---------------------------");let g=XLSX.utils.sheet_to_json(i,{defval:""});Object.defineProperty(g,"__worksheet",{value:i,enumerable:!1}),Object.defineProperty(g,"__headers",{value:o,enumerable:!1}),e(g)}catch(d){l(d)}},a.onerror=s=>l(s),a.readAsArrayBuffer(t)})}function B(t,e){try{let l=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,l,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(l){throw console.error("Error generating Excel file:",l),l}}function Re(t){return`
    <div class="excel-actions" style="display:flex;gap:0.5rem;margin-bottom:1rem;">
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
  `}var K=at(()=>{});M();var rt={},Je=null;function J(t,e){rt[t]=e}function Se(t){window.location.hash=t}function ft(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[l,...a]=e.split("?"),s=rt[l];if(!s){for(let[r,n]of Object.entries(rt))if(r.endsWith("/*")&&l.startsWith(r.slice(0,-2))){s=n;break}}Je&&(Je(),Je=null);let d=document.getElementById("main-content");if(d&&(d.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),s){let r=new URLSearchParams(a.join("?")),n=l.split("/").filter(Boolean),i=await s({path:l,params:r,segments:n,main:d});i&&(Je=i)}else{let r=d||document.getElementById("app");r&&(r.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var Ae;function ca(){return Ae||(Ae=document.createElement("div"),Ae.id="toast-container",document.body.appendChild(Ae)),Ae}function vt(t,e="info",l=3500){let a=ca(),s=document.createElement("div");s.className=`toast toast-${e}`;let d={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};s.innerHTML=`<span class="toast-icon">${d[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(s),requestAnimationFrame(()=>s.classList.add("show")),setTimeout(()=>{s.classList.remove("show"),setTimeout(()=>s.remove(),350)},l)}var Z=t=>vt(t,"success"),G=t=>vt(t,"error");ye();M();M();ye();function Ue({columns:t,data:e,onEdit:l,onDelete:a,onView:s,actions:d=[],emptyText:r="Tidak ada data",bulkSelect:n=null}){let i=document.createElement("div");if(i.className="table-wrapper",!e||e.length===0)return i.innerHTML=`<div class="empty-state"><p>${r}</p></div>`,i;let c=document.createElement("table");c.className="data-table";let m=document.createElement("thead"),p=document.createElement("tr");if(n){let g=document.createElement("th");g.style.width="40px",g.style.textAlign="center";let y=document.createElement("input");y.type="checkbox",y.id="select-all-checkbox",y.title="Pilih semua",y.addEventListener("change",()=>{e.forEach(u=>{y.checked?n.selectedIds.add(u.id):n.selectedIds.delete(u.id)}),i.querySelectorAll(".row-checkbox").forEach(u=>u.checked=y.checked),n.onToggle()}),g.appendChild(y),p.appendChild(g)}if(t.forEach(g=>{let y=document.createElement("th");y.textContent=g.label,g.width&&(y.style.width=g.width),p.appendChild(y)}),l||a||s||d.length>0){let g=document.createElement("th");g.textContent="Aksi",g.style.width="120px",p.appendChild(g)}m.appendChild(p),c.appendChild(m);let o=document.createElement("tbody");return e.forEach(g=>{let y=document.createElement("tr");if(n){let u=document.createElement("td");u.style.textAlign="center",u.style.width="40px";let h=document.createElement("input");h.type="checkbox",h.className="row-checkbox",h.checked=n.selectedIds.has(g.id),h.addEventListener("change",()=>{if(h.checked)n.selectedIds.add(g.id);else{n.selectedIds.delete(g.id);let b=document.getElementById("select-all-checkbox");b&&(b.checked=!1)}n.onToggle()}),u.appendChild(h),y.appendChild(u)}if(t.forEach(u=>{let h=document.createElement("td");if(u.render){let b=u.render(g[u.key],g);b instanceof HTMLElement?h.appendChild(b):h.innerHTML=b||""}else h.textContent=g[u.key]!==null&&g[u.key]!==void 0&&g[u.key]!==""?g[u.key]:"";u.nowrap&&(h.style.whiteSpace="nowrap"),y.appendChild(h)}),l||a||s||d.length>0){let u=document.createElement("td");u.className="actions-cell";let h=document.createElement("div");if(h.className="btn-group",s){let b=document.createElement("button");b.className="btn btn-xs btn-ghost",b.innerHTML="\u{1F441}",b.title="Lihat",b.addEventListener("click",()=>s(g)),h.appendChild(b)}if(l){let b=document.createElement("button");b.className="btn btn-xs btn-secondary",b.innerHTML="\u270F\uFE0F",b.title="Edit",b.addEventListener("click",()=>l(g)),h.appendChild(b)}d.forEach(b=>{let C=document.createElement("button");C.className=`btn btn-xs ${b.class||"btn-ghost"}`,C.innerHTML=b.icon||b.label,C.title=b.label,C.addEventListener("click",()=>b.handler(g)),h.appendChild(C)}),u.appendChild(h),y.appendChild(u)}o.appendChild(y)}),c.appendChild(o),i.appendChild(c),i}function Ge({page:t,pages:e,total:l,limit:a,onPage:s}){if(e<=1)return null;let d=document.createElement("div");d.className="pagination";let r=document.createElement("span");r.className="pagination-info",r.textContent=`Total: ${l} data`,d.appendChild(r);let n=document.createElement("div");n.className="pagination-btns";let i=(p,o,g=!1,y=!1)=>{let u=document.createElement("button");u.className=`btn btn-sm ${y?"btn-primary":"btn-ghost"} pagination-btn`,u.textContent=p,u.disabled=g,u.addEventListener("click",()=>s(o)),n.appendChild(u)};i("\xAB",1,t===1),i("\u2039",t-1,t===1);let c=Math.max(1,t-2),m=Math.min(e,t+2);for(let p=c;p<=m;p++)i(p,p,!1,p===t);return i("\u203A",t+1,t===e),i("\xBB",e,t===e),d.appendChild(n),d}ye();function Me(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${Me(e.fields)}</div>`;let l=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",s="";switch(e.type){case"textarea":s=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${l} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let r=(e.options||[]).map(p=>{let o=typeof p=="object"?p.value:p,g=typeof p=="object"?p.label:p,y=e.value==o?"selected":"";return`<option value="${o}" ${y}>${g}</option>`}).join("");s=`<select name="${e.name}" class="form-control" ${l}><option value="">-- Pilih ${e.label||""} --</option>${r}</select>`;break;case"combobox":let n=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,i=(e.options||[]).map(p=>{let o=typeof p=="object"?p.value:p,g=typeof p=="object"?p.label||p.value||"":p||"";return(g==="undefined"||g==="[object Object]"||g==="null")&&(g=""),g?`<option value="${g}"></option>`:""}).join(""),c=e.value||"";if(e.value){let p=(e.options||[]).find(o=>(typeof o=="object"?o.value:o)==e.value);if(p){let o=typeof p=="object"?p.label||p.value||"":p||"";o&&o!=="undefined"&&o!=="[object Object]"&&o!=="null"&&(c=o)}}s=`
          <input type="text" name="${e.name}" list="${n}" class="form-control" value="${c}" placeholder="Pilih atau ketik baru..." ${l} autocomplete="off">
          <datalist id="${n}">${i}</datalist>
        `;break;case"checkbox":s=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let m=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";s=`<input type="date" name="${e.name}" class="form-control" value="${m}" ${l}>`;break;case"number":s=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${l}>`;break;case"email":s=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${l}>`;break;case"url":s=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${l}>`;break;default:s=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${l} autocomplete="off">`}let d=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${s}${d}</div>`}).join("")}function Qe(t){let e={},l=new FormData(t);for(let[a,s]of l.entries())e[a]=s===""?null:s;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function ze(t,e){e&&Object.entries(e).forEach(([l,a])=>{let s=t.querySelector(`[name="${l}"]`);s&&(s.hasAttribute("list")||(s.type==="checkbox"?s.checked=!!a:s.type==="date"&&a&&window.parseFlexibleDate?s.value=window.parseFlexibleDate(a):s.value=a??""))})}K();function N({container:t,title:e,icon:l,apiPath:a,columns:s,formFields:d,filterFields:r,defaultFilters:n={},itemLabel:i="Data",canCreate:c=!0,canEdit:m=!0,canDelete:p=!0,onBeforeSubmit:o,onAfterLoad:g,onDataLoaded:y,extraActions:u=[],initialSearch:h="",exportOptions:b=null,bulkDelete:C=!1,paginationMode:w="server"}){let S=1,_={...n};h&&(_.search=h);let x=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${l} ${e}</h1>
      <div class="page-actions">
        ${c?`<button class="btn btn-primary" id="btn-create">+ Tambah ${i}</button>`:""}
      </div>
    </div>

    ${C?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${b?Re(b.moduleName):""}

    ${r&&r.length>0?`
    <div class="filter-bar card" style="padding: 1rem;">
        ${r.map(v=>{if(v.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${_.search||""}"></div>`;if(v.type==="search-combo"){let f="dl-filter-search",$=(v.options||[]).map(P=>`<option value="${typeof P=="object"?P.label:P}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${f}" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${_.search||""}"><datalist id="${f}">${$}</datalist></div>`}return v.type==="select"?`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${v.name}" id="filter-${v.name}"><option value="">-- ${v.label} --</option>${(v.options||[]).map(f=>`<option value="${typeof f=="object"?f.value:f}" ${_[v.name]===(typeof f=="object"?f.value:f)?"selected":""}>${typeof f=="object"?f.label:f}</option>`).join("")}</select>`:""}).join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function L(){if(!document.getElementById("bulk-toolbar"))return;let f=document.getElementById("bulk-count"),$=document.getElementById("btn-bulk-delete"),P=document.getElementById("btn-bulk-cancel");f.textContent=`${x.size} item dipilih`,x.size>0?($.disabled=!1,P.disabled=!1):($.disabled=!0,P.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{x.clear(),document.querySelectorAll(".row-checkbox").forEach(f=>f.checked=!1);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1),L()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(x.size===0)return;let v=[...x],f=document.createElement("div");f.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",f.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${v.length} ${i}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${v.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(f),f.querySelector("#bulk-cancel-btn").addEventListener("click",()=>f.remove()),f.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let $=f.querySelector("#bulk-confirm-btn");$.disabled=!0,$.textContent="Menghapus...";let P=await k(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:v})});f.remove(),P.ok?(Z(`${v.length} ${i} berhasil dihapus.`),x.clear(),L(),R()):G(P.data?.error||"Gagal menghapus data.")})});let T=document.getElementById("filter-search"),O;if(T?.addEventListener("input",v=>{clearTimeout(O),O=setTimeout(()=>{_.search=v.target.value,S=1,x.clear(),R()},400)}),r?.forEach(v=>{v.type==="select"&&document.getElementById(`filter-${v.name}`)?.addEventListener("change",f=>{_[v.name]=f.target.value,S=1,x.clear(),R()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{_={...n},T&&(T.value=""),r?.forEach(v=>{let f=document.getElementById(`filter-${v.name}`);f&&(f.value="")}),S=1,x.clear(),R()}),document.getElementById("btn-create")?.addEventListener("click",()=>ge(null)),b){document.getElementById(`btn-export-${b.moduleName}`)?.addEventListener("click",async f=>{let $=f.target,P=$.innerHTML;$.innerHTML="\u23F3 Loading...",$.disabled=!0;try{await b.onExport()}catch{G("Gagal export data")}finally{$.innerHTML=P,$.disabled=!1}}),document.getElementById(`btn-template-${b.moduleName}`)?.addEventListener("click",()=>{b.onTemplate()});let v=document.getElementById(`input-import-${b.moduleName}`);v?.addEventListener("change",async f=>{let $=f.target.files[0];if(!$)return;v.disabled=!0;let P=document.createElement("div");P.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",P.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(P);let V=P.querySelector("#import-progress-text"),U=P.querySelector("#import-progress-bar"),D=P.querySelector("#import-summary"),I=P.querySelector("#import-close-btn");I.addEventListener("click",()=>{P.remove(),R()});try{let W=await Oe($);if(W.length===0)throw new Error("File kosong atau format salah");let X=500,ie=0,ae=0,A=0,F=W.length;V.textContent=`Ditemukan ${F} baris data. Memulai import...`;for(let E=0;E<F;E+=X){let z=W.slice(E,E+X);V.textContent=`Mengimport baris ${E+1} - ${Math.min(E+X,F)} dari ${F}...`,U.style.width=`${Math.round(E/F*100)}%`;try{let j=await b.onImport(z);j?(ie+=j.inserted||j.metrics?.inserted||z.length,ae+=j.skipped||j.metrics?.updated||0):ie+=z.length}catch(j){console.error("Chunk import failed:",j),A+=z.length}}U.style.width="100%",V.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',D.style.display="block",D.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${F}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ie}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${A}</strong></div>
        `,A>0&&(D.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),I.style.display="block",v.value=""}catch(W){V.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${W.message}`,U.style.background="var(--danger)",U.style.width="100%",I.style.display="block",v.value=""}finally{v.disabled=!1}})}async function R(){L();let v=document.getElementById("table-container");if(!v)return;v.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let f=w==="client",$=f?1:S,P=f?be:20,V=new URLSearchParams({page:$,limit:P,...Object.fromEntries(Object.entries(_).filter(([,A])=>A))}),U=await k(`${a}?${V}`);if(!U.ok){v.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${U.data?.error||"Error"}</p></div>`;return}let D=U.data?.data||U.data||[],I=U.data?.pagination,W=D.length,X=D;if(f){D=y(D),X=D;let A=D.length,F=20,E=Math.ceil(A/F);S>E&&E>0&&(S=E);let z=(S-1)*F,j=S*F;D=D.slice(z,j),I={page:S,limit:F,total:A,pages:E}}!1,g&&g(D);let ie=Ue({columns:s,data:D,fullData:X,onEdit:m?A=>ge(A):null,actions:u.map(A=>({...A,handler:F=>A.handler(F,R)})),emptyText:`Tidak ada ${String(i||"").toLowerCase()}`,bulkSelect:C?{selectedIds:x,onToggle:L}:null});v.innerHTML="",v.appendChild(ie);let ae=document.getElementById("pagination-container");if(ae&&(ae.innerHTML="",I&&I.pages>1)){let A=Ge({page:I.page,pages:I.pages,total:I.total,limit:I.limit,onPage:F=>{S=F,R()}});A&&ae.appendChild(A)}}function fe(v){let f=typeof d=="function"?d(v):d;return Me(f)}function ge(v){let f=!!v,$=document.createElement("form");if($.noValidate=!0,$.innerHTML=fe(v),f){let V=typeof d=="function"?d(v):d;ze($,v)}let{close:P}=se({title:f?`Edit ${i}`:`Tambah ${i}`,content:$,size:"lg",confirmText:f?"Simpan Perubahan":`Tambah ${i}`,onConfirm:async(V,U)=>{if(!$.reportValidity())return;let D=V.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let I=Qe($),W=typeof d=="function"?d(v):d,X=async F=>{for(let E of F)if(E.type==="row")await X(E.fields);else if(E.type==="combobox"&&I[E.name]){let z=I[E.name],j=(E.options||[]).find(Y=>{let ne=String(typeof Y=="object"?Y.value:Y),tt=String(typeof Y=="object"?Y.label:Y);return ne===z||tt===z});if(j)I[E.name]=typeof j=="object"?j.value:j;else if(E.createApi){let Y={};Y[E.createApi.field]=z,E.createApi.extra&&Object.assign(Y,E.createApi.extra);let ne=await k(E.createApi.path,{method:"POST",body:JSON.stringify(Y)});if(ne.ok&&ne.data?.id)I[E.name]=ne.data.id;else if(ne.ok&&!ne.data?.id)I[E.name]=z;else throw new Error(`Gagal membuat master data: ${ne.data?.error||"Unknown error"}`)}}};try{await X(W)}catch(F){G(F.message),D.disabled=!1,D.textContent=f?"Simpan Perubahan":`Tambah ${i}`;return}o&&(I=await o(I,v));let ie=f?"PUT":"POST",ae=f?`${a}/${v.id}`:a,A=await k(ae,{method:ie,body:JSON.stringify(I)});A.ok?(Z(f?`${i} berhasil diperbarui.`:`${i} berhasil ditambahkan.`),U(),R()):(G(A.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=f?"Simpan Perubahan":`Tambah ${i}`)}})}function ve(v){Fe(`Hapus ${i} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let f=await k(`${a}/${v.id}`,{method:"DELETE"});f.ok?(Z(`${i} berhasil dihapus.`),R()):G(f.data?.error||"Gagal menghapus.")},`Hapus ${i}`)}return R(),R}M();M();var xe=null,Ve=null;async function we(t=!1){if(xe&&!t)return console.log("Employees Raw (Cache Hit)",xe.slice(0,5)),xe;let e=await k(`/api/employees?limit=${be}&status=Aktif`);return xe=(e.data?.data||[]).map(l=>({value:l.id,label:l.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",xe.slice(0,5)),xe}async function ee(t=!1){let l=(await we(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",l.slice(0,5)),l}async function H(t=!1){return Ve&&!t||(Ve=((await k("/api/branches?all=1")).data?.data||[]).map(l=>({value:l.id,label:l.full_name}))),Ve}function q(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function lt(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function _e(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function st(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function oe(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}K();function ot(t,e){if(t.period!=="Q3")return!1;let l=String(t.status||"").toLowerCase();if(l!=="selesai"&&l!=="completed"&&l!=="done")return!1;let a=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?a.includes("inspeksi"):e==="gcdc"?a.includes("general cleaning")||a.includes("deep cleaning"):!1}M();K();function St(t,e){let l=String(t.status||"").toLowerCase();return e==="active"?l==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&l==="aktif":!1}M();K();function dt(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let s=new Date(a);s.setDate(a.getDate()+30);let d=new Date(t.end_date);return d.setHours(0,0,0,0),d>=a&&d<=s}return!1}M();K();function xt(t,e){let l=String(t.status||"").toLowerCase();return e==="open"?l==="open":!1}M();function wt(t,e){let l=String(t.status||"").toLowerCase();return e==="pending"?l==="pending":!1}var me={};function Ee(t){if(me[t]){try{me[t].destroy()}catch{}delete me[t]}}function pa(){Object.keys(me).forEach(Ee)}var pe=(t,e=0)=>{let l=Number(t);return isNaN(l)||t===null||t===void 0?e:l},Ce=(t,e="\u2014")=>{if(t==null||t==="")return e;let l=String(t).trim();return l===""||l==="[object Object]"?e:l};function Ct(t,e,l=900){if(!t)return;let a=Math.max(0,Math.round(pe(e)));if(a===0){t.textContent="0";return}let s=Date.now(),d=()=>{let r=Math.min((Date.now()-s)/l,1),n=1-Math.pow(1-r,3);t.textContent=Math.round(n*a).toLocaleString("id-ID"),r<1?requestAnimationFrame(d):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(d)}var ua={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},ma=t=>{let e=Ce(t,"\u2014");return`<span class="status-pill ${ua[e]||"pill-neutral"}">${e}</span>`};var de={family:"Inter",size:11},ue="#94A3B8",Te="#F1F5F9",ct=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],ga=()=>window.innerWidth<768;function Ye(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:ga()?"bottom":"top",labels:{font:de,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:de,titleFont:{...de,weight:"700"}}},scales:{x:{grid:{color:Te},ticks:{font:de,color:ue,maxRotation:0}},y:{grid:{color:Te},ticks:{font:de,color:ue},beginAtZero:!0}},...t}}var ba=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),ha=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function _t(t=3){return Array(t).fill(0).map((e,l)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${l<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function re(t,e,l=8e3){try{let a=new AbortController,s=setTimeout(()=>a.abort(),l),d=await k(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(s),!d||!d.ok)return e;let r=d.data;return r?r.data!==void 0?r.data??e:r:e}catch{return e}}function ya(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(a=>{let s=document.getElementById(a);s&&(s.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(a=>{let s=document.getElementById(a);if(s&&s.style.display==="none"){s.style.display="block";let d=s.parentElement;if(d&&!d.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent="Belum ada data",s.style.display="none",d.appendChild(r)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Et({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&$t({}),["table-contracts","table-issues"].forEach(a=>{let s=document.getElementById(a);s&&s.querySelector(".skeleton")&&(s.innerHTML='<div class="chart-empty">Belum ada data</div>')});let l=document.getElementById("activity-log");l&&l.querySelector(".skeleton")&&(l.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Tt(t){pa(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">


      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${ba()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${ha()}</div>

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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${_t(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${_t(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>pt(t)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async l=>{let a=l.target.value,s=document.getElementById("jadwal-year-label");s&&(s.textContent=a);let d=document.getElementById("skel-jadwal"),r=document.getElementById("chart-jadwal");d&&(d.style.display="block",d.style.position="absolute"),r&&(r.style.display="none");let n=await re(`/api/dashboard/schedule-chart?year=${a}`,{},8e3);try{Dt(n)}catch(i){console.warn("ScheduleChart render:",i),ce("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async l=>{let a=l.target.value,s=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",d=document.getElementById("skel-insp"),r=document.getElementById("chart-insp");d&&(d.style.display="block",d.style.position="absolute"),r&&(r.style.display="none");let n=await re(s,{},8e3);try{It(n)}catch(i){console.warn("InspBar render:",i),ce("skel-insp","chart-insp")}}),t._skelTimeout=setTimeout(()=>ya(),5e3),await pt(t),t._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?pt(t):clearInterval(t._dashRefresh)},6e4)}async function pt(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,l,a,s,d,r,n,i,c,m,p,o]=await Promise.all([re("/api/dashboard/kpi",{},8e3),re("/api/dashboard/issues-trend",{},8e3),re("/api/dashboard/issues-summary",{},8e3),re("/api/dashboard/stats",{},8e3),re("/api/dashboard/calendar",[],8e3),re("/api/schedule?limit=10000",{data:[]},8e3),re("/api/employees?limit=10000",{data:[]},8e3),re("/api/contracts?limit=10000",{data:[]},8e3),re("/api/issues?limit=10000",{data:[]},8e3),re("/api/one-on-one?limit=10000",{data:[]},8e3),re("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),re(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3)]),g=document.getElementById("filter-insp-month"),y=g?g.value:"",u=y?`/api/dashboard/inspection-bar?month=${y}`:"/api/dashboard/inspection-bar",h=await re(u,{},8e3);if(e){let b=Array.isArray(r?.data)?r.data:Array.isArray(r)?r:[],C=Array.isArray(n?.data)?n.data:Array.isArray(n)?n:[],w=Array.isArray(i?.data)?i.data:Array.isArray(i)?i:[],S=Array.isArray(c?.data)?c.data:Array.isArray(c)?c:[],_=Array.isArray(m?.data)?m.data:Array.isArray(m)?m:[];e.employees&&(e.employees.current=C.filter(x=>St(x,"active")).length),e.contracts&&(e.contracts.current=w.filter(x=>dt(x,"active")).length),e.expiring30&&(e.expiring30={current:w.filter(x=>dt(x,"expiring30")).length}),e.issues&&(e.issues.current=S.filter(x=>xt(x,"open")).length),e.one_on_one&&(e.one_on_one.current=_.filter(x=>wt(x,"pending")).length),e.inspection_month&&(e.inspection_month.current=b.filter(x=>ot(x,"inspeksi")).length),e.cleaning_month&&(e.cleaning_month.current=b.filter(x=>ot(x,"gcdc")).length)}try{Et(e)}catch(b){console.warn("KPI render:",b)}try{$t(e)}catch(b){console.warn("MiniStats render:",b)}try{Dt(o)}catch(b){console.warn("ScheduleChart render:",b),ce("skel-jadwal","chart-jadwal")}try{fa(Array.isArray(a?.by_category)?a.by_category:[])}catch(b){console.warn("Donut render:",b),ce("skel-donut","chart-donut")}try{va(l)}catch(b){console.warn("Trend render:",b),ce("skel-trend","chart-trend")}try{It(h)}catch(b){console.warn("InspBar render:",b),ce("skel-insp","chart-insp")}try{let b=Array.isArray(s)?s:Array.isArray(s?.recent_issues)?s.recent_issues:[];Sa(b)}catch(b){console.warn("IssuesTable render:",b)}try{let b=Array.isArray(s?.expiring_contracts)?s.expiring_contracts:[];ka(p)}catch(b){console.warn("ContractsTable render:",b)}try{xa(Array.isArray(d)?d:[])}catch(b){console.warn("Agenda render:",b)}try{wa()}catch(b){console.warn("Quick Actions render:",b)}}function Et(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let l=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=l.map(a=>{let s=pe(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${s}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${s}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{Ct(a,parseInt(a.dataset.target)||0)})}function $t(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let l=[{icon:"\u{1F4C5}",label:"Jadwal",val:t.schedule?.current,href:"#/timeline",color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F504}",label:"Report Reliefer",val:t.reliever_completed?.current,href:"#/relievers?dash_filter=reliever",color:"mini-teal"},{icon:"\u{1F50D}",label:"Report Inspeksi",val:t.inspection_month?.current,href:"#/timeline?dash_filter=inspeksi",color:"mini-blue"},{icon:"\u{1F9F9}",label:"Report GCDC",val:t.cleaning_month?.current,href:"#/timeline?dash_filter=gcdc",color:"mini-green"},{icon:"\u{1F4A8}",label:"Report Fogging",val:t.fogging_month?.current,href:"#/reports/fogging?dash_filter=fogging",color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=l.map(a=>`
    <a href="${a.href}" class="mini-stat ${a.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${a.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${pe(a.val)}">0</div>
        <div class="mini-stat-text">${a.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(a=>Ct(a,parseInt(a.dataset.target)||0,700))}function fa(t){ce("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),l=document.getElementById("donut-legend");if(!e||!l)return;Ee("donut");let a=(t||[]).filter(i=>pe(i.count)>0);if(!a.length){Ke(e,"Belum ada data permasalahan");return}let s=a.map(i=>`${Ce(i.category,"Lainnya")}`),d=a.map(i=>pe(i.count)),r=d.reduce((i,c)=>i+c,0);l.innerHTML=a.map((i,c)=>{let m=ct[c%ct.length],p=r>0?Math.round(i.count/r*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${m}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${i.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${p}%)</span></div>
          <div class="donut-legend-label">${s[c]}</div>
        </div>
      </div>
    `}).join("");let n={id:"centerText",beforeDraw:function(i){let c=i.width,m=i.height,p=i.ctx;p.restore();let o=(m/80).toFixed(2);p.font="bold "+o+"em Inter",p.textBaseline="middle",p.fillStyle="#1E293B";let g=r.toString(),y=Math.round((c-p.measureText(g).width)/2),u=m/2;p.fillText(g,y,u-4),p.font="600 "+(o*.35).toFixed(2)+"em Inter",p.fillStyle="#64748B";let h="Total",b=Math.round((c-p.measureText(h).width)/2);p.fillText(h,b,u+10),p.save()}};me.donut=new Chart(e,{type:"doughnut",data:{labels:s,datasets:[{data:d,backgroundColor:ct,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:de,titleFont:{...de,weight:"700"},callbacks:{label:i=>` ${i.label}: ${i.parsed} kasus`}}},cutout:"75%"},plugins:[n]})}function va(t){ce("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;Ee("trend"),t=t||{};let l=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(t.labels||[]).map(r=>{if(!r||typeof r!="string")return"";try{let[n,i]=r.split("-");return(l[Number(i)-1]||i)+" "+String(n).slice(-2)}catch{return r}}),s=(t.open||[]).map(r=>pe(r)),d=(t.closed||[]).map(r=>pe(r));if(!a.length){Ke(e,"Belum ada data trend");return}me.trend=new Chart(e,{type:"line",data:{labels:a,datasets:[{label:"Open",data:s,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:d,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:Ye({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:ue,maxRotation:0,autoSkip:!1}},y:{grid:{color:Te},ticks:{font:{family:"Inter",size:9},color:ue},beginAtZero:!0}}})})}function Dt(t){ce("skel-jadwal","chart-jadwal");let e=document.getElementById("chart-jadwal");if(!e)return;Ee("jadwal"),t=t||{};let l=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(t).some(i=>Array.isArray(i)&&i.some(c=>c>0))){Ke(e,"Belum ada data jadwal");return}let s=t["Inspeksi Hygiene"]||Array(12).fill(0),d=t["General Cleaning"]||Array(12).fill(0),r=t["Deep Cleaning"]||Array(12).fill(0),n=t.Fogging||Array(12).fill(0);me.jadwal=new Chart(e,{type:"bar",data:{labels:l,datasets:[{label:"Inspeksi",data:s,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:d,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:r,backgroundColor:"#F59E0B"},{label:"Fogging",data:n,backgroundColor:"#EF4444"}]},options:Ye({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:ue,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Te},ticks:{font:{family:"Inter",size:9},color:ue},min:0}}})})}function It(t){ce("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;Ee("inspBar"),t=t||{};let l=t.labels||[],a=(t.fc||[]).map(d=>pe(d)),s=(t.spv||[]).map(d=>pe(d));if(!l.length){Ke(e,"Belum ada data inspeksi");return}me.inspBar=new Chart(e,{type:"bar",data:{labels:l,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:s,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:Ye({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:de,color:ue,maxRotation:45,minRotation:30}},y:{grid:{color:Te},ticks:{font:de,color:ue},min:0,max:100}}})})}function ka(t){ce("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;Ee("contractMiniBar"),t=t||{};let l={"06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(r=>{let n=r.split("-")[1];return l[n]||r}),s=(t.data||[]).map(r=>pe(r));if(!a.length){Ke(e,"Belum ada data");return}let d=e.getContext("2d");me.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:s,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:Ye({onClick:(r,n)=>{if(n&&n.length>0){let i=n[0].index,c=(t.labels||[])[i];c&&(window.location.hash="#/contracts?month_expiry="+c)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:de,color:ue,maxRotation:0,autoSkip:!1}},y:{grid:{color:Te,borderDash:[4,4],drawBorder:!1},ticks:{font:de,color:ue,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function Sa(t){let e=document.getElementById("table-issues");if(!e)return;let l=(t||[]).slice(0,8);if(!l.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${l.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${ma(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Ce(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Ce(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function xa(t){let e=document.getElementById("widget-agenda");if(!e)return;let l=new Date,a=`${l.getFullYear()}-${String(l.getMonth()+1).padStart(2,"0")}-${String(l.getDate()).padStart(2,"0")}`,d=(t||[]).filter(r=>(r.event_date||"").startsWith(a)).slice(0,10);if(!d.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${d.map(r=>{let n="#3B82F6",i="#EFF6FF",c="Agenda",m=(r.title||"").toLowerCase();return m.includes("inspeksi")?(n="#10B981",i="#ECFDF5",c="Inspeksi"):m.includes("cleaning")||m.includes("gcdc")?(n="#3B82F6",i="#EFF6FF",c="Cleaning"):m.includes("reliefer")?(n="#F59E0B",i="#FFFBEB",c="Reliefer"):m.includes("fogging")&&(n="#8B5CF6",i="#F5F3FF",c="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(r.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${n};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Ce(r.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Ce(r.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${i};color:${n}">${c}</div>
        </div>
      `}).join("")}
    </div>
  `}function wa(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(l=>`
    <a href="${l.href}" class="action-btn">
      <div class="action-icon" style="background:${l.bg}">${l.icon}</div>
      ${l.label}
    </a>
  `).join("")}function ce(t,e){let l=document.getElementById(t),a=document.getElementById(e);if(l&&(l.style.display="none",l.style.position=""),a){a.style.display="block";let s=a.parentElement;if(s){let d=s.querySelector(".chart-empty");d&&d.remove()}}}function Ke(t,e="Belum ada data"){if(!t)return;t.style.display="none";let l=t.parentElement;if(!l)return;if(!l.querySelector(".chart-empty")){let s=document.createElement("div");s.className="chart-empty",s.textContent=e,l.appendChild(s)}}M();async function Pt(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),l=document.getElementById("login-error"),a=document.getElementById("login-btn"),s=document.getElementById("toggle-password"),d=document.getElementById("login-password");s?.addEventListener("click",()=>{let r=d.type==="text";d.type=r?"password":"text",s.style.color=r?"":"var(--primary)"}),e?.addEventListener("submit",async r=>{r.preventDefault(),l.style.display="none";let n=e.username.value.trim(),i=e.password.value;if(!n||!i){l.textContent="Username dan password wajib diisi.",l.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let c=await k("/api/auth/login",{method:"POST",body:JSON.stringify({username:n,password:i})});c.ok&&c.data.success?(it(c.data.data.token),Ne(c.data.data.user),Z("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(l.textContent=c.data.error||"Username atau password salah.",l.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{l.textContent="Gagal terhubung ke server. Periksa koneksi internet.",l.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}M();K();async function _a(){return await H()}function Ca(t,e){let l=String(t.status||"").toLowerCase();return e==="active"?l==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&l==="aktif":!1}async function Bt(t,e){let l=await _a(),a=e?e.get("dash_filter"):null;N({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:s=>a?s.filter(d=>Ca(d,a)):s,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:s=>_e(s)},{key:"phone",label:"No. HP",render:s=>s?`<a href="tel:${s}">${s}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:s=>window.formatDate(s)},{key:"status",label:"Status",render:s=>q(s)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:l},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:s=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:s?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:s?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:l,value:s?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:s?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:s?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:s?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:s?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let s=await k(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(s.ok){let d=s.data.data.map(r=>({"Nama Lengkap":r.full_name,Cabang:r.branch_name||"",Divisi:r.division||"","No. HP":r.phone||"","Tgl Masuk":r.join_date||"",Status:r.status||""}));B(d,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async s=>{let d=i=>{if(!i)return null;let c=String(i||"").toLowerCase(),m=l.find(p=>String(p.label||"").toLowerCase()===c);return m?m.value:null},r=s.map(i=>({full_name:String(i["Nama Lengkap"]||"").trim(),branch_id:d(String(i.Cabang||"").trim()),division:String(i.Divisi||"").trim()||"FACILITY CARE",phone:String(i["No. HP"]||"").trim(),join_date:String(i["Tgl Masuk"]||"").trim(),status:String(i.Status||"").trim(),notes:String(i.Catatan||"").trim()})).filter(i=>i.full_name),n=await k("/api/import/employees",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}M();K();var mt=[],Lt=[];async function Ta(){mt=await H(),Lt=await we()}var ut=async t=>{let e=[],l=1;for(;;){let s=await(await Promise.resolve().then(()=>(M(),ke))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${l}`);if(!s.ok)break;let d=s.data?.data||s.data||[],r=Array.isArray(d)?d:[];if(e=e.concat(r),r.length<100||s.data?.pagination&&l>=s.data.pagination.pages)break;l++}return e};function Ea(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let s=new Date(a);s.setDate(a.getDate()+30);let d=new Date(t.end_date);return d.setHours(0,0,0,0),d>=a&&d<=s}return!1}async function Nt(t,e){await Ta();let l=e?e.get("dash_filter"):null;N({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",onDataLoaded:a=>l?a.filter(s=>Ea(s,l)):a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>_e(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,s)=>s.end_date&&String(s.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':lt(a)},{key:"status",label:"Status",render:a=>q(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:mt},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"month_expiry",label:"Bulan Habis",options:[{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"},{value:"2027-01",label:"Jan 2027"},{value:"2027-02",label:"Feb 2027"},{value:"2027-03",label:"Mar 2027"},{value:"2027-04",label:"Apr 2027"},{value:"2027-05",label:"Mei 2027"},{value:"2027-06",label:"Jun 2027"},{value:"2027-07",label:"Jul 2027"}]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[d,r]=await Promise.all([ut("/api/employees?status=Aktif"),ut("/api/contracts")]);if(d.length>0){let n=r.filter(p=>p.status==="Aktif"),i=new Set(n.map(p=>p.employee_id)),c=d.filter(p=>!i.has(p.id)),m=`<p style="margin-bottom:12px">Data yang terbaca: <b>${d.length}</b> Karyawan Aktif, dan <b>${n.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${c.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;c.forEach(p=>{let o=r.filter(y=>y.employee_id===p.id),g='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(o.length>0){let y=o[0];g=`Pernah ada kontrak (Status: <b style="color:#EF4444">${y.status}</b>, Selesai: ${window.formatDate(y.end_date)})`}m+=`<li style="margin-bottom:8px"><b>${p.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${p.branch_name||"-"} | ${g}</span></li>`}),m+="</ul>",Promise.resolve().then(()=>(ye(),kt)).then(p=>p.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:m,cancelText:"Tutup"}))}}catch(d){console.error(d)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let s=document.querySelector(".page-actions");s&&s.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:Lt,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:mt,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await k(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let s=a.data.data.map(d=>({"Nama Lengkap":d.employee_name,Cabang:d.branch_name||"","Div / Bagian":d.division||"","Tanggal Mulai":d.start_date||"","Tanggal Selesai":d.end_date&&String(d.end_date).startsWith("2099")?"":d.end_date||"","Sisa Kontrak":d.end_date&&String(d.end_date).startsWith("2099")?"Tetap":d.days_remaining!==null&&d.days_remaining!==void 0?`${d.days_remaining} Hari`:"",Status:d.status||""}));B(s,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[s,d]=await Promise.all([k("/api/branches?limit=10000"),ut("/api/employees")]),r=s.data?.data||[],n=d||[];console.log(`Total employee yang berhasil dimuat dari database : ${n.length}`),n.length>0&&(console.log("Contoh 5 employee pertama:"),n.slice(0,5).forEach((u,h)=>{console.log(`${h+1}. ID: ${u.id}, Name: ${u.full_name}, Status: ${u.status}`)}));let i=u=>{if(!u)return null;let h=String(u||"").replace(/\s+/g," ").toLowerCase().trim(),b=r.find(C=>String(C.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===h||String(C.code||"").replace(/\s+/g," ").toLowerCase().trim()===h||String(C.name||"").replace(/\s+/g," ").toLowerCase().trim()===h);return b?b.id:null},c=(u,h)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${h}`),console.log(`Nama dari Excel : "${u}"`),!u)return console.log("Alasan gagal mapping : Nama kosong"),null;let b=String(u||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${b}"`),console.log(`Jumlah employee di database : ${n.length}`);let C=n.find(w=>String(w.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===b);return C?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${C.id}`),C.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},m=u=>{if(!u)return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let h=String(u).trim();if(/^\d{4,5}(\.\d+)?$/.test(h)){let C=Math.floor(Number(h));if(C>2e4&&C<99999){let w=new Date(Date.UTC(1899,11,30)+C*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);let b=h.split(/[\/\-\.]/);if(b.length===3){let[C,w,S]=b.map(_=>_.trim());if(C.length===4&&w.length<=2&&S.length<=2)return`${C}-${w.padStart(2,"0")}-${S.padStart(2,"0")}`;if(S.length===4&&w.length<=2&&C.length<=2)return`${S}-${w.padStart(2,"0")}-${C.padStart(2,"0")}`}return h},p=a.map((u,h)=>{let b=h+2,C=String(u["Nama Lengkap"]||"").trim(),w=u["Tanggal Mulai"],S=m(w);if(!S){let L=a.__worksheet,T=a.__headers||[],O=T.indexOf("Tanggal Mulai"),R="N/A",fe="N/A",ge="N/A";if(O!==-1&&L&&window.XLSX){let v=window.XLSX.utils.encode_cell({c:O,r:b-1});ge=v;let f=L[v];f?(R=f.t||"undefined",fe=f.w||"undefined"):R="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let ve="Unknown";w==null||w===""?ve="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":w instanceof Date&&isNaN(w.getTime())?ve="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":ve="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${b}`),console.log(`Employee Name : ${C}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${O})`),console.log(`Raw Cell Value : "${w}"`),console.log(`JavaScript Type : ${typeof w}`),console.log(`SheetJS Cell Type : ${R}`),console.log(`SheetJS Formatted Value : "${fe}"`),console.log(`Value After Trim : "${String(w||"").trim()}"`),console.log(`Value After Date Parser : "${S}"`),console.log(`Is Empty : ${!w}`),console.log(`Is Invalid Date : ${w instanceof Date?isNaN(w.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${ve}`),console.log(`Workbook Sheet : ${L?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${ge}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(u,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(T)),console.log(`==========================
`)}let _=c(C,b),x=null;return _||(x="Karyawan tidak ditemukan di Database"),{isValid:!!_,invalidReason:x,rowNum:b,data:{employee_id:_,branch_id:i(String(u.Cabang||"").trim()),division:String(u["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:S,end_date:m(u["Tanggal Selesai"])||"2099-12-31",status:String(u.Status||"").trim(),_rawName:C}}}),o=[],g=[];if(p.forEach(u=>{u.isValid?o.push(u.data):g.push({rowNum:u.rowNum,name:u.data._rawName,reason:u.invalidReason})}),console.log(`Split Validation - Valid: ${o.length}, Invalid: ${g.length}`),o.length===0)return{inserted:0,skipped:a.length,failed:a.length};let y=await k("/api/import/contracts",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!y.ok)throw new Error(y.data?.error||"Import gagal");return y.data}}})}M();K();var gt=[],He=[];function $a(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let l of e)if(t.some(a=>a.period===l))return l;return"Q3"}async function At(t,e){gt=await H();let l=await ee();He=["BERLIN ARIANSYAH","ADE SURAHMAN"];let a=o=>o&&!He.find(g=>(typeof g=="object"?g.value:g)===o)?[...He,o]:He,s=await k(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),d=o=>{if(!o||o==="-"||String(o).trim()==="")return"";let g=String(o).split("-");return g.length===3&&g[0].length===4?`${g[2]}-${g[1]}-${g[0]}`:o},r=s.data?.data||[],n=$a(r),i=e?e.get("dash_filter"):null,c=new Date,m=`${c.getFullYear()}-${String(c.getMonth()+1).padStart(2,"0")}`,p={period:n};i==="inspeksi"?p={status:"Done",activity_type:"Inspeksi Hygiene",month:m}:i==="gcdc"&&(p={status:"Done",activity_type:"General Cleaning",month:m}),N({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",defaultFilters:p,onDataLoaded:o=>o.sort((g,y)=>{let u=g.opening_date?new Date(g.opening_date).getTime():0;return(y.opening_date?new Date(y.opening_date).getTime():0)-u}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:o=>st(o)},{key:"period",label:"Periode",render:o=>oe(o)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:o=>d(o)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:o=>d(o)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:o=>d(o)},{key:"status",label:"Status",render:o=>q(o)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:gt},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:He}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:gt,value:o?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:o?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:o?.period},{name:"pic",label:"PIC",type:"combobox",options:a(o?.pic),value:o?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:o?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:o?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:o?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:o?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let o=await k(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(o.ok){let g=o.data.data.map(y=>({Cabang:y.branch_name||"",Kegiatan:y.activity_type||"",Periode:y.period||"",PIC:y.pic||"","Tgl Opening":y.opening_date||"","Tgl Target":y.target_date||"","Tgl Selesai":y.completion_date||"",Status:y.status||""}));B(g,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async o=>{let y=(await k("/api/branches?all=1")).data?.data||[],u=w=>{if(!w)return null;let S=String(w||"").toLowerCase(),_=y.find(x=>String(x.full_name||"").toLowerCase()===S||String(x.code||"").toLowerCase()===S||String(x.name||"").toLowerCase()===S);return _?_.id:null},h=w=>{if(w==null||w==="")return"";if(w instanceof Date&&!isNaN(w.getTime()))return w.toISOString().slice(0,10);let S=String(w).trim();if(S===""||S==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(S))return S.slice(0,10);if(/^\d{4,5}$/.test(S)){let x=Number(S);if(x>2e4&&x<99999){let L=new Date(Date.UTC(1899,11,30)+x*864e5);return isNaN(L.getTime())?"":L.toISOString().slice(0,10)}}let _=S.split(/[\/\-\.]/);if(_.length===3){let[x,L,T]=_.map(O=>O.trim());if(x.length===4&&L.length<=2&&T.length<=2)return`${x}-${L.padStart(2,"0")}-${T.padStart(2,"0")}`;if(T.length===4&&L.length<=2&&x.length<=2)return`${T}-${L.padStart(2,"0")}-${x.padStart(2,"0")}`}return S},b=o.map(w=>({branch_id:u(String(w.Cabang||"").trim()),activity_type:String(w.Kegiatan||"").trim(),period:String(w.Periode||"").trim(),pic:String(w.PIC||w.Pic||"").trim(),opening_date:h(w["Tgl Opening"]||w["Tanggal Opening"]||w["Tgl Openir"]),target_date:h(w["Tgl Target"]||w["Tanggal Target"]),completion_date:h(w["Tgl Selesai"]||w["Tanggal Selesai"]),status:String(w.Status||"").trim(),notes:String(w.Catatan||w.Keterangan||"").trim()})).filter(w=>w.activity_type&&w.period),C=await k("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:b,onDuplicate:"update"})});if(!C.ok)throw new Error(C.data?.error||"Import gagal");return C.data}}})}M();K();var bt=[],We=[];function Da(t,e){let l=String(t.status||"").toLowerCase();return e==="open"?l==="open":!1}async function Ft(t,e){let l=e?e.get("dash_filter"):null;bt=await H(),We=await ee();let a=r=>r&&!We.find(n=>n.value===r)?[...We,{value:r,label:r}]:We,s=new Date().getFullYear();N({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:r=>l?r.filter(n=>Da(n,l)):r,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:r=>`<span class="badge badge-secondary">${r}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:r=>`<span title="${r}">${r?.length>50?r.slice(0,50)+"\u2026":r}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:r=>`<span title="${r||""}">${r?.length>40?r.slice(0,40)+"\u2026":r||"-"}</span>`},{key:"status",label:"Status",render:r=>q(r)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:r=>window.formatDate(r)},{key:"day_count",label:"Hari",render:r=>r??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:bt},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:["2025","2026","2027","2028","2029","2030"]}],formFields:r=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:r?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:bt,value:r?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:r?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:r?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:r?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(r?.employee_name),value:r?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(r?.fc_specialist),value:r?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:r?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:r?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:r?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let r=await k(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let n=r.data.data.map(i=>({Tanggal:i.report_date||"",Cabang:i.branch_name||"",Kategori:i.category||"",Sumber:i.source||"",Keluhan:i.complaint||"","Nama FC":i.employee_name||"","FC Spesialis":i.fc_specialist||"",Solusi:i.solution||"","Tgl Selesai":i.completion_date||"",Status:i.status||""}));B(n,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async r=>{let i=(await k("/api/branches?all=1")).data?.data||[],c=o=>{if(!o)return null;let g=String(o||"").toLowerCase(),y=i.find(u=>String(u.full_name||"").toLowerCase()===g||String(u.code||"").toLowerCase()===g||String(u.name||"").toLowerCase()===g);return y?y.id:null},m=r.map(o=>({branch_id:c(String(o.Cabang||"").trim()),report_date:String(o.Tanggal||"").trim(),category:String(o.Kategori||"").trim(),source:String(o.Sumber||"").trim(),complaint:String(o.Keluhan||"").trim(),employee_name:String(o["Nama FC"]||"").trim(),fc_specialist:String(o["FC Spesialis"]||"").trim(),solution:String(o.Solusi||"").trim(),completion_date:String(o["Tgl Selesai"]||"").trim(),status:String(o.Status||"").trim()})).filter(o=>o.report_date&&o.complaint&&o.category),p=await k("/api/import/issues",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}}})}M();var $e=[];function Ia(t,e){let l=String(t.status||"").toLowerCase();return e==="pending"?l==="pending":!1}async function Mt(t,e){let l=e?e.get("dash_filter"):null;$e=await H();let a=await ee(),s=["Ade","Berlin"],d=n=>n&&!a.find(i=>i.value===n)?[...a,{value:n,label:n}]:a,r=n=>n&&!s.find(i=>(typeof i=="object"?i.value:i)===n)?[...s,n]:s;N({container:t,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:n=>l?n.filter(i=>Ia(i,l)):n,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:n=>`<span title="${n||""}">${n?.length>50?n.slice(0,50)+"\u2026":n||"-"}</span>`},{key:"solution",label:"Solusi",render:n=>`<span title="${n||""}">${n?.length>40?n.slice(0,40)+"\u2026":n||"-"}</span>`},{key:"status",label:"Status",render:n=>q(n)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"combobox",name:"branch_id",label:"Cabang",options:$e},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async n=>{let i=new URLSearchParams(n||{}).toString(),c=await k(`/api/one-on-one?limit=10000&${i}`);if(c.ok){let m=c.data.data.map(o=>({Tanggal:o.meeting_date||"",Cabang:o.branch_name||"","Nama Karyawan":o.employee_name||"",PIC:o.pic||"",Masalah:o.problem||"",Solusi:o.solution||"",Status:o.status||"","Tgl Selesai":o.completion_date||"",Dokumen:o.document_link||""})),{downloadExcel:p}=await Promise.resolve().then(()=>(K(),le));p(m,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(K(),le));i(n,"Template_Import_OneOnOne")},onImport:async n=>{let i=o=>{if(!o)return null;let g=String(o||"").toLowerCase(),y=$e.find(u=>String(u.label||"").toLowerCase()===g);return y?y.value:null},c=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let g=String(o).trim();if(/^\d{4,5}$/.test(g)){let u=Number(g);if(u>2e4&&u<99999){let h=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);let y=g.split(/[\/\-\.]/);if(y.length===3){let[u,h,b]=y.map(C=>C.trim());if(u.length===4&&h.length<=2&&b.length<=2)return`${u}-${h.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&h.length<=2&&u.length<=2)return`${b}-${h.padStart(2,"0")}-${u.padStart(2,"0")}`}return g},m=n.map(o=>({meeting_date:c(o.Tanggal),employee_name:String(o["Nama Karyawan"]||"").trim(),branch_id:i(String(o.Cabang||"").trim()),pic:String(o.PIC||"").trim(),problem:String(o.Masalah||"").trim(),solution:String(o.Solusi||"").trim(),status:String(o.Status||"").trim(),completion_date:c(o["Tgl Selesai"]),document_link:String(o.Dokumen||"").trim()})).filter(o=>o.meeting_date&&o.employee_name&&o.branch_id),p=await k("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}},formFields:n=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:n?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:n?.branch_id&&!$e.find(i=>i.value==n.branch_id)?[...$e,{value:n.branch_id,label:n.branch_name||n.branch_id}]:$e,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"combobox",required:!0,options:d(n?.employee_name),value:n?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:r(n?.pic),createApi:{path:"/api/pic",field:"name"},value:n?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:n?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:n?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:n?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link}]})}M();async function Ot(t){let e=await H(),l=await ee(),a=["Ade","Berlin"],s=n=>n&&!l.find(i=>i.value===n)?[...l,{value:n,label:n}]:l,d=n=>n&&!a.find(i=>(typeof i=="object"?i.value:i)===n)?[...a,n]:a,r=Array.from({length:5},(n,i)=>String(new Date().getFullYear()-i));N({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:n=>{try{let i=JSON.parse(n);return Array.isArray(i)?i.join(", "):n||"-"}catch{return n||"-"}}},{key:"score",label:"Nilai",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:r}],exportOptions:{moduleName:"training",onExport:async n=>{let i=new URLSearchParams(n||{}).toString(),c=await k(`/api/training?limit=10000&${i}`);if(c.ok){let m=c.data.data.map(o=>{let g=o.participants||"";try{let y=JSON.parse(g);g=Array.isArray(y)?y.join(", "):g}catch{}return{Tanggal:o.training_date||"",Batch:o.batch||"",Materi:o.subject||"",Cabang:o.branch_name||"",Trainer:o.trainer||"",Peserta:g,Nilai:o.score!==null&&o.score!==void 0?o.score:"",Dokumen:o.document_link||""}}),{downloadExcel:p}=await Promise.resolve().then(()=>(K(),le));p(m,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(K(),le));i(n,"Template_Import_Training")},onImport:async n=>{let i=o=>{if(!o)return null;let g=String(o||"").toLowerCase(),y=e.find(u=>String(u.label||"").toLowerCase()===g);return y?y.value:null},c=o=>{if(!o)return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let g=String(o).trim();if(/^\d{4,5}$/.test(g)){let u=Number(g);if(u>2e4&&u<99999){let h=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);let y=g.split(/[\/\-\.]/);if(y.length===3){let[u,h,b]=y.map(C=>C.trim());if(u.length===4&&h.length<=2&&b.length<=2)return`${u}-${h.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&h.length<=2&&u.length<=2)return`${b}-${h.padStart(2,"0")}-${u.padStart(2,"0")}`}return g},m=n.map(o=>({training_date:c(o.Tanggal),batch:String(o.Batch||"").trim(),subject:String(o.Materi||"").trim(),branch_id:i(String(o.Cabang||"").trim()),trainer:String(o.Trainer||"").trim(),participants:String(o.Peserta||"").trim(),score:o.Nilai?Number(o.Nilai):null,document_link:String(o.Dokumen||"").trim()})).filter(o=>o.training_date&&o.subject&&o.branch_id),p=await k("/api/import/training",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}},formFields:n=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:n?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:n?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:n?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:n?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:d(n?.trainer),value:n?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let i=JSON.parse(n?.participants);return Array.isArray(i)?i.join(", "):n?.participants||""}catch{return n?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:n?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:n?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],onBeforeSubmit:async n=>(n.participants&&(n.participants=JSON.stringify(n.participants.split(",").map(i=>i.trim()).filter(Boolean))),n)})}M();ye();K();function Rt({container:t,title:e,icon:l,apiPath:a,columns:s,formFields:d,filterFields:r,defaultFilters:n={},itemLabel:i="Data",canCreate:c=!0,canEdit:m=!0,canDelete:p=!0,onBeforeSubmit:o,onAfterLoad:g,onDataLoaded:y,extraActions:u=[],initialSearch:h="",exportOptions:b=null,bulkDelete:C=!1,paginationMode:w="server"}){let S=1,_={...n};h&&(_.search=h);let x=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${l} ${e}</h1>
      <div class="page-actions">
        ${c?`<button class="btn btn-primary" id="btn-create">+ Tambah ${i}</button>`:""}
      </div>
    </div>

    ${C?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${b?Re(b.moduleName):""}

    ${r&&r.length>0?`
    <div class="filter-bar card" style="padding: 1rem;">
        ${r.map(v=>{if(v.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${_.search||""}"></div>`;if(v.type==="search-combo"){let f="dl-filter-search",$=(v.options||[]).map(P=>`<option value="${typeof P=="object"?P.label:P}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${f}" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${_.search||""}"><datalist id="${f}">${$}</datalist></div>`}return v.type==="select"?`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${v.name}" id="filter-${v.name}"><option value="">-- ${v.label} --</option>${(v.options||[]).map(f=>`<option value="${typeof f=="object"?f.value:f}" ${_[v.name]===(typeof f=="object"?f.value:f)?"selected":""}>${typeof f=="object"?f.label:f}</option>`).join("")}</select>`:""}).join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function L(){if(!document.getElementById("bulk-toolbar"))return;let f=document.getElementById("bulk-count"),$=document.getElementById("btn-bulk-delete"),P=document.getElementById("btn-bulk-cancel");f.textContent=`${x.size} item dipilih`,x.size>0?($.disabled=!1,P.disabled=!1):($.disabled=!0,P.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{x.clear(),document.querySelectorAll(".row-checkbox").forEach(f=>f.checked=!1);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1),L()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(x.size===0)return;let v=[...x],f=document.createElement("div");f.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",f.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${v.length} ${i}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${v.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(f),f.querySelector("#bulk-cancel-btn").addEventListener("click",()=>f.remove()),f.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let $=f.querySelector("#bulk-confirm-btn");$.disabled=!0,$.textContent="Menghapus...";let P=await k(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:v})});f.remove(),P.ok?(Z(`${v.length} ${i} berhasil dihapus.`),x.clear(),L(),R()):G(P.data?.error||"Gagal menghapus data.")})});let T=document.getElementById("filter-search"),O;if(T?.addEventListener("input",v=>{clearTimeout(O),O=setTimeout(()=>{_.search=v.target.value,S=1,x.clear(),R()},400)}),r?.forEach(v=>{v.type==="select"&&document.getElementById(`filter-${v.name}`)?.addEventListener("change",f=>{_[v.name]=f.target.value,S=1,x.clear(),R()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{_={...n},T&&(T.value=""),r?.forEach(v=>{let f=document.getElementById(`filter-${v.name}`);f&&(f.value="")}),S=1,x.clear(),R()}),document.getElementById("btn-create")?.addEventListener("click",()=>ge(null)),b){document.getElementById(`btn-export-${b.moduleName}`)?.addEventListener("click",async f=>{let $=f.target,P=$.innerHTML;$.innerHTML="\u23F3 Loading...",$.disabled=!0;try{await b.onExport()}catch{G("Gagal export data")}finally{$.innerHTML=P,$.disabled=!1}}),document.getElementById(`btn-template-${b.moduleName}`)?.addEventListener("click",()=>{b.onTemplate()});let v=document.getElementById(`input-import-${b.moduleName}`);v?.addEventListener("change",async f=>{let $=f.target.files[0];if(!$)return;v.disabled=!0;let P=document.createElement("div");P.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",P.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(P);let V=P.querySelector("#import-progress-text"),U=P.querySelector("#import-progress-bar"),D=P.querySelector("#import-summary"),I=P.querySelector("#import-close-btn");I.addEventListener("click",()=>{P.remove(),R()});try{let W=await Oe($);if(W.length===0)throw new Error("File kosong atau format salah");let X=500,ie=0,ae=0,A=0,F=W.length;V.textContent=`Ditemukan ${F} baris data. Memulai import...`;for(let E=0;E<F;E+=X){let z=W.slice(E,E+X);V.textContent=`Mengimport baris ${E+1} - ${Math.min(E+X,F)} dari ${F}...`,U.style.width=`${Math.round(E/F*100)}%`;try{let j=await b.onImport(z);j?(ie+=j.inserted||j.metrics?.inserted||z.length,ae+=j.skipped||j.metrics?.updated||0):ie+=z.length}catch(j){console.error("Chunk import failed:",j),A+=z.length}}U.style.width="100%",V.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',D.style.display="block",D.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${F}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ie}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${A}</strong></div>
        `,A>0&&(D.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),I.style.display="block",v.value=""}catch(W){V.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${W.message}`,U.style.background="var(--danger)",U.style.width="100%",I.style.display="block",v.value=""}finally{v.disabled=!1}})}async function R(){L();let v=document.getElementById("table-container");if(!v)return;v.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let f=w==="client",$=f?1:S,P=f?be:20,V=new URLSearchParams({page:$,limit:P,...Object.fromEntries(Object.entries(_).filter(([,A])=>A))}),U=await k(`${a}?${V}`);if(!U.ok){v.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${U.data?.error||"Error"}</p></div>`;return}let D=U.data?.data||U.data||[],I=U.data?.pagination,W=D.length,X=D;if(f){D=y(D),X=D;let A=D.length,F=20,E=Math.ceil(A/F);S>E&&E>0&&(S=E);let z=(S-1)*F,j=S*F;D=D.slice(z,j),I={page:S,limit:F,total:A,pages:E}}!1,g&&g(D);let ie=Ue({columns:s,data:D,fullData:X,onEdit:m?A=>ge(A):null,actions:u.map(A=>({...A,handler:F=>A.handler(F,R)})),emptyText:`Tidak ada ${String(i||"").toLowerCase()}`,bulkSelect:C?{selectedIds:x,onToggle:L}:null});v.innerHTML="",v.appendChild(ie);let ae=document.getElementById("pagination-container");if(ae&&(ae.innerHTML="",I&&I.pages>1)){let A=Ge({page:I.page,pages:I.pages,total:I.total,limit:I.limit,onPage:F=>{S=F,R()}});A&&ae.appendChild(A)}}function fe(v){let f=typeof d=="function"?d(v):d;return Me(f)}function ge(v){let f=!!v,$=document.createElement("form");if($.noValidate=!0,$.innerHTML=fe(v),f){let V=typeof d=="function"?d(v):d;ze($,v)}let{close:P}=se({title:f?`Edit ${i}`:`Tambah ${i}`,content:$,size:"lg",confirmText:f?"Simpan Perubahan":`Tambah ${i}`,onConfirm:async(V,U)=>{if(!$.reportValidity())return;let D=V.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let I=Qe($),W=typeof d=="function"?d(v):d,X=async F=>{for(let E of F)if(E.type==="row")await X(E.fields);else if(E.type==="combobox"&&I[E.name]){let z=I[E.name],j=(E.options||[]).find(Y=>{let ne=String(typeof Y=="object"?Y.value:Y),tt=String(typeof Y=="object"?Y.label:Y);return ne===z||tt===z});if(j)I[E.name]=typeof j=="object"?j.value:j;else if(E.createApi){let Y={};Y[E.createApi.field]=z,E.createApi.extra&&Object.assign(Y,E.createApi.extra);let ne=await k(E.createApi.path,{method:"POST",body:JSON.stringify(Y)});if(ne.ok&&ne.data?.id)I[E.name]=ne.data.id;else if(ne.ok&&!ne.data?.id)I[E.name]=z;else throw new Error(`Gagal membuat master data: ${ne.data?.error||"Unknown error"}`)}}};try{await X(W)}catch(F){G(F.message),D.disabled=!1,D.textContent=f?"Simpan Perubahan":`Tambah ${i}`;return}o&&(I=await o(I,v));let ie=f?"PUT":"POST",ae=f?`${a}/${v.id}`:a,A=await k(ae,{method:ie,body:JSON.stringify(I)});A.ok?(Z(f?`${i} berhasil diperbarui.`:`${i} berhasil ditambahkan.`),U(),R()):(G(A.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=f?"Simpan Perubahan":`Tambah ${i}`)}})}function ve(v){Fe(`Hapus ${i} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let f=await k(`${a}/${v.id}`,{method:"DELETE"});f.ok?(Z(`${i} berhasil dihapus.`),R()):G(f.data?.error||"Gagal menghapus.")},`Hapus ${i}`)}return R(),R}M();K();async function Kt(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let l=await H(),a=await ee(),s=e?e.get("dash_filter"):null,d={};if(s==="reliever"){let c=new Date;d={status:"Done",month:`${c.getFullYear()}-${String(c.getMonth()+1).padStart(2,"0")}`}}console.log("RAW",await we()),console.log("OPTIONS",a);let r=c=>c&&!a.find(m=>m.value===c)?[...a,{value:c,label:c}]:a,n=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],i=c=>c&&!n.includes(c)?[...n,c]:n;Rt({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",defaultFilters:d,onDataLoaded:c=>c.sort((m,p)=>{let o=m.backup_date?new Date(m.backup_date).getTime():0;return(p.backup_date?new Date(p.backup_date).getTime():0)-o}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:c=>oe(c)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:c=>window.formatDate(c)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:c=>window.formatDate(c)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:c=>c?`<span class="badge badge-info">${c}</span>`:"-"},{key:"status",label:"Status",render:c=>q(c)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:n},{type:"select",name:"branch_id",label:"Cabang",options:l},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:c=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:l,value:c?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:c?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:r(c?.original_fc_name),value:c?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:i(c?.reliever_name),value:c?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:c?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:c?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:c?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:c?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:c?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let c=await k(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(c.ok){let m=c.data.data.map(p=>({Cabang:p.branch_name||"","Nama Facility care":p.original_fc_name||"",Periode:p.period||"",Relifer:p.reliever_name||"","Tanggal Back Up":p.backup_date||"","Tanggal Selesai":p.completion_date||"",Keterangan:p.reason||"",Shift:p.shift||"",Status:p.status||""}));m.length===0&&m.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),B(m,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async c=>{let p=(await k("/api/branches?all=1")).data?.data||[],o=u=>{if(!u)return null;let h=String(u||"").toLowerCase(),b=p.find(C=>String(C.full_name||"").toLowerCase()===h||String(C.code||"").toLowerCase()===h||String(C.name||"").toLowerCase()===h);return b?b.id:null},g=c.map(u=>({branch_name:String(u.Cabang||"").trim(),backup_date:String(u["Tanggal Back Up"]||u["Tanggal Backup"]||"").trim(),original_fc_name:String(u["Nama Facility care"]||u["FC Digantikan"]||"").trim(),reliever_name:String(u.Relifer||u.Reliefer||"").trim(),period:String(u.Periode||"").trim(),reason:String(u.Keterangan||"").trim(),shift:String(u.Shift||"").trim(),completion_date:String(u["Tanggal Selesai"]||"").trim(),status:String(u.Status||"").trim()})).filter(u=>u.reliever_name&&u.backup_date),y=await k("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:g,onDuplicate:"update"})});if(!y.ok)throw new Error(y.data?.error||"Import gagal");return y.data}}})}M();K();async function Ht(t){let e=await H(),l=Array.from({length:4},(a,s)=>String(new Date().getFullYear()-s));N({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>oe(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>q(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:l}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let s=new URLSearchParams(a||{}).toString(),d=await k(`/api/reports/inspection?limit=10000&${s}`);if(d.ok){let r=d.data.data.map(n=>({Cabang:n.branch_name||"",Periode:n.period||"",Tanggal:n.inspection_date||"","Point FC":n.fc_score!==null&&n.fc_score!==void 0?n.fc_score:"","Point SPV":n.spv_score!==null&&n.spv_score!==void 0?n.spv_score:"",Status:n.status||"","Link Dokumen":n.document_link||""}));B(r,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let s=i=>{if(!i)return null;let c=String(i||"").toLowerCase(),m=e.find(p=>String(p.label||"").toLowerCase()===c);return m?m.value:null},d=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let c=String(i).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let p=Number(c);if(p>2e4&&p<99999){let o=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(o.getTime())?"":o.toISOString().slice(0,10)}}let m=c.split(/[\/\-\.]/);if(m.length===3){let[p,o,g]=m.map(y=>y.trim());if(p.length===4&&o.length<=2&&g.length<=2)return`${p}-${o.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&o.length<=2&&p.length<=2)return`${g}-${o.padStart(2,"0")}-${p.padStart(2,"0")}`}return c},r=a.map(i=>({branch_id:s(String(i.Cabang||"").trim()),period:String(i.Periode||"").trim(),inspection_date:d(i.Tanggal),fc_score:i["Point FC"]!==void 0&&i["Point FC"]!==""?Number(i["Point FC"]):null,spv_score:i["Point SPV"]!==void 0&&i["Point SPV"]!==""?Number(i["Point SPV"]):null,status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.period&&i.inspection_date),n=await k("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}M();K();async function jt(t){let e=await H(),l=Array.from({length:4},(a,s)=>String(new Date().getFullYear()-s));N({container:t,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>oe(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>q(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:l}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let s=new URLSearchParams(a||{}).toString(),d=await k(`/api/reports/cleaning?limit=10000&${s}`);if(d.ok){let r=d.data.data.map(n=>({Cabang:n.branch_name||"",Jenis:n.activity_type||"",Periode:n.period||"",Tanggal:n.activity_date||"",Status:n.status||"","Link Dokumen":n.document_link||""}));B(r,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let s=i=>{if(!i)return null;let c=String(i||"").toLowerCase(),m=e.find(p=>String(p.label||"").toLowerCase()===c);return m?m.value:null},d=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let c=String(i).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let p=Number(c);if(p>2e4&&p<99999){let o=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(o.getTime())?"":o.toISOString().slice(0,10)}}let m=c.split(/[\/\-\.]/);if(m.length===3){let[p,o,g]=m.map(y=>y.trim());if(p.length===4&&o.length<=2&&g.length<=2)return`${p}-${o.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&o.length<=2&&p.length<=2)return`${g}-${o.padStart(2,"0")}-${p.padStart(2,"0")}`}return c},r=a.map(i=>({branch_id:s(String(i.Cabang||"").trim()),activity_type:String(i.Jenis||i.Kegiatan||"").trim(),period:String(i.Periode||"").trim(),activity_date:d(i.Tanggal),status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.activity_type&&i.period&&i.activity_date),n=await k("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}M();K();async function qt(t,e){let l=await H(),a=Array.from({length:4},(r,n)=>String(new Date().getFullYear()-n)),s=e?e.get("dash_filter"):null,d={};if(s==="fogging"){let r=new Date,n=String(r.getMonth()+1).padStart(2,"0"),i=String(r.getFullYear());d={status:"Done",month:n,year:i}}N({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,defaultFilters:d,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:r=>`<span class="badge badge-warning">${r}</span>`},{key:"period",label:"Periode",render:r=>oe(r)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"status",label:"Status",render:r=>q(r)},{key:"document_link",label:"Dokumen",render:r=>r?`<a href="${r}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:r=>r||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:l},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:r=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:l,value:r?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:r?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:r?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:r?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:r?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async r=>{let n=new URLSearchParams(r||{}).toString(),i=await k(`/api/reports/fogging?limit=10000&${n}`);if(i.ok){let c=i.data.data.map(m=>({Cabang:m.branch_name||"",Jenis:m.activity_type||"Fogging",Periode:m.period||"",Tanggal:m.activity_date||"",Status:m.status||"","Link Dokumen":m.document_link||""}));B(c,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async r=>{let n=p=>{if(!p)return null;let o=String(p||"").toLowerCase(),g=l.find(y=>String(y.label||"").toLowerCase()===o);return g?g.value:null},i=p=>{if(p==null||p==="")return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let o=String(p).trim();if(o===""||o==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(o))return o.slice(0,10);if(/^\d{4,5}$/.test(o)){let y=Number(o);if(y>2e4&&y<99999){let u=new Date(Date.UTC(1899,11,30)+y*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}let g=o.split(/[\/\-\.]/);if(g.length===3){let[y,u,h]=g.map(b=>b.trim());if(y.length===4&&u.length<=2&&h.length<=2)return`${y}-${u.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&u.length<=2&&y.length<=2)return`${h}-${u.padStart(2,"0")}-${y.padStart(2,"0")}`}return o},c=r.map(p=>({branch_id:n(String(p.Cabang||"").trim()),activity_type:String(p.Jenis||p.Kegiatan||"Fogging").trim(),period:String(p.Periode||"").trim(),activity_date:i(p.Tanggal),status:String(p.Status||"").trim(),document_link:String(p["Link Dokumen"]||"").trim(),notes:String(p.Catatan||p.Keterangan||"").trim()})).filter(p=>p.branch_id&&p.period&&p.activity_date),m=await k("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(c)});if(!m.ok)throw new Error(m.data?.error||"Import gagal");return m.data}}})}M();K();async function Jt(t){let e=await H(),l=await ee(),a=l,s=Array.from({length:4},(n,i)=>String(new Date().getFullYear()-i)),d=n=>n&&!l.find(i=>i.value===n)?[...l,{value:n,label:n}]:l,r=n=>n&&!a.find(i=>i.value===n)?[...a,{value:n,label:n}]:a;N({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:n=>`<span title="${n||""}">${n?.length>60?n.slice(0,60)+"\u2026":n||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:n=>window.formatDate(n)},{key:"status",label:"Status",render:n=>q(n)},{key:"notes",label:"Keterangan",render:n=>n?.length>40?n.slice(0,40)+"\u2026":n||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade","Mizwar"]},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:s}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:n?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:r(n?.pic),value:n?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:n?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:n?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:n?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:n?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async n=>{let i=new URLSearchParams(n||{}).toString(),c=await k(`/api/reports/basecamp?limit=10000&${i}`);if(c.ok){let m=c.data.data.map(p=>({"Tgl Info":p.info_date||"",Cabang:p.branch_name||"",Permasalahan:p.problem||"",PIC:p.pic||"","Tgl Done":p.done_date||"",Status:p.status||"",Keterangan:p.notes||""}));B(m,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async n=>{let i=o=>{if(!o)return null;let g=String(o||"").toLowerCase(),y=e.find(u=>String(u.label||"").toLowerCase()===g);return y?y.value:null},c=o=>{if(o==null||o==="")return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let g=String(o).trim();if(g===""||g==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);if(/^\d{4,5}$/.test(g)){let u=Number(g);if(u>2e4&&u<99999){let h=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}let y=g.split(/[\/\-\.]/);if(y.length===3){let[u,h,b]=y.map(C=>C.trim());if(u.length===4&&h.length<=2&&b.length<=2)return`${u}-${h.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&h.length<=2&&u.length<=2)return`${b}-${h.padStart(2,"0")}-${u.padStart(2,"0")}`}return g},m=n.map(o=>({info_date:c(o["Tgl Info"]||o["Tanggal Info"]),branch_id:i(String(o.Cabang||"").trim()),problem:String(o.Permasalahan||"").trim(),pic:String(o.PIC||"").trim(),done_date:c(o["Tgl Done"]||o["Tanggal Done"]),status:String(o.Status||"").trim(),notes:String(o.Keterangan||o.Catatan||"").trim()})).filter(o=>o.info_date&&o.branch_id&&o.problem),p=await k("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(m)});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}}})}async function Ut(t){N({container:t,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let l=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(M(),ke)),s=await a(`/api/sop?limit=10000&${l}`);if(s.ok){let d=s.data.data.map(n=>({"Nama SOP":n.name||"",Kategori:n.category||"",Dokumen:n.document_link||"",Catatan:n.notes||n.description||""})),{downloadExcel:r}=await Promise.resolve().then(()=>(K(),le));r(d,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:l}=await Promise.resolve().then(()=>(K(),le));l(e,"Template_Import_SOP")},onImport:async e=>{let l=e.map(d=>({name:String(d["Nama SOP"]||"").trim(),category:String(d.Kategori||"").trim(),document_link:String(d.Dokumen||"").trim(),description:String(d.Catatan||"").trim()})).filter(d=>d.name),{apiFetch:a}=await Promise.resolve().then(()=>(M(),ke)),s=await a("/api/sop/import",{method:"POST",body:JSON.stringify(l)});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function Gt(t){N({container:t,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let l=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(M(),ke)),s=await a(`/api/checklist?limit=10000&${l}`);if(s.ok){let d=s.data.data.map(n=>({"Nama Checklist":n.name||"",Kategori:n.category||"",Dokumen:n.document_link||"",Deskripsi:n.description||""})),{downloadExcel:r}=await Promise.resolve().then(()=>(K(),le));r(d,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:l}=await Promise.resolve().then(()=>(K(),le));l(e,"Template_Import_Checklist")},onImport:async e=>{let l=e.map(d=>({name:String(d["Nama Checklist"]||"").trim(),category:String(d.Kategori||"").trim(),document_link:String(d.Dokumen||"").trim(),description:String(d.Deskripsi||"").trim()})).filter(d=>d.name),{apiFetch:a}=await Promise.resolve().then(()=>(M(),ke)),s=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(l)});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}M();ye();K();async function ht(t,e="forms"){if(e==="supply")return Ba(t);Pa(t)}function Pa(t){N({container:t,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let l=new URLSearchParams(e||{}).toString(),a=await k(`/api/forms?limit=10000&${l}`);a.data?.data?B(a.data.data,"Data_Master_Form"):G("Gagal export data master form")},onImport:async e=>{let l=await k("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!l.ok)throw new Error(l.data?.error||"Import failed");return l.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function Ba(t){let l=((await k("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));N({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,s)=>s.branch_name_ref||s.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let s=JSON.parse(a);return Array.isArray(s)?s.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let s=JSON.parse(a);return Array.isArray(s)?s.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>q(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let s=a?.tools_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}let d=a?.chemical_items;try{d=Array.isArray(JSON.parse(d))?JSON.parse(d).join(", "):d}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:a?.branch_id&&!l.find(r=>r.value==a.branch_id)?[...l,{value:a.branch_id,label:a.branch_name||a.branch_id}]:l,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:s},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:d},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let s=new URLSearchParams(a||{}).toString(),d=await k(`/api/reports/supply?limit=10000&${s}`);if(d.ok){let r=d.data.data.map(n=>{let i=n.tools_items;try{i=Array.isArray(JSON.parse(i))?JSON.parse(i).join(", "):i}catch{}let c=n.chemical_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}return{Waktu:n.submitted_at||"",Pengirim:n.submitter_name||"",Cabang:n.branch_name_ref||n.branch_name||"","Alat/Barang":i||"",Chemical:c||"",Catatan:n.additional_notes||"",Status:n.status||"","Diproses Oleh":n.processed_by||""}});B(r,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let d=(await k("/api/branches?all=1")).data?.data||[],r=m=>{if(!m)return null;let p=String(m||"").toLowerCase(),o=d.find(g=>String(g.full_name||"").toLowerCase()===p||String(g.code||"").toLowerCase()===p||String(g.name||"").toLowerCase()===p);return o?o.id:null},n=m=>{if(m==null||m==="")return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let p=String(m).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let g=Number(p);if(g>2e4&&g<99999){let y=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}let o=p.split(/[\/\-\.]/);if(o.length===3){let[g,y,u]=o.map(h=>h.trim());if(g.length===4&&y.length<=2&&u.length<=2)return`${g}-${y.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&y.length<=2&&g.length<=2)return`${u}-${y.padStart(2,"0")}-${g.padStart(2,"0")}`}return p},i=a.map(m=>({submitted_at:n(m.Waktu||m.Tanggal),submitter_name:String(m.Pengirim||"").trim(),branch_id:r(String(m.Cabang||"").trim()),tools_items:String(m["Alat/Barang"]||m.Alat||"").trim(),chemical_items:String(m.Chemical||"").trim(),additional_notes:String(m.Catatan||m.Keterangan||"").trim(),status:String(m.Status||"").trim(),processed_by:String(m["Diproses Oleh"]||m.PIC||"").trim()})).filter(m=>m.submitted_at&&m.submitter_name&&m.branch_id),c=await k("/api/reports/supply/import",{method:"POST",body:JSON.stringify(i)});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,s)=>{let d=se({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(r,n)=>{let i=r.querySelector("#supply-status").value,c=r.querySelector("#supply-processed-by").value;(await k(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:i,processed_by:c})})).ok?(Z("Status diperbarui."),n(),s()):G("Gagal update status.")}})}}]})}M();K();async function Qt(t){let e=he();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}N({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:l=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[l]||"badge-neutral"}">${l}</span>`},{key:"is_active",label:"Status",render:l=>l?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:l=>l?new Date(l).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:l=>{let a=!!l;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:l?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:l?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:l?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:l?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?l?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let l=await k(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(l.ok){let a=l.data.data.map(s=>({"Nama Lengkap":s.full_name||"",Username:s.username||"",Email:s.email||"",Role:s.role||"",Status:s.is_active?"Aktif":"Nonaktif"}));B(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async l=>{let a=l.map(d=>({full_name:String(d["Nama Lengkap"]||"").trim(),username:String(d.Username||"").trim(),email:String(d.Email||"").trim(),role:String(d.Role||"").trim()||"viewer",password:String(d.Password||"").trim()})).filter(d=>d.username&&d.password&&d.email&&d.full_name),s=await k("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}M();K();async function zt(t){N({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await k(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)B(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let l=e.map(s=>({code:String(s["Kode Cabang"]||"").trim(),name:String(s["Nama Pendek"]||"").trim(),full_name:String(s["Nama Lengkap"]||"").trim(),city:String(s.Kota||"").trim()})).filter(s=>s.code&&s.name),a=await k("/api/branches/import",{method:"POST",body:JSON.stringify(l)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}M();async function Vt(t){let e=new Date,l=[];t.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),s()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),s()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(d=>d.addEventListener("change",s));async function a(){try{let d=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;l=(await k(`/api/dashboard/calendar?month=${d}`)).data?.data||[]}catch(d){console.warn("[Calendar] Failed to load events, rendering empty grid:",d),l=[]}}async function s(){let d=document.getElementById("calendar-grid");if(d){d.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let r=e.getFullYear(),n=e.getMonth(),i=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),c=document.getElementById("cal-month-label");c&&(c.textContent=i);let m=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(S=>S.value)),p=l.filter(S=>m.has(S.type)),o={};p.forEach(S=>{let _=(S.event_date||"").slice(0,10);o[_]||(o[_]=[]),o[_].push(S)});let g=new Date(r,n,1).getDay(),y=new Date(r,n+1,0).getDate(),u=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],h=new Date().toISOString().slice(0,10),b='<div class="calendar-grid">';u.forEach(S=>{b+=`<div class="cal-day-header">${S}</div>`});for(let S=0;S<g;S++)b+='<div class="cal-cell cal-cell-empty"></div>';for(let S=1;S<=y;S++){let _=`${r}-${String(n+1).padStart(2,"0")}-${String(S).padStart(2,"0")}`,x=o[_]||[],L=_===h;b+=`
          <div class="cal-cell ${L?"cal-today":""} ${x.length?"cal-has-events":""}"
               data-date="${_}" tabindex="0" role="button" aria-label="${_}">
            <div class="cal-day-num ${L?"today-num":""}">${S}</div>
            <div class="cal-events-preview">
              ${x.slice(0,3).map(T=>`
                <div class="cal-event-dot cal-color-${T.color||"gray"}" title="${Xe(T.title||T.type)}">
                  <span class="cal-event-dot-label">${La(T.title||T.branch_name||T.type,18)}</span>
                </div>
              `).join("")}
              ${x.length>3?`<div class="cal-more">+${x.length-3} lagi</div>`:""}
            </div>
          </div>`}let w=(g+y)%7;if(w!==0)for(let S=0;S<7-w;S++)b+='<div class="cal-cell cal-cell-empty"></div>';b+="</div>",d.innerHTML=b,d.querySelectorAll(".cal-cell[data-date]").forEach(S=>{S.addEventListener("click",()=>{let _=S.dataset.date,x=o[_]||[];if(!x.length)return;let L=document.getElementById("cal-event-list"),T=new Date(_+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=T,document.getElementById("cal-event-items").innerHTML=x.map(O=>`
            <div class="cal-event-item cal-color-border-${O.color||"gray"}">
              <div class="cal-event-type">${Na(O.type)}</div>
              <div class="cal-event-title">${Xe(O.title||"-")}</div>
              <div class="cal-event-branch">${Xe(O.branch_name||"")}</div>
              ${O.status?`<div class="cal-event-status">${Xe(O.status)}</div>`:""}
              ${O.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${O.days_remaining} hari</div>`:""}
            </div>
          `).join(""),L.style.display="block"})})}catch(r){console.error("[Calendar] Render error:",r),d&&(d.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}s()}function La(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function Xe(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Na(t){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[t]||t}M();async function Yt(t){let e=he(),l=(e?.full_name||e?.username||"U")[0].toUpperCase(),s={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${s},${s}99)">
            ${l}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${e?.full_name||"\u2014"}</div>
            <div class="profile-username">@${e?.username||"\u2014"}</div>
            <span class="badge badge-info" style="background:${s}18;color:${s};margin-top:6px">
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
            <span class="info-value" style="color:${s};font-weight:700">${e?.role||"\u2014"}</span>
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
  `;let d=localStorage.getItem("fm_token"),r=document.getElementById("session-info");if(d&&r)try{let n=JSON.parse(atob(d.split(".")[1])),i=new Date(n.exp*1e3);r.textContent=`Berakhir: ${i.toLocaleString("id-ID")}`}catch{r.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async n=>{n.preventDefault();let i=document.getElementById("pwd-error"),c=document.getElementById("pwd-success"),m=document.getElementById("btn-save-pwd");i.style.display="none",c.style.display="none";let p=n.target,o=p.current_password.value,g=p.new_password.value,y=p.confirm_password.value;if(g!==y){i.textContent="\u274C Konfirmasi password tidak cocok.",i.style.display="block";return}if(g.length<6){i.textContent="\u274C Password baru minimal 6 karakter.",i.style.display="block";return}m.disabled=!0,m.textContent="\u23F3 Menyimpan...";let u=await k("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:o,new_password:g})});m.disabled=!1,m.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',u.ok?(c.textContent="\u2705 Password berhasil diubah.",c.style.display="block",p.reset(),Z("Password berhasil diubah.")):(i.textContent=u.data?.error||"Gagal mengubah password.",i.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}M();var Ze={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function te(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let s=Number(e);if(s>2e4&&s<99999){let d=new Date(Date.UTC(1899,11,30)+s*864e5);return isNaN(d.getTime())?null:d.toISOString().slice(0,10)}}let l=e.split(/[\/\-\.]/);if(l.length===3){let[s,d,r]=l.map(m=>m.trim()),n=Number(s),i=Number(d),c=Number(r);if(s.length===4&&n>1900)return`${s}-${d.padStart(2,"0")}-${r.padStart(2,"0")}`;if(r.length===4&&c>1900)return n>12?`${r}-${d.padStart(2,"0")}-${s.padStart(2,"0")}`:i>12?`${r}-${s.padStart(2,"0")}-${d.padStart(2,"0")}`:`${r}-${d.padStart(2,"0")}-${s.padStart(2,"0")}`;if(r.length===2&&!isNaN(c)){let m=c>=50?`19${r}`:`20${r}`;return n>12?`${m}-${d.padStart(2,"0")}-${s.padStart(2,"0")}`:`${m}-${d.padStart(2,"0")}-${s.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function Wt(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Aa={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",phone:t["No. Hp"],join_date:te(t["Tanggal Masuk"]),status:t.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:te(t["Tanggal Mulai"]),end_date:te(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:te(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:te(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:te(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:te(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:te(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:te(t["Tanggal Target"]||t["Tgl Target"]),completion_date:te(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:te(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:te(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:te(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:te(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:te(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:te(t["Tanggal Back Up"]),completion_date:te(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:te(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:te(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Fa(t,e){let l=Ze[t];if(!l)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Aa[l.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let s=[],d=[],r=[];return e.filter(i=>!Wt(i)).forEach((i,c)=>{let m=e.indexOf(i)+2,p=[];a.required.forEach(({key:g,label:y})=>{let u=i[g];if(u==null||String(u).trim()===""){let h=Object.keys(i).filter(b=>b.trim()).join(", ");p.push({column:y,originalValue:u||"",reason:`Kolom "${y}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${h.slice(0,120)}`})}});let o=a.map(i);p.length>0?d.push({row:m,data:o,raw:i,errors:p}):(s.push(i),r.push(o))}),{valid:s,errors:d,mapped:r}}function Xt(t){let e=[];return t.SheetNames.forEach(l=>{let a=Ze[l];if(!a)return;let s=t.Sheets[l],d=window.XLSX.utils.sheet_to_json(s,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),r=Fa(l,d),n=d.filter(i=>!Wt(i));e.push({sheetName:l,module:a.module,label:a.label,total:n.length,valid:r.mapped.length,errorCount:r.errors.length,errors:r.errors,mapped:r.mapped,skipped:!1})}),e}function Zt(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,s])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(s),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function ea(t){let e=window.XLSX,l=e.utils.book_new(),a=!1;return t.forEach(s=>{if(!s.errors||s.errors.length===0)return;a=!0;let d=s.errors.map(n=>({"No. Baris":n.row,"Kolom Gagal":(n.errors||[]).map(i=>i.column||i).join("; "),"Alasan Error":(n.errors||[]).map(i=>i.reason||i).join("; "),...Object.fromEntries(Object.entries(n.data||{}).map(([i,c])=>[i,c??""]))})),r=e.utils.json_to_sheet(d);e.utils.book_append_sheet(l,r,s.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(l,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Ma=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function ta(t){t.innerHTML=`
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
              ${Object.entries(Ze).map(([u,{label:h}])=>`<span class="import-sheet-tag">\u{1F4C4} ${u} \u2192 ${h}</span>`).join("")}
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
  `;let e=null,l=null,a=0,s={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function d(u){Object.entries(s).forEach(([h,b])=>{b.style.display=h===u?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let u=document.getElementById("btn-backup-db");u.disabled=!0,u.textContent="\u23F3 Memproses Backup...";try{let h=await k("/api/import/backup");if(h.ok){if(!window.XLSX){G("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let b=window.XLSX,C=b.utils.book_new();Object.entries(h.data.database).forEach(([w,S])=>{let _=S.length>0?S:[{}],x=b.utils.json_to_sheet(_);b.utils.book_append_sheet(C,x,w.substring(0,31))}),b.writeFile(C,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),Z("Backup berhasil diunduh!")}else G("Gagal memproses backup: "+(h.data?.error||"Unknown error"))}catch(h){G("Gagal memproses backup: "+h.message)}finally{u.disabled=!1,u.textContent="\u{1F4E6} Backup Database"}});let r=document.getElementById("btn-sync-google");r&&r.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let u=r.innerHTML;r.innerHTML='<span class="spinner"></span> Menyinkronkan...',r.disabled=!0;try{let h=await k("/api/sync/google-sheets",{method:"POST"});h.ok?alert("Sinkronisasi Berhasil: "+(h.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(h.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{r.innerHTML=u,r.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Zt(),Z("Template Excel berhasil didownload!")});let n=document.getElementById("file-input"),i=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",u=>{u.stopPropagation(),n.click()}),n.addEventListener("change",u=>{u.target.files[0]&&c(u.target.files[0])}),i.addEventListener("dragover",u=>{u.preventDefault(),i.classList.add("drag-over")}),i.addEventListener("dragleave",()=>i.classList.remove("drag-over")),i.addEventListener("drop",u=>{u.preventDefault(),i.classList.remove("drag-over");let h=u.dataTransfer.files[0];h&&h.name.match(/\.xlsx?$/i)?c(h):G("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,n.value="",document.getElementById("file-info").style.display="none",i.style.display="",d("upload")});async function c(u){e=u,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${u.name} (${(u.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",i.style.display="none",await m(u)}async function m(u){d("validating");let h=document.getElementById("validation-status"),b=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");h.textContent="Membaca file Excel...",b.style.width="20%",await je(200);let C=await u.arrayBuffer(),w=window.XLSX.read(C,{type:"array",cellDates:!0});h.textContent=`Memvalidasi ${w.SheetNames.length} sheet...`,b.style.width="50%",await je(100),l=Xt(w),b.style.width="100%",h.textContent="Validasi selesai!",await je(300),p()}catch(C){d("upload"),G("Gagal memproses file: "+C.message),document.getElementById("file-info").style.display="flex",i.style.display="none"}}function p(){d("preview");let u=l.filter(T=>!T.skipped).length,h=l.reduce((T,O)=>T+O.total,0),b=l.reduce((T,O)=>T+O.valid,0),C=l.reduce((T,O)=>T+O.errorCount,0),w=h>0?Math.round(b/h*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${u} sheet</span>
      <span class="badge badge-secondary">${h} baris</span>
      <span class="badge badge-success">${b} valid (${w}%)</span>
      ${C>0?`<span class="badge badge-danger">${C} error</span>`:""}
    `;let S=document.getElementById("preview-table-container");S.innerHTML=`
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
          ${l.map((T,O)=>`
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
                ${T.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${O}">\u{1F50D} ${T.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,S.querySelectorAll(".btn-detail-error").forEach(T=>{T.addEventListener("click",()=>{let O=l[Number(T.dataset.idx)];o(O)})});let _=document.getElementById("error-detail-section"),x=document.getElementById("error-detail-container");x.innerHTML="",_.style.display="none";let L=document.getElementById("btn-start-import");b===0?(L.disabled=!0,L.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(L.disabled=!1,C>0?(L.innerHTML=`\u{1F680} Import ${b} Data Valid (${C} dilewati)`,L.title="Baris error akan dilewati, baris valid tetap diimport"):L.innerHTML=`\u{1F680} Mulai Import ${b} Data`)}function o(u){let h=document.getElementById("error-detail-section"),b=document.getElementById("error-detail-container");h.style.display="";let C=u.errors.slice(0,100).map(w=>(Array.isArray(w.errors)?w.errors:[]).map(_=>{let x=typeof _=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${w.row}</span></td>
            <td><strong>${x?_.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${x&&_.originalValue!==void 0?_.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${x?_.reason:_}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${x&&_.aliases?`Gunakan salah satu nama kolom:<br><em>${_.aliases}</em>`:x&&_.hint?_.hint:""}
            </td>
          </tr>
        `}).join("")).join("");b.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${u.sheetName} \u2014 ${u.errorCount} baris error dari ${u.total} total
          ${u.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
            <tbody>${C||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${u.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,h.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{d("upload"),document.getElementById("file-info").style.display="none",i.style.display="",e=null,n.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!l)return;ea(l)?Z("Log error berhasil didownload."):Z("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let u=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";g(u)});async function g(u){d("importing"),a=Date.now();let h=[];Ma.forEach(_=>{let x=l?.find(L=>L.module===_&&L.mapped?.length>0);x&&h.push(x)});let b=document.getElementById("import-steps-list");b.innerHTML=h.map(_=>`
      <div class="import-step-item" id="step-item-${_.module}">
        <span class="step-item-icon" id="step-icon-${_.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${_.label} <span class="step-item-count">(${_.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${_.module}"></span>
      </div>
    `).join("");let C=document.getElementById("import-bar"),w=document.getElementById("import-current-status"),S={totalSheets:h.length,totalRows:h.reduce((_,x)=>_+x.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let _=0;_<h.length;_++){let x=h[_],L=document.getElementById(`step-icon-${x.module}`),T=document.getElementById(`step-status-${x.module}`);L.textContent="\u{1F504}",T.textContent="Mengimport...",w.textContent=`Mengimport ${x.label}...`,C.style.width=`${Math.round(_/h.length*100)}%`;try{let O=await k(`/api/import/${x.module}`,{method:"POST",body:JSON.stringify({rows:x.mapped,onDuplicate:u})});if(O.ok){let R=O.data;S.inserted+=R.inserted||0,S.skipped+=R.skipped||0,S.moduleResults.push({label:x.label,inserted:R.inserted||0,skipped:R.skipped||0,status:"ok"}),L.textContent="\u2705",T.innerHTML=`<span class="badge badge-success">${R.inserted||0} berhasil</span>${R.skipped>0?` <span class="badge badge-neutral">${R.skipped} skip</span>`:""}`}else S.failed++,S.moduleResults.push({label:x.label,inserted:0,skipped:0,status:"error",error:O.data?.error}),L.textContent="\u274C",T.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(O){S.failed++,S.moduleResults.push({label:x.label,inserted:0,skipped:0,status:"error",error:O.message}),L.textContent="\u274C",T.innerHTML='<span class="badge badge-danger">Gagal</span>'}await je(150)}C.style.width="100%",w.textContent="Selesai!",await je(400),y(S)}function y(u){d("summary");let h=((Date.now()-a)/1e3).toFixed(1),b=u.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${b?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${b?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${u.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${u.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${u.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${u.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${u.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${u.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
      <div class="summary-stat-card info">
        <div class="stat-value">${h}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${u.moduleResults.map(C=>`
            <tr>
              <td>${C.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${C.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${C.skipped}</span></td>
              <td style="text-align:center">
                ${C.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${C.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,l=null,n.value="",document.getElementById("file-info").style.display="none",i.style.display="",d("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function je(t){return new Promise(e=>setTimeout(e,t))}M();var et=[],aa=[];async function na(t){et=await H(),aa=await ee(),N({container:t,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:et}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let l=new URLSearchParams(e||{}).toString(),a=await k(`/api/sp?limit=10000&${l}`);if(a.ok){let s=a.data.data.map(r=>({"Nama Karyawan":r.employee_name||"",Divisi:r.division||"",Cabang:r.branch_name||"","Tanggal Sp":r.tanggal||"","Akhir Sp":r.akhir_sp||"","Jenis Sp":r.sp_type||"","Link Document / Foto":r.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(K(),le));d(s,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:l}=await Promise.resolve().then(()=>(K(),le));l(e,"Template_Import_SP")},onImport:async e=>{let l=r=>{if(!r)return null;let n=String(r||"").toLowerCase(),i=et.find(c=>String(c.label||"").toLowerCase()===n);return i?i.value:null},a=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let n=String(r).trim();if(/^\d{4,5}$/.test(n)){let c=Number(n);if(c>2e4&&c<99999){let m=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(n))return n.slice(0,10);let i=n.split(/[\/\-\.]/);if(i.length===3){let[c,m,p]=i.map(o=>o.trim());if(c.length===4&&m.length<=2&&p.length<=2)return`${c}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&m.length<=2&&c.length<=2)return`${p}-${m.padStart(2,"0")}-${c.padStart(2,"0")}`}return n},s=e.map(r=>({employee_name:String(r["Nama Karyawan"]||"").trim(),division:String(r.Divisi||"").trim(),branch_id:l(String(r.Cabang||"").trim()),tanggal:a(r["Tanggal Sp"]),akhir_sp:a(r["Akhir Sp"]),sp_type:String(r["Jenis Sp"]||"").trim(),document_link:String(r["Link Document / Foto"]||"").trim()})).filter(r=>r.employee_name&&r.branch_id),d=await k("/api/import/sp",{method:"POST",body:JSON.stringify({rows:s,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:[{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:aa},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"combobox",name:"branch_id",label:"Cabang",required:!0,options:et,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}M();var De=[],ia=[];async function ra(t){De=await H(),ia=await ee(),N({container:t,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",options:De},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",options:De}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let l=new URLSearchParams(e||{}).toString(),a=await k(`/api/mutasi?limit=10000&${l}`);if(a.ok){let s=a.data.data.map(r=>({Tanggal:r.tanggal||"","Nama Karyawan":r.employee_name||"","Cabang Asal":r.from_branch_name||"","Cabang Tujuan":r.to_branch_name||"",Status:r.status||"",Dokumen:r.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(K(),le));d(s,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:l}=await Promise.resolve().then(()=>(K(),le));l(e,"Template_Import_Mutasi")},onImport:async e=>{let l=r=>{if(!r)return null;let n=String(r||"").toLowerCase(),i=De.find(c=>String(c.label||"").toLowerCase()===n);return i?i.value:null},a=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let n=String(r).trim();if(/^\d{4,5}$/.test(n)){let c=Number(n);if(c>2e4&&c<99999){let m=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(n))return n.slice(0,10);let i=n.split(/[\/\-\.]/);if(i.length===3){let[c,m,p]=i.map(o=>o.trim());if(c.length===4&&m.length<=2&&p.length<=2)return`${c}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&m.length<=2&&c.length<=2)return`${p}-${m.padStart(2,"0")}-${c.padStart(2,"0")}`}return n},s=e.map(r=>({tanggal:a(r.Tanggal),employee_name:String(r["Nama Karyawan"]||"").trim(),from_branch_id:l(String(r["Cabang Asal"]||"").trim()),to_branch_id:l(String(r["Cabang Tujuan"]||"").trim()),status:String(r.Status||"").trim(),document_link:String(r.Dokumen||"").trim()})).filter(r=>r.tanggal&&r.employee_name&&r.from_branch_id&&r.to_branch_id),d=await k("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:s,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:ia},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:De,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:De,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}M();async function la(t){t.innerHTML=`
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
  `;let e=document.getElementById("btnRefreshSync"),l=document.getElementById("queueStatusFilter");e.addEventListener("click",s),l.addEventListener("change",n),document.getElementById("btnRetryAll").addEventListener("click",()=>a("retry",{allFailed:!0})),document.getElementById("btnResetStuck").addEventListener("click",()=>a("reset-stuck",{})),document.getElementById("btnPauseSync").addEventListener("click",()=>a("pause",{})),document.getElementById("btnResumeSync").addEventListener("click",()=>a("resume",{})),document.getElementById("btnReconcile").addEventListener("click",()=>a("reconcile",{module:"ALL",repairMode:!0})),document.getElementById("btnRetrySelected").addEventListener("click",()=>{let c=Array.from(document.querySelectorAll(".chk-queue:checked")).map(m=>m.value);if(c.length===0)return alert("No items selected");a("retry",{ids:c})}),document.getElementById("chkAllQueue").addEventListener("change",c=>{document.querySelectorAll(".chk-queue").forEach(m=>m.checked=c.target.checked)});async function a(c,m){if(confirm(`Are you sure you want to execute action: ${c}?`)){showLoading();try{let p=await k(`/api/sync/actions/${c}`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(m)});p.ok?(alert(p.data?.message||"Success"),s()):G(p.error||"Action failed")}catch(p){G(p.message)}hideLoading()}}await s();async function s(){showLoading(),await Promise.all([r(),n(),d(),i()]),hideLoading()}async function d(){try{let c=await k("/api/sync/performance");if(!c.ok)return;let{webhook:m,google_api:p,d1:o,queue:g,throughput:y}=c.data;document.getElementById("latencyHistograms").innerHTML=`
        <table class="w-full text-left">
          <thead><tr class="text-gray-500 border-b"><th>Metric</th><th>P50</th><th>P95</th><th>P99</th><th>Max</th></tr></thead>
          <tbody class="divide-y">
            <tr><td class="py-1">Webhook</td><td>${m.P50}ms</td><td>${m.P95}ms</td><td>${m.P99}ms</td><td>${m.Max}ms</td></tr>
            <tr><td class="py-1">Google API</td><td>${p.P50}ms</td><td>${p.P95}ms</td><td>${p.P99}ms</td><td>${p.Max}ms</td></tr>
            <tr><td class="py-1">D1 Execute</td><td>${o.P50}ms</td><td>${o.P95}ms</td><td>${o.P99}ms</td><td>${o.Max}ms</td></tr>
            <tr><td class="py-1">Queue Delay</td><td>${g.P50}ms</td><td>${g.P95}ms</td><td>${g.P99}ms</td><td>${g.Max}ms</td></tr>
          </tbody>
        </table>
        <div class="mt-4 pt-3 border-t text-gray-600 flex justify-between">
          <span>Throughput: <b>${y.events_per_sec}</b> ev/sec</span>
          <span><b>${y.events_per_min}</b> ev/min</span>
        </div>
      `,document.getElementById("lblSnapshotStatus").innerHTML='Checked 10 modules. Status: <b>COMPLETED</b><br><span class="text-xs text-gray-500">Run today at 02:00</span>'}catch(c){console.error(c)}}async function r(){try{let c=await k("/api/sync/health");if(!c.ok)return G("Failed to fetch sync health");let{status:m,queue:p,circuit_breaker:o}=c.data,g=`
        <div class="bg-white p-4 rounded-lg shadow border-l-4 ${m==="HEALTHY"?"border-green-500":m==="WARNING"?"border-yellow-500":"border-red-500"}">
          <p class="text-sm text-gray-500">System Health</p>
          <p class="text-2xl font-bold ${m==="HEALTHY"?"text-green-600":m==="WARNING"?"text-yellow-600":"text-red-600"}">${m}</p>
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
      `;document.getElementById("syncOverviewCards").innerHTML=g;let y=document.getElementById("cbStateBadge"),u=document.getElementById("cbStateDesc"),h=document.getElementById("cbStatusCard");h.className="bg-white rounded-lg shadow p-6 border-l-4",o==="CLOSED"?(y.className="px-3 py-1 rounded-full text-sm font-medium bg-green-100 text-green-800",y.textContent="CLOSED",u.textContent="Traffic is flowing normally to Google Sheets.",h.classList.add("border-green-500")):o==="OPEN"?(y.className="px-3 py-1 rounded-full text-sm font-medium bg-red-100 text-red-800",y.textContent="OPEN",u.textContent="Failing fast. Traffic to Google Sheets is paused due to repeated failures.",h.classList.add("border-red-500")):o==="HALF_OPEN"?(y.className="px-3 py-1 rounded-full text-sm font-medium bg-yellow-100 text-yellow-800",y.textContent="HALF-OPEN",u.textContent="Testing recovery. Permitting limited traffic to verify stability.",h.classList.add("border-yellow-500")):y.textContent=o||"UNKNOWN"}catch(c){console.error(c)}}async function n(){try{let c=document.getElementById("queueStatusFilter").value,m=await k("/api/sync/queue?limit=15"+(c?"&status="+c:""));if(!m.ok)return;let p=document.getElementById("queueTableBody"),o=m.data?.data||m.data||[];if(o.length===0){p.innerHTML='<tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">No events found</td></tr>';return}p.innerHTML=o.map(g=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2"><input type="checkbox" class="chk-queue" value="${g.id}" /></td>
          <td class="px-4 py-2 font-mono text-xs text-gray-500" title="${g.id}">${g.id.split("-")[0]}...</td>
          <td class="px-4 py-2 font-medium">${g.entity_name}</td>
          <td class="px-4 py-2">
             <span class="px-2 py-0.5 rounded text-xs ${g.action==="INSERT"?"bg-blue-100 text-blue-800":g.action==="UPDATE"?"bg-purple-100 text-purple-800":"bg-red-100 text-red-800"}">${g.action}</span>
          </td>
          <td class="px-4 py-2">
             <span class="px-2 py-0.5 rounded text-xs ${g.status==="PENDING"?"bg-yellow-100 text-yellow-800":g.status==="PROCESSING"?"bg-blue-100 text-blue-800":g.status==="DEAD_LETTER"?"bg-red-100 text-red-800":"bg-gray-100 text-gray-800"}">${g.status}</span>
             ${g.last_error?`<br><span class="text-xs text-red-500 max-w-xs block truncate" title="${g.last_error}">${g.last_error}</span>`:""}
          </td>
          <td class="px-4 py-2 text-gray-600">${g.retry_count||0}</td>
          <td class="px-4 py-2 text-gray-500 whitespace-nowrap">${window.formatDate(g.created_at)} ${new Date(g.created_at).toLocaleTimeString("id-ID")}</td>
        </tr>
      `).join("")}catch(c){console.error(c)}}async function i(){try{let c=await k("/api/sync/metrics");if(!c.ok)return;let m=document.getElementById("metricsTableBody"),p=c.data||[];if(p.length===0){m.innerHTML='<tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">No metrics found in last 24h</td></tr>';return}m.innerHTML=p.map(o=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2 font-medium text-gray-800">${o.module}</td>
          <td class="px-4 py-2 text-gray-600">${o.total_events}</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(o.avg_webhook_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(o.avg_d1_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(o.avg_queue_wait_ms||0)} ms</td>
        </tr>
      `).join("")}catch(c){console.error(c)}}}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let e=Math.floor(Number(t)-25569);return new Date(e*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let e=t.split(/[\/\-]/);return`${e[2]}-${e[1]}-${e[0]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let l=e.split("-");if(l.length===3&&l[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],s=parseInt(l[2],10),d=a[parseInt(l[1],10)-1];return`${s} ${d} ${l[0]}`}return e};function Q(t){return async e=>{if(!Be()){Se("/login");return}return t(e)}}var qe=null;function Oa(){qe&&clearInterval(qe);let t=()=>{let e=new Date,l=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),s=document.getElementById("header-clock-time"),d=document.getElementById("header-clock-date");s&&(s.textContent=l),d&&(d.textContent=a)};t(),qe=setInterval(t,1e3)}async function Ra(){try{let t=await k("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},l=(a,s)=>{let d=document.getElementById(a);d&&(d.textContent=s>0?s:"",d.style.display=s>0?"inline-flex":"none")};l("badge-issues",e.issues?.current||0),l("badge-contracts",e.expiring30?.current||0),l("badge-oo1",e.one_on_one?.current||0),l("badge-schedule",e.schedule?.current||0),l("badge-supply",e.supply?.current||0)}catch{}}var Ie=[];async function Ka(){try{let t=await k("/api/dashboard/notifications");if(!t.ok)return;Ie=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Ie.length>0?"block":"none",e.textContent=Ie.length)}catch{}}function Ha(){if(!Ie.length){se({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,l)=>l()});return}let t=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${Ie.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;se({title:`Notifikasi (${Ie.length})`,content:t,confirmText:"Tutup",onConfirm:(e,l)=>l()})}function sa(){let t=he(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
          </div>`:""}
        </nav>

        <div class="sidebar-footer">
          <div class="sidebar-user">
            <div class="sidebar-avatar">BA</div>
            <div class="sidebar-user-info">
              <div class="sidebar-user-name">Berlin Ariansyah</div>
              <div class="sidebar-user-role">Administrator</div>
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
                <span class="topbar-greeting-time">${(()=>{let c=new Date().getHours();return c>=4&&c<11?"Selamat Pagi":c>=11&&c<15?"Selamat Siang":c>=15&&c<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">Berlin Ariansyah</span> \u{1F44B}
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
              <img src="https://ui-avatars.com/api/?name=Berlin+Ariansyah&background=2563EB&color=fff&bold=true" class="topbar-avatar" alt="Avatar" />
              <div class="topbar-user-text">
                <span class="topbar-user-name">Berlin Ariansyah</span>
                <span class="topbar-user-role-mini">Administrator</span>
              </div>
              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" style="margin-left:4px;color:var(--gray-400)"><polyline points="6 9 12 15 18 9"/></svg>
            </a>
          </div>
        </header>

        <main id="main-content" class="main-content"></main>
      </div>
    </div>
  `;let l=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),s=document.getElementById("topbar-menu-btn"),d=document.getElementById("sidebar-close"),r=()=>{l.classList.add("open"),a.classList.add("show")},n=()=>{l.classList.remove("open"),a.classList.remove("show")};s?.addEventListener("click",r),d?.addEventListener("click",n),a?.addEventListener("click",n),document.querySelectorAll(".nav-item").forEach(c=>c.addEventListener("click",n));function i(){let c=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(o=>{let g=o.dataset.route;o.classList.toggle("active",c===g||g!=="/dashboard"&&c.startsWith(g))});let m=document.getElementById("topbar-title"),p=document.querySelector(".nav-item.active .nav-label");m&&p&&(m.textContent=p.textContent)}window.addEventListener("hashchange",i),i(),Oa(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await k("/api/auth/logout",{method:"POST"}),Le(),qe&&clearInterval(qe),Se("/login")}),Ra(),Ka(),document.getElementById("btn-notif")?.addEventListener("click",c=>{c.preventDefault(),Ha()})}async function ja(){J("/login",({main:e})=>Pt(e)),J("/dashboard",Q(({main:e})=>Tt(e))),J("/calendar",Q(({main:e})=>Vt(e))),J("/employees",Q(({main:e,params:l})=>Bt(e,l))),J("/contracts",Q(({main:e,params:l})=>Nt(e,l))),J("/sp",Q(({main:e})=>na(e))),J("/mutasi",Q(({main:e})=>ra(e))),J("/sync-dashboard",Q(({main:e})=>la(e))),J("/timeline",Q(({main:e,params:l})=>At(e,l))),J("/issues",Q(({main:e,params:l})=>Ft(e,l))),J("/one-on-one",Q(({main:e,params:l})=>Mt(e,l))),J("/training",Q(({main:e})=>Ot(e))),J("/relievers",Q(({main:e,params:l})=>Kt(e,l))),J("/reports/inspection",Q(({main:e})=>Ht(e))),J("/reports/cleaning",Q(({main:e})=>jt(e))),J("/reports/fogging",Q(({main:e})=>qt(e))),J("/reports/basecamp",Q(({main:e})=>Jt(e))),J("/reports/supply",Q(({main:e})=>ht(e,"supply"))),J("/sop",Q(({main:e})=>Ut(e))),J("/checklist",Q(({main:e})=>Gt(e))),J("/forms",Q(({main:e})=>ht(e))),J("/users",Q(({main:e})=>Qt(e))),J("/branches",Q(({main:e})=>zt(e))),J("/profile",Q(({main:e})=>Yt(e))),J("/settings/import",Q(({main:e})=>ta(e)));let t=Be();if(!t&&window.location.hash!=="#/login"&&Se("/login"),t){let e=await k("/api/auth/me");e.ok?(Ne(e.data.data),sa()):(Le(),Se("/login"))}window.addEventListener("fm:login",()=>{sa(),Se("/dashboard")}),ft()}ja();
