var ra=Object.defineProperty;var Je=(t,e)=>()=>(t&&(e=t(t=0)),e);var Ge=(t,e)=>{for(var r in e)ra(t,r,{get:e[r],enumerable:!0})};var de={};Ge(de,{API:()=>gt,CLIENT_SIDE_MAX_ROWS:()=>_e,IS_DEVELOPMENT:()=>Qe,apiFetch:()=>f,clearToken:()=>Ce,getToken:()=>xe,getUser:()=>le,setToken:()=>Ve,setUser:()=>Te});function xe(){return localStorage.getItem("fm_token")}function Ve(t){localStorage.setItem("fm_token",t)}function Ce(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function le(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Te(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function f(t,e={}){let r=xe(),a={"Content-Type":"application/json",...r?{Authorization:`Bearer ${r}`}:{},...e.headers||{}};try{let l=`cb=${Date.now()}`,o=t.includes("?")?"&":"?",i=`${gt}${t}${o}${l}`,s=await fetch(i,{...e,headers:a}),n;try{let c=await s.text();try{n=JSON.parse(c)}catch{n={error:`Server Error (${s.status}): ${c.substring(0,80)}...`}}}catch{n={error:"Gagal membaca respon dari server"}}return s.status===401&&(Ce(),window.location.hash="#/login"),{ok:s.ok,status:s.status,data:n}}catch(l){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${l.message})`}}}}var Qe,sa,gt,_e,D=Je(()=>{Qe=!1,sa="https://fm-operations-api.facilitycare-audydental.workers.dev",gt=sa,_e=1e4});var yt={};Ge(yt,{confirmDialog:()=>ze,createModal:()=>ae});function ae({title:t,content:e,onConfirm:r,onCancel:a,confirmText:l="Simpan",cancelText:o="Batal",size:i="md",confirmClass:s="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},c=document.createElement("div");c.className="modal-overlay",c.innerHTML=`
    <div class="modal" style="max-width:${n[i]||n.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${o}</button>
        ${r?`<button class="btn ${s} modal-confirm">${l}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&c.querySelector(".modal-body").appendChild(e);let m=()=>{c.classList.remove("show"),setTimeout(()=>c.remove(),250)};return c.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),m()}),c.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),m()}),r&&c.querySelector(".modal-confirm").addEventListener("click",()=>r(c,m)),c.addEventListener("click",d=>{d.target===c&&(a&&a(),m())}),document.body.appendChild(c),requestAnimationFrame(()=>c.classList.add("show")),{overlay:c,close:m}}function ze(t,e,r="Konfirmasi"){return ae({title:r,content:`<p>${t}</p>`,onConfirm:(a,l)=>{e(),l()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var Ee=Je(()=>{});var V={};Ge(V,{downloadExcel:()=>E,parseExcel:()=>Xe,renderExcelButtons:()=>Ze});function Xe(t){return new Promise((e,r)=>{let a=new FileReader;a.onload=l=>{try{let o=new Uint8Array(l.target.result),i=XLSX.read(o,{type:"array"}),s=i.SheetNames[0],n=i.Sheets[s];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${i.SheetNames.join(", ")}`),console.log(`Sheet Used: ${s}`);let c=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),m=c.e.r-c.s.r+1,d=c.e.c-c.s.c+1;console.log(`Total Rows (including empty): ${m}`),console.log(`Total Columns: ${d}`);let p=[];for(let h=c.s.c;h<=c.e.c;++h){let b=n[XLSX.utils.encode_cell({c:h,r:c.s.r})];b&&b.v&&p.push(b.v)}console.log(`Headers Found: ${p.join(", ")}`),console.log("---------------------------");let u=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(u,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(u,"__headers",{value:p,enumerable:!1}),e(u)}catch(o){r(o)}},a.onerror=l=>r(l),a.readAsArrayBuffer(t)})}function E(t,e){try{let r=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,r,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(r){throw console.error("Error generating Excel file:",r),r}}function Ze(t){return`
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
  `}var B=Je(()=>{});D();var Ye={},Ne=null;function O(t,e){Ye[t]=e}function pe(t){window.location.hash=t}function bt(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[r,...a]=e.split("?"),l=Ye[r];if(!l){for(let[i,s]of Object.entries(Ye))if(i.endsWith("/*")&&r.startsWith(i.slice(0,-2))){l=s;break}}Ne&&(Ne(),Ne=null);let o=document.getElementById("main-content");if(o&&(o.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),l){let i=new URLSearchParams(a.join("?")),s=r.split("/").filter(Boolean),n=await l({path:r,params:i,segments:s,main:o});n&&(Ne=n)}else{let i=o||document.getElementById("app");i&&(i.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var $e;function oa(){return $e||($e=document.createElement("div"),$e.id="toast-container",document.body.appendChild($e)),$e}function ht(t,e="info",r=3500){let a=oa(),l=document.createElement("div");l.className=`toast toast-${e}`;let o={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};l.innerHTML=`<span class="toast-icon">${o[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(l),requestAnimationFrame(()=>l.classList.add("show")),setTimeout(()=>{l.classList.remove("show"),setTimeout(()=>l.remove(),350)},r)}var J=t=>ht(t,"success"),Q=t=>ht(t,"error");Ee();D();D();function ft({columns:t,data:e,onEdit:r,onDelete:a,onView:l,actions:o=[],emptyText:i="Tidak ada data",bulkSelect:s=null}){let n=document.createElement("div");if(n.className="table-wrapper",!e||e.length===0)return n.innerHTML=`<div class="empty-state"><p>${i}</p></div>`,n;let c=document.createElement("table");c.className="data-table";let m=document.createElement("thead"),d=document.createElement("tr");if(s){let u=document.createElement("th");u.style.width="40px",u.style.textAlign="center";let h=document.createElement("input");h.type="checkbox",h.id="select-all-checkbox",h.title="Pilih semua",h.addEventListener("change",()=>{e.forEach(b=>{h.checked?s.selectedIds.add(b.id):s.selectedIds.delete(b.id)}),n.querySelectorAll(".row-checkbox").forEach(b=>b.checked=h.checked),s.onToggle()}),u.appendChild(h),d.appendChild(u)}if(t.forEach(u=>{let h=document.createElement("th");h.textContent=u.label,u.width&&(h.style.width=u.width),d.appendChild(h)}),r||a||l||o.length>0){let u=document.createElement("th");u.textContent="Aksi",u.style.width="120px",d.appendChild(u)}m.appendChild(d),c.appendChild(m);let p=document.createElement("tbody");return e.forEach(u=>{let h=document.createElement("tr");if(s){let b=document.createElement("td");b.style.textAlign="center",b.style.width="40px";let g=document.createElement("input");g.type="checkbox",g.className="row-checkbox",g.checked=s.selectedIds.has(u.id),g.addEventListener("change",()=>{if(g.checked)s.selectedIds.add(u.id);else{s.selectedIds.delete(u.id);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1)}s.onToggle()}),b.appendChild(g),h.appendChild(b)}if(t.forEach(b=>{let g=document.createElement("td");if(b.render){let y=b.render(u[b.key],u);y instanceof HTMLElement?g.appendChild(y):g.innerHTML=y||""}else g.textContent=u[b.key]!==null&&u[b.key]!==void 0&&u[b.key]!==""?u[b.key]:"";b.nowrap&&(g.style.whiteSpace="nowrap"),h.appendChild(g)}),r||a||l||o.length>0){let b=document.createElement("td");b.className="actions-cell";let g=document.createElement("div");if(g.className="btn-group",l){let y=document.createElement("button");y.className="btn btn-xs btn-ghost",y.innerHTML="\u{1F441}",y.title="Lihat",y.addEventListener("click",()=>l(u)),g.appendChild(y)}if(r){let y=document.createElement("button");y.className="btn btn-xs btn-secondary",y.innerHTML="\u270F\uFE0F",y.title="Edit",y.addEventListener("click",()=>r(u)),g.appendChild(y)}o.forEach(y=>{let k=document.createElement("button");k.className=`btn btn-xs ${y.class||"btn-ghost"}`,k.innerHTML=y.icon||y.label,k.title=y.label,k.addEventListener("click",()=>y.handler(u)),g.appendChild(k)}),b.appendChild(g),h.appendChild(b)}p.appendChild(h)}),c.appendChild(p),n.appendChild(c),n}function vt({page:t,pages:e,total:r,limit:a,onPage:l}){if(e<=1)return null;let o=document.createElement("div");o.className="pagination";let i=document.createElement("span");i.className="pagination-info",i.textContent=`Total: ${r} data`,o.appendChild(i);let s=document.createElement("div");s.className="pagination-btns";let n=(d,p,u=!1,h=!1)=>{let b=document.createElement("button");b.className=`btn btn-sm ${h?"btn-primary":"btn-ghost"} pagination-btn`,b.textContent=d,b.disabled=u,b.addEventListener("click",()=>l(p)),s.appendChild(b)};n("\xAB",1,t===1),n("\u2039",t-1,t===1);let c=Math.max(1,t-2),m=Math.min(e,t+2);for(let d=c;d<=m;d++)n(d,d,!1,d===t);return n("\u203A",t+1,t===e),n("\xBB",e,t===e),o.appendChild(s),o}Ee();function We(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${We(e.fields)}</div>`;let r=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",l="";switch(e.type){case"textarea":l=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${r} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let i=(e.options||[]).map(d=>{let p=typeof d=="object"?d.value:d,u=typeof d=="object"?d.label:d,h=e.value==p?"selected":"";return`<option value="${p}" ${h}>${u}</option>`}).join("");l=`<select name="${e.name}" class="form-control" ${r}><option value="">-- Pilih ${e.label||""} --</option>${i}</select>`;break;case"combobox":let s=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,n=(e.options||[]).map(d=>{let p=typeof d=="object"?d.value:d,u=typeof d=="object"?d.label||d.value||"":d||"";return(u==="undefined"||u==="[object Object]"||u==="null")&&(u=""),u?`<option value="${u}"></option>`:""}).join(""),c=e.value||"";if(e.value){let d=(e.options||[]).find(p=>(typeof p=="object"?p.value:p)==e.value);if(d){let p=typeof d=="object"?d.label||d.value||"":d||"";p&&p!=="undefined"&&p!=="[object Object]"&&p!=="null"&&(c=p)}}l=`
          <input type="text" name="${e.name}" list="${s}" class="form-control" value="${c}" placeholder="Pilih atau ketik baru..." ${r} autocomplete="off">
          <datalist id="${s}">${n}</datalist>
        `;break;case"checkbox":l=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let m=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";l=`<input type="date" name="${e.name}" class="form-control" value="${m}" ${r}>`;break;case"number":l=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${r}>`;break;case"email":l=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${r}>`;break;case"url":l=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${r}>`;break;default:l=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${r} autocomplete="off">`}let o=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${l}${o}</div>`}).join("")}function kt(t){let e={},r=new FormData(t);for(let[a,l]of r.entries())e[a]=l===""?null:l;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function St(t,e){e&&Object.entries(e).forEach(([r,a])=>{let l=t.querySelector(`[name="${r}"]`);l&&(l.hasAttribute("list")||(l.type==="checkbox"?l.checked=!!a:l.type==="date"&&a&&window.parseFlexibleDate?l.value=window.parseFlexibleDate(a):l.value=a??""))})}B();function $({container:t,title:e,icon:r,apiPath:a,columns:l,formFields:o,filterFields:i,defaultFilters:s={},itemLabel:n="Data",canCreate:c=!0,canEdit:m=!0,canDelete:d=!0,onBeforeSubmit:p,onAfterLoad:u,onDataLoaded:h,extraActions:b=[],initialSearch:g="",exportOptions:y=null,bulkDelete:k=!1,paginationMode:S="server"}){let v=1,w={...s};g&&(w.search=g);let T=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${r} ${e}</h1>
      <div class="page-actions">
        ${c?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
      </div>
    </div>

    ${k?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${y?Ze(y.moduleName):""}

    ${i&&i.length>0?`
    <div class="filter-bar card">
      <div class="filter-bar-inner">
        ${i.map(_=>_.type==="search"?`<div class="filter-search"><input type="search" class="form-control" placeholder="${_.placeholder||"Cari..."}" id="filter-search" value="${w.search||""}"></div>`:_.type==="select"||_.type==="combobox"?`<select class="form-control filter-select" name="${_.name}" id="filter-${_.name}"><option value="">-- ${_.label} --</option>${(_.options||[]).map(x=>`<option value="${typeof x=="object"?x.value:x}" ${w[_.name]===(typeof x=="object"?x.value:x)?"selected":""}>${typeof x=="object"?x.label:x}</option>`).join("")}</select>`:"").join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
      </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function L(){if(!document.getElementById("bulk-toolbar"))return;let x=document.getElementById("bulk-count"),A=document.getElementById("btn-bulk-delete"),j=document.getElementById("btn-bulk-cancel");x.textContent=`${T.size} item dipilih`,T.size>0?(A.disabled=!1,j.disabled=!1):(A.disabled=!0,j.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{T.clear(),document.querySelectorAll(".row-checkbox").forEach(x=>x.checked=!1);let _=document.getElementById("select-all-checkbox");_&&(_.checked=!1),L()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(T.size===0)return;let _=[...T],x=document.createElement("div");x.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",x.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${_.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${_.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(x),x.querySelector("#bulk-cancel-btn").addEventListener("click",()=>x.remove()),x.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let A=x.querySelector("#bulk-confirm-btn");A.disabled=!0,A.textContent="Menghapus...";let j=await f(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:_})});x.remove(),j.ok?(J(`${_.length} ${n} berhasil dihapus.`),T.clear(),L(),F()):Q(j.data?.error||"Gagal menghapus data.")})});let C=document.getElementById("filter-search"),P;if(C?.addEventListener("input",_=>{clearTimeout(P),P=setTimeout(()=>{w.search=_.target.value,v=1,F()},400)}),i?.forEach(_=>{(_.type==="select"||_.type==="combobox")&&document.getElementById(`filter-${_.name}`)?.addEventListener("change",x=>{w[_.name]=x.target.value,v=1,F()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{w={},C&&(C.value=""),i?.forEach(_=>{let x=document.getElementById(`filter-${_.name}`);x&&(x.value="")}),v=1,F()}),document.getElementById("btn-create")?.addEventListener("click",()=>Se(null)),y){document.getElementById(`btn-export-${y.moduleName}`)?.addEventListener("click",async x=>{let A=x.target,j=A.innerHTML;A.innerHTML="\u23F3 Loading...",A.disabled=!0;try{await y.onExport()}catch{Q("Gagal export data")}finally{A.innerHTML=j,A.disabled=!1}}),document.getElementById(`btn-template-${y.moduleName}`)?.addEventListener("click",()=>{y.onTemplate()});let _=document.getElementById(`input-import-${y.moduleName}`);_?.addEventListener("change",async x=>{let A=x.target.files[0];if(!A)return;let j=document.getElementById(`label-import-${y.moduleName}`),Y=j?j.querySelector(".import-text"):null,te=Y?Y.innerText:"";Y&&(Y.innerText="\u231B Memproses..."),j&&(j.style.pointerEvents="none"),_.disabled=!0;try{let M=await Xe(A);if(M.length===0)throw new Error("File kosong atau format salah");await y.onImport(M),J("Import berhasil!"),F()}catch(M){Q(M.message||"Gagal import data")}finally{Y&&(Y.innerText=te),j&&(j.style.pointerEvents="auto"),_.disabled=!1,_.value=""}})}async function F(){T.clear(),L();let _=document.getElementById("table-container");if(!_)return;_.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let x=S==="client",A=x?1:v,j=x?_e:20,Y=new URLSearchParams({page:A,limit:j,...Object.fromEntries(Object.entries(w).filter(([,G])=>G))}),te=await f(`${a}?${Y}`);if(!te.ok){_.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${te.data?.error||"Error"}</p></div>`;return}let M=te.data?.data||te.data||[],R=te.data?.pagination,Ue=M.length;if(x){M=h(M);let G=M.length,W=20,ie=Math.ceil(G/W);v>ie&&ie>0&&(v=ie);let U=(v-1)*W,ce=v*W;M=M.slice(U,ce),R={page:v,limit:W,total:G,pages:ie}}!1,u&&u(M);let Le=ft({columns:l,data:M,onEdit:m?G=>Se(G):null,actions:b.map(G=>({...G,handler:W=>G.handler(W,F)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:k?{selectedIds:T,onToggle:L}:null});_.innerHTML="",_.appendChild(Le);let we=document.getElementById("pagination-container");if(we&&(we.innerHTML="",R&&R.pages>1)){let G=vt({page:R.page,pages:R.pages,total:R.total,limit:R.limit,onPage:W=>{v=W,F()}});G&&we.appendChild(G)}}function ke(_){let x=typeof o=="function"?o(_):o;return We(x)}function Se(_){let x=!!_,A=document.createElement("form");if(A.noValidate=!0,A.innerHTML=ke(_),x){let Y=typeof o=="function"?o(_):o;St(A,_)}let{close:j}=ae({title:x?`Edit ${n}`:`Tambah ${n}`,content:A,size:"lg",confirmText:x?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(Y,te)=>{if(!A.reportValidity())return;let M=Y.querySelector(".modal-confirm");M.disabled=!0,M.textContent="Menyimpan...";let R=kt(A),Ue=typeof o=="function"?o(_):o,Le=async ie=>{for(let U of ie)if(U.type==="row")await Le(U.fields);else if(U.type==="combobox"&&R[U.name]){let ce=R[U.name],Ae=(U.options||[]).find(ee=>{let re=String(typeof ee=="object"?ee.value:ee),ia=String(typeof ee=="object"?ee.label:ee);return re===ce||ia===ce});if(Ae)R[U.name]=typeof Ae=="object"?Ae.value:Ae;else if(U.createApi){let ee={};ee[U.createApi.field]=ce,U.createApi.extra&&Object.assign(ee,U.createApi.extra);let re=await f(U.createApi.path,{method:"POST",body:JSON.stringify(ee)});if(re.ok&&re.data?.id)R[U.name]=re.data.id;else if(re.ok&&!re.data?.id)R[U.name]=ce;else throw new Error(`Gagal membuat master data: ${re.data?.error||"Unknown error"}`)}}};try{await Le(Ue)}catch(ie){Q(ie.message),M.disabled=!1,M.textContent=x?"Simpan Perubahan":`Tambah ${n}`;return}p&&(R=await p(R,_));let we=x?"PUT":"POST",G=x?`${a}/${_.id}`:a,W=await f(G,{method:we,body:JSON.stringify(R)});W.ok?(J(x?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),te(),F()):(Q(W.data?.error||"Gagal menyimpan data."),M.disabled=!1,M.textContent=x?"Simpan Perubahan":`Tambah ${n}`)}})}function je(_){ze(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let x=await f(`${a}/${_.id}`,{method:"DELETE"});x.ok?(J(`${n} berhasil dihapus.`),F()):Q(x.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return F(),F}D();D();var me=null,Oe=null;async function ue(t=!1){if(me&&!t)return console.log("Employees Raw (Cache Hit)",me.slice(0,5)),me;let e=await f(`/api/employees?limit=${_e}&status=Aktif`);return me=(e.data?.data||[]).map(r=>({value:r.id,label:r.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",me.slice(0,5)),me}async function q(t=!1){let r=(await ue(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",r.slice(0,5)),r}async function I(t=!1){return Oe&&!t||(Oe=((await f("/api/branches?all=1")).data?.data||[]).map(r=>({value:r.id,label:r.full_name}))),Oe}function N(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function et(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function ge(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function tt(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function X(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}B();function at(t,e){if(t.period!=="Q3")return!1;let r=String(t.status||"").toLowerCase();if(r!=="selesai"&&r!=="completed"&&r!=="done")return!1;let a=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?a.includes("inspeksi"):e==="gcdc"?a.includes("general cleaning")||a.includes("deep cleaning"):!1}D();B();function wt(t,e){let r=String(t.status||"").toLowerCase();return e==="active"?r==="aktif":!1}D();B();function nt(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let l=new Date(a);l.setDate(a.getDate()+30);let o=new Date(t.end_date);return o.setHours(0,0,0,0),o>=a&&o<=l}return!1}D();B();function _t(t,e){let r=String(t.status||"").toLowerCase();return e==="open"?r==="open":!1}D();function xt(t,e){let r=String(t.status||"").toLowerCase();return e==="pending"?r==="pending":!1}var oe={};function De(t){if(oe[t]){try{oe[t].destroy()}catch{}delete oe[t]}}function la(){Object.keys(oe).forEach(De)}var ne=(t,e=0)=>{let r=Number(t);return isNaN(r)||t===null||t===void 0?e:r},be=(t,e="\u2014")=>{if(t==null||t==="")return e;let r=String(t).trim();return r===""||r==="[object Object]"?e:r};var ca=t=>{if(!t||typeof t!="string")return"";try{let[e,r]=t.split("-");return new Date(Number(e),Number(r)-1).toLocaleDateString("id-ID",{month:"short",year:"2-digit"})}catch{return t}};function Tt(t,e,r=900){if(!t)return;let a=Math.max(0,Math.round(ne(e)));if(a===0){t.textContent="0";return}let l=Date.now(),o=()=>{let i=Math.min((Date.now()-l)/r,1),s=1-Math.pow(1-i,3);t.textContent=Math.round(s*a).toLocaleString("id-ID"),i<1?requestAnimationFrame(o):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(o)}var da={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},pa=t=>{let e=be(t,"\u2014");return`<span class="status-pill ${da[e]||"pill-neutral"}">${e}</span>`};var z={family:"Inter",size:11},he="#94A3B8",Fe="#F1F5F9",it=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],ma=()=>window.innerWidth<768;function ot(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:ma()?"bottom":"top",labels:{font:z,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:z,titleFont:{...z,weight:"700"}}},scales:{x:{grid:{color:Fe},ticks:{font:z,color:he,maxRotation:0}},y:{grid:{color:Fe},ticks:{font:z,color:he},beginAtZero:!0}},...t}}var ua=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),ga=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function rt(t=3){return Array(t).fill(0).map((e,r)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${r<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function Z(t,e,r=8e3){try{let a=new AbortController,l=setTimeout(()=>a.abort(),r),o=await f(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(l),!o||!o.ok)return e;let i=o.data;return i?i.data!==void 0?i.data??e:i:e}catch{return e}}function ba(){["skel-donut","skel-trend","skel-insp","skel-contract"].forEach(a=>{let l=document.getElementById(a);l&&(l.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract"].forEach(a=>{let l=document.getElementById(a);if(l&&l.style.display==="none"){l.style.display="block";let o=l.parentElement;if(o&&!o.querySelector(".chart-empty")){let i=document.createElement("div");i.className="chart-empty",i.textContent="Belum ada data",l.style.display="none",o.appendChild(i)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Et({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Dt({}),["table-contracts","table-issues"].forEach(a=>{let l=document.getElementById(a);l&&l.querySelector(".skeleton")&&(l.innerHTML='<div class="chart-empty">Belum ada data</div>')});let r=document.getElementById("activity-log");r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function $t(t){la(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">


      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${ua()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${ga()}</div>

      <!-- Charts Row -->
      <div class="charts-row">
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan per Kategori</div>
          </div>
          <div style="display:flex; gap:20px; align-items:center; height:140px">
            <div class="chart-canvas-wrap" style="flex:1;height:100%;position:relative">
              <div id="skel-donut" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
              <canvas id="chart-donut" style="display:none"></canvas>
            </div>
            <div id="donut-legend" class="donut-legend" style="width:110px"></div>
          </div>
          <div style="text-align:center; font-size:0.75rem; color:var(--text-3); margin-top:16px">
            Periode: 22 Juni - 22 Juli 2026
          </div>
        </div>
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Trend Permasalahan 12 Bulan</div>
            <div style="display:flex;align-items:center;gap:16px;font-size:0.75rem;font-weight:600;color:var(--text-2)">
               <div style="display:flex;align-items:center;gap:6px"><div style="width:16px;height:8px;border:2px solid #EF4444;border-radius:2px"></div> Open</div>
               <div style="display:flex;align-items:center;gap:6px"><div style="width:16px;height:8px;border:2px solid #10B981;border-radius:2px"></div> Closed</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:140px;position:relative">
            <div id="skel-trend" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-trend" style="display:none"></canvas>
          </div>
        </div>
        <div class="chart-card" style="grid-column: 1 / -1;">
          <div class="chart-card-header" style="align-items:flex-start">
            <div>
              <a href="#/reports/inspection" class="chart-card-title" style="text-decoration:none; display:inline-block">Rata-rata Skor Inspeksi per Cabang <span style="font-size:0.8rem; color:var(--primary); font-weight:600; margin-left:8px">Lihat Laporan &rarr;</span></a>
              <div class="chart-card-subtitle" style="font-size:0.65rem">Skor rata-rata SCM & Cleaning</div>
            </div>
            <select id="insp-month-filter" class="btn-ghost" style="padding:4px;font-size:0.7rem;border:1px solid var(--border);border-radius:4px;cursor:pointer"><option value="">Bulan Ini</option></select>
          </div>
            <div class="chart-canvas-wrap" style="height:220px;position:relative;margin-top:10px;padding: 10px 0;">
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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${rt(3)}</div>
        </div>
        <!-- KPI Kebersihan -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">KPI Kebersihan</div>
          </div>
          <div id="widget-kpi-kebersihan" style="margin-top:0px">${rt(4)}</div>
        </div>
        <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${rt(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>st(t)),t._skelTimeout=setTimeout(()=>ba(),5e3),await st(t),t._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?st(t):clearInterval(t._dashRefresh)},6e4)}async function st(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,r,a,l,o,i,s,n,c,m,d]=await Promise.all([Z("/api/dashboard/kpi",{},8e3),Z("/api/dashboard/issues-trend",{},8e3),Z("/api/dashboard/issues-summary",{},8e3),Z("/api/dashboard/inspection-bar",{},8e3),Z("/api/dashboard/stats",{},8e3),Z("/api/dashboard/calendar",[],8e3),Z("/api/schedule?limit=10000",{data:[]},8e3),Z("/api/employees?limit=10000",{data:[]},8e3),Z("/api/contracts?limit=10000",{data:[]},8e3),Z("/api/issues?limit=10000",{data:[]},8e3),Z("/api/one_on_one?limit=10000",{data:[]},8e3)]);if(e){let u=Array.isArray(s?.data)?s.data:Array.isArray(s)?s:[],h=Array.isArray(n?.data)?n.data:Array.isArray(n)?n:[],b=Array.isArray(c?.data)?c.data:Array.isArray(c)?c:[],g=Array.isArray(m?.data)?m.data:Array.isArray(m)?m:[],y=Array.isArray(d?.data)?d.data:Array.isArray(d)?d:[];e.employees&&(e.employees.current=h.filter(k=>wt(k,"active")).length),e.contracts&&(e.contracts.current=b.filter(k=>nt(k,"active")).length),e.expiring30&&(e.expiring30={current:b.filter(k=>nt(k,"expiring30")).length}),e.issues&&(e.issues.current=g.filter(k=>_t(k,"open")).length),e.one_on_one&&(e.one_on_one.current=y.filter(k=>xt(k,"pending")).length),e.inspection_month&&(e.inspection_month.current=u.filter(k=>at(k,"inspeksi")).length),e.cleaning_month&&(e.cleaning_month.current=u.filter(k=>at(k,"gcdc")).length)}try{Et(e)}catch(u){console.warn("KPI render:",u)}try{Dt(e)}catch(u){console.warn("MiniStats render:",u)}try{ha(Array.isArray(a?.by_category)?a.by_category:[])}catch(u){console.warn("Donut render:",u),se("skel-donut","chart-donut")}try{ya(r)}catch(u){console.warn("Trend render:",u),se("skel-trend","chart-trend")}try{Ct(l)}catch(u){console.warn("InspBar render:",u),se("skel-insp","chart-insp")}try{let u=Array.isArray(o)?o:Array.isArray(o?.recent_issues)?o.recent_issues:[];va(u)}catch(u){console.warn("IssuesTable render:",u)}try{let u=Array.isArray(o?.expiring_contracts)?o.expiring_contracts:[];fa()}catch(u){console.warn("ContractsTable render:",u)}try{ka(Array.isArray(i)?i:[])}catch(u){console.warn("Agenda render:",u)}try{Sa(e)}catch(u){console.warn("KPI Kebersihan render:",u)}try{wa()}catch(u){console.warn("Quick Actions render:",u)}let p=document.getElementById("insp-month-filter");if(p){let u=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];(async()=>{try{let h=await f("/api/dashboard/inspection-months"),b=h&&h.ok?h.data:null,g=Array.isArray(b)?b:b&&Array.isArray(b.data)?b.data:[];p.innerHTML='<option value="">6 Bulan Terakhir</option>',g.forEach(y=>{let[k,S]=y.split("-"),v=document.createElement("option");v.value=y,v.textContent=u[parseInt(S,10)-1]+" "+k,p.appendChild(v)})}catch{p.innerHTML='<option value="">6 Bulan Terakhir</option>';let b=new Date().getFullYear();u.forEach((g,y)=>{let k=`${b}-${String(y+1).padStart(2,"0")}`,S=document.createElement("option");S.value=k,S.textContent=g+" "+b,p.appendChild(S)})}})(),p.addEventListener("change",async h=>{let b=h.target.value,g=document.getElementById("skel-insp"),y=document.getElementById("chart-insp");g&&(g.style.display="block"),y&&(y.style.display="none");let k=b?`?month=${b}`:"",S=await f("/api/dashboard/inspection-bar"+k);if(S&&S.ok){let v=S.data?.data!==void 0?S.data.data:S.data;Ct(v)}else se("skel-insp","chart-insp")})}}function Et(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let r=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=r.map(a=>{let l=ne(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${l}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${l}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{Tt(a,parseInt(a.dataset.target)||0)})}function Dt(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let r=[{icon:"\u{1F4C5}",label:"Jadwal",val:t.schedule?.current,href:"#/timeline",color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F504}",label:"Reliefer",val:t.reliever_total?.current,href:"#/relievers?dash_filter=reliever",color:"mini-teal"},{icon:"\u{1F50D}",label:"Inspeksi",val:t.inspection_month?.current,href:"#/timeline?dash_filter=inspeksi",color:"mini-blue"},{icon:"\u{1F9F9}",label:"GCDC",val:t.cleaning_month?.current,href:"#/timeline?dash_filter=gcdc",color:"mini-green"},{icon:"\u{1F4A8}",label:"Fogging",val:t.fogging_month?.current,href:"#/reports/fogging",color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=r.map(a=>`
    <a href="${a.href}" class="mini-stat ${a.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${a.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${ne(a.val)}">0</div>
        <div class="mini-stat-text">${a.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(a=>Tt(a,parseInt(a.dataset.target)||0,700))}function ha(t){se("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),r=document.getElementById("donut-legend");if(!e||!r)return;De("donut");let a=(t||[]).filter(n=>ne(n.count)>0);if(!a.length){lt(e,"Belum ada data permasalahan");return}let l=a.map(n=>`${be(n.category,"Lainnya")}`),o=a.map(n=>ne(n.count)),i=o.reduce((n,c)=>n+c,0);r.innerHTML=a.map((n,c)=>{let m=it[c%it.length],d=i>0?Math.round(n.count/i*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${m}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${d}%)</span></div>
          <div class="donut-legend-label">${l[c]}</div>
        </div>
      </div>
    `}).join("");let s={id:"centerText",beforeDraw:function(n){let c=n.width,m=n.height,d=n.ctx;d.restore();let p=(m/80).toFixed(2);d.font="bold "+p+"em Inter",d.textBaseline="middle",d.fillStyle="#1E293B";let u=i.toString(),h=Math.round((c-d.measureText(u).width)/2),b=m/2;d.fillText(u,h,b-10),d.font="600 "+(p*.35).toFixed(2)+"em Inter",d.fillStyle="#64748B";let g="Total",y=Math.round((c-d.measureText(g).width)/2);d.fillText(g,y,b+15),d.save()}};oe.donut=new Chart(e,{type:"doughnut",data:{labels:l,datasets:[{data:o,backgroundColor:it,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:z,titleFont:{...z,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutout:"75%"},plugins:[s]})}function ya(t){se("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;De("trend"),t=t||{};let r=(t.labels||[]).map(ca),a=(t.open||[]).map(o=>ne(o)),l=(t.closed||[]).map(o=>ne(o));if(!r.length){lt(e,"Belum ada data trend");return}oe.trend=new Chart(e,{type:"line",data:{labels:r,datasets:[{label:"Open",data:a,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:l,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#10B981",borderWidth:2}]},options:ot({plugins:{legend:{display:!1}}})})}function Ct(t){se("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;let r=e.parentElement;if(r){let i=r.querySelector(".chart-empty");i&&i.remove()}De("inspBar"),e.style.display="block",t=t||{};let a=t.labels||[],l=(t.fc||[]).map(i=>ne(i)),o=(t.spv||[]).map(i=>ne(i));if(!a.length){lt(e,"Belum ada data inspeksi");return}oe.inspBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Skor FC",data:l,backgroundColor:"rgba(37,99,235,0.85)",borderRadius:6,borderSkipped:!1,barPercentage:.5,categoryPercentage:.7},{label:"Skor SPV",data:o,backgroundColor:"rgba(16,185,129,0.85)",borderRadius:6,borderSkipped:!1,barPercentage:.5,categoryPercentage:.7}]},options:ot({maintainAspectRatio:!1,plugins:{legend:{position:"top",labels:{font:{...z,size:13,weight:"bold"},padding:20,usePointStyle:!0,pointStyle:"circle"}},tooltip:{padding:12,cornerRadius:8,backgroundColor:"rgba(15,23,42,0.9)",titleFont:{size:14,family:"'Inter', sans-serif"},bodyFont:{size:13,family:"'Inter', sans-serif"}}},scales:{x:{grid:{display:!1},ticks:{font:z,color:he,maxRotation:45,minRotation:30}},y:{grid:{color:Fe,borderDash:[4,4]},ticks:{font:z,color:he,stepSize:20},min:0,max:100}}})})}function fa(){se("skel-contract-mini","chart-contract-mini");let t=document.getElementById("chart-contract-mini");if(!t)return;De("contractMiniBar");let e=["Jun","Jul","Agu","Sep","Okt","Nov","Des"],r=[12,18,9,24,15,30,42],l=t.getContext("2d").createLinearGradient(0,0,0,200);l.addColorStop(0,"#60A5FA"),l.addColorStop(1,"#2563EB"),oe.contractMiniBar=new Chart(t,{type:"bar",data:{labels:e,datasets:[{label:"Kontrak Habis",data:r,backgroundColor:l,borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:ot({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:z,color:he,maxRotation:0}},y:{grid:{color:Fe,borderDash:[4,4],drawBorder:!1},ticks:{font:z,color:he,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function va(t){let e=document.getElementById("table-issues");if(!e)return;let r=(t||[]).slice(0,8);if(!r.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${r.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${pa(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${be(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${be(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function ka(t){let e=document.getElementById("widget-agenda");if(!e)return;let r=new Date,a=`${r.getFullYear()}-${String(r.getMonth()+1).padStart(2,"0")}-${String(r.getDate()).padStart(2,"0")}`,o=(t||[]).filter(i=>(i.event_date||"").startsWith(a)).slice(0,10);if(!o.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada agenda hari ini</div>';return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${o.map(i=>{let s="#3B82F6",n="#EFF6FF",c="Agenda",m=(i.title||"").toLowerCase();return m.includes("inspeksi")?(s="#10B981",n="#ECFDF5",c="Inspeksi"):m.includes("cleaning")||m.includes("gcdc")?(s="#3B82F6",n="#EFF6FF",c="Cleaning"):m.includes("reliefer")?(s="#F59E0B",n="#FFFBEB",c="Reliefer"):m.includes("fogging")&&(s="#8B5CF6",n="#F5F3FF",c="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(i.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${s};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${be(i.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${be(i.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${s}">${c}</div>
        </div>
      `}).join("")}
    </div>
  `}function Sa(t){let e=document.getElementById("widget-kpi-kebersihan");if(!e)return;let r=[{label:"Kebersihan Area",val:"97%",target:"Target 95%",icon:"\u{1F9F9}",bg:"#ECFDF5",color:"#10B981"},{label:"Penyelesaian Complaint",val:"100%",target:"Target 100%",icon:"\u23F1\uFE0F",bg:"#ECFDF5",color:"#10B981"},{label:"Kepatuhan Jadwal Cleaning",val:"99%",target:"Target 100%",icon:"\u23F1\uFE0F",bg:"#EFF6FF",color:"#3B82F6"},{label:"Kepatuhan GCDC",val:"100%",target:"Target 100%",icon:"\u{1F9F9}",bg:"#EFF6FF",color:"#3B82F6"},{label:"Complaint Cleaning (\u226410)",val:"4",target:"Target \u226410",icon:"\u{1F4DD}",bg:"#F5F3FF",color:"#8B5CF6"},{label:"Pelaksanaan Fogging",val:"100%",target:"Target 100%",icon:"\u{1F4A8}",bg:"#F5F3FF",color:"#8B5CF6"}];e.innerHTML=`
    <div style="display:grid;grid-template-columns:1fr;gap:12px;padding-bottom:12px;height:160px;overflow-y:auto;padding-right:8px;">
      ${r.map(a=>{let l=a.val.includes("%")?parseInt(a.val):Math.min(100,parseInt(a.val)*10);return`
        <div class="prog-item">
          <div class="prog-header">
            <div class="prog-title">
              <div class="prog-title-icon" style="background:${a.bg};color:${a.color}">${a.icon}</div>
              ${a.label}
            </div>
            <div class="prog-val">${a.val}</div>
          </div>
          <span class="prog-target">${a.target}</span>
          <div class="prog-bar-bg">
            <div class="prog-bar-fill" style="width:${l}%;background:${a.color}"></div>
          </div>
        </div>
      `}).join("")}
    </div>
  `}function wa(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(r=>`
    <a href="${r.href}" class="action-btn">
      <div class="action-icon" style="background:${r.bg}">${r.icon}</div>
      ${r.label}
    </a>
  `).join("")}function se(t,e){let r=document.getElementById(t),a=document.getElementById(e);r&&(r.style.display="none",r.style.position=""),a&&(a.style.display="block")}function lt(t,e="Belum ada data"){if(!t)return;t.style.display="none";let r=t.parentElement;if(!r)return;if(!r.querySelector(".chart-empty")){let l=document.createElement("div");l.className="chart-empty",l.textContent=e,r.appendChild(l)}}D();async function It(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),r=document.getElementById("login-error"),a=document.getElementById("login-btn"),l=document.getElementById("toggle-password"),o=document.getElementById("login-password");l?.addEventListener("click",()=>{let i=o.type==="text";o.type=i?"password":"text",l.style.color=i?"":"var(--primary)"}),e?.addEventListener("submit",async i=>{i.preventDefault(),r.style.display="none";let s=e.username.value.trim(),n=e.password.value;if(!s||!n){r.textContent="Username dan password wajib diisi.",r.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let c=await f("/api/auth/login",{method:"POST",body:JSON.stringify({username:s,password:n})});c.ok&&c.data.success?(Ve(c.data.data.token),Te(c.data.data.user),J("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(r.textContent=c.data.error||"Username atau password salah.",r.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{r.textContent="Gagal terhubung ke server. Periksa koneksi internet.",r.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}D();B();async function _a(){return await I()}async function Pt(t,e){let r=e?e.get("dash_filter"):null,a=await _a();$({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",defaultFilters:{status:r==="active"?"Aktif":""},onDataLoaded:o=>o,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:o=>ge(o)},{key:"phone",label:"No. HP",render:o=>o?`<a href="tel:${o}">${o}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>N(o)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:a},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:o=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:o?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:o?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:a,value:o?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:o?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:o?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:o?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let o=await f(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(o.ok){let i=o.data.data.map(s=>({"Nama Lengkap":s.full_name,Cabang:s.branch_name||"",Divisi:s.division||"","No. HP":s.phone||"","Tgl Masuk":s.join_date||"",Status:s.status||""}));E(i,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async o=>{let i=c=>{if(!c)return null;let m=String(c||"").toLowerCase(),d=a.find(p=>String(p.label||"").toLowerCase()===m);return d?d.value:null},s=o.map(c=>({full_name:String(c["Nama Lengkap"]||"").trim(),branch_id:i(String(c.Cabang||"").trim()),division:String(c.Divisi||"").trim()||"FACILITY CARE",phone:String(c["No. HP"]||"").trim(),join_date:String(c["Tgl Masuk"]||"").trim(),status:String(c.Status||"").trim(),notes:String(c.Catatan||"").trim()})).filter(c=>c.full_name),n=await f("/api/employees/import",{method:"POST",body:JSON.stringify(s)});if(!n.ok)throw new Error(n.data?.error||"Import gagal")}}})}D();B();var dt=[],Bt=[];async function xa(){dt=await I(),Bt=await ue()}var ct=async t=>{let e=[],r=1;for(;;){let l=await(await Promise.resolve().then(()=>(D(),de))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${r}`);if(!l.ok)break;let o=l.data?.data||l.data||[],i=Array.isArray(o)?o:[];if(e=e.concat(i),i.length<100||l.data?.pagination&&r>=l.data.pagination.pages)break;r++}return e};async function Me(t,e){await xa(),$({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",defaultFilters:{},onDataLoaded:a=>a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>ge(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,l)=>l.end_date&&String(l.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':et(a)},{key:"status",label:"Status",render:a=>N(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:dt},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[o,i]=await Promise.all([ct("/api/employees?status=Aktif"),ct("/api/contracts")]);if(o.length>0){let s=i.filter(d=>d.status==="Aktif"),n=new Set(s.map(d=>d.employee_id)),c=o.filter(d=>!n.has(d.id)),m=`<p style="margin-bottom:12px">Data yang terbaca: <b>${o.length}</b> Karyawan Aktif, dan <b>${s.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${c.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;c.forEach(d=>{let p=i.filter(h=>h.employee_id===d.id),u='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(p.length>0){let h=p[0];u=`Pernah ada kontrak (Status: <b style="color:#EF4444">${h.status}</b>, Selesai: ${window.formatDate(h.end_date)})`}m+=`<li style="margin-bottom:8px"><b>${d.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${d.branch_name||"-"} | ${u}</span></li>`}),m+="</ul>",Promise.resolve().then(()=>(Ee(),yt)).then(d=>d.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:m,cancelText:"Tutup"}))}}catch(o){console.error(o)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let l=document.querySelector(".page-actions");l&&l.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:Bt,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:dt,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await f(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let l=a.data.data.map(o=>({"Nama Lengkap":o.employee_name,Cabang:o.branch_name||"","Div / Bagian":o.division||"","Tanggal Mulai":o.start_date||"","Tanggal Selesai":o.end_date&&String(o.end_date).startsWith("2099")?"":o.end_date||"","Sisa Kontrak":o.end_date&&String(o.end_date).startsWith("2099")?"Tetap":o.days_remaining!==null&&o.days_remaining!==void 0?`${o.days_remaining} Hari`:"",Status:o.status||""}));E(l,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[l,o]=await Promise.all([f("/api/branches?limit=10000"),ct("/api/employees")]),i=l.data?.data||[],s=o||[];console.log(`Total employee yang berhasil dimuat dari database : ${s.length}`),s.length>0&&(console.log("Contoh 5 employee pertama:"),s.slice(0,5).forEach((g,y)=>{console.log(`${y+1}. ID: ${g.id}, Name: ${g.full_name}, Status: ${g.status}`)}));let n=g=>{if(!g)return null;let y=String(g||"").replace(/\s+/g," ").toLowerCase().trim(),k=i.find(S=>String(S.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===y||String(S.code||"").replace(/\s+/g," ").toLowerCase().trim()===y||String(S.name||"").replace(/\s+/g," ").toLowerCase().trim()===y);return k?k.id:null},c=(g,y)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${y}`),console.log(`Nama dari Excel : "${g}"`),!g)return console.log("Alasan gagal mapping : Nama kosong"),null;let k=String(g||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${k}"`),console.log(`Jumlah employee di database : ${s.length}`);let S=s.find(v=>String(v.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===k);return S?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${S.id}`),S.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},m=g=>{if(!g)return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let y=String(g).trim();if(/^\d{4,5}(\.\d+)?$/.test(y)){let S=Math.floor(Number(y));if(S>2e4&&S<99999){let v=new Date(Date.UTC(1899,11,30)+S*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);let k=y.split(/[\/\-\.]/);if(k.length===3){let[S,v,w]=k.map(T=>T.trim());if(S.length===4&&v.length<=2&&w.length<=2)return`${S}-${v.padStart(2,"0")}-${w.padStart(2,"0")}`;if(w.length===4&&v.length<=2&&S.length<=2)return`${w}-${v.padStart(2,"0")}-${S.padStart(2,"0")}`}return y},d=a.map((g,y)=>{let k=y+2,S=String(g["Nama Lengkap"]||"").trim(),v=g["Tanggal Mulai"],w=m(v);if(!w){let C=a.__worksheet,P=a.__headers||[],F=P.indexOf("Tanggal Mulai"),ke="N/A",Se="N/A",je="N/A";if(F!==-1&&C&&window.XLSX){let x=window.XLSX.utils.encode_cell({c:F,r:k-1});je=x;let A=C[x];A?(ke=A.t||"undefined",Se=A.w||"undefined"):ke="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let _="Unknown";v==null||v===""?_="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":v instanceof Date&&isNaN(v.getTime())?_="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":_="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${k}`),console.log(`Employee Name : ${S}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${F})`),console.log(`Raw Cell Value : "${v}"`),console.log(`JavaScript Type : ${typeof v}`),console.log(`SheetJS Cell Type : ${ke}`),console.log(`SheetJS Formatted Value : "${Se}"`),console.log(`Value After Trim : "${String(v||"").trim()}"`),console.log(`Value After Date Parser : "${w}"`),console.log(`Is Empty : ${!v}`),console.log(`Is Invalid Date : ${v instanceof Date?isNaN(v.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${_}`),console.log(`Workbook Sheet : ${C?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${je}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(g,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(P)),console.log(`==========================
`)}let T=c(S,k),L=null;return T?w||(L="Tanggal Mulai kosong atau tidak berformat tanggal"):L="Karyawan tidak ditemukan di Database",{isValid:!!(T&&w),invalidReason:L,rowNum:k,data:{employee_id:T,branch_id:n(String(g.Cabang||"").trim()),division:String(g["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:w,end_date:m(g["Tanggal Selesai"])||"2099-12-31",status:String(g.Status||"").trim(),_rawName:S}}}),p=[],u=[];if(d.forEach(g=>{g.isValid?p.push(g.data):u.push({rowNum:g.rowNum,name:g.data._rawName,reason:g.invalidReason})}),console.log(`Split Validation - Valid: ${p.length}, Invalid: ${u.length}`),p.length===0){let g=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${a.length}
Valid: 0
Invalid: ${u.length}

Daftar Kegagalan (Contoh):
`;u.slice(0,10).forEach(y=>{g+=`- Row ${y.rowNum} | Nama: ${y.name} | Alasan: ${y.reason}
`}),u.length>10&&(g+=`- ... dan ${u.length-10} lainnya.
`),alert(g);return}let h=await f("/api/contracts/import",{method:"POST",body:JSON.stringify(p)}),b=`IMPORT SUMMARY
======================
`;b+=`Total Baris Excel : ${a.length}
`,b+=`Baris Valid       : ${p.length}
`,b+=`Baris Invalid     : ${u.length}

`,h&&h.data&&h.data.metrics?(b+=`Berhasil INSERT   : ${h.data.metrics.inserted}
`,b+=`Berhasil UPDATE   : ${h.data.metrics.updated}
`):b+=`Berhasil diproses : ${p.length}
`,u.length>0&&(b+=`
DAFTAR DATA DILEWATI:
`,u.forEach(g=>{b+=`- Row ${g.rowNum} | ${g.name} | ${g.reason}
`})),alert(b),typeof Me=="function"&&Me()}}})}D();B();var pt=[],Ie=[];function Ca(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let r of e)if(t.some(a=>a.period===r))return r;return"Q3"}function Ta(t,e){if(t.period!=="Q3")return!1;let r=String(t.status||"").toLowerCase();if(r!=="selesai"&&r!=="completed"&&r!=="done")return!1;let a=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?a.includes("inspeksi"):e==="gcdc"?a.includes("general cleaning")||a.includes("deep cleaning"):!1}async function Lt(t,e){pt=await I();let r=await q();Ie=["Ade","Berlin"];let a=c=>c&&!Ie.find(m=>(typeof m=="object"?m.value:m)===c)?[...Ie,c]:Ie,l=await f(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),o=c=>{if(!c||c==="-"||String(c).trim()==="")return"";let m=String(c).split("-");return m.length===3&&m[0].length===4?`${m[2]}-${m[1]}-${m[0]}`:c},i=l.data?.data||[],s=Ca(i),n=e?e.get("dash_filter"):null;$({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",defaultFilters:n?{period:"Q3"}:{period:s},onDataLoaded:c=>(n&&(c=c.filter(m=>Ta(m,n))),c.sort((m,d)=>{let p=m.opening_date?new Date(m.opening_date).getTime():0;return(d.opening_date?new Date(d.opening_date).getTime():0)-p})),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:c=>tt(c)},{key:"period",label:"Periode",render:c=>X(c)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:c=>o(c)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:c=>o(c)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:c=>o(c)},{key:"status",label:"Status",render:c=>N(c)}],filterFields:[{type:"combobox",name:"branch_id",label:"Cabang",options:pt},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"combobox",name:"pic",label:"PIC",options:Ie}],formFields:c=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:pt,value:c?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"],value:c?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:c?.period},{name:"pic",label:"PIC",type:"combobox",options:a(c?.pic),value:c?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:c?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:c?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:c?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:c?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:c?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let c=await f(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(c.ok){let m=c.data.data.map(d=>({Cabang:d.branch_name||"",Kegiatan:d.activity_type||"",Periode:d.period||"",PIC:d.pic||"","Tgl Opening":d.opening_date||"","Tgl Target":d.target_date||"","Tgl Selesai":d.completion_date||"",Status:d.status||""}));E(m,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async c=>{let d=(await f("/api/branches?all=1")).data?.data||[],p=g=>{if(!g)return null;let y=String(g||"").toLowerCase(),k=d.find(S=>String(S.full_name||"").toLowerCase()===y||String(S.code||"").toLowerCase()===y||String(S.name||"").toLowerCase()===y);return k?k.id:null},u=g=>{if(g==null||g==="")return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let y=String(g).trim();if(y===""||y==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);if(/^\d{4,5}$/.test(y)){let S=Number(y);if(S>2e4&&S<99999){let v=new Date(Date.UTC(1899,11,30)+S*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}let k=y.split(/[\/\-\.]/);if(k.length===3){let[S,v,w]=k.map(T=>T.trim());if(S.length===4&&v.length<=2&&w.length<=2)return`${S}-${v.padStart(2,"0")}-${w.padStart(2,"0")}`;if(w.length===4&&v.length<=2&&S.length<=2)return`${w}-${v.padStart(2,"0")}-${S.padStart(2,"0")}`}return y},h=c.map(g=>({branch_id:p(String(g.Cabang||"").trim()),activity_type:String(g.Kegiatan||"").trim(),period:String(g.Periode||"").trim(),pic:String(g.PIC||g.Pic||"").trim(),opening_date:u(g["Tgl Opening"]||g["Tanggal Opening"]||g["Tgl Openir"]),target_date:u(g["Tgl Target"]||g["Tanggal Target"]),completion_date:u(g["Tgl Selesai"]||g["Tanggal Selesai"]),status:String(g.Status||"").trim(),notes:String(g.Catatan||g.Keterangan||"").trim()})).filter(g=>g.activity_type&&g.period),b=await f("/api/schedule/import",{method:"POST",body:JSON.stringify(h)});if(!b.ok)throw new Error(b.data?.error||"Import gagal")}}})}D();B();var mt=[],Ke=[];async function At(t,e){let r=e?e.get("dash_filter"):null;mt=await I(),Ke=await q();let a=i=>i&&!Ke.find(s=>s.value===i)?[...Ke,{value:i,label:i}]:Ke,l=new Date().getFullYear(),o=Array.from({length:5},(i,s)=>String(l-s));$({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",paginationMode:"client",defaultFilters:{status:r==="open"?"Open":""},onDataLoaded:i=>i,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:i=>`<span class="badge badge-secondary">${i}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:i=>`<span title="${i}">${i?.length>50?i.slice(0,50)+"\u2026":i}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:i=>`<span title="${i||""}">${i?.length>40?i.slice(0,40)+"\u2026":i||"-"}</span>`},{key:"status",label:"Status",render:i=>N(i)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:i=>window.formatDate(i)},{key:"day_count",label:"Hari",render:i=>i??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:mt},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:o}],formFields:i=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:i?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:mt,value:i?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:i?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:i?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:i?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(i?.employee_name),value:i?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(i?.fc_specialist),value:i?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:i?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:i?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:i?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let i=await f(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let s=i.data.data.map(n=>({Tanggal:n.report_date||"",Cabang:n.branch_name||"",Kategori:n.category||"",Sumber:n.source||"",Keluhan:n.complaint||"","Nama FC":n.employee_name||"","FC Spesialis":n.fc_specialist||"",Solusi:n.solution||"","Tgl Selesai":n.completion_date||"",Status:n.status||""}));E(s,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async i=>{let n=(await f("/api/branches?all=1")).data?.data||[],c=p=>{if(!p)return null;let u=String(p||"").toLowerCase(),h=n.find(b=>String(b.full_name||"").toLowerCase()===u||String(b.code||"").toLowerCase()===u||String(b.name||"").toLowerCase()===u);return h?h.id:null},m=i.map(p=>({branch_id:c(String(p.Cabang||"").trim()),report_date:String(p.Tanggal||"").trim(),category:String(p.Kategori||"").trim(),source:String(p.Sumber||"").trim(),complaint:String(p.Keluhan||"").trim(),employee_name:String(p["Nama FC"]||"").trim(),fc_specialist:String(p["FC Spesialis"]||"").trim(),solution:String(p.Solusi||"").trim(),completion_date:String(p["Tgl Selesai"]||"").trim(),status:String(p.Status||"").trim()})).filter(p=>p.report_date&&p.complaint&&p.category),d=await f("/api/issues/import",{method:"POST",body:JSON.stringify(m)});if(!d.ok)throw new Error(d.data?.error||"Import gagal")}}})}D();var ye=[];async function Nt(t,e){let r=e?e.get("dash_filter"):null;ye=await I();let a=await q(),l=["Ade","Berlin"],o=s=>s&&!a.find(n=>n.value===s)?[...a,{value:s,label:s}]:a,i=s=>s&&!l.find(n=>(typeof n=="object"?n.value:n)===s)?[...l,s]:l;$({container:t,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",paginationMode:"client",defaultFilters:{status:r==="pending"?"Open":""},onDataLoaded:s=>s,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:s=>`<span title="${s||""}">${s?.length>50?s.slice(0,50)+"\u2026":s||"-"}</span>`},{key:"solution",label:"Solusi",render:s=>`<span title="${s||""}">${s?.length>40?s.slice(0,40)+"\u2026":s||"-"}</span>`},{key:"status",label:"Status",render:s=>N(s)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:s=>window.formatDate(s)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:s=>s?`<a href="${s}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"combobox",name:"branch_id",label:"Cabang",options:ye},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async s=>{let n=new URLSearchParams(s||{}).toString(),c=await f(`/api/one-on-one?limit=10000&${n}`);if(c.ok){let m=c.data.data.map(p=>({Tanggal:p.meeting_date||"",Cabang:p.branch_name||"","Nama Karyawan":p.employee_name||"",PIC:p.pic||"",Masalah:p.problem||"",Solusi:p.solution||"",Status:p.status||"","Tgl Selesai":p.completion_date||"",Dokumen:p.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(B(),V));d(m,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let s=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(B(),V));n(s,"Template_Import_OneOnOne")},onImport:async s=>{let n=p=>{if(!p)return null;let u=String(p||"").toLowerCase(),h=ye.find(b=>String(b.label||"").toLowerCase()===u);return h?h.value:null},c=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let u=String(p).trim();if(/^\d{4,5}$/.test(u)){let b=Number(u);if(b>2e4&&b<99999){let g=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);let h=u.split(/[\/\-\.]/);if(h.length===3){let[b,g,y]=h.map(k=>k.trim());if(b.length===4&&g.length<=2&&y.length<=2)return`${b}-${g.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&g.length<=2&&b.length<=2)return`${y}-${g.padStart(2,"0")}-${b.padStart(2,"0")}`}return u},m=s.map(p=>({meeting_date:c(p.Tanggal),employee_name:String(p["Nama Karyawan"]||"").trim(),branch_id:n(String(p.Cabang||"").trim()),pic:String(p.PIC||"").trim(),problem:String(p.Masalah||"").trim(),solution:String(p.Solusi||"").trim(),status:String(p.Status||"").trim(),completion_date:c(p["Tgl Selesai"]),document_link:String(p.Dokumen||"").trim()})).filter(p=>p.meeting_date&&p.employee_name&&p.branch_id),d=await f("/api/one-on-one/import",{method:"POST",body:JSON.stringify(m)});if(!d.ok)throw new Error(d.data?.error||"Import gagal")}},formFields:s=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:s?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:s?.branch_id&&!ye.find(n=>n.value==s.branch_id)?[...ye,{value:s.branch_id,label:s.branch_name||s.branch_id}]:ye,createApi:{path:"/api/branches",field:"full_name"},value:s?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"combobox",required:!0,options:o(s?.employee_name),value:s?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:i(s?.pic),createApi:{path:"/api/pic",field:"name"},value:s?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:s?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:s?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:s?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:s?.document_link}]})}D();async function Ot(t){let e=await I(),r=await q(),a=["Ade","Berlin"],l=s=>s&&!r.find(n=>n.value===s)?[...r,{value:s,label:s}]:r,o=s=>s&&!a.find(n=>(typeof n=="object"?n.value:n)===s)?[...a,s]:a,i=Array.from({length:5},(s,n)=>String(new Date().getFullYear()-n));$({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:s=>window.formatDate(s)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:s=>{try{let n=JSON.parse(s);return Array.isArray(n)?n.join(", "):s||"-"}catch{return s||"-"}}},{key:"score",label:"Nilai",render:s=>s!=null?`<strong>${s}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:s=>s?`<a href="${s}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:i}],exportOptions:{moduleName:"training",onExport:async s=>{let n=new URLSearchParams(s||{}).toString(),c=await f(`/api/training?limit=10000&${n}`);if(c.ok){let m=c.data.data.map(p=>{let u=p.participants||"";try{let h=JSON.parse(u);u=Array.isArray(h)?h.join(", "):u}catch{}return{Tanggal:p.training_date||"",Batch:p.batch||"",Materi:p.subject||"",Cabang:p.branch_name||"",Trainer:p.trainer||"",Peserta:u,Nilai:p.score!==null&&p.score!==void 0?p.score:"",Dokumen:p.document_link||""}}),{downloadExcel:d}=await Promise.resolve().then(()=>(B(),V));d(m,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let s=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(B(),V));n(s,"Template_Import_Training")},onImport:async s=>{let n=p=>{if(!p)return null;let u=String(p||"").toLowerCase(),h=e.find(b=>String(b.label||"").toLowerCase()===u);return h?h.value:null},c=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let u=String(p).trim();if(/^\d{4,5}$/.test(u)){let b=Number(u);if(b>2e4&&b<99999){let g=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);let h=u.split(/[\/\-\.]/);if(h.length===3){let[b,g,y]=h.map(k=>k.trim());if(b.length===4&&g.length<=2&&y.length<=2)return`${b}-${g.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&g.length<=2&&b.length<=2)return`${y}-${g.padStart(2,"0")}-${b.padStart(2,"0")}`}return u},m=s.map(p=>({training_date:c(p.Tanggal),batch:String(p.Batch||"").trim(),subject:String(p.Materi||"").trim(),branch_id:n(String(p.Cabang||"").trim()),trainer:String(p.Trainer||"").trim(),participants:String(p.Peserta||"").trim(),score:p.Nilai?Number(p.Nilai):null,document_link:String(p.Dokumen||"").trim()})).filter(p=>p.training_date&&p.subject&&p.branch_id),d=await f("/api/training/import",{method:"POST",body:JSON.stringify(m)});if(!d.ok)throw new Error(d.data?.error||"Import gagal")}},formFields:s=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:s?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:s?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:s?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:s?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:o(s?.trainer),value:s?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let n=JSON.parse(s?.participants);return Array.isArray(n)?n.join(", "):s?.participants||""}catch{return s?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:s?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:s?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:s?.notes}],onBeforeSubmit:async s=>(s.participants&&(s.participants=JSON.stringify(s.participants.split(",").map(n=>n.trim()).filter(Boolean))),s)})}D();B();async function Ft(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let r=await I(),a=await q(),l=e?e.get("dash_filter"):null;console.log("RAW",await ue()),console.log("OPTIONS",a);let o=n=>n&&!a.find(c=>c.value===n)?[...a,{value:n,label:n}]:a,i=["Agung Septiadi","Wasrikin","IQBAL AL BANNA","Muhammad Tri Ismandanu"],s=n=>n&&!i.includes(n)?[...i,n]:i;$({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",onDataLoaded:n=>{if(l==="reliever"){let c=new Date,m=c.getFullYear(),d=String(c.getMonth()+1).padStart(2,"0");return n.filter(p=>{if(String(p.status||"").toLowerCase()!=="done")return!1;let u=p.backup_date||"";if(u.includes("/")){let h=u.split("/");if(h.length===3&&(h[2].length===4?h[2]:`20${h[2]}`)==m&&h[1].padStart(2,"0")==d)return!0}else if(u.includes("-")&&u.startsWith(`${m}-${d}`))return!0;return!1})}return n},columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:n=>X(n)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:n=>window.formatDate(n)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:n=>n?`<span class="badge badge-info">${n}</span>`:"-"},{key:"status",label:"Status",render:n=>N(n)}],filterFields:[{type:"search",placeholder:"Cari reliefer / FC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:r,value:n?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:o(n?.original_fc_name),value:n?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:s(n?.reliever_name),value:n?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:n?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:n?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:n?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:n?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let n=await f(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let c=n.data.data.map(m=>({Cabang:m.branch_name||"","Nama Facility care":m.original_fc_name||"",Periode:m.period||"",Relifer:m.reliever_name||"","Tanggal Back Up":m.backup_date||"","Tanggal Selesai":m.completion_date||"",Keterangan:m.reason||"",Shift:m.shift||"",Status:m.status||""}));c.length===0&&c.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),E(c,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async n=>{let m=(await f("/api/branches?all=1")).data?.data||[],d=h=>{if(!h)return null;let b=String(h||"").toLowerCase(),g=m.find(y=>String(y.full_name||"").toLowerCase()===b||String(y.code||"").toLowerCase()===b||String(y.name||"").toLowerCase()===b);return g?g.id:null},p=n.map(h=>({branch_name:String(h.Cabang||"").trim(),backup_date:String(h["Tanggal Back Up"]||h["Tanggal Backup"]||"").trim(),original_fc_name:String(h["Nama Facility care"]||h["FC Digantikan"]||"").trim(),reliever_name:String(h.Relifer||h.Reliefer||"").trim(),period:String(h.Periode||"").trim(),reason:String(h.Keterangan||"").trim(),shift:String(h.Shift||"").trim(),completion_date:String(h["Tanggal Selesai"]||"").trim(),status:String(h.Status||"").trim()})).filter(h=>h.reliever_name&&h.backup_date),u=await f("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:p})});if(!u.ok)throw new Error(u.data?.error||"Import gagal")}}})}D();B();async function Mt(t){let e=await I(),r=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));$({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>X(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>N(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),o=await f(`/api/reports/inspection?limit=10000&${l}`);if(o.ok){let i=o.data.data.map(s=>({Cabang:s.branch_name||"",Periode:s.period||"",Tanggal:s.inspection_date||"","Point FC":s.fc_score!==null&&s.fc_score!==void 0?s.fc_score:"","Point SPV":s.spv_score!==null&&s.spv_score!==void 0?s.spv_score:"",Status:s.status||"","Link Dokumen":s.document_link||""}));E(i,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Inspeksi")},onImport:async a=>{let l=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),m=e.find(d=>String(d.label||"").toLowerCase()===c);return m?m.value:null},o=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let c=String(n).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let d=Number(c);if(d>2e4&&d<99999){let p=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let m=c.split(/[\/\-\.]/);if(m.length===3){let[d,p,u]=m.map(h=>h.trim());if(d.length===4&&p.length<=2&&u.length<=2)return`${d}-${p.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&p.length<=2&&d.length<=2)return`${u}-${p.padStart(2,"0")}-${d.padStart(2,"0")}`}return c},i=a.map(n=>({branch_id:l(String(n.Cabang||"").trim()),period:String(n.Periode||"").trim(),inspection_date:o(n.Tanggal),fc_score:n["Point FC"]!==void 0&&n["Point FC"]!==""?Number(n["Point FC"]):null,spv_score:n["Point SPV"]!==void 0&&n["Point SPV"]!==""?Number(n["Point SPV"]):null,status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.inspection_date),s=await f("/api/reports/inspection/import",{method:"POST",body:JSON.stringify(i)});if(!s.ok)throw new Error(s.data?.error||"Import gagal")}}})}D();B();async function Kt(t){let e=await I(),r=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));$({container:t,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>X(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>N(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),o=await f(`/api/reports/cleaning?limit=10000&${l}`);if(o.ok){let i=o.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||""}));E(i,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_GCDC")},onImport:async a=>{let l=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),m=e.find(d=>String(d.label||"").toLowerCase()===c);return m?m.value:null},o=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let c=String(n).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let d=Number(c);if(d>2e4&&d<99999){let p=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let m=c.split(/[\/\-\.]/);if(m.length===3){let[d,p,u]=m.map(h=>h.trim());if(d.length===4&&p.length<=2&&u.length<=2)return`${d}-${p.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&p.length<=2&&d.length<=2)return`${u}-${p.padStart(2,"0")}-${d.padStart(2,"0")}`}return c},i=a.map(n=>({branch_id:l(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:o(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),s=await f("/api/reports/cleaning/import",{method:"POST",body:JSON.stringify(i)});if(!s.ok)throw new Error(s.data?.error||"Import gagal")}}})}D();B();async function Rt(t){let e=await I(),r=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));$({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge badge-warning">${a}</span>`},{key:"period",label:"Periode",render:a=>X(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>N(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:a=>a||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:a?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),o=await f(`/api/reports/fogging?limit=10000&${l}`);if(o.ok){let i=o.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"Fogging",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||""}));E(i,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async a=>{let l=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),m=e.find(d=>String(d.label||"").toLowerCase()===c);return m?m.value:null},o=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let c=String(n).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let d=Number(c);if(d>2e4&&d<99999){let p=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let m=c.split(/[\/\-\.]/);if(m.length===3){let[d,p,u]=m.map(h=>h.trim());if(d.length===4&&p.length<=2&&u.length<=2)return`${d}-${p.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&p.length<=2&&d.length<=2)return`${u}-${p.padStart(2,"0")}-${d.padStart(2,"0")}`}return c},i=a.map(n=>({branch_id:l(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"Fogging").trim(),period:String(n.Periode||"").trim(),activity_date:o(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.activity_date),s=await f("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(i)});if(!s.ok)throw new Error(s.data?.error||"Import gagal")}}})}D();B();async function qt(t){let e=await I(),r=await q(),a=r,l=i=>i&&!r.find(s=>s.value===i)?[...r,{value:i,label:i}]:r,o=i=>i&&!a.find(s=>s.value===i)?[...a,{value:i,label:i}]:a;$({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:i=>window.formatDate(i)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:i=>`<span title="${i||""}">${i?.length>60?i.slice(0,60)+"\u2026":i||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:i=>window.formatDate(i)},{key:"status",label:"Status",render:i=>N(i)},{key:"notes",label:"Keterangan",render:i=>i?.length>40?i.slice(0,40)+"\u2026":i||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]}],formFields:i=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:i?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:o(i?.pic),value:i?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:i?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:i?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:i?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:i?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:i?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async i=>{let s=new URLSearchParams(i||{}).toString(),n=await f(`/api/reports/basecamp?limit=10000&${s}`);if(n.ok){let c=n.data.data.map(m=>({"Tgl Info":m.info_date||"",Cabang:m.branch_name||"",Permasalahan:m.problem||"",PIC:m.pic||"","Tgl Done":m.done_date||"",Status:m.status||"",Keterangan:m.notes||""}));E(c,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async i=>{let s=d=>{if(!d)return null;let p=String(d||"").toLowerCase(),u=e.find(h=>String(h.label||"").toLowerCase()===p);return u?u.value:null},n=d=>{if(d==null||d==="")return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let p=String(d).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let h=Number(p);if(h>2e4&&h<99999){let b=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}let u=p.split(/[\/\-\.]/);if(u.length===3){let[h,b,g]=u.map(y=>y.trim());if(h.length===4&&b.length<=2&&g.length<=2)return`${h}-${b.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&b.length<=2&&h.length<=2)return`${g}-${b.padStart(2,"0")}-${h.padStart(2,"0")}`}return p},c=i.map(d=>({info_date:n(d["Tgl Info"]||d["Tanggal Info"]),branch_id:s(String(d.Cabang||"").trim()),problem:String(d.Permasalahan||"").trim(),pic:String(d.PIC||"").trim(),done_date:n(d["Tgl Done"]||d["Tanggal Done"]),status:String(d.Status||"").trim(),notes:String(d.Keterangan||d.Catatan||"").trim()})).filter(d=>d.info_date&&d.branch_id&&d.problem),m=await f("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(c)});if(!m.ok)throw new Error(m.data?.error||"Import gagal")}}})}async function Ht(t){$({container:t,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(D(),de)),l=await a(`/api/sop?limit=10000&${r}`);if(l.ok){let o=l.data.data.map(s=>({"Nama SOP":s.name||"",Kategori:s.category||"",Dokumen:s.document_link||"",Catatan:s.notes||s.description||""})),{downloadExcel:i}=await Promise.resolve().then(()=>(B(),V));i(o,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:r}=await Promise.resolve().then(()=>(B(),V));r(e,"Template_Import_SOP")},onImport:async e=>{let r=e.map(o=>({name:String(o["Nama SOP"]||"").trim(),category:String(o.Kategori||"").trim(),document_link:String(o.Dokumen||"").trim(),description:String(o.Catatan||"").trim()})).filter(o=>o.name),{apiFetch:a}=await Promise.resolve().then(()=>(D(),de)),l=await a("/api/sop/import",{method:"POST",body:JSON.stringify(r)});if(!l.ok)throw new Error(l.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function jt(t){$({container:t,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(D(),de)),l=await a(`/api/checklist?limit=10000&${r}`);if(l.ok){let o=l.data.data.map(s=>({"Nama Checklist":s.name||"",Kategori:s.category||"",Dokumen:s.document_link||"",Deskripsi:s.description||""})),{downloadExcel:i}=await Promise.resolve().then(()=>(B(),V));i(o,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:r}=await Promise.resolve().then(()=>(B(),V));r(e,"Template_Import_Checklist")},onImport:async e=>{let r=e.map(o=>({name:String(o["Nama Checklist"]||"").trim(),category:String(o.Kategori||"").trim(),document_link:String(o.Dokumen||"").trim(),description:String(o.Deskripsi||"").trim()})).filter(o=>o.name),{apiFetch:a}=await Promise.resolve().then(()=>(D(),de)),l=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(r)});if(!l.ok)throw new Error(l.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}D();Ee();B();async function ut(t,e="forms"){if(e==="supply")return Ea(t);$a(t)}function $a(t){$({container:t,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}]})}async function Ea(t){let r=((await f("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));$({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,l)=>l.branch_name_ref||l.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>N(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let l=a?.tools_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}let o=a?.chemical_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:a?.branch_id&&!r.find(i=>i.value==a.branch_id)?[...r,{value:a.branch_id,label:a.branch_name||a.branch_id}]:r,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:l},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:o},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),o=await f(`/api/reports/supply?limit=10000&${l}`);if(o.ok){let i=o.data.data.map(s=>{let n=s.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let c=s.chemical_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}return{Waktu:s.submitted_at||"",Pengirim:s.submitter_name||"",Cabang:s.branch_name_ref||s.branch_name||"","Alat/Barang":n||"",Chemical:c||"",Catatan:s.additional_notes||"",Status:s.status||"","Diproses Oleh":s.processed_by||""}});E(i,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let o=(await f("/api/branches?all=1")).data?.data||[],i=m=>{if(!m)return null;let d=String(m||"").toLowerCase(),p=o.find(u=>String(u.full_name||"").toLowerCase()===d||String(u.code||"").toLowerCase()===d||String(u.name||"").toLowerCase()===d);return p?p.id:null},s=m=>{if(m==null||m==="")return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let d=String(m).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let u=Number(d);if(u>2e4&&u<99999){let h=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}let p=d.split(/[\/\-\.]/);if(p.length===3){let[u,h,b]=p.map(g=>g.trim());if(u.length===4&&h.length<=2&&b.length<=2)return`${u}-${h.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&h.length<=2&&u.length<=2)return`${b}-${h.padStart(2,"0")}-${u.padStart(2,"0")}`}return d},n=a.map(m=>({submitted_at:s(m.Waktu||m.Tanggal),submitter_name:String(m.Pengirim||"").trim(),branch_id:i(String(m.Cabang||"").trim()),tools_items:String(m["Alat/Barang"]||m.Alat||"").trim(),chemical_items:String(m.Chemical||"").trim(),additional_notes:String(m.Catatan||m.Keterangan||"").trim(),status:String(m.Status||"").trim(),processed_by:String(m["Diproses Oleh"]||m.PIC||"").trim()})).filter(m=>m.submitted_at&&m.submitter_name&&m.branch_id),c=await f("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,l)=>{let o=ae({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(i,s)=>{let n=i.querySelector("#supply-status").value,c=i.querySelector("#supply-processed-by").value;(await f(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:c})})).ok?(J("Status diperbarui."),s(),l()):Q("Gagal update status.")}})}}]})}D();B();async function Ut(t){let e=le();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}$({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:r=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[r]||"badge-neutral"}">${r}</span>`},{key:"is_active",label:"Status",render:r=>r?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:r=>r?new Date(r).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:r=>{let a=!!r;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:r?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:r?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:r?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:r?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?r?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let r=await f(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let a=r.data.data.map(l=>({"Nama Lengkap":l.full_name||"",Username:l.username||"",Email:l.email||"",Role:l.role||"",Status:l.is_active?"Aktif":"Nonaktif"}));E(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async r=>{let a=r.map(o=>({full_name:String(o["Nama Lengkap"]||"").trim(),username:String(o.Username||"").trim(),email:String(o.Email||"").trim(),role:String(o.Role||"").trim()||"viewer",password:String(o.Password||"").trim()})).filter(o=>o.username&&o.password&&o.email&&o.full_name),l=await f("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!l.ok)throw new Error(l.data?.error||"Import gagal")}}})}D();B();async function Jt(t){$({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await f(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)E(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let r=e.map(l=>({code:String(l["Kode Cabang"]||"").trim(),name:String(l["Nama Pendek"]||"").trim(),full_name:String(l["Nama Lengkap"]||"").trim(),city:String(l.Kota||"").trim()})).filter(l=>l.code&&l.name),a=await f("/api/branches/import",{method:"POST",body:JSON.stringify(r)});if(!a.ok)throw new Error(a.data?.error||"Import gagal")}}})}D();async function Gt(t){let e=new Date,r=[];t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F4C5} Kalender</h1>
    </div>
    <div class="card">
      <div class="card-header calendar-nav">
        <button class="btn btn-ghost btn-sm" id="cal-prev">\u2039 Prev</button>
        <span class="calendar-month-label" id="cal-month-label"></span>
        <button class="btn btn-ghost btn-sm" id="cal-next">Next \u203A</button>
        <div class="calendar-filters" style="display: flex; flex-wrap: wrap; gap: 10px;">
          <label class="filter-check"><input type="checkbox" value="schedule"        checked class="cal-filter"> Jadwal</label>
          <label class="filter-check"><input type="checkbox" value="reliever"        checked class="cal-filter"> Reliefer</label>
          <label class="filter-check"><input type="checkbox" value="cleaning"        checked class="cal-filter"> Cleaning</label>
          <label class="filter-check"><input type="checkbox" value="inspection"      checked class="cal-filter"> Inspeksi</label>
          <label class="filter-check"><input type="checkbox" value="fogging"         checked class="cal-filter"> Fogging</label>
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),l()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),l()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(o=>o.addEventListener("change",l));async function a(){try{let o=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;r=(await f(`/api/dashboard/calendar?month=${o}`)).data?.data||[]}catch(o){console.warn("[Calendar] Failed to load events, rendering empty grid:",o),r=[]}}async function l(){let o=document.getElementById("calendar-grid");if(o){o.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let i=e.getFullYear(),s=e.getMonth(),n=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),c=document.getElementById("cal-month-label");c&&(c.textContent=n);let m=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(v=>v.value)),d=r.filter(v=>m.has(v.type)),p={};d.forEach(v=>{let w=(v.event_date||"").slice(0,10);p[w]||(p[w]=[]),p[w].push(v)});let u=new Date(i,s,1).getDay(),h=new Date(i,s+1,0).getDate(),b=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],g=new Date().toISOString().slice(0,10),y='<div class="calendar-grid">';b.forEach(v=>{y+=`<div class="cal-day-header">${v}</div>`});for(let v=0;v<u;v++)y+='<div class="cal-cell cal-cell-empty"></div>';for(let v=1;v<=h;v++){let w=`${i}-${String(s+1).padStart(2,"0")}-${String(v).padStart(2,"0")}`,T=p[w]||[],L=w===g;y+=`
          <div class="cal-cell ${L?"cal-today":""} ${T.length?"cal-has-events":""}"
               data-date="${w}" tabindex="0" role="button" aria-label="${w}">
            <div class="cal-day-num ${L?"today-num":""}">${v}</div>
            <div class="cal-events-preview">
              ${T.slice(0,3).map(C=>`
                <div class="cal-event-dot cal-color-${C.color||"gray"}" title="${Re(C.title||C.type)}">
                  <span class="cal-event-dot-label">${Da(C.title||C.branch_name||C.type,18)}</span>
                </div>
              `).join("")}
              ${T.length>3?`<div class="cal-more">+${T.length-3} lagi</div>`:""}
            </div>
          </div>`}let S=(u+h)%7;if(S!==0)for(let v=0;v<7-S;v++)y+='<div class="cal-cell cal-cell-empty"></div>';y+="</div>",o.innerHTML=y,o.querySelectorAll(".cal-cell[data-date]").forEach(v=>{v.addEventListener("click",()=>{let w=v.dataset.date,T=p[w]||[];if(!T.length)return;let L=document.getElementById("cal-event-list"),C=new Date(w+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=C,document.getElementById("cal-event-items").innerHTML=T.map(P=>`
            <div class="cal-event-item cal-color-border-${P.color||"gray"}">
              <div class="cal-event-type">${Ia(P.type)}</div>
              <div class="cal-event-title">${Re(P.title||"-")}</div>
              <div class="cal-event-branch">${Re(P.branch_name||"")}</div>
              ${P.status?`<div class="cal-event-status">${Re(P.status)}</div>`:""}
              ${P.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${P.days_remaining} hari</div>`:""}
            </div>
          `).join(""),L.style.display="block"})})}catch(i){console.error("[Calendar] Render error:",i),o&&(o.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}l()}function Da(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function Re(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Ia(t){return{schedule:"\u{1F5D3} Jadwal",reliever:"\u{1F504} Reliefer",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50E} Inspeksi",fogging:"\u{1F4A8} Fogging",contract_expiry:"\u{1F4CB} Kontrak Habis"}[t]||t}D();async function Qt(t){let e=le(),r=(e?.full_name||e?.username||"U")[0].toUpperCase(),l={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
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
  `;let o=localStorage.getItem("fm_token"),i=document.getElementById("session-info");if(o&&i)try{let s=JSON.parse(atob(o.split(".")[1])),n=new Date(s.exp*1e3);i.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{i.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async s=>{s.preventDefault();let n=document.getElementById("pwd-error"),c=document.getElementById("pwd-success"),m=document.getElementById("btn-save-pwd");n.style.display="none",c.style.display="none";let d=s.target,p=d.current_password.value,u=d.new_password.value,h=d.confirm_password.value;if(u!==h){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(u.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}m.disabled=!0,m.textContent="\u23F3 Menyimpan...";let b=await f("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:p,new_password:u})});m.disabled=!1,m.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',b.ok?(c.textContent="\u2705 Password berhasil diubah.",c.style.display="block",d.reset(),J("Password berhasil diubah.")):(n.textContent=b.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}D();var qe={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function H(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let l=Number(e);if(l>2e4&&l<99999){let o=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(o.getTime())?null:o.toISOString().slice(0,10)}}let r=e.split(/[\/\-\.]/);if(r.length===3){let[l,o,i]=r.map(m=>m.trim()),s=Number(l),n=Number(o),c=Number(i);if(l.length===4&&s>1900)return`${l}-${o.padStart(2,"0")}-${i.padStart(2,"0")}`;if(i.length===4&&c>1900)return s>12?`${i}-${o.padStart(2,"0")}-${l.padStart(2,"0")}`:n>12?`${i}-${l.padStart(2,"0")}-${o.padStart(2,"0")}`:`${i}-${o.padStart(2,"0")}-${l.padStart(2,"0")}`;if(i.length===2&&!isNaN(c)){let m=c>=50?`19${i}`:`20${i}`;return s>12?`${m}-${o.padStart(2,"0")}-${l.padStart(2,"0")}`:`${m}-${o.padStart(2,"0")}-${l.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function Vt(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Pa={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",phone:t["No. Hp"],join_date:H(t["Tanggal Masuk"]),status:t.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:H(t["Tanggal Mulai"]),end_date:H(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:H(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:H(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:H(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:H(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:H(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:H(t["Tanggal Target"]||t["Tgl Target"]),completion_date:H(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:H(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:H(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:H(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:H(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:H(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:H(t["Tanggal Back Up"]),completion_date:H(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:H(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:H(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Ba(t,e){let r=qe[t];if(!r)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Pa[r.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let l=[],o=[],i=[];return e.filter(n=>!Vt(n)).forEach((n,c)=>{let m=e.indexOf(n)+2,d=[];a.required.forEach(({key:u,label:h})=>{let b=n[u];if(b==null||String(b).trim()===""){let g=Object.keys(n).filter(y=>y.trim()).join(", ");d.push({column:h,originalValue:b||"",reason:`Kolom "${h}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${g.slice(0,120)}`})}});let p=a.map(n);d.length>0?o.push({row:m,data:p,raw:n,errors:d}):(l.push(n),i.push(p))}),{valid:l,errors:o,mapped:i}}function Yt(t){let e=[];return t.SheetNames.forEach(r=>{let a=qe[r];if(!a)return;let l=t.Sheets[r],o=window.XLSX.utils.sheet_to_json(l,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),i=Ba(r,o),s=o.filter(n=>!Vt(n));e.push({sheetName:r,module:a.module,label:a.label,total:s.length,valid:i.mapped.length,errorCount:i.errors.length,errors:i.errors,mapped:i.mapped,skipped:!1})}),e}function zt(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu","NAMA KARYAWAN":"Budi Santoso",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02",Day:"1"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06",Day:"1","Link Document":""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}]}).forEach(([a,l])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(l),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function Wt(t){let e=window.XLSX,r=e.utils.book_new(),a=!1;return t.forEach(l=>{if(!l.errors||l.errors.length===0)return;a=!0;let o=l.errors.map(s=>({"No. Baris":s.row,"Kolom Gagal":(s.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(s.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(s.data||{}).map(([n,c])=>[n,c??""]))})),i=e.utils.json_to_sheet(o);e.utils.book_append_sheet(r,i,l.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(r,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var La=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function Xt(t){t.innerHTML=`
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
              ${Object.entries(qe).map(([b,{label:g}])=>`<span class="import-sheet-tag">\u{1F4C4} ${b} \u2192 ${g}</span>`).join("")}
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
  `;let e=null,r=null,a=0,l={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function o(b){Object.entries(l).forEach(([g,y])=>{y.style.display=g===b?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let b=document.getElementById("btn-backup-db");b.disabled=!0,b.textContent="\u23F3 Memproses Backup...";try{let g=await f("/api/import/backup");if(g.ok){let y=new Blob([JSON.stringify(g.data,null,2)],{type:"application/json"}),k=URL.createObjectURL(y),S=document.createElement("a");S.href=k,S.download=`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.json`,document.body.appendChild(S),S.click(),document.body.removeChild(S),URL.revokeObjectURL(k),J("Backup berhasil diunduh!")}else Q("Gagal memproses backup: "+(g.data?.error||"Unknown error"))}catch(g){Q("Gagal memproses backup: "+g.message)}finally{b.disabled=!1,b.textContent="\u{1F4E6} Backup Database"}});let i=document.getElementById("btn-sync-google");i&&i.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let b=i.innerHTML;i.innerHTML='<span class="spinner"></span> Menyinkronkan...',i.disabled=!0;try{let g=await f("/api/sync/google-sheets",{method:"POST"});g.ok?alert("Sinkronisasi Berhasil: "+(g.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(g.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{i.innerHTML=b,i.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{zt(),J("Template Excel berhasil didownload!")});let s=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",b=>{b.stopPropagation(),s.click()}),s.addEventListener("change",b=>{b.target.files[0]&&c(b.target.files[0])}),n.addEventListener("dragover",b=>{b.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",b=>{b.preventDefault(),n.classList.remove("drag-over");let g=b.dataTransfer.files[0];g&&g.name.match(/\.xlsx?$/i)?c(g):Q("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,s.value="",document.getElementById("file-info").style.display="none",n.style.display="",o("upload")});async function c(b){e=b,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${b.name} (${(b.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await m(b)}async function m(b){o("validating");let g=document.getElementById("validation-status"),y=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");g.textContent="Membaca file Excel...",y.style.width="20%",await Pe(200);let k=await b.arrayBuffer(),S=window.XLSX.read(k,{type:"array",cellDates:!0});g.textContent=`Memvalidasi ${S.SheetNames.length} sheet...`,y.style.width="50%",await Pe(100),r=Yt(S),y.style.width="100%",g.textContent="Validasi selesai!",await Pe(300),d()}catch(k){o("upload"),Q("Gagal memproses file: "+k.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function d(){o("preview");let b=r.filter(C=>!C.skipped).length,g=r.reduce((C,P)=>C+P.total,0),y=r.reduce((C,P)=>C+P.valid,0),k=r.reduce((C,P)=>C+P.errorCount,0),S=g>0?Math.round(y/g*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${b} sheet</span>
      <span class="badge badge-secondary">${g} baris</span>
      <span class="badge badge-success">${y} valid (${S}%)</span>
      ${k>0?`<span class="badge badge-danger">${k} error</span>`:""}
    `;let v=document.getElementById("preview-table-container");v.innerHTML=`
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
          ${r.map((C,P)=>`
            <tr class="${C.errorCount>0?"row-error":C.skipped?"row-skipped":"row-ok"}">
              <td><strong>${C.sheetName}</strong></td>
              <td>${C.label}</td>
              <td style="text-align:center">${C.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${C.valid}</span></td>
              <td style="text-align:center">${C.errorCount>0?`<span class="badge badge-danger">${C.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${C.skipped?'<span class="badge badge-neutral">Dilewati</span>':C.errorCount>0&&C.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':C.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':C.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${C.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${P}">\u{1F50D} ${C.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,v.querySelectorAll(".btn-detail-error").forEach(C=>{C.addEventListener("click",()=>{let P=r[Number(C.dataset.idx)];p(P)})});let w=document.getElementById("error-detail-section"),T=document.getElementById("error-detail-container");T.innerHTML="",w.style.display="none";let L=document.getElementById("btn-start-import");y===0?(L.disabled=!0,L.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(L.disabled=!1,k>0?(L.innerHTML=`\u{1F680} Import ${y} Data Valid (${k} dilewati)`,L.title="Baris error akan dilewati, baris valid tetap diimport"):L.innerHTML=`\u{1F680} Mulai Import ${y} Data`)}function p(b){let g=document.getElementById("error-detail-section"),y=document.getElementById("error-detail-container");g.style.display="";let k=b.errors.slice(0,100).map(S=>(Array.isArray(S.errors)?S.errors:[]).map(w=>{let T=typeof w=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${S.row}</span></td>
            <td><strong>${T?w.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${T&&w.originalValue!==void 0?w.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${T?w.reason:w}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${T&&w.aliases?`Gunakan salah satu nama kolom:<br><em>${w.aliases}</em>`:T&&w.hint?w.hint:""}
            </td>
          </tr>
        `}).join("")).join("");y.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${b.sheetName} \u2014 ${b.errorCount} baris error dari ${b.total} total
          ${b.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
        ${b.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,g.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{o("upload"),document.getElementById("file-info").style.display="none",n.style.display="",e=null,s.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!r)return;Wt(r)?J("Log error berhasil didownload."):J("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let b=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";u(b)});async function u(b){o("importing"),a=Date.now();let g=[];La.forEach(w=>{let T=r?.find(L=>L.module===w&&L.mapped?.length>0);T&&g.push(T)});let y=document.getElementById("import-steps-list");y.innerHTML=g.map(w=>`
      <div class="import-step-item" id="step-item-${w.module}">
        <span class="step-item-icon" id="step-icon-${w.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${w.label} <span class="step-item-count">(${w.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${w.module}"></span>
      </div>
    `).join("");let k=document.getElementById("import-bar"),S=document.getElementById("import-current-status"),v={totalSheets:g.length,totalRows:g.reduce((w,T)=>w+T.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let w=0;w<g.length;w++){let T=g[w],L=document.getElementById(`step-icon-${T.module}`),C=document.getElementById(`step-status-${T.module}`);L.textContent="\u{1F504}",C.textContent="Mengimport...",S.textContent=`Mengimport ${T.label}...`,k.style.width=`${Math.round(w/g.length*100)}%`;try{let P=await f(`/api/import/${T.module}`,{method:"POST",body:JSON.stringify({rows:T.mapped,onDuplicate:b})});if(P.ok){let F=P.data;v.inserted+=F.inserted||0,v.skipped+=F.skipped||0,v.moduleResults.push({label:T.label,inserted:F.inserted||0,skipped:F.skipped||0,status:"ok"}),L.textContent="\u2705",C.innerHTML=`<span class="badge badge-success">${F.inserted||0} berhasil</span>${F.skipped>0?` <span class="badge badge-neutral">${F.skipped} skip</span>`:""}`}else v.failed++,v.moduleResults.push({label:T.label,inserted:0,skipped:0,status:"error",error:P.data?.error}),L.textContent="\u274C",C.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(P){v.failed++,v.moduleResults.push({label:T.label,inserted:0,skipped:0,status:"error",error:P.message}),L.textContent="\u274C",C.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Pe(150)}k.style.width="100%",S.textContent="Selesai!",await Pe(400),h(v)}function h(b){o("summary");let g=((Date.now()-a)/1e3).toFixed(1),y=b.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${y?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${y?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${b.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${b.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${b.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${b.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${b.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${b.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
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
          ${b.moduleResults.map(k=>`
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,r=null,s.value="",document.getElementById("file-info").style.display="none",n.style.display="",o("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Pe(t){return new Promise(e=>setTimeout(e,t))}D();var He=[],Zt=[];async function ea(t){He=await I(),Zt=await q(),$({container:t,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:He}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await f(`/api/sp?limit=10000&${r}`);if(a.ok){let l=a.data.data.map(i=>({"Nama Karyawan":i.employee_name||"",Divisi:i.division||"",Cabang:i.branch_name||"","Tanggal Sp":i.tanggal||"","Akhir Sp":i.akhir_sp||"","Jenis Sp":i.sp_type||"","Link Document / Foto":i.document_link||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(B(),V));o(l,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(B(),V));r(e,"Template_Import_SP")},onImport:async e=>{let r=i=>{if(!i)return null;let s=String(i||"").toLowerCase(),n=He.find(c=>String(c.label||"").toLowerCase()===s);return n?n.value:null},a=i=>{if(!i)return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let s=String(i).trim();if(/^\d{4,5}$/.test(s)){let c=Number(s);if(c>2e4&&c<99999){let m=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);let n=s.split(/[\/\-\.]/);if(n.length===3){let[c,m,d]=n.map(p=>p.trim());if(c.length===4&&m.length<=2&&d.length<=2)return`${c}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&m.length<=2&&c.length<=2)return`${d}-${m.padStart(2,"0")}-${c.padStart(2,"0")}`}return s},l=e.map(i=>({employee_name:String(i["Nama Karyawan"]||"").trim(),division:String(i.Divisi||"").trim(),branch_id:r(String(i.Cabang||"").trim()),tanggal:a(i["Tanggal Sp"]),akhir_sp:a(i["Akhir Sp"]),sp_type:String(i["Jenis Sp"]||"").trim(),document_link:String(i["Link Document / Foto"]||"").trim()})).filter(i=>i.employee_name&&i.branch_id),o=await f("/api/sp/import",{method:"POST",body:JSON.stringify(l)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}},formFields:[{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:Zt},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"combobox",name:"branch_id",label:"Cabang",required:!0,options:He,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}D();var fe=[],ta=[];async function aa(t){fe=await I(),ta=await q(),$({container:t,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",options:fe},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",options:fe}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await f(`/api/mutasi?limit=10000&${r}`);if(a.ok){let l=a.data.data.map(i=>({Tanggal:i.tanggal||"","Nama Karyawan":i.employee_name||"","Cabang Asal":i.from_branch_name||"","Cabang Tujuan":i.to_branch_name||"",Status:i.status||"",Dokumen:i.document_link||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(B(),V));o(l,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(B(),V));r(e,"Template_Import_Mutasi")},onImport:async e=>{let r=i=>{if(!i)return null;let s=String(i||"").toLowerCase(),n=fe.find(c=>String(c.label||"").toLowerCase()===s);return n?n.value:null},a=i=>{if(!i)return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let s=String(i).trim();if(/^\d{4,5}$/.test(s)){let c=Number(s);if(c>2e4&&c<99999){let m=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);let n=s.split(/[\/\-\.]/);if(n.length===3){let[c,m,d]=n.map(p=>p.trim());if(c.length===4&&m.length<=2&&d.length<=2)return`${c}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&m.length<=2&&c.length<=2)return`${d}-${m.padStart(2,"0")}-${c.padStart(2,"0")}`}return s},l=e.map(i=>({tanggal:a(i.Tanggal),employee_name:String(i["Nama Karyawan"]||"").trim(),from_branch_id:r(String(i["Cabang Asal"]||"").trim()),to_branch_id:r(String(i["Cabang Tujuan"]||"").trim()),status:String(i.Status||"").trim(),document_link:String(i.Dokumen||"").trim()})).filter(i=>i.tanggal&&i.employee_name&&i.from_branch_id&&i.to_branch_id),o=await f("/api/mutasi/import",{method:"POST",body:JSON.stringify(l)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:ta},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:fe,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:fe,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let r=Math.floor(Number(t)-25569);return new Date(r*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let r=t.split(/[\/\-]/);return`${r[2]}-${r[1]}-${r[0]}`}let e=t.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(e){let r=e[1],a=parseInt(e[2],10),l=parseInt(e[3],10);if(a>12&&l<=12)return`${r}-${e[3]}-${e[2]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let r=e.split("-");if(r.length===3&&r[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],l=parseInt(r[2],10),o=a[parseInt(r[1],10)-1];return`${l} ${o} ${r[0]}`}return e};function K(t){return async e=>{if(!xe()){pe("/login");return}return t(e)}}var Be=null;function Aa(){Be&&clearInterval(Be);let t=()=>{let e=new Date,r=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),l=document.getElementById("header-clock-time"),o=document.getElementById("header-clock-date");l&&(l.textContent=r),o&&(o.textContent=a)};t(),Be=setInterval(t,1e3)}async function Na(){try{let t=await f("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},r=(a,l)=>{let o=document.getElementById(a);o&&(o.textContent=l>0?l:"",o.style.display=l>0?"inline-flex":"none")};r("badge-issues",e.issues?.current||0),r("badge-contracts",e.expiring30?.current||0),r("badge-oo1",e.one_on_one?.current||0),r("badge-schedule",e.schedule?.current||0),r("badge-supply",e.supply?.current||0)}catch{}}var ve=[];async function Oa(){try{let t=await f("/api/dashboard/notifications");if(!t.ok)return;ve=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=ve.length>0?"block":"none",e.textContent=ve.length)}catch{}}function Fa(){if(!ve.length){ae({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,r)=>r()});return}let t=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${ve.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;ae({title:`Notifikasi (${ve.length})`,content:t,confirmText:"Tutup",onConfirm:(e,r)=>r()})}function na(){let t=le(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
  `;let r=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),l=document.getElementById("topbar-menu-btn"),o=document.getElementById("sidebar-close"),i=()=>{r.classList.add("open"),a.classList.add("show")},s=()=>{r.classList.remove("open"),a.classList.remove("show")};l?.addEventListener("click",i),o?.addEventListener("click",s),a?.addEventListener("click",s),document.querySelectorAll(".nav-item").forEach(c=>c.addEventListener("click",s));function n(){let c=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(p=>{let u=p.dataset.route;p.classList.toggle("active",c===u||u!=="/dashboard"&&c.startsWith(u))});let m=document.getElementById("topbar-title"),d=document.querySelector(".nav-item.active .nav-label");m&&d&&(m.textContent=d.textContent)}window.addEventListener("hashchange",n),n(),Aa(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await f("/api/auth/logout",{method:"POST"}),Ce(),Be&&clearInterval(Be),pe("/login")}),Na(),Oa(),document.getElementById("btn-notif")?.addEventListener("click",c=>{c.preventDefault(),Fa()})}async function Ma(){O("/login",({main:e})=>It(e)),O("/dashboard",K(({main:e})=>$t(e))),O("/calendar",K(({main:e})=>Gt(e))),O("/employees",K(({main:e,params:r})=>Pt(e,r))),O("/contracts",K(({main:e,params:r})=>Me(e,r))),O("/sp",K(({main:e})=>ea(e))),O("/mutasi",K(({main:e})=>aa(e))),O("/timeline",K(({main:e,params:r})=>Lt(e,r))),O("/issues",K(({main:e,params:r})=>At(e,r))),O("/one-on-one",K(({main:e,params:r})=>Nt(e,r))),O("/training",K(({main:e})=>Ot(e))),O("/relievers",K(({main:e,params:r})=>Ft(e,r))),O("/reports/inspection",K(({main:e})=>Mt(e))),O("/reports/cleaning",K(({main:e})=>Kt(e))),O("/reports/fogging",K(({main:e})=>Rt(e))),O("/reports/basecamp",K(({main:e})=>qt(e))),O("/reports/supply",K(({main:e})=>ut(e,"supply"))),O("/sop",K(({main:e})=>Ht(e))),O("/checklist",K(({main:e})=>jt(e))),O("/forms",K(({main:e})=>ut(e))),O("/users",K(({main:e})=>Ut(e))),O("/branches",K(({main:e})=>Jt(e))),O("/profile",K(({main:e})=>Qt(e))),O("/settings/import",K(({main:e})=>Xt(e)));let t=xe();if(!t&&window.location.hash!=="#/login"&&pe("/login"),t){let e=await f("/api/auth/me");e.ok?(Te(e.data.data),na()):(Ce(),pe("/login"))}window.addEventListener("fm:login",()=>{na(),pe("/dashboard")}),bt()}Ma();
