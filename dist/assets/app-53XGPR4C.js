var Xa=Object.defineProperty;var qe=(e,a)=>()=>(e&&(a=e(e=0)),a);var He=(e,a)=>{for(var r in a)Xa(e,r,{get:a[r],enumerable:!0})};var oe={};He(oe,{API:()=>pa,CLIENT_SIDE_MAX_ROWS:()=>Ue,IS_DEVELOPMENT:()=>je,apiFetch:()=>v,clearToken:()=>ye,getToken:()=>fe,getUser:()=>re,setToken:()=>Je,setUser:()=>ve});function fe(){return localStorage.getItem("fm_token")}function Je(e){localStorage.setItem("fm_token",e)}function ye(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function re(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function ve(e){localStorage.setItem("fm_user",JSON.stringify(e))}async function v(e,a={}){let r=fe(),t={"Content-Type":"application/json",...r?{Authorization:`Bearer ${r}`}:{},...a.headers||{}};try{let n=`cb=${Date.now()}`,l=e.includes("?")?"&":"?",b=`${pa}${e}${l}${n}`,h=await fetch(b,{...a,headers:t}),o;try{let u=await h.text();try{o=JSON.parse(u)}catch{o={error:`Server Error (${h.status}): ${u.substring(0,80)}...`}}}catch{o={error:"Gagal membaca respon dari server"}}return h.status===401&&(ye(),window.location.hash="#/login"),{ok:h.ok,status:h.status,data:o}}catch(n){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${n.message})`}}}}var Za,pa,je,Ue,I=qe(()=>{Za="",pa=Za,je=window.location.hostname==="localhost"||window.location.hostname==="127.0.0.1",Ue=1e4});var ga={};He(ga,{confirmDialog:()=>Qe,createModal:()=>ee});function ee({title:e,content:a,onConfirm:r,onCancel:t,confirmText:n="Simpan",cancelText:l="Batal",size:b="md",confirmClass:h="btn-primary"}){let o={sm:"400px",md:"560px",lg:"720px",xl:"900px"},u=document.createElement("div");u.className="modal-overlay",u.innerHTML=`
    <div class="modal" style="max-width:${o[b]||o.md}">
      <div class="modal-header">
        <h3 class="modal-title">${e}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof a=="string"?a:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${l}</button>
        ${r?`<button class="btn ${h} modal-confirm">${n}</button>`:""}
      </div>
    </div>
  `,a instanceof HTMLElement&&u.querySelector(".modal-body").appendChild(a);let c=()=>{u.classList.remove("show"),setTimeout(()=>u.remove(),250)};return u.querySelector(".modal-close").addEventListener("click",()=>{t&&t(),c()}),u.querySelector(".modal-cancel").addEventListener("click",()=>{t&&t(),c()}),r&&u.querySelector(".modal-confirm").addEventListener("click",()=>r(u,c)),u.addEventListener("click",i=>{i.target===u&&(t&&t(),c())}),document.body.appendChild(u),requestAnimationFrame(()=>u.classList.add("show")),{overlay:u,close:c}}function Qe(e,a,r="Konfirmasi"){return ee({title:r,content:`<p>${e}</p>`,onConfirm:(t,n)=>{a(),n()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var Se=qe(()=>{});var G={};He(G,{downloadExcel:()=>E,parseExcel:()=>Ye,renderExcelButtons:()=>We});function Ye(e){return new Promise((a,r)=>{let t=new FileReader;t.onload=n=>{try{let l=new Uint8Array(n.target.result),b=XLSX.read(l,{type:"array"}),h=b.SheetNames[0],o=b.Sheets[h];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${e.name}`),console.log(`File Size: ${(e.size/1024).toFixed(2)} KB`),console.log(`File Type: ${e.type||"unknown"}`),console.log(`Sheets Found: ${b.SheetNames.join(", ")}`),console.log(`Sheet Used: ${h}`);let u=XLSX.utils.decode_range(o["!ref"]||"A1:A1"),c=u.e.r-u.s.r+1,i=u.e.c-u.s.c+1;console.log(`Total Rows (including empty): ${c}`),console.log(`Total Columns: ${i}`);let d=[];for(let g=u.s.c;g<=u.e.c;++g){let m=o[XLSX.utils.encode_cell({c:g,r:u.s.r})];m&&m.v&&d.push(m.v)}console.log(`Headers Found: ${d.join(", ")}`),console.log("---------------------------");let s=XLSX.utils.sheet_to_json(o,{defval:""});Object.defineProperty(s,"__worksheet",{value:o,enumerable:!1}),Object.defineProperty(s,"__headers",{value:d,enumerable:!1}),a(s)}catch(l){r(l)}},t.onerror=n=>r(n),t.readAsArrayBuffer(e)})}function E(e,a){try{let r=XLSX.utils.json_to_sheet(e),t=XLSX.utils.book_new();XLSX.utils.book_append_sheet(t,r,"Data"),XLSX.writeFile(t,`${a}.xlsx`)}catch(r){throw console.error("Error generating Excel file:",r),r}}function We(e){return`
    <div class="excel-actions" style="display:flex;gap:0.5rem;margin-bottom:1rem;">
      <button class="btn btn-secondary btn-sm" id="btn-export-${e}">
        \u{1F4E5} Export Excel
      </button>
      <button class="btn btn-secondary btn-sm" id="btn-template-${e}">
        \u{1F4C4} Download Template
      </button>
      <label class="btn btn-primary btn-sm" style="cursor:pointer;margin:0;" id="label-import-${e}">
        <span class="import-text">\u{1F4E4} Import Excel</span>
        <input type="file" id="input-import-${e}" accept=".xlsx, .xls, .csv" style="display:none;">
      </label>
    </div>
  `}var L=qe(()=>{});I();var Ge={},Le=null;function N(e,a){Ge[e]=a}function ce(e){window.location.hash=e}function ma(){async function e(){let a=window.location.hash.replace("#","")||"/dashboard",[r,...t]=a.split("?"),n=Ge[r];if(!n){for(let[b,h]of Object.entries(Ge))if(b.endsWith("/*")&&r.startsWith(b.slice(0,-2))){n=h;break}}Le&&(Le(),Le=null);let l=document.getElementById("main-content");if(l&&(l.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),n){let b=new URLSearchParams(t.join("?")),h=r.split("/").filter(Boolean),o=await n({path:r,params:b,segments:h,main:l});o&&(Le=o)}else{let b=l||document.getElementById("app");b&&(b.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",e),e()}var ke;function et(){return ke||(ke=document.createElement("div"),ke.id="toast-container",document.body.appendChild(ke)),ke}function ua(e,a="info",r=3500){let t=et(),n=document.createElement("div");n.className=`toast toast-${a}`;let l={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};n.innerHTML=`<span class="toast-icon">${l[a]||"\u2139"}</span><span class="toast-msg">${e}</span>`,t.appendChild(n),requestAnimationFrame(()=>n.classList.add("show")),setTimeout(()=>{n.classList.remove("show"),setTimeout(()=>n.remove(),350)},r)}var j=e=>ua(e,"success"),J=e=>ua(e,"error");Se();I();I();function ba({columns:e,data:a,onEdit:r,onDelete:t,onView:n,actions:l=[],emptyText:b="Tidak ada data",bulkSelect:h=null}){let o=document.createElement("div");if(o.className="table-wrapper",!a||a.length===0)return o.innerHTML=`<div class="empty-state"><p>${b}</p></div>`,o;let u=document.createElement("table");u.className="data-table";let c=document.createElement("thead"),i=document.createElement("tr");if(h){let s=document.createElement("th");s.style.width="40px",s.style.textAlign="center";let g=document.createElement("input");g.type="checkbox",g.id="select-all-checkbox",g.title="Pilih semua",g.addEventListener("change",()=>{a.forEach(m=>{g.checked?h.selectedIds.add(m.id):h.selectedIds.delete(m.id)}),o.querySelectorAll(".row-checkbox").forEach(m=>m.checked=g.checked),h.onToggle()}),s.appendChild(g),i.appendChild(s)}if(e.forEach(s=>{let g=document.createElement("th");g.textContent=s.label,s.width&&(g.style.width=s.width),i.appendChild(g)}),r||t||n||l.length>0){let s=document.createElement("th");s.textContent="Aksi",s.style.width="120px",i.appendChild(s)}c.appendChild(i),u.appendChild(c);let d=document.createElement("tbody");return a.forEach(s=>{let g=document.createElement("tr");if(h){let m=document.createElement("td");m.style.textAlign="center",m.style.width="40px";let p=document.createElement("input");p.type="checkbox",p.className="row-checkbox",p.checked=h.selectedIds.has(s.id),p.addEventListener("change",()=>{if(p.checked)h.selectedIds.add(s.id);else{h.selectedIds.delete(s.id);let f=document.getElementById("select-all-checkbox");f&&(f.checked=!1)}h.onToggle()}),m.appendChild(p),g.appendChild(m)}if(e.forEach(m=>{let p=document.createElement("td");if(m.render){let f=m.render(s[m.key],s);f instanceof HTMLElement?p.appendChild(f):p.innerHTML=f||""}else p.textContent=s[m.key]!==null&&s[m.key]!==void 0&&s[m.key]!==""?s[m.key]:"";m.nowrap&&(p.style.whiteSpace="nowrap"),g.appendChild(p)}),r||t||n||l.length>0){let m=document.createElement("td");m.className="actions-cell";let p=document.createElement("div");if(p.className="btn-group",n){let f=document.createElement("button");f.className="btn btn-xs btn-ghost",f.innerHTML="\u{1F441}",f.title="Lihat",f.addEventListener("click",()=>n(s)),p.appendChild(f)}if(r){let f=document.createElement("button");f.className="btn btn-xs btn-secondary",f.innerHTML="\u270F\uFE0F",f.title="Edit",f.addEventListener("click",()=>r(s)),p.appendChild(f)}l.forEach(f=>{let _=document.createElement("button");_.className=`btn btn-xs ${f.class||"btn-ghost"}`,_.innerHTML=f.icon||f.label,_.title=f.label,_.addEventListener("click",()=>f.handler(s)),p.appendChild(_)}),m.appendChild(p),g.appendChild(m)}d.appendChild(g)}),u.appendChild(d),o.appendChild(u),o}function ha({page:e,pages:a,total:r,limit:t,onPage:n}){if(a<=1)return null;let l=document.createElement("div");l.className="pagination";let b=document.createElement("span");b.className="pagination-info",b.textContent=`Total: ${r} data`,l.appendChild(b);let h=document.createElement("div");h.className="pagination-btns";let o=(i,d,s=!1,g=!1)=>{let m=document.createElement("button");m.className=`btn btn-sm ${g?"btn-primary":"btn-ghost"} pagination-btn`,m.textContent=i,m.disabled=s,m.addEventListener("click",()=>n(d)),h.appendChild(m)};o("\xAB",1,e===1),o("\u2039",e-1,e===1);let u=Math.max(1,e-2),c=Math.min(a,e+2);for(let i=u;i<=c;i++)o(i,i,!1,i===e);return o("\u203A",e+1,e===a),o("\xBB",a,e===a),l.appendChild(h),l}Se();function Ve(e){return e.map(a=>{if(a.type==="hidden")return`<input type="hidden" name="${a.name}" value="${a.value||""}">`;if(a.type==="row")return`<div class="form-row">${Ve(a.fields)}</div>`;let r=a.required?"required":"",t=a.label?`<label class="form-label">${a.label}${a.required?' <span class="required">*</span>':""}</label>`:"",n="";switch(a.type){case"textarea":n=`<textarea name="${a.name}" class="form-control" placeholder="${a.placeholder||""}" ${r} rows="${a.rows||3}">${a.value||""}</textarea>`;break;case"select":let b=(a.options||[]).map(i=>{let d=typeof i=="object"?i.value:i,s=typeof i=="object"?i.label:i,g=a.value==d?"selected":"";return`<option value="${d}" ${g}>${s}</option>`}).join("");n=`<select name="${a.name}" class="form-control" ${r}><option value="">-- Pilih ${a.label||""} --</option>${b}</select>`;break;case"combobox":let h=`dl-${a.name}-${Math.random().toString(36).substring(7)}`,o=(a.options||[]).map(i=>{let d=typeof i=="object"?i.value:i;return`<option value="${typeof i=="object"?i.label:i}"></option>`}).join(""),u=a.value||"";if(a.value){let i=(a.options||[]).find(d=>(typeof d=="object"?d.value:d)==a.value);i&&(u=typeof i=="object"?i.label:i)}n=`
          <input type="text" name="${a.name}" list="${h}" class="form-control" value="${u}" placeholder="Pilih atau ketik baru..." ${r} autocomplete="off">
          <datalist id="${h}">${o}</datalist>
        `;break;case"checkbox":n=`<label class="checkbox-label"><input type="checkbox" name="${a.name}" value="1" ${a.value?"checked":""}> ${a.checkLabel||a.label}</label>`;break;case"date":let c=window.parseFlexibleDate&&a.value?window.parseFlexibleDate(a.value):a.value||"";n=`<input type="date" name="${a.name}" class="form-control" value="${c}" ${r}>`;break;case"number":n=`<input type="number" name="${a.name}" class="form-control" value="${a.value||""}" placeholder="${a.placeholder||""}" min="${a.min||""}" max="${a.max||""}" step="${a.step||"1"}" ${r}>`;break;case"email":n=`<input type="email" name="${a.name}" class="form-control" value="${a.value||""}" placeholder="${a.placeholder||""}" ${r}>`;break;case"url":n=`<input type="url" name="${a.name}" class="form-control" value="${a.value||""}" placeholder="${a.placeholder||"https://..."}" ${r}>`;break;default:n=`<input type="${a.type||"text"}" name="${a.name}" class="form-control" value="${a.value||""}" placeholder="${a.placeholder||""}" ${r} autocomplete="off">`}let l=a.hint?`<div class="form-hint">${a.hint}</div>`:"";return`<div class="form-group ${a.class||""}">${t}${n}${l}</div>`}).join("")}function fa(e){let a={},r=new FormData(e);for(let[t,n]of r.entries())a[t]=n===""?null:n;return e.querySelectorAll("input[type=checkbox]").forEach(t=>{t.checked||(a[t.name]=null)}),a}function ya(e,a){a&&Object.entries(a).forEach(([r,t])=>{let n=e.querySelector(`[name="${r}"]`);n&&(n.hasAttribute("list")||(n.type==="checkbox"?n.checked=!!t:n.type==="date"&&t&&window.parseFlexibleDate?n.value=window.parseFlexibleDate(t):n.value=t??""))})}L();function $({container:e,title:a,icon:r,apiPath:t,columns:n,formFields:l,filterFields:b,defaultFilters:h={},itemLabel:o="Data",canCreate:u=!0,canEdit:c=!0,canDelete:i=!0,onBeforeSubmit:d,onAfterLoad:s,onDataLoaded:g,extraActions:m=[],initialSearch:p="",exportOptions:f=null,bulkDelete:_=!1,paginationMode:y="server"}){let k=1,S={...h};p&&(S.search=p);let w=new Set;e.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${r} ${a}</h1>
      <div class="page-actions">
        ${u?`<button class="btn btn-primary" id="btn-create">+ Tambah ${o}</button>`:""}
      </div>
    </div>

    ${_?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${f?We(f.moduleName):""}

    ${b&&b.length>0?`
    <div class="filter-bar card">
      <div class="filter-bar-inner">
        ${b.map(T=>T.type==="search"?`<div class="filter-search"><input type="search" class="form-control" placeholder="${T.placeholder||"Cari..."}" id="filter-search" value="${S.search||""}"></div>`:T.type==="select"?`<select class="form-control filter-select" name="${T.name}" id="filter-${T.name}"><option value="">-- ${T.label} --</option>${(T.options||[]).map(x=>`<option value="${typeof x=="object"?x.value:x}" ${S[T.name]===(typeof x=="object"?x.value:x)?"selected":""}>${typeof x=="object"?x.label:x}</option>`).join("")}</select>`:"").join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
      </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function D(){if(!document.getElementById("bulk-toolbar"))return;let x=document.getElementById("bulk-count"),B=document.getElementById("btn-bulk-delete"),q=document.getElementById("btn-bulk-cancel");x.textContent=`${w.size} item dipilih`,w.size>0?(B.disabled=!1,q.disabled=!1):(B.disabled=!0,q.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{w.clear(),document.querySelectorAll(".row-checkbox").forEach(x=>x.checked=!1);let T=document.getElementById("select-all-checkbox");T&&(T.checked=!1),D()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(w.size===0)return;let T=[...w],x=document.createElement("div");x.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",x.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${T.length} ${o}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${T.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(x),x.querySelector("#bulk-cancel-btn").addEventListener("click",()=>x.remove()),x.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let B=x.querySelector("#bulk-confirm-btn");B.disabled=!0,B.textContent="Menghapus...";let q=await v(`${t}/bulk`,{method:"DELETE",body:JSON.stringify({ids:T})});x.remove(),q.ok?(j(`${T.length} ${o} berhasil dihapus.`),w.clear(),D(),O()):J(q.data?.error||"Gagal menghapus data.")})});let C=document.getElementById("filter-search"),P;if(C?.addEventListener("input",T=>{clearTimeout(P),P=setTimeout(()=>{S.search=T.target.value,k=1,O()},400)}),b?.forEach(T=>{T.type==="select"&&document.getElementById(`filter-${T.name}`)?.addEventListener("change",x=>{S[T.name]=x.target.value,k=1,O()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{S={...h},C&&(C.value=""),b?.forEach(T=>{let x=document.getElementById(`filter-${T.name}`);x&&(x.value="")}),k=1,O()}),document.getElementById("btn-create")?.addEventListener("click",()=>be(null)),f){document.getElementById(`btn-export-${f.moduleName}`)?.addEventListener("click",async x=>{let B=x.target,q=B.innerHTML;B.innerHTML="\u23F3 Loading...",B.disabled=!0;try{await f.onExport()}catch{J("Gagal export data")}finally{B.innerHTML=q,B.disabled=!1}}),document.getElementById(`btn-template-${f.moduleName}`)?.addEventListener("click",()=>{f.onTemplate()});let T=document.getElementById(`input-import-${f.moduleName}`);T?.addEventListener("change",async x=>{let B=x.target.files[0];if(!B)return;let q=document.getElementById(`label-import-${f.moduleName}`),Q=q?q.querySelector(".import-text"):null,Z=Q?Q.innerText:"";Q&&(Q.innerText="\u231B Memproses..."),q&&(q.style.pointerEvents="none"),T.disabled=!0;try{let F=await Ye(B);if(F.length===0)throw new Error("File kosong atau format salah");await f.onImport(F),j("Import berhasil!"),O()}catch(F){J(F.message||"Gagal import data")}finally{Q&&(Q.innerText=Z),q&&(q.style.pointerEvents="auto"),T.disabled=!1,T.value=""}})}async function O(){let T=document.getElementById("table-container");if(!T)return;T.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let x=y==="client",B=x?1:k,q=x?Ue:20,Q=new URLSearchParams({page:B,limit:q,...Object.fromEntries(Object.entries(S).filter(([,U])=>U))}),Z=await v(`${t}?${Q}`);if(!Z.ok){T.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${Z.data?.error||"Error"}</p></div>`;return}let F=Z.data?.data||Z.data||[],R=Z.data?.pagination,Ke=F.length;if(x){F=g(F);let U=F.length,V=20,te=Math.ceil(U/V);k>te&&te>0&&(k=te);let H=(k-1)*V,se=k*V;F=F.slice(H,se),R={page:k,limit:V,total:U,pages:te}}je&&console.log({mode:x?"Client-Side":"Server-Side",module:t,totalData:Ke,filteredData:F.length,currentPage:k,pageSize:R?R.limit:20,totalPages:R?R.pages:1,startIndex:x?(k-1)*20:0,endIndex:x?k*20:F.length,rowsRendered:F.length}),s&&s(F);let Ie=ba({columns:n,data:F,onEdit:c?U=>be(U):null,actions:m.map(U=>({...U,handler:V=>U.handler(V,O)})),emptyText:`Tidak ada ${String(o||"").toLowerCase()}`,bulkSelect:_?{selectedIds:w,onToggle:D}:null});T.innerHTML="",T.appendChild(Ie);let he=document.getElementById("pagination-container");if(he&&(he.innerHTML="",R&&R.pages>1)){let U=ha({page:R.page,pages:R.pages,total:R.total,limit:R.limit,onPage:V=>{k=V,O()}});U&&he.appendChild(U)}}function ge(T){let x=typeof l=="function"?l(T):l;return Ve(x)}function be(T){let x=!!T,B=document.createElement("form");if(B.noValidate=!0,B.innerHTML=ge(T),x){let Q=typeof l=="function"?l(T):l;ya(B,T)}let{close:q}=ee({title:x?`Edit ${o}`:`Tambah ${o}`,content:B,size:"lg",confirmText:x?"Simpan Perubahan":`Tambah ${o}`,onConfirm:async(Q,Z)=>{if(!B.reportValidity())return;let F=Q.querySelector(".modal-confirm");F.disabled=!0,F.textContent="Menyimpan...";let R=fa(B),Ke=typeof l=="function"?l(T):l,Ie=async te=>{for(let H of te)if(H.type==="row")await Ie(H.fields);else if(H.type==="combobox"&&R[H.name]){let se=R[H.name],Pe=(H.options||[]).find(X=>{let ne=String(typeof X=="object"?X.value:X),za=String(typeof X=="object"?X.label:X);return ne===se||za===se});if(Pe)R[H.name]=typeof Pe=="object"?Pe.value:Pe;else if(H.createApi){let X={};X[H.createApi.field]=se,H.createApi.extra&&Object.assign(X,H.createApi.extra);let ne=await v(H.createApi.path,{method:"POST",body:JSON.stringify(X)});if(ne.ok&&ne.data?.id)R[H.name]=ne.data.id;else if(ne.ok&&!ne.data?.id)R[H.name]=se;else throw new Error(`Gagal membuat master data: ${ne.data?.error||"Unknown error"}`)}}};try{await Ie(Ke)}catch(te){J(te.message),F.disabled=!1,F.textContent=x?"Simpan Perubahan":`Tambah ${o}`;return}d&&(R=await d(R,T));let he=x?"PUT":"POST",U=x?`${t}/${T.id}`:t,V=await v(U,{method:he,body:JSON.stringify(R)});V.ok?(j(x?`${o} berhasil diperbarui.`:`${o} berhasil ditambahkan.`),Z(),O()):(J(V.data?.error||"Gagal menyimpan data."),F.disabled=!1,F.textContent=x?"Simpan Perubahan":`Tambah ${o}`)}})}function Re(T){Qe(`Hapus ${o} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let x=await v(`${t}/${T.id}`,{method:"DELETE"});x.ok?(j(`${o} berhasil dihapus.`),O()):J(x.data?.error||"Gagal menghapus.")},`Hapus ${o}`)}return O(),O}I();function A(e){let a={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!e||e==="-"||String(e).trim()===""?"":`<span class="badge ${a[e]||"badge-neutral"}">${e}</span>`}function ze(e){return e==null?'<span class="badge badge-neutral">-</span>':e<0?`<span class="badge badge-danger">Expired (${Math.abs(e)}h)</span>`:e<=14?`<span class="badge badge-danger">${e} hari</span>`:e<=30?`<span class="badge badge-warning">${e} hari</span>`:`<span class="badge badge-success">${e} hari</span>`}function de(e){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[e]||"badge-neutral"}">${e||"-"}</span>`}function Xe(e){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[e]||"badge-neutral"}">${e||"-"}</span>`}function Y(e){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[e]||"badge-neutral"}">${e||"-"}</span>`}L();function Ze(e,a){if(e.period!=="Q3")return!1;let r=String(e.status||"").toLowerCase();if(r!=="selesai"&&r!=="completed"&&r!=="done")return!1;let t=String(e.activity_type||"").toLowerCase();return a==="inspeksi"?t.includes("inspeksi"):a==="gcdc"?t.includes("general cleaning")||t.includes("deep cleaning"):!1}I();L();function va(e,a){let r=String(e.status||"").toLowerCase();return a==="active"?r==="aktif":!1}I();L();function ea(e,a){if(String(e.status||"").toLowerCase()!=="aktif")return!1;if(a==="active")return!0;if(a==="expiring30"){if(!e.end_date)return!1;let t=new Date;t.setHours(0,0,0,0);let n=new Date(t);n.setDate(t.getDate()+30);let l=new Date(e.end_date);return l.setHours(0,0,0,0),l>=t&&l<=n}return!1}I();L();function ka(e,a){let r=String(e.status||"").toLowerCase();return a==="open"?r==="open":!1}I();function Sa(e,a){let r=String(e.status||"").toLowerCase();return a==="pending"?r==="pending":!1}var ie={};function _e(e){if(ie[e]){try{ie[e].destroy()}catch{}delete ie[e]}}function at(){Object.keys(ie).forEach(_e)}var ae=(e,a=0)=>{let r=Number(e);return isNaN(r)||e===null||e===void 0?a:r},pe=(e,a="\u2014")=>{if(e==null||e==="")return a;let r=String(e).trim();return r===""||r==="[object Object]"?a:r};var tt=e=>{if(!e||typeof e!="string")return"";try{let[a,r]=e.split("-");return new Date(Number(a),Number(r)-1).toLocaleDateString("id-ID",{month:"short",year:"2-digit"})}catch{return e}};function _a(e,a,r=900){if(!e)return;let t=Math.max(0,Math.round(ae(a)));if(t===0){e.textContent="0";return}let n=Date.now(),l=()=>{let b=Math.min((Date.now()-n)/r,1),h=1-Math.pow(1-b,3);e.textContent=Math.round(h*t).toLocaleString("id-ID"),b<1?requestAnimationFrame(l):e.textContent=t.toLocaleString("id-ID")};requestAnimationFrame(l)}var nt={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},it=e=>{let a=pe(e,"\u2014");return`<span class="status-pill ${nt[a]||"pill-neutral"}">${a}</span>`};var z={family:"Inter",size:11},me="#94A3B8",Be="#F1F5F9",aa=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],rt=()=>window.innerWidth<768;function ia(e={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:rt()?"bottom":"top",labels:{font:z,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:z,titleFont:{...z,weight:"700"}}},scales:{x:{grid:{color:Be},ticks:{font:z,color:me,maxRotation:0}},y:{grid:{color:Be},ticks:{font:z,color:me},beginAtZero:!0}},...e}}var lt=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),st=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function ta(e=3){return Array(e).fill(0).map((a,r)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${r<e-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function W(e,a,r=8e3){try{let t=new AbortController,n=setTimeout(()=>t.abort(),r),l=await v(e,{signal:t.signal}).catch(()=>null);if(clearTimeout(n),!l||!l.ok)return a;let b=l.data;return b?b.data!==void 0?b.data??a:b:a}catch{return a}}function ot(){["skel-donut","skel-trend","skel-insp","skel-contract"].forEach(t=>{let n=document.getElementById(t);n&&(n.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract"].forEach(t=>{let n=document.getElementById(t);if(n&&n.style.display==="none"){n.style.display="block";let l=n.parentElement;if(l&&!l.querySelector(".chart-empty")){let b=document.createElement("div");b.className="chart-empty",b.textContent="Belum ada data",n.style.display="none",l.appendChild(b)}}});let e=document.getElementById("kpi-row");e&&e.querySelector(".skeleton")&&xa({});let a=document.getElementById("mini-stats-row");a&&a.querySelector(".skeleton")&&Ca({}),["table-contracts","table-issues"].forEach(t=>{let n=document.getElementById(t);n&&n.querySelector(".skeleton")&&(n.innerHTML='<div class="chart-empty">Belum ada data</div>')});let r=document.getElementById("activity-log");r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function wa(e){at(),e._dashRefresh&&clearInterval(e._dashRefresh),e._skelTimeout&&clearTimeout(e._skelTimeout),e.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">


      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${lt()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${st()}</div>

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
        <div class="chart-card">
          <div class="chart-card-header" style="align-items:flex-start">
            <div>
              <a href="#/reports/inspection" class="chart-card-title" style="text-decoration:none; display:inline-block">Rata-rata Skor Inspeksi per Cabang <span style="font-size:0.8rem; color:var(--primary); font-weight:600; margin-left:8px">Lihat Laporan &rarr;</span></a>
              <div class="chart-card-subtitle" style="font-size:0.65rem">Skor rata-rata SCM & Cleaning</div>
            </div>
            <select class="btn-ghost" style="padding:4px;font-size:0.7rem;border:1px solid var(--border);border-radius:4px;cursor:pointer"><option>Bulan Ini</option></select>
          </div>
          <div class="chart-canvas-wrap" style="height:140px;position:relative;margin-top:10px">
            <div id="skel-insp" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-insp" style="display:none"></canvas>
          </div>
        </div>
      </div>

      <!-- Bottom Row -->
      <div class="bottom-row" style="margin-top:24px;">
        <!-- Jadwal Hari Ini -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Jadwal Hari Ini</div>
            <a href="#/calendar" class="chart-link">Lihat Kalender</a>
          </div>
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto">${ta(3)}</div>
        </div>
        <!-- KPI Kebersihan -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">KPI Kebersihan</div>
          </div>
          <div id="widget-kpi-kebersihan" style="margin-top:0px">${ta(4)}</div>
        </div>
        <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${ta(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>na(e)),e._skelTimeout=setTimeout(()=>ot(),5e3),await na(e),e._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?na(e):clearInterval(e._dashRefresh)},6e4)}async function na(e){e._skelTimeout&&(clearTimeout(e._skelTimeout),e._skelTimeout=null);let[a,r,t,n,l,b,h,o,u,c,i]=await Promise.all([W("/api/dashboard/kpi",{},8e3),W("/api/dashboard/issues-trend",{},8e3),W("/api/dashboard/issues-summary",{},8e3),W("/api/dashboard/inspection-bar",{},8e3),W("/api/dashboard/stats",{},8e3),W("/api/dashboard/calendar",[],8e3),W("/api/schedule?limit=10000",{data:[]},8e3),W("/api/employees?limit=10000",{data:[]},8e3),W("/api/contracts?limit=10000",{data:[]},8e3),W("/api/issues?limit=10000",{data:[]},8e3),W("/api/one_on_one?limit=10000",{data:[]},8e3)]);if(a){let d=Array.isArray(h?.data)?h.data:Array.isArray(h)?h:[],s=Array.isArray(o?.data)?o.data:Array.isArray(o)?o:[],g=Array.isArray(u?.data)?u.data:Array.isArray(u)?u:[],m=Array.isArray(c?.data)?c.data:Array.isArray(c)?c:[],p=Array.isArray(i?.data)?i.data:Array.isArray(i)?i:[];a.employees&&(a.employees.current=s.filter(f=>va(f,"active")).length),a.contracts&&(a.contracts.current=g.filter(f=>ea(f,"active")).length),a.expiring30&&(a.expiring30={current:g.filter(f=>ea(f,"expiring30")).length}),a.issues&&(a.issues.current=m.filter(f=>ka(f,"open")).length),a.one_on_one&&(a.one_on_one.current=p.filter(f=>Sa(f,"pending")).length),a.inspection_month&&(a.inspection_month.current=d.filter(f=>Ze(f,"inspeksi")).length),a.cleaning_month&&(a.cleaning_month.current=d.filter(f=>Ze(f,"gcdc")).length)}try{xa(a)}catch(d){console.warn("KPI render:",d)}try{Ca(a)}catch(d){console.warn("MiniStats render:",d)}try{ct(Array.isArray(t?.by_category)?t.by_category:[])}catch(d){console.warn("Donut render:",d),le("skel-donut","chart-donut")}try{dt(r)}catch(d){console.warn("Trend render:",d),le("skel-trend","chart-trend")}try{pt(n)}catch(d){console.warn("InspBar render:",d),le("skel-insp","chart-insp")}try{let d=Array.isArray(l)?l:Array.isArray(l?.recent_issues)?l.recent_issues:[];ut(d)}catch(d){console.warn("IssuesTable render:",d)}try{let d=Array.isArray(l?.expiring_contracts)?l.expiring_contracts:[];mt()}catch(d){console.warn("ContractsTable render:",d)}try{gt(Array.isArray(b)?b:[])}catch(d){console.warn("Agenda render:",d)}try{bt(a)}catch(d){console.warn("KPI Kebersihan render:",d)}try{ht()}catch(d){console.warn("Quick Actions render:",d)}}function xa(e){let a=document.getElementById("kpi-row");if(!a)return;e=e||{};let r=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];a.innerHTML=r.map(t=>{let n=ae(e[t.key]?.current,0);return`
      <a href="${t.href}" class="kpi-card ${t.color}" style="text-decoration:none;padding:12px 16px">
        <div style="display:flex; gap:16px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:48px;height:48px;border-radius:12px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${t.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${n}" style="font-size:1.8rem; font-weight:800; line-height:1; color:var(--text-1)">${n}</div>
            <div class="kpi-label" style="font-size:0.85rem; font-weight:700; color:var(--text-2); margin-top:6px">${t.label}</div>
            <div class="kpi-subtitle" style="font-size:0.7rem; color:var(--text-3); margin-top:2px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis">${t.sub}</div>
          </div>
        </div>
      </a>`}).join(""),a.querySelectorAll(".kpi-value").forEach(t=>{_a(t,parseInt(t.dataset.target)||0)})}function Ca(e){let a=document.getElementById("mini-stats-row");if(!a)return;e=e||{};let r=[{icon:"\u{1F4C5}",label:"Jadwal",val:e.schedule?.current,href:"#/timeline",color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:e.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F4E6}",label:"Permintaan",val:e.supply?.current,href:"#/reports/supply",color:"mini-orange"},{icon:"\u{1F50D}",label:"Inspeksi",val:e.inspection_month?.current,href:"#/timeline?dash_filter=inspeksi",color:"mini-blue"},{icon:"\u{1F9F9}",label:"GCDC",val:e.cleaning_month?.current,href:"#/timeline?dash_filter=gcdc",color:"mini-green"},{icon:"\u{1F4A8}",label:"Fogging",val:e.fogging_month?.current,href:"#/reports/fogging",color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:e.branches?.current,href:"#/branches",color:"mini-teal"}];a.innerHTML=r.map(t=>`
    <a href="${t.href}" class="mini-stat ${t.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${t.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${ae(t.val)}">0</div>
        <div class="mini-stat-text">${t.label}</div>
      </div>
    </a>`).join(""),a.querySelectorAll(".mini-stat-value").forEach(t=>_a(t,parseInt(t.dataset.target)||0,700))}function ct(e){le("skel-donut","chart-donut");let a=document.getElementById("chart-donut"),r=document.getElementById("donut-legend");if(!a||!r)return;_e("donut");let t=(e||[]).filter(o=>ae(o.count)>0);if(!t.length){ra(a,"Belum ada data permasalahan");return}let n=t.map(o=>`${pe(o.category,"Lainnya")}`),l=t.map(o=>ae(o.count)),b=l.reduce((o,u)=>o+u,0);r.innerHTML=t.map((o,u)=>{let c=aa[u%aa.length],i=b>0?Math.round(o.count/b*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${c}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${o.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${i}%)</span></div>
          <div class="donut-legend-label">${n[u]}</div>
        </div>
      </div>
    `}).join("");let h={id:"centerText",beforeDraw:function(o){let u=o.width,c=o.height,i=o.ctx;i.restore();let d=(c/80).toFixed(2);i.font="bold "+d+"em Inter",i.textBaseline="middle",i.fillStyle="#1E293B";let s=b.toString(),g=Math.round((u-i.measureText(s).width)/2),m=c/2;i.fillText(s,g,m-10),i.font="600 "+(d*.35).toFixed(2)+"em Inter",i.fillStyle="#64748B";let p="Total",f=Math.round((u-i.measureText(p).width)/2);i.fillText(p,f,m+15),i.save()}};ie.donut=new Chart(a,{type:"doughnut",data:{labels:n,datasets:[{data:l,backgroundColor:aa,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:z,titleFont:{...z,weight:"700"},callbacks:{label:o=>` ${o.label}: ${o.parsed} kasus`}}},cutout:"75%"},plugins:[h]})}function dt(e){le("skel-trend","chart-trend");let a=document.getElementById("chart-trend");if(!a)return;_e("trend"),e=e||{};let r=(e.labels||[]).map(tt),t=(e.open||[]).map(l=>ae(l)),n=(e.closed||[]).map(l=>ae(l));if(!r.length){ra(a,"Belum ada data trend");return}ie.trend=new Chart(a,{type:"line",data:{labels:r,datasets:[{label:"Open",data:t,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:n,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#10B981",borderWidth:2}]},options:ia({plugins:{legend:{display:!1}}})})}function pt(e){le("skel-insp","chart-insp");let a=document.getElementById("chart-insp");if(!a)return;_e("inspBar"),e=e||{};let r=e.labels||[],t=(e.fc||[]).map(l=>ae(l)),n=(e.spv||[]).map(l=>ae(l));if(!r.length){ra(a,"Belum ada data inspeksi");return}ie.inspBar=new Chart(a,{type:"bar",data:{labels:r,datasets:[{label:"Skor FC",data:t,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:n,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:ia({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:z,color:me,maxRotation:45,minRotation:30}},y:{grid:{color:Be},ticks:{font:z,color:me},min:0,max:100}}})})}function mt(){le("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;_e("contractMiniBar");let a=["Jun","Jul","Agu","Sep","Okt","Nov","Des"],r=[12,18,9,24,15,30,42],n=e.getContext("2d").createLinearGradient(0,0,0,200);n.addColorStop(0,"#60A5FA"),n.addColorStop(1,"#2563EB"),ie.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:r,backgroundColor:n,borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:ia({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:z,color:me,maxRotation:0}},y:{grid:{color:Be,borderDash:[4,4],drawBorder:!1},ticks:{font:z,color:me,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function ut(e){let a=document.getElementById("table-issues");if(!a)return;let r=(e||[]).slice(0,8);if(!r.length){a.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}a.innerHTML=`
    <div class="dash-list">
      ${r.map(t=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${it(t.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${pe(t.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${pe(t.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function gt(e){let a=document.getElementById("widget-agenda");if(!a)return;let r=(e||[]).slice(0,10);if(!r.length){a.innerHTML=""return}a.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${r.map(t=>{let n="#3B82F6",l="#EFF6FF",b="Agenda",h=(t.title||"").toLowerCase();return h.includes("inspeksi")?(n="#10B981",l="#ECFDF5",b="Inspeksi"):h.includes("cleaning")||h.includes("gcdc")?(n="#3B82F6",l="#EFF6FF",b="Cleaning"):h.includes("reliefer")?(n="#F59E0B",l="#FFFBEB",b="Reliefer"):h.includes("fogging")&&(n="#8B5CF6",l="#F5F3FF",b="Fogging"),`
        <div style="display:flex;gap:16px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px">${new Date(t.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${n};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${pe(t.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3)">${pe(t.branch_name)}</div>
          </div>
          <div style="font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${l};color:${n}">${b}</div>
        </div>
      `}).join("")}
    </div>
  `}function bt(e){let a=document.getElementById("widget-kpi-kebersihan");if(!a)return;let r=[{label:"Kebersihan Area",val:"97%",target:"Target 95%",icon:"\u{1F9F9}",bg:"#ECFDF5",color:"#10B981"},{label:"Penyelesaian Complaint",val:"100%",target:"Target 100%",icon:"\u23F1\uFE0F",bg:"#ECFDF5",color:"#10B981"},{label:"Kepatuhan Jadwal Cleaning",val:"99%",target:"Target 100%",icon:"\u23F1\uFE0F",bg:"#EFF6FF",color:"#3B82F6"},{label:"Kepatuhan GCDC",val:"100%",target:"Target 100%",icon:"\u{1F9F9}",bg:"#EFF6FF",color:"#3B82F6"},{label:"Complaint Cleaning (\u226410)",val:"4",target:"Target \u226410",icon:"\u{1F4DD}",bg:"#F5F3FF",color:"#8B5CF6"},{label:"Pelaksanaan Fogging",val:"100%",target:"Target 100%",icon:"\u{1F4A8}",bg:"#F5F3FF",color:"#8B5CF6"}];a.innerHTML=`
    <div style="display:grid;grid-template-columns:1fr;gap:12px;padding-bottom:12px;height:160px;overflow-y:auto;padding-right:8px;">
      ${r.map(t=>{let n=t.val.includes("%")?parseInt(t.val):Math.min(100,parseInt(t.val)*10);return`
        <div class="prog-item">
          <div class="prog-header">
            <div class="prog-title">
              <div class="prog-title-icon" style="background:${t.bg};color:${t.color}">${t.icon}</div>
              ${t.label}
            </div>
            <div class="prog-val">${t.val}</div>
          </div>
          <span class="prog-target">${t.target}</span>
          <div class="prog-bar-bg">
            <div class="prog-bar-fill" style="width:${n}%;background:${t.color}"></div>
          </div>
        </div>
      `}).join("")}
    </div>
  `}function ht(){let e=document.getElementById("quick-actions");if(!e)return;let a=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];e.innerHTML=a.map(r=>`
    <a href="${r.href}" class="action-btn">
      <div class="action-icon" style="background:${r.bg}">${r.icon}</div>
      ${r.label}
    </a>
  `).join("")}function le(e,a){let r=document.getElementById(e),t=document.getElementById(a);r&&(r.style.display="none",r.style.position=""),t&&(t.style.display="block")}function ra(e,a="Belum ada data"){if(!e)return;e.style.display="none";let r=e.parentElement;if(!r)return;if(!r.querySelector(".chart-empty")){let n=document.createElement("div");n.className="chart-empty",n.textContent=a,r.appendChild(n)}}I();async function Ta(e){document.getElementById("app").innerHTML=`
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
  `;let a=document.getElementById("login-form"),r=document.getElementById("login-error"),t=document.getElementById("login-btn"),n=document.getElementById("toggle-password"),l=document.getElementById("login-password");n?.addEventListener("click",()=>{let b=l.type==="text";l.type=b?"password":"text",n.style.color=b?"":"var(--primary)"}),a?.addEventListener("submit",async b=>{b.preventDefault(),r.style.display="none";let h=a.username.value.trim(),o=a.password.value;if(!h||!o){r.textContent="Username dan password wajib diisi.",r.style.display="block";return}t.querySelector(".btn-text").style.display="none",t.querySelector(".btn-spinner").style.display="",t.disabled=!0;try{let u=await v("/api/auth/login",{method:"POST",body:JSON.stringify({username:h,password:o})});u.ok&&u.data.success?(Je(u.data.data.token),ve(u.data.data.user),j("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(r.textContent=u.data.error||"Username atau password salah.",r.style.display="block",t.classList.add("shake"),setTimeout(()=>t.classList.remove("shake"),600))}catch{r.textContent="Gagal terhubung ke server. Periksa koneksi internet.",r.style.display="block"}finally{t.querySelector(".btn-text").style.display="",t.querySelector(".btn-spinner").style.display="none",t.disabled=!1}})}I();L();var we=[],la=[];async function ft(){la=(await v("/api/branches?all=1")).data?.data||[],we=la.map(a=>({value:a.id,label:a.full_name}))}function yt(e,a){let r=String(e.status||"").toLowerCase();return a==="active"?r==="aktif":!1}async function $a(e,a){await ft();let r=a?a.get("dash_filter"):null;$({container:e,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:t=>r?t.filter(n=>yt(n,r)):t,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:t=>de(t)},{key:"phone",label:"No. HP",render:t=>t?`<a href="tel:${t}">${t}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:t=>window.formatDate(t)},{key:"status",label:"Status",render:t=>A(t)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:we},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:t=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:t?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:t?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:t?.branch_id&&!we.find(n=>n.value==t.branch_id)?[...we,{value:t.branch_id,label:t.branch_name||t.branch_id}]:we,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:t?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:t?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:t?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let t=await v(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(t.ok){let n=t.data.data.map(l=>({"Nama Lengkap":l.full_name,Cabang:l.branch_name||"",Divisi:l.division||"","No. HP":l.phone||"","Tgl Masuk":l.join_date||"",Status:l.status||""}));E(n,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async t=>{let n=h=>{if(!h)return null;let o=String(h||"").toLowerCase(),u=la.find(c=>String(c.full_name||"").toLowerCase()===o||String(c.code||"").toLowerCase()===o||String(c.name||"").toLowerCase()===o);return u?u.id:null},l=t.map(h=>({full_name:String(h["Nama Lengkap"]||"").trim(),branch_id:n(String(h.Cabang||"").trim()),division:String(h.Divisi||"").trim()||"FACILITY CARE",phone:String(h["No. HP"]||"").trim(),join_date:String(h["Tgl Masuk"]||"").trim(),status:String(h.Status||"").trim(),notes:String(h.Catatan||"").trim()})).filter(h=>h.full_name),b=await v("/api/employees/import",{method:"POST",body:JSON.stringify(l)});if(!b.ok)throw new Error(b.data?.error||"Import gagal")}}})}I();L();var xe=[],Ae=[];async function vt(){let[e,a]=await Promise.all([v("/api/branches?all=1"),v("/api/employees?limit=10000&status=Aktif")]);xe=(e.data?.data||[]).map(r=>({value:r.id,label:r.full_name})),Ae=(a.data?.data||[]).map(r=>({value:r.id,label:r.full_name}))}var sa=async e=>{let a=[],r=1;for(;;){let n=await(await Promise.resolve().then(()=>(I(),oe))).apiFetch(`${e}${e.includes("?")?"&":"?"}limit=100&page=${r}`);if(!n.ok)break;let l=n.data?.data||n.data||[],b=Array.isArray(l)?l:[];if(a=a.concat(b),b.length<100||n.data?.pagination&&r>=n.data.pagination.pages)break;r++}return a};function kt(e,a){if(String(e.status||"").toLowerCase()!=="aktif")return!1;if(a==="active")return!0;if(a==="expiring30"){if(!e.end_date)return!1;let t=new Date;t.setHours(0,0,0,0);let n=new Date(t);n.setDate(t.getDate()+30);let l=new Date(e.end_date);return l.setHours(0,0,0,0),l>=t&&l<=n}return!1}async function Ne(e,a){await vt();let r=a?a.get("dash_filter"):null;$({container:e,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",onDataLoaded:t=>r?t.filter(n=>kt(n,r)):t,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:t=>de(t)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:t=>window.formatDate(t)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:t=>!t||String(t).startsWith("2099")?"Tetap / PKWTT":window.formatDate(t)},{key:"days_remaining",label:"Sisa Kontrak",render:(t,n)=>n.end_date&&String(n.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':ze(t)},{key:"status",label:"Status",render:t=>A(t)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:xe},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:t=>(t.end_date||(t.end_date="2099-12-31"),t),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let t=document.createElement("button");t.id="btn-find-missing",t.className="btn btn-ghost",t.innerHTML="\u{1F50D} Cek Selisih Karyawan",t.style.marginLeft="8px",t.style.color="#EF4444",t.style.border="1px solid currentColor",t.onclick=async()=>{t.innerHTML="\u231B Mencari...",t.disabled=!0;try{let[l,b]=await Promise.all([sa("/api/employees?status=Aktif"),sa("/api/contracts")]);if(l.length>0){let h=b.filter(i=>i.status==="Aktif"),o=new Set(h.map(i=>i.employee_id)),u=l.filter(i=>!o.has(i.id)),c=`<p style="margin-bottom:12px">Data yang terbaca: <b>${l.length}</b> Karyawan Aktif, dan <b>${h.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${u.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;u.forEach(i=>{let d=b.filter(g=>g.employee_id===i.id),s='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(d.length>0){let g=d[0];s=`Pernah ada kontrak (Status: <b style="color:#EF4444">${g.status}</b>, Selesai: ${window.formatDate(g.end_date)})`}c+=`<li style="margin-bottom:8px"><b>${i.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${i.branch_name||"-"} | ${s}</span></li>`}),c+="</ul>",Promise.resolve().then(()=>(Se(),ga)).then(i=>i.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:c,cancelText:"Tutup"}))}}catch(l){console.error(l)}t.innerHTML="\u{1F50D} Cek Selisih Karyawan",t.disabled=!1};let n=document.querySelector(".page-actions");n&&n.appendChild(t)}},formFields:t=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:t?.employee_id&&!Ae.find(n=>n.value==t.employee_id)?[...Ae,{value:t.employee_id,label:t.employee_name||t.employee_id}]:Ae,createApi:{path:"/api/employees",field:"full_name"},value:t?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:t?.branch_id&&!xe.find(n=>n.value==t.branch_id)?[...xe,{value:t.branch_id,label:t.branch_name||t.branch_id}]:xe,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:t?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:t?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:t?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:t?.end_date&&!String(t.end_date).startsWith("2099")?t.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:t?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:t?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let t=await v(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(t.ok){let n=t.data.data.map(l=>({"Nama Lengkap":l.employee_name,Cabang:l.branch_name||"","Div / Bagian":l.division||"","Tanggal Mulai":l.start_date||"","Tanggal Selesai":l.end_date&&String(l.end_date).startsWith("2099")?"":l.end_date||"","Sisa Kontrak":l.end_date&&String(l.end_date).startsWith("2099")?"Tetap":l.days_remaining!==null&&l.days_remaining!==void 0?`${l.days_remaining} Hari`:"",Status:l.status||""}));E(n,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async t=>{let[n,l]=await Promise.all([v("/api/branches?limit=10000"),sa("/api/employees")]),b=n.data?.data||[],h=l||[];console.log(`Total employee yang berhasil dimuat dari database : ${h.length}`),h.length>0&&(console.log("Contoh 5 employee pertama:"),h.slice(0,5).forEach((p,f)=>{console.log(`${f+1}. ID: ${p.id}, Name: ${p.full_name}, Status: ${p.status}`)}));let o=p=>{if(!p)return null;let f=String(p||"").replace(/\s+/g," ").toLowerCase().trim(),_=b.find(y=>String(y.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===f||String(y.code||"").replace(/\s+/g," ").toLowerCase().trim()===f||String(y.name||"").replace(/\s+/g," ").toLowerCase().trim()===f);return _?_.id:null},u=(p,f)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${f}`),console.log(`Nama dari Excel : "${p}"`),!p)return console.log("Alasan gagal mapping : Nama kosong"),null;let _=String(p||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${_}"`),console.log(`Jumlah employee di database : ${h.length}`);let y=h.find(k=>String(k.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===_);return y?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${y.id}`),y.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},c=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let f=String(p).trim();if(/^\d{4,5}(\.\d+)?$/.test(f)){let y=Math.floor(Number(f));if(y>2e4&&y<99999){let k=new Date(Date.UTC(1899,11,30)+y*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(f))return f.slice(0,10);let _=f.split(/[\/\-\.]/);if(_.length===3){let[y,k,S]=_.map(w=>w.trim());if(y.length===4&&k.length<=2&&S.length<=2)return`${y}-${k.padStart(2,"0")}-${S.padStart(2,"0")}`;if(S.length===4&&k.length<=2&&y.length<=2)return`${S}-${k.padStart(2,"0")}-${y.padStart(2,"0")}`}return f},i=t.map((p,f)=>{let _=f+2,y=String(p["Nama Lengkap"]||"").trim(),k=p["Tanggal Mulai"],S=c(k);if(!S){let C=t.__worksheet,P=t.__headers||[],O=P.indexOf("Tanggal Mulai"),ge="N/A",be="N/A",Re="N/A";if(O!==-1&&C&&window.XLSX){let x=window.XLSX.utils.encode_cell({c:O,r:_-1});Re=x;let B=C[x];B?(ge=B.t||"undefined",be=B.w||"undefined"):ge="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let T="Unknown";k==null||k===""?T="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":k instanceof Date&&isNaN(k.getTime())?T="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":T="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${_}`),console.log(`Employee Name : ${y}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${O})`),console.log(`Raw Cell Value : "${k}"`),console.log(`JavaScript Type : ${typeof k}`),console.log(`SheetJS Cell Type : ${ge}`),console.log(`SheetJS Formatted Value : "${be}"`),console.log(`Value After Trim : "${String(k||"").trim()}"`),console.log(`Value After Date Parser : "${S}"`),console.log(`Is Empty : ${!k}`),console.log(`Is Invalid Date : ${k instanceof Date?isNaN(k.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${T}`),console.log(`Workbook Sheet : ${C?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${Re}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(p,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(P)),console.log(`==========================
`)}let w=u(y,_),D=null;return w?S||(D="Tanggal Mulai kosong atau tidak berformat tanggal"):D="Karyawan tidak ditemukan di Database",{isValid:!!(w&&S),invalidReason:D,rowNum:_,data:{employee_id:w,branch_id:o(String(p.Cabang||"").trim()),division:String(p["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:S,end_date:c(p["Tanggal Selesai"])||"2099-12-31",status:String(p.Status||"").trim(),_rawName:y}}}),d=[],s=[];if(i.forEach(p=>{p.isValid?d.push(p.data):s.push({rowNum:p.rowNum,name:p.data._rawName,reason:p.invalidReason})}),console.log(`Split Validation - Valid: ${d.length}, Invalid: ${s.length}`),d.length===0){let p=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${t.length}
Valid: 0
Invalid: ${s.length}

Daftar Kegagalan (Contoh):
`;s.slice(0,10).forEach(f=>{p+=`- Row ${f.rowNum} | Nama: ${f.name} | Alasan: ${f.reason}
`}),s.length>10&&(p+=`- ... dan ${s.length-10} lainnya.
`),alert(p);return}let g=await v("/api/contracts/import",{method:"POST",body:JSON.stringify(d)}),m=`IMPORT SUMMARY
======================
`;m+=`Total Baris Excel : ${t.length}
`,m+=`Baris Valid       : ${d.length}
`,m+=`Baris Invalid     : ${s.length}

`,g&&g.data&&g.data.metrics?(m+=`Berhasil INSERT   : ${g.data.metrics.inserted}
`,m+=`Berhasil UPDATE   : ${g.data.metrics.updated}
`):m+=`Berhasil diproses : ${d.length}
`,s.length>0&&(m+=`
DAFTAR DATA DILEWATI:
`,s.forEach(p=>{m+=`- Row ${p.rowNum} | ${p.name} | ${p.reason}
`})),alert(m),typeof Ne=="function"&&Ne()}}})}I();L();var Ce=[],oa=[];function St(e){if(!Array.isArray(e))return"Q3";let a=["Q4","Q3","Q2","Q1"];for(let r of a)if(e.some(t=>t.period===r))return r;return"Q3"}function _t(e,a){if(e.period!=="Q3")return!1;let r=String(e.status||"").toLowerCase();if(r!=="selesai"&&r!=="completed"&&r!=="done")return!1;let t=String(e.activity_type||"").toLowerCase();return a==="inspeksi"?t.includes("inspeksi"):a==="gcdc"?t.includes("general cleaning")||t.includes("deep cleaning"):!1}async function Ea(e,a){let[r,t,n]=await Promise.all([v("/api/branches?all=1"),v(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),v(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`)]);Ce=(r.data?.data||[]).map(d=>({value:d.id,label:d.full_name}));let l=(t.data?.data||[]).map(d=>({value:d.full_name,label:d.full_name}));oa=[...(n.data?.data||[]).filter(d=>d.role==="FC Spesialis").map(d=>({value:d.name,label:d.name}))];let h=d=>d&&!l.find(s=>s.value===d)?[...l,{value:d,label:d}]:l,o=d=>{if(!d||d==="-"||String(d).trim()==="")return"";let s=String(d).split("-");return s.length===3&&s[0].length===4?`${s[2]}-${s[1]}-${s[0]}`:d},u=n.data?.data||[],c=St(u),i=a?a.get("dash_filter"):null;$({container:e,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",defaultFilters:i?{period:"Q3"}:{period:c},onDataLoaded:d=>(i&&(d=d.filter(s=>_t(s,i))),d.sort((s,g)=>{let m=s.opening_date?new Date(s.opening_date).getTime():0;return(g.opening_date?new Date(g.opening_date).getTime():0)-m})),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:d=>Xe(d)},{key:"period",label:"Periode",render:d=>Y(d)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:d=>o(d)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:d=>o(d)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:d=>o(d)},{key:"status",label:"Status",render:d=>A(d)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:Ce},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:oa}],formFields:d=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:d?.branch_id&&!Ce.find(s=>s.value==d.branch_id)?[...Ce,{value:d.branch_id,label:d.branch_name||d.branch_id}]:Ce,createApi:{path:"/api/branches",field:"full_name"},value:d?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"],value:d?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:d?.period},{name:"pic",label:"PIC",type:"combobox",options:oa,createApi:{path:"/api/pic",field:"name"},value:d?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:d?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:d?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:d?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:d?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:d?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let d=await v(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(d.ok){let s=d.data.data.map(g=>({Cabang:g.branch_name||"",Kegiatan:g.activity_type||"",Periode:g.period||"",PIC:g.pic||"","Tgl Opening":g.opening_date||"","Tgl Target":g.target_date||"","Tgl Selesai":g.completion_date||"",Status:g.status||""}));E(s,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async d=>{let g=(await v("/api/branches?all=1")).data?.data||[],m=y=>{if(!y)return null;let k=String(y||"").toLowerCase(),S=g.find(w=>String(w.full_name||"").toLowerCase()===k||String(w.code||"").toLowerCase()===k||String(w.name||"").toLowerCase()===k);return S?S.id:null},p=y=>{if(y==null||y==="")return"";if(y instanceof Date&&!isNaN(y.getTime()))return y.toISOString().slice(0,10);let k=String(y).trim();if(k===""||k==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(k))return k.slice(0,10);if(/^\d{4,5}$/.test(k)){let w=Number(k);if(w>2e4&&w<99999){let D=new Date(Date.UTC(1899,11,30)+w*864e5);return isNaN(D.getTime())?"":D.toISOString().slice(0,10)}}let S=k.split(/[\/\-\.]/);if(S.length===3){let[w,D,C]=S.map(P=>P.trim());if(w.length===4&&D.length<=2&&C.length<=2)return`${w}-${D.padStart(2,"0")}-${C.padStart(2,"0")}`;if(C.length===4&&D.length<=2&&w.length<=2)return`${C}-${D.padStart(2,"0")}-${w.padStart(2,"0")}`}return k},f=d.map(y=>({branch_id:m(String(y.Cabang||"").trim()),activity_type:String(y.Kegiatan||"").trim(),period:String(y.Periode||"").trim(),pic:String(y.PIC||y.Pic||"").trim(),opening_date:p(y["Tgl Opening"]||y["Tanggal Opening"]||y["Tgl Openir"]),target_date:p(y["Tgl Target"]||y["Tanggal Target"]),completion_date:p(y["Tgl Selesai"]||y["Tanggal Selesai"]),status:String(y.Status||"").trim(),notes:String(y.Catatan||y.Keterangan||"").trim()})).filter(y=>y.activity_type&&y.period),_=await v("/api/schedule/import",{method:"POST",body:JSON.stringify(f)});if(!_.ok)throw new Error(_.data?.error||"Import gagal")}}})}I();L();var Te=[],Oe=[];function wt(e,a){let r=String(e.status||"").toLowerCase();return a==="open"?r==="open":!1}async function Da(e,a){let r=a?a.get("dash_filter"):null,[t,n,l]=await Promise.all([v("/api/branches?all=1"),v(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`),v(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`)]);Te=(t.data?.data||[]).map(s=>({value:s.id,label:s.full_name})),Oe=(n.data?.data||[]).map(s=>({value:s.full_name,label:s.full_name}));let b=(l.data?.data||[]).filter(s=>s.role==="FC Spesialis").map(s=>({value:s.name,label:s.name})),h=(l.data?.data||[]).filter(s=>s.role==="Pelapor").map(s=>({value:s.name,label:s.name})),o=s=>s&&!Oe.find(g=>g.value===s)?[...Oe,{value:s,label:s}]:Oe,u=s=>s&&!b.find(g=>g.value===s)?[...b,{value:s,label:s}]:b,c=s=>s&&!h.find(g=>g.value===s)?[...h,{value:s,label:s}]:h,i=new Date().getFullYear(),d=Array.from({length:5},(s,g)=>String(i-g));$({container:e,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:s=>r?s.filter(g=>wt(g,r)):s,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:s=>`<span class="badge badge-secondary">${s}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:s=>`<span title="${s}">${s?.length>50?s.slice(0,50)+"\u2026":s}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:s=>`<span title="${s||""}">${s?.length>40?s.slice(0,40)+"\u2026":s||"-"}</span>`},{key:"status",label:"Status",render:s=>A(s)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:s=>window.formatDate(s)},{key:"day_count",label:"Hari",render:s=>s??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:Te},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:d}],formFields:s=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:s?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:s?.branch_id&&!Te.find(g=>g.value==s.branch_id)?[...Te,{value:s.branch_id,label:s.branch_name||s.branch_id}]:Te,createApi:{path:"/api/branches",field:"full_name"},value:s?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:s?.category},{name:"source",label:"Sumber Laporan",type:"select",options:[...c(s?.source),{value:"Lainnya",label:"Lainnya"}],value:s?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:s?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"select",options:o(s?.employee_name),value:s?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"select",options:u(s?.fc_specialist),value:s?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:s?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:s?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let s=await v(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(s.ok){let g=s.data.data.map(m=>({Tanggal:m.report_date||"",Cabang:m.branch_name||"",Kategori:m.category||"",Sumber:m.source||"",Keluhan:m.complaint||"","Nama FC":m.employee_name||"","FC Spesialis":m.fc_specialist||"",Solusi:m.solution||"","Tgl Selesai":m.completion_date||"",Status:m.status||""}));E(g,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async s=>{let m=(await v("/api/branches?all=1")).data?.data||[],p=y=>{if(!y)return null;let k=String(y||"").toLowerCase(),S=m.find(w=>String(w.full_name||"").toLowerCase()===k||String(w.code||"").toLowerCase()===k||String(w.name||"").toLowerCase()===k);return S?S.id:null},f=s.map(y=>({branch_id:p(String(y.Cabang||"").trim()),report_date:String(y.Tanggal||"").trim(),category:String(y.Kategori||"").trim(),source:String(y.Sumber||"").trim(),complaint:String(y.Keluhan||"").trim(),employee_name:String(y["Nama FC"]||"").trim(),fc_specialist:String(y["FC Spesialis"]||"").trim(),solution:String(y.Solusi||"").trim(),completion_date:String(y["Tgl Selesai"]||"").trim(),status:String(y.Status||"").trim()})).filter(y=>y.report_date&&y.complaint&&y.category),_=await v("/api/issues/import",{method:"POST",body:JSON.stringify(f)});if(!_.ok)throw new Error(_.data?.error||"Import gagal")}}})}I();function xt(e,a){let r=String(e.status||"").toLowerCase();return a==="pending"?r==="pending":!1}async function Ia(e,a){let r=a?a.get("dash_filter"):null,[t,n,l]=await Promise.all([v("/api/branches?all=1"),v(`/api/one_on_one${window.location.search?window.location.search+"&":"?"}limit=10000`),v(`/api/one_on_one${window.location.search?window.location.search+"&":"?"}limit=10000`)]),b=(t.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),h=(n.data?.data||[]).map(i=>({value:i.full_name,label:i.full_name})),o=(l.data?.data||[]).filter(i=>i.role==="FC Spesialis").map(i=>({value:i.name,label:i.name})),u=i=>i&&!h.find(d=>d.value===i)?[...h,{value:i,label:i}]:h,c=i=>i&&!o.find(d=>d.value===i)?[...o,{value:i,label:i}]:o;$({container:e,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:i=>r?i.filter(d=>xt(d,r)):i,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:i=>`<span title="${i||""}">${i?.length>50?i.slice(0,50)+"\u2026":i||"-"}</span>`},{key:"solution",label:"Solusi",render:i=>`<span title="${i||""}">${i?.length>40?i.slice(0,40)+"\u2026":i||"-"}</span>`},{key:"status",label:"Status",render:i=>A(i)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:i=>window.formatDate(i)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:i=>i?`<a href="${i}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:b},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async i=>{let d=new URLSearchParams(i||{}).toString(),s=await v(`/api/one-on-one?limit=10000&${d}`);if(s.ok){let g=s.data.data.map(p=>({Tanggal:p.meeting_date||"",Cabang:p.branch_name||"","Nama Karyawan":p.employee_name||"",PIC:p.pic||"",Masalah:p.problem||"",Solusi:p.solution||"",Status:p.status||"","Tgl Selesai":p.completion_date||"",Dokumen:p.document_link||""})),{downloadExcel:m}=await Promise.resolve().then(()=>(L(),G));m(g,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let i=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:d}=await Promise.resolve().then(()=>(L(),G));d(i,"Template_Import_OneOnOne")},onImport:async i=>{let d=p=>{if(!p)return null;let f=String(p||"").toLowerCase(),_=t.data?.data.find(y=>String(y.full_name||"").toLowerCase()===f||String(y.code||"").toLowerCase()===f||String(y.name||"").toLowerCase()===f);return _?_.id:null},s=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let f=String(p).trim();if(/^\d{4,5}$/.test(f)){let y=Number(f);if(y>2e4&&y<99999){let k=new Date(Date.UTC(1899,11,30)+y*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(f))return f.slice(0,10);let _=f.split(/[\/\-\.]/);if(_.length===3){let[y,k,S]=_.map(w=>w.trim());if(y.length===4&&k.length<=2&&S.length<=2)return`${y}-${k.padStart(2,"0")}-${S.padStart(2,"0")}`;if(S.length===4&&k.length<=2&&y.length<=2)return`${S}-${k.padStart(2,"0")}-${y.padStart(2,"0")}`}return f},g=i.map(p=>({meeting_date:s(p.Tanggal),employee_name:String(p["Nama Karyawan"]||"").trim(),branch_id:d(String(p.Cabang||"").trim()),pic:String(p.PIC||"").trim(),problem:String(p.Masalah||"").trim(),solution:String(p.Solusi||"").trim(),status:String(p.Status||"").trim(),completion_date:s(p["Tgl Selesai"]),document_link:String(p.Dokumen||"").trim()})).filter(p=>p.meeting_date&&p.employee_name&&p.branch_id),m=await v("/api/one-on-one/import",{method:"POST",body:JSON.stringify(g)});if(!m.ok)throw new Error(m.data?.error||"Import gagal")}},formFields:i=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:i?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:i?.branch_id&&!b.find(d=>d.value==i.branch_id)?[...b,{value:i.branch_id,label:i.branch_name||i.branch_id}]:b,createApi:{path:"/api/branches",field:"full_name"},value:i?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:u(i?.employee_name),value:i?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:c(i?.pic),createApi:{path:"/api/pic",field:"name"},value:i?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:i?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:i?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:i?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:i?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:i?.document_link}]})}I();async function Pa(e){let[a,r,t]=await Promise.all([v("/api/branches?all=1"),v(`/api/training${window.location.search?window.location.search+"&":"?"}limit=10000`),v(`/api/training${window.location.search?window.location.search+"&":"?"}limit=10000`)]),n=(a.data?.data||[]).map(c=>({value:c.id,label:c.full_name})),l=(r.data?.data||[]).map(c=>({value:c.full_name,label:c.full_name})),b=(t.data?.data||[]).filter(c=>c.role==="FC Spesialis").map(c=>({value:c.name,label:c.name})),h=c=>c&&!l.find(i=>i.value===c)?[...l,{value:c,label:c}]:l,o=c=>c&&!b.find(i=>i.value===c)?[...b,{value:c,label:c}]:b,u=Array.from({length:5},(c,i)=>String(new Date().getFullYear()-i));$({container:e,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:c=>window.formatDate(c)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:c=>{try{let i=JSON.parse(c);return Array.isArray(i)?i.join(", "):c||"-"}catch{return c||"-"}}},{key:"score",label:"Nilai",render:c=>c!=null?`<strong>${c}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:c=>c?`<a href="${c}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:u}],exportOptions:{moduleName:"training",onExport:async c=>{let i=new URLSearchParams(c||{}).toString(),d=await v(`/api/training?limit=10000&${i}`);if(d.ok){let s=d.data.data.map(m=>{let p=m.participants||"";try{let f=JSON.parse(p);p=Array.isArray(f)?f.join(", "):p}catch{}return{Tanggal:m.training_date||"",Batch:m.batch||"",Materi:m.subject||"",Cabang:m.branch_name||"",Trainer:m.trainer||"",Peserta:p,Nilai:m.score!==null&&m.score!==void 0?m.score:"",Dokumen:m.document_link||""}}),{downloadExcel:g}=await Promise.resolve().then(()=>(L(),G));g(s,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let c=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(L(),G));i(c,"Template_Import_Training")},onImport:async c=>{let i=m=>{if(!m)return null;let p=String(m||"").toLowerCase(),f=a.data?.data.find(_=>String(_.full_name||"").toLowerCase()===p||String(_.code||"").toLowerCase()===p||String(_.name||"").toLowerCase()===p);return f?f.id:null},d=m=>{if(!m)return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let p=String(m).trim();if(/^\d{4,5}$/.test(p)){let _=Number(p);if(_>2e4&&_<99999){let y=new Date(Date.UTC(1899,11,30)+_*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);let f=p.split(/[\/\-\.]/);if(f.length===3){let[_,y,k]=f.map(S=>S.trim());if(_.length===4&&y.length<=2&&k.length<=2)return`${_}-${y.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&y.length<=2&&_.length<=2)return`${k}-${y.padStart(2,"0")}-${_.padStart(2,"0")}`}return p},s=c.map(m=>({training_date:d(m.Tanggal),batch:String(m.Batch||"").trim(),subject:String(m.Materi||"").trim(),branch_id:i(String(m.Cabang||"").trim()),trainer:String(m.Trainer||"").trim(),participants:String(m.Peserta||"").trim(),score:m.Nilai?Number(m.Nilai):null,document_link:String(m.Dokumen||"").trim()})).filter(m=>m.training_date&&m.subject&&m.branch_id),g=await v("/api/training/import",{method:"POST",body:JSON.stringify(s)});if(!g.ok)throw new Error(g.data?.error||"Import gagal")}},formFields:c=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:c?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:c?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:c?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:c?.branch_id&&!n.find(i=>i.value==c.branch_id)?[...n,{value:c.branch_id,label:c.branch_name||c.branch_id}]:n,createApi:{path:"/api/branches",field:"full_name"},value:c?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:o(c?.trainer),createApi:{path:"/api/pic",field:"name"},value:c?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let i=JSON.parse(c?.participants);return Array.isArray(i)?i.join(", "):c?.participants||""}catch{return c?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:c?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:c?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:c?.notes}],onBeforeSubmit:async c=>(c.participants&&(c.participants=JSON.stringify(c.participants.split(",").map(i=>i.trim()).filter(Boolean))),c)})}I();L();async function La(e){let[a,r]=await Promise.all([v("/api/branches?all=1"),v(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`)]),t=(a.data?.data||[]).map(o=>({value:o.id,label:o.full_name})),n=(r.data?.data||[]).map(o=>({value:o.full_name,label:o.full_name})),l=o=>o&&!n.find(u=>u.value===o)?[...n,{value:o,label:o}]:n,b=["Krishna Aryaan Permana","Agung Septiadi","Indra Saputro","Wariskin","Iqbal"],h=o=>{let u=b.map(c=>({value:c,label:c}));return o&&!u.find(c=>c.value===o)?[...u,{value:o,label:o}]:u};$({container:e,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",columns:[{key:"backup_date",label:"Tanggal Backup",nowrap:!0,render:o=>window.formatDate(o)},{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"FC Digantikan"},{key:"period",label:"Periode",render:o=>Y(o)},{key:"reliever_name",label:"Reliefer"},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:o=>o?`<span class="badge badge-info">${o}</span>`:"-"},{key:"status",label:"Status",render:o=>A(o)}],filterFields:[{type:"search",placeholder:"Cari reliefer / FC..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:o?.branch_id&&!t.find(u=>u.value==o.branch_id)?[...t,{value:o.branch_id,label:o.branch_name||o.branch_id}]:t,createApi:{path:"/api/branches",field:"full_name"},value:o?.branch_id},{name:"period",label:"Periode",type:"select",options:["Q1","Q2","Q3","Q4"],value:o?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"FC yang Digantikan",type:"select",options:[{value:"",label:"BELUM ADA FC"},...l(o?.original_fc_name)],value:o?.original_fc_name},{name:"reliever_name",label:"Nama Reliefer",type:"select",required:!0,options:h(o?.reliever_name),value:o?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Backup",type:"date",required:!0,value:o?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:o?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"select",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:o?.reason},{name:"shift",label:"Shift",type:"select",options:["Pagi","Siang","Full Shift","Middle"],value:o?.shift}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done","Tidak Datang"],value:o?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let o=await v(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(o.ok){let u=o.data.data.map(c=>({"Tanggal Backup":c.backup_date||"",Cabang:c.branch_name||"","FC Digantikan":c.original_fc_name||"",Periode:c.period||"",Reliefer:c.reliever_name||"",Keterangan:c.reason||"",Shift:c.shift||"","Tanggal Selesai":c.completion_date||"",Status:c.status||""}));E(u,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Tanggal Backup":"2024-03-10",Cabang:"001. Pondok Bambu","FC Digantikan":"Budi Santoso",Periode:"Q1",Reliefer:"Andi",Keterangan:"Sakit",Shift:"Pagi","Tanggal Selesai":"2024-03-10",Status:"Done"}],"Template_Import_Reliefer")},onImport:async o=>{let c=(await v("/api/branches?all=1")).data?.data||[],i=g=>{if(!g)return null;let m=String(g||"").toLowerCase(),p=c.find(f=>String(f.full_name||"").toLowerCase()===m||String(f.code||"").toLowerCase()===m||String(f.name||"").toLowerCase()===m);return p?p.id:null},d=o.map(g=>({branch_id:i(String(g.Cabang||"").trim()),backup_date:String(g["Tanggal Backup"]||"").trim(),original_fc_name:String(g["FC Digantikan"]||"").trim(),reliever_name:String(g.Reliefer||"").trim(),period:String(g.Periode||"").trim(),reason:String(g.Keterangan||"").trim(),shift:String(g.Shift||"").trim(),completion_date:String(g["Tanggal Selesai"]||"").trim(),status:String(g.Status||"").trim()})).filter(g=>g.reliever_name&&g.backup_date),s=await v("/api/relievers/import",{method:"POST",body:JSON.stringify(d)});if(!s.ok)throw new Error(s.data?.error||"Import gagal")}}})}I();L();async function Ba(e){let r=((await v("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name})),t=Array.from({length:4},(n,l)=>String(new Date().getFullYear()-l));$({container:e,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:n=>Y(n)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"fc_score",label:"Point FC",render:n=>n!=null?`<strong class="${n>=80?"text-success":n>=60?"text-warning":"text-danger"}">${n}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"status",label:"Status",render:n=>A(n)},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"select",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:t}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:n?.branch_id&&!r.find(l=>l.value==n.branch_id)?[...r,{value:n.branch_id,label:n.branch_name||n.branch_id}]:r,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:n?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:n?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:n?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:n?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async n=>{let l=new URLSearchParams(n||{}).toString(),b=await v(`/api/reports/inspection?limit=10000&${l}`);if(b.ok){let h=b.data.data.map(o=>({Cabang:o.branch_name||"",Periode:o.period||"",Tanggal:o.inspection_date||"","Point FC":o.fc_score!==null&&o.fc_score!==void 0?o.fc_score:"","Point SPV":o.spv_score!==null&&o.spv_score!==void 0?o.spv_score:"",Status:o.status||"","Link Dokumen":o.document_link||""}));E(h,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Inspeksi")},onImport:async n=>{let b=(await v("/api/branches?all=1")).data?.data||[],h=i=>{if(!i)return null;let d=String(i||"").toLowerCase(),s=b.find(g=>String(g.full_name||"").toLowerCase()===d||String(g.code||"").toLowerCase()===d||String(g.name||"").toLowerCase()===d);return s?s.id:null},o=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let d=String(i).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let g=Number(d);if(g>2e4&&g<99999){let m=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}let s=d.split(/[\/\-\.]/);if(s.length===3){let[g,m,p]=s.map(f=>f.trim());if(g.length===4&&m.length<=2&&p.length<=2)return`${g}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&m.length<=2&&g.length<=2)return`${p}-${m.padStart(2,"0")}-${g.padStart(2,"0")}`}return d},u=n.map(i=>({branch_id:h(String(i.Cabang||"").trim()),period:String(i.Periode||"").trim(),inspection_date:o(i.Tanggal),fc_score:i["Point FC"]!==void 0&&i["Point FC"]!==""?Number(i["Point FC"]):null,spv_score:i["Point SPV"]!==void 0&&i["Point SPV"]!==""?Number(i["Point SPV"]):null,status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.period&&i.inspection_date),c=await v("/api/reports/inspection/import",{method:"POST",body:JSON.stringify(u)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}}})}I();L();async function Aa(e){let r=((await v("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name})),t=Array.from({length:4},(n,l)=>String(new Date().getFullYear()-l));$({container:e,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:n=>`<span class="badge ${n==="Deep Cleaning"?"badge-purple":"badge-success"}">${n}</span>`},{key:"period",label:"Periode",render:n=>Y(n)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"status",label:"Status",render:n=>A(n)},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:t}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:n?.branch_id&&!r.find(l=>l.value==n.branch_id)?[...r,{value:n.branch_id,label:n.branch_name||n.branch_id}]:r,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:n?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:n?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:n?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:n?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async n=>{let l=new URLSearchParams(n||{}).toString(),b=await v(`/api/reports/cleaning?limit=10000&${l}`);if(b.ok){let h=b.data.data.map(o=>({Cabang:o.branch_name||"",Jenis:o.activity_type||"",Periode:o.period||"",Tanggal:o.activity_date||"",Status:o.status||"","Link Dokumen":o.document_link||""}));E(h,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_GCDC")},onImport:async n=>{let b=(await v("/api/branches?all=1")).data?.data||[],h=i=>{if(!i)return null;let d=String(i||"").toLowerCase(),s=b.find(g=>String(g.full_name||"").toLowerCase()===d||String(g.code||"").toLowerCase()===d||String(g.name||"").toLowerCase()===d);return s?s.id:null},o=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let d=String(i).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let g=Number(d);if(g>2e4&&g<99999){let m=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}let s=d.split(/[\/\-\.]/);if(s.length===3){let[g,m,p]=s.map(f=>f.trim());if(g.length===4&&m.length<=2&&p.length<=2)return`${g}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&m.length<=2&&g.length<=2)return`${p}-${m.padStart(2,"0")}-${g.padStart(2,"0")}`}return d},u=n.map(i=>({branch_id:h(String(i.Cabang||"").trim()),activity_type:String(i.Jenis||i.Kegiatan||"").trim(),period:String(i.Periode||"").trim(),activity_date:o(i.Tanggal),status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.activity_type&&i.period&&i.activity_date),c=await v("/api/reports/cleaning/import",{method:"POST",body:JSON.stringify(u)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}}})}I();L();async function Na(e){let r=((await v("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name})),t=Array.from({length:4},(n,l)=>String(new Date().getFullYear()-l));$({container:e,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:n=>`<span class="badge badge-warning">${n}</span>`},{key:"period",label:"Periode",render:n=>Y(n)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"status",label:"Status",render:n=>A(n)},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:n=>n||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:t}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:n?.branch_id&&!r.find(l=>l.value==n.branch_id)?[...r,{value:n.branch_id,label:n.branch_name||n.branch_id}]:r,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:n?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:n?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:n?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async n=>{let l=new URLSearchParams(n||{}).toString(),b=await v(`/api/reports/fogging?limit=10000&${l}`);if(b.ok){let h=b.data.data.map(o=>({Cabang:o.branch_name||"",Jenis:o.activity_type||"Fogging",Periode:o.period||"",Tanggal:o.activity_date||"",Status:o.status||"","Link Dokumen":o.document_link||""}));E(h,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async n=>{let b=(await v("/api/branches?all=1")).data?.data||[],h=i=>{if(!i)return null;let d=String(i||"").toLowerCase(),s=b.find(g=>String(g.full_name||"").toLowerCase()===d||String(g.code||"").toLowerCase()===d||String(g.name||"").toLowerCase()===d);return s?s.id:null},o=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let d=String(i).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let g=Number(d);if(g>2e4&&g<99999){let m=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}let s=d.split(/[\/\-\.]/);if(s.length===3){let[g,m,p]=s.map(f=>f.trim());if(g.length===4&&m.length<=2&&p.length<=2)return`${g}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&m.length<=2&&g.length<=2)return`${p}-${m.padStart(2,"0")}-${g.padStart(2,"0")}`}return d},u=n.map(i=>({branch_id:h(String(i.Cabang||"").trim()),activity_type:String(i.Jenis||i.Kegiatan||"Fogging").trim(),period:String(i.Periode||"").trim(),activity_date:o(i.Tanggal),status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.period&&i.activity_date),c=await v("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(u)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}}})}I();L();async function Oa(e){let[a,r,t]=await Promise.all([v("/api/branches?all=1"),v(`/api/basecamp_reports${window.location.search?window.location.search+"&":"?"}limit=10000`),v(`/api/basecamp_reports${window.location.search?window.location.search+"&":"?"}limit=10000`)]),n=(a.data?.data||[]).map(u=>({value:u.id,label:u.full_name})),l=(r.data?.data||[]).map(u=>({value:u.full_name,label:u.full_name})),b=(t.data?.data||[]).filter(u=>u.role==="FC Spesialis").map(u=>({value:u.name,label:u.name})),h=u=>u&&!l.find(c=>c.value===u)?[...l,{value:u,label:u}]:l,o=u=>u&&!b.find(c=>c.value===u)?[...b,{value:u,label:u}]:b;$({container:e,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:u=>window.formatDate(u)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:u=>`<span title="${u||""}">${u?.length>60?u.slice(0,60)+"\u2026":u||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:u=>window.formatDate(u)},{key:"status",label:"Status",render:u=>A(u)},{key:"notes",label:"Keterangan",render:u=>u?.length>40?u.slice(0,40)+"\u2026":u||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan / PIC..."},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]}],formFields:u=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:u?.branch_id&&!n.find(c=>c.value==u.branch_id)?[...n,{value:u.branch_id,label:u.branch_name||u.branch_id}]:n,createApi:{path:"/api/branches",field:"full_name"},value:u?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:o(u?.pic),createApi:{path:"/api/pic",field:"name"},value:u?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:u?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:u?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:u?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:u?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:u?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async u=>{let c=new URLSearchParams(u||{}).toString(),i=await v(`/api/reports/basecamp?limit=10000&${c}`);if(i.ok){let d=i.data.data.map(s=>({"Tgl Info":s.info_date||"",Cabang:s.branch_name||"",Permasalahan:s.problem||"",PIC:s.pic||"","Tgl Done":s.done_date||"",Status:s.status||"",Keterangan:s.notes||""}));E(d,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async u=>{let i=(await v("/api/branches?all=1")).data?.data||[],d=p=>{if(!p)return null;let f=String(p||"").toLowerCase(),_=i.find(y=>String(y.full_name||"").toLowerCase()===f||String(y.code||"").toLowerCase()===f||String(y.name||"").toLowerCase()===f);return _?_.id:null},s=p=>{if(p==null||p==="")return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let f=String(p).trim();if(f===""||f==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(f))return f.slice(0,10);if(/^\d{4,5}$/.test(f)){let y=Number(f);if(y>2e4&&y<99999){let k=new Date(Date.UTC(1899,11,30)+y*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}let _=f.split(/[\/\-\.]/);if(_.length===3){let[y,k,S]=_.map(w=>w.trim());if(y.length===4&&k.length<=2&&S.length<=2)return`${y}-${k.padStart(2,"0")}-${S.padStart(2,"0")}`;if(S.length===4&&k.length<=2&&y.length<=2)return`${S}-${k.padStart(2,"0")}-${y.padStart(2,"0")}`}return f},g=u.map(p=>({info_date:s(p["Tgl Info"]||p["Tanggal Info"]),branch_id:d(String(p.Cabang||"").trim()),problem:String(p.Permasalahan||"").trim(),pic:String(p.PIC||"").trim(),done_date:s(p["Tgl Done"]||p["Tanggal Done"]),status:String(p.Status||"").trim(),notes:String(p.Keterangan||p.Catatan||"").trim()})).filter(p=>p.info_date&&p.branch_id&&p.problem),m=await v("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(g)});if(!m.ok)throw new Error(m.data?.error||"Import gagal")}}})}async function Fa(e){$({container:e,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),{apiFetch:t}=await Promise.resolve().then(()=>(I(),oe)),n=await t(`/api/sop?limit=10000&${r}`);if(n.ok){let l=n.data.data.map(h=>({"Nama SOP":h.name||"",Kategori:h.category||"",Dokumen:h.document_link||"",Catatan:h.notes||h.description||""})),{downloadExcel:b}=await Promise.resolve().then(()=>(L(),G));b(l,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let a=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:r}=await Promise.resolve().then(()=>(L(),G));r(a,"Template_Import_SOP")},onImport:async a=>{let r=a.map(l=>({name:String(l["Nama SOP"]||"").trim(),category:String(l.Kategori||"").trim(),document_link:String(l.Dokumen||"").trim(),description:String(l.Catatan||"").trim()})).filter(l=>l.name),{apiFetch:t}=await Promise.resolve().then(()=>(I(),oe)),n=await t("/api/sop/import",{method:"POST",body:JSON.stringify(r)});if(!n.ok)throw new Error(n.data?.error||"Import gagal")}},formFields:a=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:a?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:a?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:a?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:a?.description}]})}async function Ma(e){$({container:e,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async a=>{let r=new URLSearchParams(a||{}).toString(),{apiFetch:t}=await Promise.resolve().then(()=>(I(),oe)),n=await t(`/api/checklist?limit=10000&${r}`);if(n.ok){let l=n.data.data.map(h=>({"Nama Checklist":h.name||"",Kategori:h.category||"",Dokumen:h.document_link||"",Deskripsi:h.description||""})),{downloadExcel:b}=await Promise.resolve().then(()=>(L(),G));b(l,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let a=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:r}=await Promise.resolve().then(()=>(L(),G));r(a,"Template_Import_Checklist")},onImport:async a=>{let r=a.map(l=>({name:String(l["Nama Checklist"]||"").trim(),category:String(l.Kategori||"").trim(),document_link:String(l.Dokumen||"").trim(),description:String(l.Deskripsi||"").trim()})).filter(l=>l.name),{apiFetch:t}=await Promise.resolve().then(()=>(I(),oe)),n=await t("/api/checklist/import",{method:"POST",body:JSON.stringify(r)});if(!n.ok)throw new Error(n.data?.error||"Import gagal")}},formFields:a=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:a?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:a?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:a?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:a?.description}]})}I();Se();L();async function ca(e,a="forms"){if(a==="supply")return Tt(e);Ct(e)}function Ct(e){$({container:e,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:a=>a?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:a=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:a?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:a?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:a?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:a?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:a?.is_public}]})}async function Tt(e){let r=((await v("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name}));$({container:e,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:t=>t?new Date(t).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(t,n)=>n.branch_name_ref||n.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:t=>{try{let n=JSON.parse(t);return Array.isArray(n)?n.join(", "):t}catch{return t||"-"}}},{key:"chemical_items",label:"Chemical",render:t=>{try{let n=JSON.parse(t);return Array.isArray(n)?n.join(", "):t}catch{return t||"-"}}},{key:"additional_notes",label:"Catatan",render:t=>t?.length>40?t.slice(0,40)+"\u2026":t||"-"},{key:"status",label:"Status",render:t=>A(t)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:t=>{let n=t?.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let l=t?.chemical_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:t?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:t?.branch_id&&!r.find(b=>b.value==t.branch_id)?[...r,{value:t.branch_id,label:t.branch_name||t.branch_id}]:r,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:n},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:t?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:l},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:t?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:t?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:t?.status||""},{name:"processed_by",label:"Diproses Oleh",value:t?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async t=>{let n=new URLSearchParams(t||{}).toString(),l=await v(`/api/reports/supply?limit=10000&${n}`);if(l.ok){let b=l.data.data.map(h=>{let o=h.tools_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}let u=h.chemical_items;try{u=Array.isArray(JSON.parse(u))?JSON.parse(u).join(", "):u}catch{}return{Waktu:h.submitted_at||"",Pengirim:h.submitter_name||"",Cabang:h.branch_name_ref||h.branch_name||"","Alat/Barang":o||"",Chemical:u||"",Catatan:h.additional_notes||"",Status:h.status||"","Diproses Oleh":h.processed_by||""}});E(b,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async t=>{let l=(await v("/api/branches?all=1")).data?.data||[],b=c=>{if(!c)return null;let i=String(c||"").toLowerCase(),d=l.find(s=>String(s.full_name||"").toLowerCase()===i||String(s.code||"").toLowerCase()===i||String(s.name||"").toLowerCase()===i);return d?d.id:null},h=c=>{if(c==null||c==="")return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let i=String(c).trim();if(i===""||i==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(i))return i.slice(0,10);if(/^\d{4,5}$/.test(i)){let s=Number(i);if(s>2e4&&s<99999){let g=new Date(Date.UTC(1899,11,30)+s*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}let d=i.split(/[\/\-\.]/);if(d.length===3){let[s,g,m]=d.map(p=>p.trim());if(s.length===4&&g.length<=2&&m.length<=2)return`${s}-${g.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&g.length<=2&&s.length<=2)return`${m}-${g.padStart(2,"0")}-${s.padStart(2,"0")}`}return i},o=t.map(c=>({submitted_at:h(c.Waktu||c.Tanggal),submitter_name:String(c.Pengirim||"").trim(),branch_id:b(String(c.Cabang||"").trim()),tools_items:String(c["Alat/Barang"]||c.Alat||"").trim(),chemical_items:String(c.Chemical||"").trim(),additional_notes:String(c.Catatan||c.Keterangan||"").trim(),status:String(c.Status||"").trim(),processed_by:String(c["Diproses Oleh"]||c.PIC||"").trim()})).filter(c=>c.submitted_at&&c.submitter_name&&c.branch_id),u=await v("/api/reports/supply/import",{method:"POST",body:JSON.stringify(o)});if(!u.ok)throw new Error(u.data?.error||"Import gagal")}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(t,n)=>{let l=ee({title:"Update Status Permintaan",content:`
              <div class="form-group">
                <label class="form-label">Status</label>
                <select class="form-control" id="supply-status">
                  <option value="Pending" ${t.status==="Pending"?"selected":""}>Pending</option>
                  <option value="Diproses" ${t.status==="Diproses"?"selected":""}>Diproses</option>
                  <option value="Selesai" ${t.status==="Selesai"?"selected":""}>Selesai</option>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label">Diproses Oleh</label>
                <input type="text" class="form-control" id="supply-processed-by" value="${t.processed_by||""}" placeholder="Nama">
              </div>
            `,onConfirm:async(b,h)=>{let o=b.querySelector("#supply-status").value,u=b.querySelector("#supply-processed-by").value;(await v(`/api/reports/supply/${t.id}`,{method:"PUT",body:JSON.stringify({status:o,processed_by:u})})).ok?(j("Status diperbarui."),h(),n()):J("Gagal update status.")}})}}]})}I();L();async function Ra(e){let a=re();if(!a||!["superadmin","admin"].includes(a.role)){e.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}$({container:e,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:r=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[r]||"badge-neutral"}">${r}</span>`},{key:"is_active",label:"Status",render:r=>r?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:r=>r?new Date(r).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:r=>{let t=!!r;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:r?.full_name},{name:"username",label:"Username",required:!t,placeholder:"username",value:r?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!t,placeholder:"email@contoh.com",value:r?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:r?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:t?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!t,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:t?r?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let r=await v(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let t=r.data.data.map(n=>({"Nama Lengkap":n.full_name||"",Username:n.username||"",Email:n.email||"",Role:n.role||"",Status:n.is_active?"Aktif":"Nonaktif"}));E(t,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async r=>{let t=r.map(l=>({full_name:String(l["Nama Lengkap"]||"").trim(),username:String(l.Username||"").trim(),email:String(l.Email||"").trim(),role:String(l.Role||"").trim()||"viewer",password:String(l.Password||"").trim()})).filter(l=>l.username&&l.password&&l.email&&l.full_name),n=await v("/api/users/import",{method:"POST",body:JSON.stringify(t)});if(!n.ok)throw new Error(n.data?.error||"Import gagal")}}})}I();L();async function Ka(e){$({container:e,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:a=>a?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:a=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:a?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:a?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:a?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:a?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:a?.is_active!==void 0?a.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let a=await v(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok)E(a.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async a=>{let r=a.map(n=>({code:String(n["Kode Cabang"]||"").trim(),name:String(n["Nama Pendek"]||"").trim(),full_name:String(n["Nama Lengkap"]||"").trim(),city:String(n.Kota||"").trim()})).filter(n=>n.code&&n.name),t=await v("/api/branches/import",{method:"POST",body:JSON.stringify(r)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}}})}I();async function qa(e){let a=new Date,r=[];e.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{a.setMonth(a.getMonth()-1),n()}),document.getElementById("cal-next").addEventListener("click",()=>{a.setMonth(a.getMonth()+1),n()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(l=>l.addEventListener("change",n));async function t(){try{let l=`${a.getFullYear()}-${String(a.getMonth()+1).padStart(2,"0")}`;r=(await v(`/api/dashboard/calendar?month=${l}`)).data?.data||[]}catch(l){console.warn("[Calendar] Failed to load events, rendering empty grid:",l),r=[]}}async function n(){let l=document.getElementById("calendar-grid");if(l){l.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await t();try{let b=a.getFullYear(),h=a.getMonth(),o=a.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),u=document.getElementById("cal-month-label");u&&(u.textContent=o);let c=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(k=>k.value)),i=r.filter(k=>c.has(k.type)),d={};i.forEach(k=>{let S=(k.event_date||"").slice(0,10);d[S]||(d[S]=[]),d[S].push(k)});let s=new Date(b,h,1).getDay(),g=new Date(b,h+1,0).getDate(),m=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],p=new Date().toISOString().slice(0,10),f='<div class="calendar-grid">';m.forEach(k=>{f+=`<div class="cal-day-header">${k}</div>`});for(let k=0;k<s;k++)f+='<div class="cal-cell cal-cell-empty"></div>';for(let k=1;k<=g;k++){let S=`${b}-${String(h+1).padStart(2,"0")}-${String(k).padStart(2,"0")}`,w=d[S]||[],D=S===p;f+=`
          <div class="cal-cell ${D?"cal-today":""} ${w.length?"cal-has-events":""}"
               data-date="${S}" tabindex="0" role="button" aria-label="${S}">
            <div class="cal-day-num ${D?"today-num":""}">${k}</div>
            <div class="cal-events-preview">
              ${w.slice(0,3).map(C=>`
                <div class="cal-event-dot cal-color-${C.color||"gray"}" title="${Fe(C.title||C.type)}">
                  <span class="cal-event-dot-label">${$t(C.title||C.branch_name||C.type,18)}</span>
                </div>
              `).join("")}
              ${w.length>3?`<div class="cal-more">+${w.length-3} lagi</div>`:""}
            </div>
          </div>`}let y=(s+g)%7;if(y!==0)for(let k=0;k<7-y;k++)f+='<div class="cal-cell cal-cell-empty"></div>';f+="</div>",l.innerHTML=f,l.querySelectorAll(".cal-cell[data-date]").forEach(k=>{k.addEventListener("click",()=>{let S=k.dataset.date,w=d[S]||[];if(!w.length)return;let D=document.getElementById("cal-event-list"),C=new Date(S+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=C,document.getElementById("cal-event-items").innerHTML=w.map(P=>`
            <div class="cal-event-item cal-color-border-${P.color||"gray"}">
              <div class="cal-event-type">${Et(P.type)}</div>
              <div class="cal-event-title">${Fe(P.title||"-")}</div>
              <div class="cal-event-branch">${Fe(P.branch_name||"")}</div>
              ${P.status?`<div class="cal-event-status">${Fe(P.status)}</div>`:""}
              ${P.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${P.days_remaining} hari</div>`:""}
            </div>
          `).join(""),D.style.display="block"})})}catch(b){console.error("[Calendar] Render error:",b),l&&(l.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}n()}function $t(e,a){return e?e.length>a?e.slice(0,a)+"\u2026":e:""}function Fe(e){return e?String(e).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Et(e){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[e]||e}I();async function Ha(e){let a=re(),r=(a?.full_name||a?.username||"U")[0].toUpperCase(),n={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[a?.role]||"#64748B";e.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${n},${n}99)">
            ${r}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${a?.full_name||"\u2014"}</div>
            <div class="profile-username">@${a?.username||"\u2014"}</div>
            <span class="badge badge-info" style="background:${n}18;color:${n};margin-top:6px">
              ${(a?.role||"viewer").toUpperCase()}
            </span>
          </div>
        </div>

        <hr class="profile-divider">

        <div class="info-list">
          <div class="info-row">
            <span class="info-key">\u{1F4E7} Email</span>
            <span class="info-value">${a?.email||"\u2014"}</span>
          </div>
          <div class="info-row">
            <span class="info-key">\u{1F464} Username</span>
            <span class="info-value">${a?.username||"\u2014"}</span>
          </div>
          <div class="info-row">
            <span class="info-key">\u{1F3AF} Role</span>
            <span class="info-value" style="color:${n};font-weight:700">${a?.role||"\u2014"}</span>
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
  `;let l=localStorage.getItem("fm_token"),b=document.getElementById("session-info");if(l&&b)try{let h=JSON.parse(atob(l.split(".")[1])),o=new Date(h.exp*1e3);b.textContent=`Berakhir: ${o.toLocaleString("id-ID")}`}catch{b.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async h=>{h.preventDefault();let o=document.getElementById("pwd-error"),u=document.getElementById("pwd-success"),c=document.getElementById("btn-save-pwd");o.style.display="none",u.style.display="none";let i=h.target,d=i.current_password.value,s=i.new_password.value,g=i.confirm_password.value;if(s!==g){o.textContent="\u274C Konfirmasi password tidak cocok.",o.style.display="block";return}if(s.length<6){o.textContent="\u274C Password baru minimal 6 karakter.",o.style.display="block";return}c.disabled=!0,c.textContent="\u23F3 Menyimpan...";let m=await v("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:d,new_password:s})});c.disabled=!1,c.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',m.ok?(u.textContent="\u2705 Password berhasil diubah.",u.style.display="block",i.reset(),j("Password berhasil diubah.")):(o.textContent=m.data?.error||"Gagal mengubah password.",o.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}I();var Me={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function K(e){if(e==null||e==="")return null;if(e instanceof Date)return isNaN(e.getTime())?null:e.toISOString().slice(0,10);let a=String(e).trim();if(a===""||a==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(a))return a.slice(0,10);if(/^\d{4,5}$/.test(a)){let n=Number(a);if(n>2e4&&n<99999){let l=new Date(Date.UTC(1899,11,30)+n*864e5);return isNaN(l.getTime())?null:l.toISOString().slice(0,10)}}let r=a.split(/[\/\-\.]/);if(r.length===3){let[n,l,b]=r.map(c=>c.trim()),h=Number(n),o=Number(l),u=Number(b);if(n.length===4&&h>1900)return`${n}-${l.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&u>1900)return h>12?`${b}-${l.padStart(2,"0")}-${n.padStart(2,"0")}`:o>12?`${b}-${n.padStart(2,"0")}-${l.padStart(2,"0")}`:`${b}-${l.padStart(2,"0")}-${n.padStart(2,"0")}`;if(b.length===2&&!isNaN(u)){let c=u>=50?`19${b}`:`20${b}`;return h>12?`${c}-${l.padStart(2,"0")}-${n.padStart(2,"0")}`:`${c}-${l.padStart(2,"0")}-${n.padStart(2,"0")}`}}let t=new Date(a);return isNaN(t.getTime())?null:t.toISOString().slice(0,10)}function ja(e){return Object.values(e).every(a=>a==null||String(a).trim()==="")}var Dt={validation:{required:[],map:e=>({cabang:e.CABANG,pic:e.PIC,kegiatan:e.KEGIATAN,quartal:e.QUARTAL,masa_pkwt:e["MASA PKWT"],pic_pelapor:e["PIC PELAPOR"],kontrak:e.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:e=>({name:e["Nama SOP"],category:e.Kategori||"Umum",document_link:e["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({full_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||"FACILITY CARE",phone:e["No. Hp"],join_date:K(e["Tanggal Masuk"]),status:e.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:e=>({employee_name:e["Nama Lengkap"],branch_name:e.Cabang,division:e["Div / Bagian"]||"FACILITY CARE",start_date:K(e["Tanggal Mulai"]),end_date:K(e["Tanggal Selesai"]),contract_type:e["Tipe Kontrak"]||"",pkwt_number:e.PKWT||"",status:e.Status||"",notes:e.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:e=>({report_date:K(e["Tanggal Info"]),branch_name:e.Cabang,category:e.Kategori,source:e["Sumber Laporan"],complaint:e.Keluhan,employee_name:e["Nama FC"],fc_specialist:e["FC Spesialis"],solution:e.Solusi,status:e.Status||"",completion_date:K(e["Tanggal Selesai"])})},one_on_one:{required:[],map:e=>({meeting_date:K(e.Tanggal),branch_name:e.Cabang,employee_name:e["Nama Karyawan"],pic:e.Pic,problem:e.Masalah,solution:e.Solusi,status:e.Status||"",completion_date:K(e["Tanggal Selesai"]),document_link:e["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:e=>({branch_name:e.Cabang,activity_type:e.Kegiatan,period:e.Periode,pic:e.Pic||e.PIC,opening_date:K(e["Tanggal Opening"]||e["Tgl Opening"]),target_date:K(e["Tanggal Target"]||e["Tgl Target"]),completion_date:K(e["Tanggal Selesai"]||e["Tgl Selesai"]),status:e.Status||"",notes:e.Keterangan||e.Catatan})},inspection:{required:[],map:e=>({inspection_date:K(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",fc_score:e["Point FC SP"]!==void 0&&e["Point FC SP"]!==null?parseFloat(String(e["Point FC SP"]).replace(",",".")):null,spv_score:e["Point SPV"]!==void 0&&e["Point SPV"]!==null?parseFloat(String(e["Point SPV"]).replace(",",".")):null,document_link:e.Link,notes:""})},cleaning:{required:[],map:e=>({activity_date:K(e.Tanggal),branch_name:e.Cabang,activity_type:e["Jenis Kegiatan"]||"General Cleaning",period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},fogging:{required:[],map:e=>({activity_date:K(e.Tanggal),branch_name:e.Cabang,period:e.Periode,status:e.Status||"",document_link:e.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:e=>({info_date:K(e["Tgl Info"]),branch_name:e.Cabang,problem:e.Permasalahan,pic:e.PIC,done_date:K(e["Tgl Done"]),status:e.Status||"",notes:e.Ket})},relievers:{required:[],map:e=>({branch_name:e.Cabang,original_fc_name:e["Nama Facility care"],period:e.Periode,reliever_name:e.Relifer,backup_date:K(e["Tanggal Back Up"]),completion_date:K(e["Tanggal Selesai"]),reason:e.Keterangan,shift:e.Shift,status:e.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:e=>({training_date:K(e.Tanggal),batch:e.Batch,subject:e.Materi,participants:e.Peserta,branch_name:e.Cabang,trainer:e.Trainer,score:e.Nilai!==void 0&&e.Nilai!==null?parseFloat(String(e.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:e=>({name:e["Master Checklist"],category:"Umum",document_link:e["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:e=>({name:e["Master Form"],category:"Umum",document_link:e["Link Document"],description:""})},supply:{required:[],map:e=>({submitted_at:K(e.Timestamp),submitter_name:e["Nama Lengkap"],branch_name:e["Kebutuhan Untuk Cabang"],tools_items:e["Alat - Alat / Barang"],tools_quantity:e["Jumlah Permintaan Alat / Barang"],chemical_items:e.Chemical,chemical_quantity:e["Jumlah Permintaan Chemical"],additional_notes:e["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:e.Status||""})}};function It(e,a){let r=Me[e];if(!r)return{valid:[],errors:[],mapped:[],skipped:!0};let t=Dt[r.module];if(!t)return{valid:[],errors:[],mapped:[],skipped:!0};let n=[],l=[],b=[];return a.filter(o=>!ja(o)).forEach((o,u)=>{let c=a.indexOf(o)+2,i=[];t.required.forEach(({key:s,label:g})=>{let m=o[s];if(m==null||String(m).trim()===""){let p=Object.keys(o).filter(f=>f.trim()).join(", ");i.push({column:g,originalValue:m||"",reason:`Kolom "${g}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${p.slice(0,120)}`})}});let d=t.map(o);i.length>0?l.push({row:c,data:d,raw:o,errors:i}):(n.push(o),b.push(d))}),{valid:n,errors:l,mapped:b}}function Ua(e){let a=[];return e.SheetNames.forEach(r=>{let t=Me[r];if(!t)return;let n=e.Sheets[r],l=window.XLSX.utils.sheet_to_json(n,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),b=It(r,l),h=l.filter(o=>!ja(o));a.push({sheetName:r,module:t.module,label:t.label,total:h.length,valid:b.mapped.length,errorCount:b.errors.length,errors:b.errors,mapped:b.mapped,skipped:!1})}),a}function Ja(){let e=window.XLSX,a=e.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu","NAMA KARYAWAN":"Budi Santoso",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02",Day:"1"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06",Day:"1","Link Document":""}]}).forEach(([t,n])=>{e.utils.book_append_sheet(a,e.utils.json_to_sheet(n),t)}),e.writeFile(a,"Template_Import_Data_Awal_FCMS.xlsx")}function Ga(e){let a=window.XLSX,r=a.utils.book_new(),t=!1;return e.forEach(n=>{if(!n.errors||n.errors.length===0)return;t=!0;let l=n.errors.map(h=>({"No. Baris":h.row,"Kolom Gagal":(h.errors||[]).map(o=>o.column||o).join("; "),"Alasan Error":(h.errors||[]).map(o=>o.reason||o).join("; "),...Object.fromEntries(Object.entries(h.data||{}).map(([o,u])=>[o,u??""]))})),b=a.utils.json_to_sheet(l);a.utils.book_append_sheet(r,b,n.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),t?(a.writeFile(r,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Pt=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function Qa(e){e.innerHTML=`
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
              ${Object.entries(Me).map(([m,{label:p}])=>`<span class="import-sheet-tag">\u{1F4C4} ${m} \u2192 ${p}</span>`).join("")}
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
  `;let a=null,r=null,t=0,n={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function l(m){Object.entries(n).forEach(([p,f])=>{f.style.display=p===m?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let m=document.getElementById("btn-backup-db");m.disabled=!0,m.textContent="\u23F3 Memproses Backup...";try{let p=await v("/api/import/backup");if(p.ok){let f=new Blob([JSON.stringify(p.data,null,2)],{type:"application/json"}),_=URL.createObjectURL(f),y=document.createElement("a");y.href=_,y.download=`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.json`,document.body.appendChild(y),y.click(),document.body.removeChild(y),URL.revokeObjectURL(_),j("Backup berhasil diunduh!")}else J("Gagal memproses backup: "+(p.data?.error||"Unknown error"))}catch(p){J("Gagal memproses backup: "+p.message)}finally{m.disabled=!1,m.textContent="\u{1F4E6} Backup Database"}});let b=document.getElementById("btn-sync-google");b&&b.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let m=b.innerHTML;b.innerHTML='<span class="spinner"></span> Menyinkronkan...',b.disabled=!0;try{let p=await v("/api/sync/google-sheets",{method:"POST"});p.ok?alert("Sinkronisasi Berhasil: "+(p.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(p.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{b.innerHTML=m,b.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Ja(),j("Template Excel berhasil didownload!")});let h=document.getElementById("file-input"),o=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",m=>{m.stopPropagation(),h.click()}),h.addEventListener("change",m=>{m.target.files[0]&&u(m.target.files[0])}),o.addEventListener("dragover",m=>{m.preventDefault(),o.classList.add("drag-over")}),o.addEventListener("dragleave",()=>o.classList.remove("drag-over")),o.addEventListener("drop",m=>{m.preventDefault(),o.classList.remove("drag-over");let p=m.dataTransfer.files[0];p&&p.name.match(/\.xlsx?$/i)?u(p):J("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{a=null,h.value="",document.getElementById("file-info").style.display="none",o.style.display="",l("upload")});async function u(m){a=m,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${m.name} (${(m.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",o.style.display="none",await c(m)}async function c(m){l("validating");let p=document.getElementById("validation-status"),f=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");p.textContent="Membaca file Excel...",f.style.width="20%",await $e(200);let _=await m.arrayBuffer(),y=window.XLSX.read(_,{type:"array",cellDates:!0});p.textContent=`Memvalidasi ${y.SheetNames.length} sheet...`,f.style.width="50%",await $e(100),r=Ua(y),f.style.width="100%",p.textContent="Validasi selesai!",await $e(300),i()}catch(_){l("upload"),J("Gagal memproses file: "+_.message),document.getElementById("file-info").style.display="flex",o.style.display="none"}}function i(){l("preview");let m=r.filter(C=>!C.skipped).length,p=r.reduce((C,P)=>C+P.total,0),f=r.reduce((C,P)=>C+P.valid,0),_=r.reduce((C,P)=>C+P.errorCount,0),y=p>0?Math.round(f/p*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${m} sheet</span>
      <span class="badge badge-secondary">${p} baris</span>
      <span class="badge badge-success">${f} valid (${y}%)</span>
      ${_>0?`<span class="badge badge-danger">${_} error</span>`:""}
    `;let k=document.getElementById("preview-table-container");k.innerHTML=`
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
    `,k.querySelectorAll(".btn-detail-error").forEach(C=>{C.addEventListener("click",()=>{let P=r[Number(C.dataset.idx)];d(P)})});let S=document.getElementById("error-detail-section"),w=document.getElementById("error-detail-container");w.innerHTML="",S.style.display="none";let D=document.getElementById("btn-start-import");f===0?(D.disabled=!0,D.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(D.disabled=!1,_>0?(D.innerHTML=`\u{1F680} Import ${f} Data Valid (${_} dilewati)`,D.title="Baris error akan dilewati, baris valid tetap diimport"):D.innerHTML=`\u{1F680} Mulai Import ${f} Data`)}function d(m){let p=document.getElementById("error-detail-section"),f=document.getElementById("error-detail-container");p.style.display="";let _=m.errors.slice(0,100).map(y=>(Array.isArray(y.errors)?y.errors:[]).map(S=>{let w=typeof S=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${y.row}</span></td>
            <td><strong>${w?S.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${w&&S.originalValue!==void 0?S.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${w?S.reason:S}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${w&&S.aliases?`Gunakan salah satu nama kolom:<br><em>${S.aliases}</em>`:w&&S.hint?S.hint:""}
            </td>
          </tr>
        `}).join("")).join("");f.innerHTML=`
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
            <tbody>${_||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${m.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,p.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{l("upload"),document.getElementById("file-info").style.display="none",o.style.display="",a=null,h.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!r)return;Ga(r)?j("Log error berhasil didownload."):j("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let m=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";s(m)});async function s(m){l("importing"),t=Date.now();let p=[];Pt.forEach(S=>{let w=r?.find(D=>D.module===S&&D.mapped?.length>0);w&&p.push(w)});let f=document.getElementById("import-steps-list");f.innerHTML=p.map(S=>`
      <div class="import-step-item" id="step-item-${S.module}">
        <span class="step-item-icon" id="step-icon-${S.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${S.label} <span class="step-item-count">(${S.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${S.module}"></span>
      </div>
    `).join("");let _=document.getElementById("import-bar"),y=document.getElementById("import-current-status"),k={totalSheets:p.length,totalRows:p.reduce((S,w)=>S+w.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let S=0;S<p.length;S++){let w=p[S],D=document.getElementById(`step-icon-${w.module}`),C=document.getElementById(`step-status-${w.module}`);D.textContent="\u{1F504}",C.textContent="Mengimport...",y.textContent=`Mengimport ${w.label}...`,_.style.width=`${Math.round(S/p.length*100)}%`;try{let P=await v(`/api/import/${w.module}`,{method:"POST",body:JSON.stringify({rows:w.mapped,onDuplicate:m})});if(P.ok){let O=P.data;k.inserted+=O.inserted||0,k.skipped+=O.skipped||0,k.moduleResults.push({label:w.label,inserted:O.inserted||0,skipped:O.skipped||0,status:"ok"}),D.textContent="\u2705",C.innerHTML=`<span class="badge badge-success">${O.inserted||0} berhasil</span>${O.skipped>0?` <span class="badge badge-neutral">${O.skipped} skip</span>`:""}`}else k.failed++,k.moduleResults.push({label:w.label,inserted:0,skipped:0,status:"error",error:P.data?.error}),D.textContent="\u274C",C.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(P){k.failed++,k.moduleResults.push({label:w.label,inserted:0,skipped:0,status:"error",error:P.message}),D.textContent="\u274C",C.innerHTML='<span class="badge badge-danger">Gagal</span>'}await $e(150)}_.style.width="100%",y.textContent="Selesai!",await $e(400),g(k)}function g(m){l("summary");let p=((Date.now()-t)/1e3).toFixed(1),f=m.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${f?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${f?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
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
        <div class="stat-value">${p}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${m.moduleResults.map(_=>`
            <tr>
              <td>${_.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${_.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${_.skipped}</span></td>
              <td style="text-align:center">
                ${_.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${_.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{a=null,r=null,h.value="",document.getElementById("file-info").style.display="none",o.style.display="",l("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function $e(e){return new Promise(a=>setTimeout(a,e))}I();var da=[],Lt=[];async function Va(e){let[a,r]=await Promise.all([v("/api/branches?all=1"),v(`/api/sp_data${window.location.search?window.location.search+"&":"?"}limit=10000`)]);da=(a.data?.data||[]).map(t=>({value:t.id,label:t.full_name})),Lt=(r.data?.data||[]).map(t=>({value:t.full_name,label:t.full_name})),$({container:e,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:t=>t?`<span class="badge badge-info">${t}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:t=>t?new Date(t).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:t=>t?new Date(t).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:t=>`<span class="badge badge-warning">${t||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:t=>t?`<a href="${t}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:da}],exportOptions:{moduleName:"sp_data",onExport:async t=>{let n=new URLSearchParams(t||{}).toString(),l=await v(`/api/sp?limit=10000&${n}`);if(l.ok){let b=l.data.data.map(o=>({"Nama Karyawan":o.employee_name||"",Divisi:o.division||"",Cabang:o.branch_name||"","Tanggal Sp":o.tanggal||"","Akhir Sp":o.akhir_sp||"","Jenis Sp":o.sp_type||"","Link Document / Foto":o.document_link||""})),{downloadExcel:h}=await Promise.resolve().then(()=>(L(),G));h(b,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let t=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(L(),G));n(t,"Template_Import_SP")},onImport:async t=>{let l=(await v("/api/branches?all=1")).data?.data||[],b=c=>{if(!c)return null;let i=String(c||"").toLowerCase(),d=l.find(s=>String(s.full_name||"").toLowerCase()===i||String(s.code||"").toLowerCase()===i||String(s.name||"").toLowerCase()===i);return d?d.id:null},h=c=>{if(!c)return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let i=String(c).trim();if(/^\d{4,5}$/.test(i)){let s=Number(i);if(s>2e4&&s<99999){let g=new Date(Date.UTC(1899,11,30)+s*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(i))return i.slice(0,10);let d=i.split(/[\/\-\.]/);if(d.length===3){let[s,g,m]=d.map(p=>p.trim());if(s.length===4&&g.length<=2&&m.length<=2)return`${s}-${g.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&g.length<=2&&s.length<=2)return`${m}-${g.padStart(2,"0")}-${s.padStart(2,"0")}`}return i},o=t.map(c=>({employee_name:String(c["Nama Karyawan"]||"").trim(),division:String(c.Divisi||"").trim(),branch_id:b(String(c.Cabang||"").trim()),tanggal:h(c["Tanggal Sp"]),akhir_sp:h(c["Akhir Sp"]),sp_type:String(c["Jenis Sp"]||"").trim(),document_link:String(c["Link Document / Foto"]||"").trim()})).filter(c=>c.employee_name&&c.branch_id),u=await v("/api/sp/import",{method:"POST",body:JSON.stringify(o)});if(!u.ok)throw new Error(u.data?.error||"Import gagal")}},formFields:[{type:"text",name:"employee_name",label:"Nama Karyawan",required:!0},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:da},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}I();var Ee=[],Bt=[];async function Ya(e){let[a,r]=await Promise.all([v("/api/branches?all=1"),v(`/api/mutasi_data${window.location.search?window.location.search+"&":"?"}limit=10000`)]);Ee=(a.data?.data||[]).map(t=>({value:t.id,label:t.full_name})),Bt=(r.data?.data||[]).map(t=>({value:t.full_name,label:t.full_name})),$({container:e,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:t=>t?new Date(t).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:t=>`<span class="badge ${t==="Selesai"?"badge-success":"badge-warning"}">${t||"-"}</span>`},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:Ee},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:Ee}],exportOptions:{moduleName:"mutasi_data",onExport:async t=>{let n=new URLSearchParams(t||{}).toString(),l=await v(`/api/mutasi?limit=10000&${n}`);if(l.ok){let b=l.data.data.map(o=>({Tanggal:o.tanggal||"","Nama Karyawan":o.employee_name||"","Cabang Asal":o.from_branch_name||"","Cabang Tujuan":o.to_branch_name||"",Status:o.status||"",Dokumen:o.document_link||""})),{downloadExcel:h}=await Promise.resolve().then(()=>(L(),G));h(b,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let t=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(L(),G));n(t,"Template_Import_Mutasi")},onImport:async t=>{let l=(await v("/api/branches?all=1")).data?.data||[],b=c=>{if(!c)return null;let i=String(c||"").toLowerCase(),d=l.find(s=>String(s.full_name||"").toLowerCase()===i||String(s.code||"").toLowerCase()===i||String(s.name||"").toLowerCase()===i);return d?d.id:null},h=c=>{if(!c)return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let i=String(c).trim();if(/^\d{4,5}$/.test(i)){let s=Number(i);if(s>2e4&&s<99999){let g=new Date(Date.UTC(1899,11,30)+s*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(i))return i.slice(0,10);let d=i.split(/[\/\-\.]/);if(d.length===3){let[s,g,m]=d.map(p=>p.trim());if(s.length===4&&g.length<=2&&m.length<=2)return`${s}-${g.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&g.length<=2&&s.length<=2)return`${m}-${g.padStart(2,"0")}-${s.padStart(2,"0")}`}return i},o=t.map(c=>({tanggal:h(c.Tanggal),employee_name:String(c["Nama Karyawan"]||"").trim(),from_branch_id:b(String(c["Cabang Asal"]||"").trim()),to_branch_id:b(String(c["Cabang Tujuan"]||"").trim()),status:String(c.Status||"").trim(),document_link:String(c.Dokumen||"").trim()})).filter(c=>c.tanggal&&c.employee_name&&c.from_branch_id&&c.to_branch_id),u=await v("/api/mutasi/import",{method:"POST",body:JSON.stringify(o)});if(!u.ok)throw new Error(u.data?.error||"Import gagal")}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"text",name:"employee_name",label:"Nama Karyawan",required:!0},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Ee,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Ee,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}window.parseFlexibleDate=e=>{if(!e||e==="-")return"";if(e=String(e).trim(),/^\d{5}$/.test(e)){let a=Math.floor(Number(e)-25569);return new Date(a*86400*1e3).toISOString().split("T")[0]}if(e.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let a=e.split(/[\/\-]/);return`${a[2]}-${a[1]}-${a[0]}`}return e.split("T")[0]};window.formatDate=e=>{let a=window.parseFlexibleDate(e);if(!a)return"";let r=a.split("-");if(r.length===3&&r[0].length===4){let t=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],n=parseInt(r[2],10),l=t[parseInt(r[1],10)-1];return`${n} ${l} ${r[0]}`}return a};function M(e){return async a=>{if(!fe()){ce("/login");return}return e(a)}}var De=null;function At(){De&&clearInterval(De);let e=()=>{let a=new Date,r=a.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),t=a.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),n=document.getElementById("header-clock-time"),l=document.getElementById("header-clock-date");n&&(n.textContent=r),l&&(l.textContent=t)};e(),De=setInterval(e,1e3)}async function Nt(){try{let e=await v("/api/dashboard/kpi");if(!e.ok)return;let a=e.data?.data||e.data||{},r=(t,n)=>{let l=document.getElementById(t);l&&(l.textContent=n>0?n:"",l.style.display=n>0?"inline-flex":"none")};r("badge-issues",a.issues?.current||0),r("badge-contracts",a.expiring30?.current||0),r("badge-oo1",a.one_on_one?.current||0),r("badge-schedule",a.schedule?.current||0),r("badge-supply",a.supply?.current||0)}catch{}}var ue=[];async function Ot(){try{let e=await v("/api/dashboard/notifications");if(!e.ok)return;ue=e.data?.data||e.data||[];let a=document.getElementById("notif-dot");a&&(a.style.display=ue.length>0?"block":"none",a.textContent=ue.length)}catch{}}function Ft(){if(!ue.length){ee({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(a,r)=>r()});return}let e=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${ue.map(a=>`
        <div class="notif-item notif-severity-${a.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${a.severity==="danger"?"danger":a.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${a.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${a.date}</span>
            <span class="badge badge-${a.severity==="danger"?"danger":a.severity==="warning"?"warning":"info"}">${a.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;ee({title:`Notifikasi (${ue.length})`,content:e,confirmText:"Tutup",onConfirm:(a,r)=>r()})}function Wa(){let e=re(),a=(e?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
                <span class="topbar-greeting-time">${(()=>{let u=new Date().getHours();return u>=4&&u<11?"Selamat Pagi":u>=11&&u<15?"Selamat Siang":u>=15&&u<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">Berlin Ariansyah</span> \u{1F44B}
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
  `;let r=document.getElementById("sidebar"),t=document.getElementById("sidebar-overlay"),n=document.getElementById("topbar-menu-btn"),l=document.getElementById("sidebar-close"),b=()=>{r.classList.add("open"),t.classList.add("show")},h=()=>{r.classList.remove("open"),t.classList.remove("show")};n?.addEventListener("click",b),l?.addEventListener("click",h),t?.addEventListener("click",h),document.querySelectorAll(".nav-item").forEach(u=>u.addEventListener("click",h));function o(){let u=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(d=>{let s=d.dataset.route;d.classList.toggle("active",u===s||s!=="/dashboard"&&u.startsWith(s))});let c=document.getElementById("topbar-title"),i=document.querySelector(".nav-item.active .nav-label");c&&i&&(c.textContent=i.textContent)}window.addEventListener("hashchange",o),o(),At(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await v("/api/auth/logout",{method:"POST"}),ye(),De&&clearInterval(De),ce("/login")}),Nt(),Ot(),document.getElementById("btn-notif")?.addEventListener("click",u=>{u.preventDefault(),Ft()})}async function Mt(){N("/login",({main:a})=>Ta(a)),N("/dashboard",M(({main:a})=>wa(a))),N("/calendar",M(({main:a})=>qa(a))),N("/employees",M(({main:a,params:r})=>$a(a,r))),N("/contracts",M(({main:a,params:r})=>Ne(a,r))),N("/sp",M(({main:a})=>Va(a))),N("/mutasi",M(({main:a})=>Ya(a))),N("/timeline",M(({main:a,params:r})=>Ea(a,r))),N("/issues",M(({main:a,params:r})=>Da(a,r))),N("/one-on-one",M(({main:a,params:r})=>Ia(a,r))),N("/training",M(({main:a})=>Pa(a))),N("/relievers",M(({main:a})=>La(a))),N("/reports/inspection",M(({main:a})=>Ba(a))),N("/reports/cleaning",M(({main:a})=>Aa(a))),N("/reports/fogging",M(({main:a})=>Na(a))),N("/reports/basecamp",M(({main:a})=>Oa(a))),N("/reports/supply",M(({main:a})=>ca(a,"supply"))),N("/sop",M(({main:a})=>Fa(a))),N("/checklist",M(({main:a})=>Ma(a))),N("/forms",M(({main:a})=>ca(a))),N("/users",M(({main:a})=>Ra(a))),N("/branches",M(({main:a})=>Ka(a))),N("/profile",M(({main:a})=>Ha(a))),N("/settings/import",M(({main:a})=>Qa(a)));let e=fe();if(!e&&window.location.hash!=="#/login"&&ce("/login"),e){let a=await v("/api/auth/me");a.ok?(ve(a.data.data),Wa()):(ye(),ce("/login"))}window.addEventListener("fm:login",()=>{Wa(),ce("/dashboard")}),ma()}Mt();
