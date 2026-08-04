var sa=Object.defineProperty;var Je=(t,e)=>()=>(t&&(e=t(t=0)),e);var Ge=(t,e)=>{for(var i in e)sa(t,i,{get:e[i],enumerable:!0})};var ge={};Ge(ge,{API:()=>ut,CLIENT_SIDE_MAX_ROWS:()=>Te,IS_DEVELOPMENT:()=>Qe,apiFetch:()=>f,clearToken:()=>$e,getToken:()=>Ee,getUser:()=>ue,setToken:()=>Ve,setUser:()=>De});function Ee(){return localStorage.getItem("fm_token")}function Ve(t){localStorage.setItem("fm_token",t)}function $e(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ue(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function De(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function f(t,e={}){let i=Ee(),a={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...e.headers||{}};try{let o=`cb=${Date.now()}`,c=t.includes("?")?"&":"?",r=`${ut}${t}${c}${o}`,s=await fetch(r,{...e,headers:a}),n;try{let l=await s.text();try{n=JSON.parse(l)}catch{n={error:`Server Error (${s.status}): ${l.substring(0,80)}...`}}}catch{n={error:"Gagal membaca respon dari server"}}return s.status===401&&($e(),window.location.hash="#/login"),{ok:s.ok,status:s.status,data:n}}catch(o){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${o.message})`}}}}var Qe,oa,ut,Te,D=Je(()=>{Qe=!1,oa="https://fm-operations-api.facilitycare-audydental.workers.dev",ut=oa,Te=1e4});var ht={};Ge(ht,{confirmDialog:()=>ze,createModal:()=>le});function le({title:t,content:e,onConfirm:i,onCancel:a,confirmText:o="Simpan",cancelText:c="Batal",size:r="md",confirmClass:s="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},l=document.createElement("div");l.className="modal-overlay",l.innerHTML=`
    <div class="modal" style="max-width:${n[r]||n.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${c}</button>
        ${i?`<button class="btn ${s} modal-confirm">${o}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&l.querySelector(".modal-body").appendChild(e);let m=()=>{l.classList.remove("show"),setTimeout(()=>l.remove(),250)};return l.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),m()}),l.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),m()}),i&&l.querySelector(".modal-confirm").addEventListener("click",()=>i(l,m)),l.addEventListener("click",d=>{d.target===l&&(a&&a(),m())}),document.body.appendChild(l),requestAnimationFrame(()=>l.classList.add("show")),{overlay:l,close:m}}function ze(t,e,i="Konfirmasi"){return le({title:i,content:`<p>${t}</p>`,onConfirm:(a,o)=>{e(),o()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var Pe=Je(()=>{});var Y={};Ge(Y,{downloadExcel:()=>$,parseExcel:()=>Xe,renderExcelButtons:()=>Ze});function Xe(t){return new Promise((e,i)=>{let a=new FileReader;a.onload=o=>{try{let c=new Uint8Array(o.target.result),r=XLSX.read(c,{type:"array"}),s=r.SheetNames[0],n=r.Sheets[s];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${r.SheetNames.join(", ")}`),console.log(`Sheet Used: ${s}`);let l=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),m=l.e.r-l.s.r+1,d=l.e.c-l.s.c+1;console.log(`Total Rows (including empty): ${m}`),console.log(`Total Columns: ${d}`);let p=[];for(let h=l.s.c;h<=l.e.c;++h){let g=n[XLSX.utils.encode_cell({c:h,r:l.s.r})];g&&g.v&&p.push(g.v)}console.log(`Headers Found: ${p.join(", ")}`),console.log("---------------------------");let u=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(u,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(u,"__headers",{value:p,enumerable:!1}),e(u)}catch(c){i(c)}},a.onerror=o=>i(o),a.readAsArrayBuffer(t)})}function $(t,e){try{let i=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,i,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function Ze(t){return`
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
  `}var L=Je(()=>{});D();var We={},Oe=null;function A(t,e){We[t]=e}function be(t){window.location.hash=t}function gt(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[i,...a]=e.split("?"),o=We[i];if(!o){for(let[r,s]of Object.entries(We))if(r.endsWith("/*")&&i.startsWith(r.slice(0,-2))){o=s;break}}Oe&&(Oe(),Oe=null);let c=document.getElementById("main-content");if(c&&(c.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),o){let r=new URLSearchParams(a.join("?")),s=i.split("/").filter(Boolean),n=await o({path:i,params:r,segments:s,main:c});n&&(Oe=n)}else{let r=c||document.getElementById("app");r&&(r.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var Ie;function la(){return Ie||(Ie=document.createElement("div"),Ie.id="toast-container",document.body.appendChild(Ie)),Ie}function bt(t,e="info",i=3500){let a=la(),o=document.createElement("div");o.className=`toast toast-${e}`;let c={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};o.innerHTML=`<span class="toast-icon">${c[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(o),requestAnimationFrame(()=>o.classList.add("show")),setTimeout(()=>{o.classList.remove("show"),setTimeout(()=>o.remove(),350)},i)}var z=t=>bt(t,"success"),Q=t=>bt(t,"error");Pe();D();D();function yt({columns:t,data:e,onEdit:i,onDelete:a,onView:o,actions:c=[],emptyText:r="Tidak ada data",bulkSelect:s=null}){let n=document.createElement("div");if(n.className="table-wrapper",!e||e.length===0)return n.innerHTML=`<div class="empty-state"><p>${r}</p></div>`,n;let l=document.createElement("table");l.className="data-table";let m=document.createElement("thead"),d=document.createElement("tr");if(s){let u=document.createElement("th");u.style.width="40px",u.style.textAlign="center";let h=document.createElement("input");h.type="checkbox",h.id="select-all-checkbox",h.title="Pilih semua",h.addEventListener("change",()=>{e.forEach(g=>{h.checked?s.selectedIds.add(g.id):s.selectedIds.delete(g.id)}),n.querySelectorAll(".row-checkbox").forEach(g=>g.checked=h.checked),s.onToggle()}),u.appendChild(h),d.appendChild(u)}if(t.forEach(u=>{let h=document.createElement("th");h.textContent=u.label,u.width&&(h.style.width=u.width),d.appendChild(h)}),i||a||o||c.length>0){let u=document.createElement("th");u.textContent="Aksi",u.style.width="120px",d.appendChild(u)}m.appendChild(d),l.appendChild(m);let p=document.createElement("tbody");return e.forEach(u=>{let h=document.createElement("tr");if(s){let g=document.createElement("td");g.style.textAlign="center",g.style.width="40px";let b=document.createElement("input");b.type="checkbox",b.className="row-checkbox",b.checked=s.selectedIds.has(u.id),b.addEventListener("change",()=>{if(b.checked)s.selectedIds.add(u.id);else{s.selectedIds.delete(u.id);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1)}s.onToggle()}),g.appendChild(b),h.appendChild(g)}if(t.forEach(g=>{let b=document.createElement("td");if(g.render){let y=g.render(u[g.key],u);y instanceof HTMLElement?b.appendChild(y):b.innerHTML=y||""}else b.textContent=u[g.key]!==null&&u[g.key]!==void 0&&u[g.key]!==""?u[g.key]:"";g.nowrap&&(b.style.whiteSpace="nowrap"),h.appendChild(b)}),i||a||o||c.length>0){let g=document.createElement("td");g.className="actions-cell";let b=document.createElement("div");if(b.className="btn-group",o){let y=document.createElement("button");y.className="btn btn-xs btn-ghost",y.innerHTML="\u{1F441}",y.title="Lihat",y.addEventListener("click",()=>o(u)),b.appendChild(y)}if(i){let y=document.createElement("button");y.className="btn btn-xs btn-secondary",y.innerHTML="\u270F\uFE0F",y.title="Edit",y.addEventListener("click",()=>i(u)),b.appendChild(y)}c.forEach(y=>{let v=document.createElement("button");v.className=`btn btn-xs ${y.class||"btn-ghost"}`,v.innerHTML=y.icon||y.label,v.title=y.label,v.addEventListener("click",()=>y.handler(u)),b.appendChild(v)}),g.appendChild(b),h.appendChild(g)}p.appendChild(h)}),l.appendChild(p),n.appendChild(l),n}function ft({page:t,pages:e,total:i,limit:a,onPage:o}){if(e<=1)return null;let c=document.createElement("div");c.className="pagination";let r=document.createElement("span");r.className="pagination-info",r.textContent=`Total: ${i} data`,c.appendChild(r);let s=document.createElement("div");s.className="pagination-btns";let n=(d,p,u=!1,h=!1)=>{let g=document.createElement("button");g.className=`btn btn-sm ${h?"btn-primary":"btn-ghost"} pagination-btn`,g.textContent=d,g.disabled=u,g.addEventListener("click",()=>o(p)),s.appendChild(g)};n("\xAB",1,t===1),n("\u2039",t-1,t===1);let l=Math.max(1,t-2),m=Math.min(e,t+2);for(let d=l;d<=m;d++)n(d,d,!1,d===t);return n("\u203A",t+1,t===e),n("\xBB",e,t===e),c.appendChild(s),c}Pe();function Ye(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${Ye(e.fields)}</div>`;let i=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",o="";switch(e.type){case"textarea":o=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${i} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let r=(e.options||[]).map(d=>{let p=typeof d=="object"?d.value:d,u=typeof d=="object"?d.label:d,h=e.value==p?"selected":"";return`<option value="${p}" ${h}>${u}</option>`}).join("");o=`<select name="${e.name}" class="form-control" ${i}><option value="">-- Pilih ${e.label||""} --</option>${r}</select>`;break;case"combobox":let s=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,n=(e.options||[]).map(d=>{let p=typeof d=="object"?d.value:d,u=typeof d=="object"?d.label||d.value||"":d||"";return(u==="undefined"||u==="[object Object]"||u==="null")&&(u=""),u?`<option value="${u}"></option>`:""}).join(""),l=e.value||"";if(e.value){let d=(e.options||[]).find(p=>(typeof p=="object"?p.value:p)==e.value);if(d){let p=typeof d=="object"?d.label||d.value||"":d||"";p&&p!=="undefined"&&p!=="[object Object]"&&p!=="null"&&(l=p)}}o=`
          <input type="text" name="${e.name}" list="${s}" class="form-control" value="${l}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${s}">${n}</datalist>
        `;break;case"checkbox":o=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let m=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";o=`<input type="date" name="${e.name}" class="form-control" value="${m}" ${i}>`;break;case"number":o=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${i}>`;break;case"email":o=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i}>`;break;case"url":o=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${i}>`;break;default:o=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i} autocomplete="off">`}let c=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${o}${c}</div>`}).join("")}function vt(t){let e={},i=new FormData(t);for(let[a,o]of i.entries())e[a]=o===""?null:o;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function kt(t,e){e&&Object.entries(e).forEach(([i,a])=>{let o=t.querySelector(`[name="${i}"]`);o&&(o.hasAttribute("list")||(o.type==="checkbox"?o.checked=!!a:o.type==="date"&&a&&window.parseFlexibleDate?o.value=window.parseFlexibleDate(a):o.value=a??""))})}L();function E({container:t,title:e,icon:i,apiPath:a,columns:o,formFields:c,filterFields:r,defaultFilters:s={},itemLabel:n="Data",canCreate:l=!0,canEdit:m=!0,canDelete:d=!0,onBeforeSubmit:p,onAfterLoad:u,onDataLoaded:h,extraActions:g=[],initialSearch:b="",exportOptions:y=null,bulkDelete:v=!1,paginationMode:k="server"}){let S=1,x={...s};b&&(x.search=b);let T=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions">
        ${l?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
      </div>
    </div>

    ${v?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${y?Ze(y.moduleName):""}

    ${r&&r.length>0?`
    <div class="filter-bar card">
      <div class="filter-bar-inner">
        ${r.map(_=>_.type==="search"?`<div class="filter-search"><input type="search" class="form-control" placeholder="${_.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}"></div>`:_.type==="select"?`<select class="form-control filter-select" name="${_.name}" id="filter-${_.name}"><option value="">-- ${_.label} --</option>${(_.options||[]).map(w=>`<option value="${typeof w=="object"?w.value:w}" ${x[_.name]===(typeof w=="object"?w.value:w)?"selected":""}>${typeof w=="object"?w.label:w}</option>`).join("")}</select>`:"").join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
      </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function B(){if(!document.getElementById("bulk-toolbar"))return;let w=document.getElementById("bulk-count"),R=document.getElementById("btn-bulk-delete"),j=document.getElementById("btn-bulk-cancel");w.textContent=`${T.size} item dipilih`,T.size>0?(R.disabled=!1,j.disabled=!1):(R.disabled=!0,j.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{T.clear(),document.querySelectorAll(".row-checkbox").forEach(w=>w.checked=!1);let _=document.getElementById("select-all-checkbox");_&&(_.checked=!1),B()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(T.size===0)return;let _=[...T],w=document.createElement("div");w.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",w.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${_.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${_.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(w),w.querySelector("#bulk-cancel-btn").addEventListener("click",()=>w.remove()),w.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let R=w.querySelector("#bulk-confirm-btn");R.disabled=!0,R.textContent="Menghapus...";let j=await f(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:_})});w.remove(),j.ok?(z(`${_.length} ${n} berhasil dihapus.`),T.clear(),B(),K()):Q(j.data?.error||"Gagal menghapus data.")})});let C=document.getElementById("filter-search"),I;if(C?.addEventListener("input",_=>{clearTimeout(I),I=setTimeout(()=>{x.search=_.target.value,S=1,K()},400)}),r?.forEach(_=>{_.type==="select"&&document.getElementById(`filter-${_.name}`)?.addEventListener("change",w=>{x[_.name]=w.target.value,S=1,K()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{x={...s},C&&(C.value=""),r?.forEach(_=>{let w=document.getElementById(`filter-${_.name}`);w&&(w.value="")}),S=1,K()}),document.getElementById("btn-create")?.addEventListener("click",()=>_e(null)),y){document.getElementById(`btn-export-${y.moduleName}`)?.addEventListener("click",async w=>{let R=w.target,j=R.innerHTML;R.innerHTML="\u23F3 Loading...",R.disabled=!0;try{await y.onExport()}catch{Q("Gagal export data")}finally{R.innerHTML=j,R.disabled=!1}}),document.getElementById(`btn-template-${y.moduleName}`)?.addEventListener("click",()=>{y.onTemplate()});let _=document.getElementById(`input-import-${y.moduleName}`);_?.addEventListener("change",async w=>{let R=w.target.files[0];if(!R)return;_.disabled=!0;let j=document.createElement("div");j.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",j.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(j);let ie=j.querySelector("#import-progress-text"),Z=j.querySelector("#import-progress-bar"),q=j.querySelector("#import-summary"),F=j.querySelector("#import-close-btn");F.addEventListener("click",()=>{j.remove(),K()});try{let se=await Xe(R);if(se.length===0)throw new Error("File kosong atau format salah");let ce=500,oe=0,G=0,V=0,W=se.length;ie.textContent=`Ditemukan ${W} baris data. Memulai import...`;for(let O=0;O<W;O+=ce){let te=se.slice(O,O+ce);ie.textContent=`Mengimport baris ${O+1} - ${Math.min(O+ce,W)} dari ${W}...`,Z.style.width=`${Math.round(O/W*100)}%`;try{let ee=await y.onImport(te);ee?(oe+=ee.inserted||ee.metrics?.inserted||te.length,G+=ee.skipped||ee.metrics?.updated||0):oe+=te.length}catch(ee){console.error("Chunk import failed:",ee),V+=te.length}}Z.style.width="100%",ie.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',q.style.display="block",q.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${W}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${oe}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${G}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${V}</strong></div>
        `,V>0&&(q.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),F.style.display="block",_.value=""}catch(se){ie.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${se.message}`,Z.style.background="var(--danger)",Z.style.width="100%",F.style.display="block",_.value=""}finally{_.disabled=!1}})}async function K(){T.clear(),B();let _=document.getElementById("table-container");if(!_)return;_.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let w=k==="client",R=w?1:S,j=w?Te:20,ie=new URLSearchParams({page:R,limit:j,...Object.fromEntries(Object.entries(x).filter(([,G])=>G))}),Z=await f(`${a}?${ie}`);if(!Z.ok){_.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${Z.data?.error||"Error"}</p></div>`;return}let q=Z.data?.data||Z.data||[],F=Z.data?.pagination,se=q.length;if(w){q=h(q);let G=q.length,V=20,W=Math.ceil(G/V);S>W&&W>0&&(S=W);let O=(S-1)*V,te=S*V;q=q.slice(O,te),F={page:S,limit:V,total:G,pages:W}}!1,u&&u(q);let ce=yt({columns:o,data:q,onEdit:m?G=>_e(G):null,actions:g.map(G=>({...G,handler:V=>G.handler(V,K)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:v?{selectedIds:T,onToggle:B}:null});_.innerHTML="",_.appendChild(ce);let oe=document.getElementById("pagination-container");if(oe&&(oe.innerHTML="",F&&F.pages>1)){let G=ft({page:F.page,pages:F.pages,total:F.total,limit:F.limit,onPage:V=>{S=V,K()}});G&&oe.appendChild(G)}}function Fe(_){let w=typeof c=="function"?c(_):c;return Ye(w)}function _e(_){let w=!!_,R=document.createElement("form");if(R.noValidate=!0,R.innerHTML=Fe(_),w){let ie=typeof c=="function"?c(_):c;kt(R,_)}let{close:j}=le({title:w?`Edit ${n}`:`Tambah ${n}`,content:R,size:"lg",confirmText:w?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(ie,Z)=>{if(!R.reportValidity())return;let q=ie.querySelector(".modal-confirm");q.disabled=!0,q.textContent="Menyimpan...";let F=vt(R),se=typeof c=="function"?c(_):c,ce=async W=>{for(let O of W)if(O.type==="row")await ce(O.fields);else if(O.type==="combobox"&&F[O.name]){let te=F[O.name],ee=(O.options||[]).find(re=>{let de=String(typeof re=="object"?re.value:re),ra=String(typeof re=="object"?re.label:re);return de===te||ra===te});if(ee)F[O.name]=typeof ee=="object"?ee.value:ee;else if(O.createApi){let re={};re[O.createApi.field]=te,O.createApi.extra&&Object.assign(re,O.createApi.extra);let de=await f(O.createApi.path,{method:"POST",body:JSON.stringify(re)});if(de.ok&&de.data?.id)F[O.name]=de.data.id;else if(de.ok&&!de.data?.id)F[O.name]=te;else throw new Error(`Gagal membuat master data: ${de.data?.error||"Unknown error"}`)}}};try{await ce(se)}catch(W){Q(W.message),q.disabled=!1,q.textContent=w?"Simpan Perubahan":`Tambah ${n}`;return}p&&(F=await p(F,_));let oe=w?"PUT":"POST",G=w?`${a}/${_.id}`:a,V=await f(G,{method:oe,body:JSON.stringify(F)});V.ok?(z(w?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),Z(),K()):(Q(V.data?.error||"Gagal menyimpan data."),q.disabled=!1,q.textContent=w?"Simpan Perubahan":`Tambah ${n}`)}})}function Ce(_){ze(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let w=await f(`${a}/${_.id}`,{method:"DELETE"});w.ok?(z(`${n} berhasil dihapus.`),K()):Q(w.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return K(),K}D();D();var he=null,Me=null;async function ye(t=!1){if(he&&!t)return console.log("Employees Raw (Cache Hit)",he.slice(0,5)),he;let e=await f(`/api/employees?limit=${Te}&status=Aktif`);return he=(e.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",he.slice(0,5)),he}async function U(t=!1){let i=(await ye(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",i.slice(0,5)),i}async function P(t=!1){return Me&&!t||(Me=((await f("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),Me}function N(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function et(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function fe(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function tt(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function ae(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}L();function at(t,e){if(t.period!=="Q3")return!1;let i=String(t.status||"").toLowerCase();if(i!=="selesai"&&i!=="completed"&&i!=="done")return!1;let a=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?a.includes("inspeksi"):e==="gcdc"?a.includes("general cleaning")||a.includes("deep cleaning"):!1}D();L();function St(t,e){let i=String(t.status||"").toLowerCase();return e==="active"?i==="aktif":!1}D();L();function nt(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let o=new Date(a);o.setDate(a.getDate()+30);let c=new Date(t.end_date);return c.setHours(0,0,0,0),c>=a&&c<=o}return!1}D();L();function xt(t,e){let i=String(t.status||"").toLowerCase();return e==="open"?i==="open":!1}D();function wt(t,e){let i=String(t.status||"").toLowerCase();return e==="pending"?i==="pending":!1}var pe={};function Le(t){if(pe[t]){try{pe[t].destroy()}catch{}delete pe[t]}}function ca(){Object.keys(pe).forEach(Le)}var H=(t,e=0)=>{let i=Number(t);return isNaN(i)||t===null||t===void 0?e:i},ve=(t,e="\u2014")=>{if(t==null||t==="")return e;let i=String(t).trim();return i===""||i==="[object Object]"?e:i};var da=t=>{if(!t||typeof t!="string")return"";try{let[e,i]=t.split("-");return new Date(Number(e),Number(i)-1).toLocaleDateString("id-ID",{month:"short",year:"2-digit"})}catch{return t}};function _t(t,e,i=900){if(!t)return;let a=Math.max(0,Math.round(H(e)));if(a===0){t.textContent="0";return}let o=Date.now(),c=()=>{let r=Math.min((Date.now()-o)/i,1),s=1-Math.pow(1-r,3);t.textContent=Math.round(s*a).toLocaleString("id-ID"),r<1?requestAnimationFrame(c):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(c)}var pa={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},ma=t=>{let e=ve(t,"\u2014");return`<span class="status-pill ${pa[e]||"pill-neutral"}">${e}</span>`};var ne={family:"Inter",size:11},ke="#94A3B8",Ke="#F1F5F9",it=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],ua=()=>window.innerWidth<768;function ot(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:ua()?"bottom":"top",labels:{font:ne,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ne,titleFont:{...ne,weight:"700"}}},scales:{x:{grid:{color:Ke},ticks:{font:ne,color:ke,maxRotation:0}},y:{grid:{color:Ke},ticks:{font:ne,color:ke},beginAtZero:!0}},...t}}var ga=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),ba=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function rt(t=3){return Array(t).fill(0).map((e,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function X(t,e,i=8e3){try{let a=new AbortController,o=setTimeout(()=>a.abort(),i),c=await f(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(o),!c||!c.ok)return e;let r=c.data;return r?r.data!==void 0?r.data??e:r:e}catch{return e}}function ha(){["skel-donut","skel-trend","skel-insp","skel-contract"].forEach(a=>{let o=document.getElementById(a);o&&(o.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract"].forEach(a=>{let o=document.getElementById(a);if(o&&o.style.display==="none"){o.style.display="block";let c=o.parentElement;if(c&&!c.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent="Belum ada data",o.style.display="none",c.appendChild(r)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Tt({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Et({}),["table-contracts","table-issues"].forEach(a=>{let o=document.getElementById(a);o&&o.querySelector(".skeleton")&&(o.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Ct(t){ca(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">


      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${ga()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${ba()}</div>

      <!-- Charts Row -->
      <div class="charts-row" style="grid-template-columns: repeat(2, minmax(0, 1fr));">
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
              <option value="">Bulan Ini</option>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>st(t)),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let a=i.target.value,o=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",c=document.getElementById("skel-insp"),r=document.getElementById("chart-insp");c&&(c.style.display="block",c.style.position="absolute"),r&&(r.style.display="none");let s=await X(o,{},8e3);try{$t(s)}catch(n){console.warn("InspBar render:",n),me("skel-insp","chart-insp")}}),t._skelTimeout=setTimeout(()=>ha(),5e3),await st(t),t._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?st(t):clearInterval(t._dashRefresh)},6e4)}async function st(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,i,a,o,c,r,s,n,l,m,d,p]=await Promise.all([X("/api/dashboard/kpi",{},8e3),X("/api/dashboard/issues-trend",{},8e3),X("/api/dashboard/issues-summary",{},8e3),X("/api/dashboard/inspection-bar",{},8e3),X("/api/dashboard/stats",{},8e3),X("/api/dashboard/calendar",[],8e3),X("/api/schedule?limit=10000",{data:[]},8e3),X("/api/employees?limit=10000",{data:[]},8e3),X("/api/contracts?limit=10000",{data:[]},8e3),X("/api/issues?limit=10000",{data:[]},8e3),X("/api/one_on_one?limit=10000",{data:[]},8e3),X("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3)]);if(e){let u=Array.isArray(s?.data)?s.data:Array.isArray(s)?s:[],h=Array.isArray(n?.data)?n.data:Array.isArray(n)?n:[],g=Array.isArray(l?.data)?l.data:Array.isArray(l)?l:[],b=Array.isArray(m?.data)?m.data:Array.isArray(m)?m:[],y=Array.isArray(d?.data)?d.data:Array.isArray(d)?d:[];e.employees&&(e.employees.current=h.filter(v=>St(v,"active")).length),e.contracts&&(e.contracts.current=g.filter(v=>nt(v,"active")).length),e.expiring30&&(e.expiring30={current:g.filter(v=>nt(v,"expiring30")).length}),e.issues&&(e.issues.current=b.filter(v=>xt(v,"open")).length),e.one_on_one&&(e.one_on_one.current=y.filter(v=>wt(v,"pending")).length),e.inspection_month&&(e.inspection_month.current=u.filter(v=>at(v,"inspeksi")).length),e.cleaning_month&&(e.cleaning_month.current=u.filter(v=>at(v,"gcdc")).length)}try{Tt(e)}catch(u){console.warn("KPI render:",u)}try{Et(e)}catch(u){console.warn("MiniStats render:",u)}try{ya(Array.isArray(a?.by_category)?a.by_category:[])}catch(u){console.warn("Donut render:",u),me("skel-donut","chart-donut")}try{fa(i)}catch(u){console.warn("Trend render:",u),me("skel-trend","chart-trend")}try{$t(o)}catch(u){console.warn("InspBar render:",u),me("skel-insp","chart-insp")}try{let u=Array.isArray(c)?c:Array.isArray(c?.recent_issues)?c.recent_issues:[];ka(u)}catch(u){console.warn("IssuesTable render:",u)}try{let u=Array.isArray(c?.expiring_contracts)?c.expiring_contracts:[];va(p)}catch(u){console.warn("ContractsTable render:",u)}try{Sa(Array.isArray(r)?r:[])}catch(u){console.warn("Agenda render:",u)}try{xa(e)}catch(u){console.warn("KPI Kebersihan render:",u)}try{wa()}catch(u){console.warn("Quick Actions render:",u)}}function Tt(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let i=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=i.map(a=>{let o=H(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${o}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${o}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{_t(a,parseInt(a.dataset.target)||0)})}function Et(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let i=[{icon:"\u{1F4C5}",label:"Jadwal",val:t.schedule?.current,href:"#/timeline",color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F504}",label:"Reliefer",val:t.reliever_total?.current,href:"#/relievers?dash_filter=reliever",color:"mini-teal"},{icon:"\u{1F50D}",label:"Inspeksi",val:t.inspection_month?.current,href:"#/timeline?dash_filter=inspeksi",color:"mini-blue"},{icon:"\u{1F9F9}",label:"GCDC",val:t.cleaning_month?.current,href:"#/timeline?dash_filter=gcdc",color:"mini-green"},{icon:"\u{1F4A8}",label:"Fogging",val:t.fogging_month?.current,href:"#/reports/fogging",color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=i.map(a=>`
    <a href="${a.href}" class="mini-stat ${a.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${a.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${H(a.val)}">0</div>
        <div class="mini-stat-text">${a.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(a=>_t(a,parseInt(a.dataset.target)||0,700))}function ya(t){me("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!e||!i)return;Le("donut");let a=(t||[]).filter(n=>H(n.count)>0);if(!a.length){Re(e,"Belum ada data permasalahan");return}let o=a.map(n=>`${ve(n.category,"Lainnya")}`),c=a.map(n=>H(n.count)),r=c.reduce((n,l)=>n+l,0);i.innerHTML=a.map((n,l)=>{let m=it[l%it.length],d=r>0?Math.round(n.count/r*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${m}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${d}%)</span></div>
          <div class="donut-legend-label">${o[l]}</div>
        </div>
      </div>
    `}).join("");let s={id:"centerText",beforeDraw:function(n){let l=n.width,m=n.height,d=n.ctx;d.restore();let p=(m/80).toFixed(2);d.font="bold "+p+"em Inter",d.textBaseline="middle",d.fillStyle="#1E293B";let u=r.toString(),h=Math.round((l-d.measureText(u).width)/2),g=m/2;d.fillText(u,h,g-10),d.font="600 "+(p*.35).toFixed(2)+"em Inter",d.fillStyle="#64748B";let b="Total",y=Math.round((l-d.measureText(b).width)/2);d.fillText(b,y,g+15),d.save()}};pe.donut=new Chart(e,{type:"doughnut",data:{labels:o,datasets:[{data:c,backgroundColor:it,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ne,titleFont:{...ne,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutout:"75%"},plugins:[s]})}function fa(t){me("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;Le("trend"),t=t||{};let i=(t.labels||[]).map(da),a=(t.open||[]).map(c=>H(c)),o=(t.closed||[]).map(c=>H(c));if(!i.length){Re(e,"Belum ada data trend");return}pe.trend=new Chart(e,{type:"line",data:{labels:i,datasets:[{label:"Open",data:a,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:o,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#10B981",borderWidth:2}]},options:ot({plugins:{legend:{display:!1}}})})}function $t(t){me("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;Le("inspBar"),t=t||{};let i=t.labels||[],a=(t.fc||[]).map(c=>H(c)),o=(t.spv||[]).map(c=>H(c));if(!i.length){Re(e,"Belum ada data inspeksi");return}pe.inspBar=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:o,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:ot({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ne,color:ke,maxRotation:45,minRotation:30}},y:{grid:{color:Ke},ticks:{font:ne,color:ke},min:0,max:100}}})})}function va(t){me("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;Le("contractMiniBar"),t=t||{};let i=t.labels||[],a=(t.data||[]).map(r=>H(r));if(!i.length){Re(e,"Belum ada data");return}let c=e.getContext("2d").createLinearGradient(0,0,0,200);c.addColorStop(0,"#60A5FA"),c.addColorStop(1,"#2563EB"),pe.contractMiniBar=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Kontrak Habis",data:a,backgroundColor:c,borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:ot({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ne,color:ke,maxRotation:0}},y:{grid:{color:Ke,borderDash:[4,4],drawBorder:!1},ticks:{font:ne,color:ke,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function ka(t){let e=document.getElementById("table-issues");if(!e)return;let i=(t||[]).slice(0,8);if(!i.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${i.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${ma(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${ve(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${ve(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function Sa(t){let e=document.getElementById("widget-agenda");if(!e)return;let i=new Date,a=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,c=(t||[]).filter(r=>(r.event_date||"").startsWith(a)).slice(0,10);if(!c.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${c.map(r=>{let s="#3B82F6",n="#EFF6FF",l="Agenda",m=(r.title||"").toLowerCase();return m.includes("inspeksi")?(s="#10B981",n="#ECFDF5",l="Inspeksi"):m.includes("cleaning")||m.includes("gcdc")?(s="#3B82F6",n="#EFF6FF",l="Cleaning"):m.includes("reliefer")?(s="#F59E0B",n="#FFFBEB",l="Reliefer"):m.includes("fogging")&&(s="#8B5CF6",n="#F5F3FF",l="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(r.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${s};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${ve(r.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${ve(r.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${s}">${l}</div>
        </div>
      `}).join("")}
    </div>
  `}function xa(t){let e=document.getElementById("widget-kpi-kebersihan");if(!e)return;let i=t?.kebersihan||{},a=H(i.area?.avg_fc,0),o=H(i.area?.avg_spv,0),c=a>0&&o>0?(a+o)/2:a||o||0,r=H(i.issues?.total,0),s=H(i.issues?.closed,0),n=r>0?Math.round(s/r*100):100,l=H(i.cleaning?.scheduled,0),m=H(i.cleaning?.reported,0),d=l>0?Math.round(m/l*100):100,p=H(i.gcdc?.scheduled,0),u=H(i.gcdc?.reported,0),h=p>0?Math.round(u/p*100):100,g=H(i.issues?.cleanTotal,0),y=H(t?.fogging_month?.current,0)>0?100:0,v=[{label:"Kebersihan Area",val:`${Math.round(c)}%`,target:"Target 95%",icon:"\u{1F9F9}",bg:c>=95?"#ECFDF5":"#FEF2F2",color:c>=95?"#10B981":"#EF4444"},{label:"Penyelesaian Complaint",val:`${n}%`,target:"Target 100%",icon:"\u23F1\uFE0F",bg:n>=100?"#ECFDF5":"#FEF2F2",color:n>=100?"#10B981":"#EF4444"},{label:"Kepatuhan Jadwal Cleaning",val:`${d}%`,target:"Target 100%",icon:"\u23F1\uFE0F",bg:d>=100?"#EFF6FF":"#FEF2F2",color:d>=100?"#3B82F6":"#EF4444"},{label:"Kepatuhan GCDC",val:`${h}%`,target:"Target 100%",icon:"\u{1F9F9}",bg:h>=100?"#EFF6FF":"#FEF2F2",color:h>=100?"#3B82F6":"#EF4444"},{label:"Complaint Cleaning (\u226410)",val:`${g}`,target:"Target \u226410",icon:"\u{1F4DD}",bg:g<=10?"#F5F3FF":"#FEF2F2",color:g<=10?"#8B5CF6":"#EF4444"},{label:"Pelaksanaan Fogging",val:`${y}%`,target:"Target 100%",icon:"\u{1F4A8}",bg:y>=100?"#F5F3FF":"#FEF2F2",color:y>=100?"#8B5CF6":"#EF4444"}];e.innerHTML=`
    <div style="display:grid;grid-template-columns:1fr;gap:12px;padding-bottom:12px;height:160px;overflow-y:auto;padding-right:8px;">
      ${v.map(k=>{let S=k.val.includes("%")?parseInt(k.val):Math.min(100,parseInt(k.val)*10);return`
        <div class="prog-item">
          <div class="prog-header">
            <div class="prog-title">
              <div class="prog-title-icon" style="background:${k.bg};color:${k.color}">${k.icon}</div>
              ${k.label}
            </div>
            <div class="prog-val">${k.val}</div>
          </div>
          <span class="prog-target">${k.target}</span>
          <div class="prog-bar-bg">
            <div class="prog-bar-fill" style="width:${S}%;background:${k.color}"></div>
          </div>
        </div>
      `}).join("")}
    </div>
  `}function wa(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function me(t,e){let i=document.getElementById(t),a=document.getElementById(e);i&&(i.style.display="none",i.style.position=""),a&&(a.style.display="block")}function Re(t,e="Belum ada data"){if(!t)return;t.style.display="none";let i=t.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent=e,i.appendChild(o)}}D();async function Dt(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),i=document.getElementById("login-error"),a=document.getElementById("login-btn"),o=document.getElementById("toggle-password"),c=document.getElementById("login-password");o?.addEventListener("click",()=>{let r=c.type==="text";c.type=r?"password":"text",o.style.color=r?"":"var(--primary)"}),e?.addEventListener("submit",async r=>{r.preventDefault(),i.style.display="none";let s=e.username.value.trim(),n=e.password.value;if(!s||!n){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let l=await f("/api/auth/login",{method:"POST",body:JSON.stringify({username:s,password:n})});l.ok&&l.data.success?(Ve(l.data.data.token),De(l.data.data.user),z("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=l.data.error||"Username atau password salah.",i.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}D();L();async function _a(){return await P()}function Ca(t,e){let i=String(t.status||"").toLowerCase();return e==="active"?i==="aktif":!1}async function It(t,e){let i=await _a(),a=e?e.get("dash_filter"):null;E({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:o=>a?o.filter(c=>Ca(c,a)):o,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:o=>fe(o)},{key:"phone",label:"No. HP",render:o=>o?`<a href="tel:${o}">${o}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>N(o)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:o=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:o?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:o?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:i,value:o?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:o?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:o?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:o?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let o=await f(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(o.ok){let c=o.data.data.map(r=>({"Nama Lengkap":r.full_name,Cabang:r.branch_name||"",Divisi:r.division||"","No. HP":r.phone||"","Tgl Masuk":r.join_date||"",Status:r.status||""}));$(c,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async o=>{let c=n=>{if(!n)return null;let l=String(n||"").toLowerCase(),m=i.find(d=>String(d.label||"").toLowerCase()===l);return m?m.value:null},r=o.map(n=>({full_name:String(n["Nama Lengkap"]||"").trim(),branch_id:c(String(n.Cabang||"").trim()),division:String(n.Divisi||"").trim()||"FACILITY CARE",phone:String(n["No. HP"]||"").trim(),join_date:String(n["Tgl Masuk"]||"").trim(),status:String(n.Status||"").trim(),notes:String(n.Catatan||"").trim()})).filter(n=>n.full_name),s=await f("/api/import/employees",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}D();L();var ct=[],Pt=[];async function Ta(){ct=await P(),Pt=await ye()}var lt=async t=>{let e=[],i=1;for(;;){let o=await(await Promise.resolve().then(()=>(D(),ge))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${i}`);if(!o.ok)break;let c=o.data?.data||o.data||[],r=Array.isArray(c)?c:[];if(e=e.concat(r),r.length<100||o.data?.pagination&&i>=o.data.pagination.pages)break;i++}return e};function Ea(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let o=new Date(a);o.setDate(a.getDate()+30);let c=new Date(t.end_date);return c.setHours(0,0,0,0),c>=a&&c<=o}return!1}async function Lt(t,e){await Ta();let i=e?e.get("dash_filter"):null;E({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",onDataLoaded:a=>i?a.filter(o=>Ea(o,i)):a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>fe(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,o)=>o.end_date&&String(o.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':et(a)},{key:"status",label:"Status",render:a=>N(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:ct},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[c,r]=await Promise.all([lt("/api/employees?status=Aktif"),lt("/api/contracts")]);if(c.length>0){let s=r.filter(d=>d.status==="Aktif"),n=new Set(s.map(d=>d.employee_id)),l=c.filter(d=>!n.has(d.id)),m=`<p style="margin-bottom:12px">Data yang terbaca: <b>${c.length}</b> Karyawan Aktif, dan <b>${s.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${l.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;l.forEach(d=>{let p=r.filter(h=>h.employee_id===d.id),u='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(p.length>0){let h=p[0];u=`Pernah ada kontrak (Status: <b style="color:#EF4444">${h.status}</b>, Selesai: ${window.formatDate(h.end_date)})`}m+=`<li style="margin-bottom:8px"><b>${d.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${d.branch_name||"-"} | ${u}</span></li>`}),m+="</ul>",Promise.resolve().then(()=>(Pe(),ht)).then(d=>d.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:m,cancelText:"Tutup"}))}}catch(c){console.error(c)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let o=document.querySelector(".page-actions");o&&o.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:Pt,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:ct,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await f(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let o=a.data.data.map(c=>({"Nama Lengkap":c.employee_name,Cabang:c.branch_name||"","Div / Bagian":c.division||"","Tanggal Mulai":c.start_date||"","Tanggal Selesai":c.end_date&&String(c.end_date).startsWith("2099")?"":c.end_date||"","Sisa Kontrak":c.end_date&&String(c.end_date).startsWith("2099")?"Tetap":c.days_remaining!==null&&c.days_remaining!==void 0?`${c.days_remaining} Hari`:"",Status:c.status||""}));$(o,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[o,c]=await Promise.all([f("/api/branches?limit=10000"),lt("/api/employees")]),r=o.data?.data||[],s=c||[];console.log(`Total employee yang berhasil dimuat dari database : ${s.length}`),s.length>0&&(console.log("Contoh 5 employee pertama:"),s.slice(0,5).forEach((g,b)=>{console.log(`${b+1}. ID: ${g.id}, Name: ${g.full_name}, Status: ${g.status}`)}));let n=g=>{if(!g)return null;let b=String(g||"").replace(/\s+/g," ").toLowerCase().trim(),y=r.find(v=>String(v.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===b||String(v.code||"").replace(/\s+/g," ").toLowerCase().trim()===b||String(v.name||"").replace(/\s+/g," ").toLowerCase().trim()===b);return y?y.id:null},l=(g,b)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${b}`),console.log(`Nama dari Excel : "${g}"`),!g)return console.log("Alasan gagal mapping : Nama kosong"),null;let y=String(g||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${y}"`),console.log(`Jumlah employee di database : ${s.length}`);let v=s.find(k=>String(k.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===y);return v?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${v.id}`),v.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},m=g=>{if(!g)return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let b=String(g).trim();if(/^\d{4,5}(\.\d+)?$/.test(b)){let v=Math.floor(Number(b));if(v>2e4&&v<99999){let k=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(b))return b.slice(0,10);let y=b.split(/[\/\-\.]/);if(y.length===3){let[v,k,S]=y.map(x=>x.trim());if(v.length===4&&k.length<=2&&S.length<=2)return`${v}-${k.padStart(2,"0")}-${S.padStart(2,"0")}`;if(S.length===4&&k.length<=2&&v.length<=2)return`${S}-${k.padStart(2,"0")}-${v.padStart(2,"0")}`}return b},d=a.map((g,b)=>{let y=b+2,v=String(g["Nama Lengkap"]||"").trim(),k=g["Tanggal Mulai"],S=m(k);if(!S){let B=a.__worksheet,C=a.__headers||[],I=C.indexOf("Tanggal Mulai"),K="N/A",Fe="N/A",_e="N/A";if(I!==-1&&B&&window.XLSX){let _=window.XLSX.utils.encode_cell({c:I,r:y-1});_e=_;let w=B[_];w?(K=w.t||"undefined",Fe=w.w||"undefined"):K="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let Ce="Unknown";k==null||k===""?Ce="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":k instanceof Date&&isNaN(k.getTime())?Ce="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":Ce="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${y}`),console.log(`Employee Name : ${v}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${I})`),console.log(`Raw Cell Value : "${k}"`),console.log(`JavaScript Type : ${typeof k}`),console.log(`SheetJS Cell Type : ${K}`),console.log(`SheetJS Formatted Value : "${Fe}"`),console.log(`Value After Trim : "${String(k||"").trim()}"`),console.log(`Value After Date Parser : "${S}"`),console.log(`Is Empty : ${!k}`),console.log(`Is Invalid Date : ${k instanceof Date?isNaN(k.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${Ce}`),console.log(`Workbook Sheet : ${B?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${_e}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(g,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(C)),console.log(`==========================
`)}let x=l(v,y),T=null;return x||(T="Karyawan tidak ditemukan di Database"),{isValid:!!x,invalidReason:T,rowNum:y,data:{employee_id:x,branch_id:n(String(g.Cabang||"").trim()),division:String(g["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:S,end_date:m(g["Tanggal Selesai"])||"2099-12-31",status:String(g.Status||"").trim(),_rawName:v}}}),p=[],u=[];if(d.forEach(g=>{g.isValid?p.push(g.data):u.push({rowNum:g.rowNum,name:g.data._rawName,reason:g.invalidReason})}),console.log(`Split Validation - Valid: ${p.length}, Invalid: ${u.length}`),p.length===0)return{inserted:0,skipped:a.length,failed:a.length};let h=await f("/api/import/contracts",{method:"POST",body:JSON.stringify({rows:p,onDuplicate:"update"})});if(!h.ok)throw new Error(h.data?.error||"Import gagal");return h.data}}})}D();L();var dt=[],Be=[];function $a(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let i of e)if(t.some(a=>a.period===i))return i;return"Q3"}function Da(t,e){if(t.period!=="Q3")return!1;let i=String(t.status||"").toLowerCase();if(i!=="selesai"&&i!=="completed"&&i!=="done")return!1;let a=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?a.includes("inspeksi"):e==="gcdc"?a.includes("general cleaning")||a.includes("deep cleaning"):!1}async function Bt(t,e){dt=await P();let i=await U();Be=["Ade","Berlin"];let a=l=>l&&!Be.find(m=>(typeof m=="object"?m.value:m)===l)?[...Be,l]:Be,o=await f(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),c=l=>{if(!l||l==="-"||String(l).trim()==="")return"";let m=String(l).split("-");return m.length===3&&m[0].length===4?`${m[2]}-${m[1]}-${m[0]}`:l},r=o.data?.data||[],s=$a(r),n=e?e.get("dash_filter"):null;E({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",defaultFilters:n?{period:"Q3"}:{period:s},onDataLoaded:l=>(n&&(l=l.filter(m=>Da(m,n))),l.sort((m,d)=>{let p=m.opening_date?new Date(m.opening_date).getTime():0;return(d.opening_date?new Date(d.opening_date).getTime():0)-p})),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:l=>tt(l)},{key:"period",label:"Periode",render:l=>ae(l)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:l=>c(l)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:l=>c(l)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:l=>c(l)},{key:"status",label:"Status",render:l=>N(l)}],filterFields:[{type:"combobox",name:"branch_id",label:"Cabang",options:dt},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"combobox",name:"pic",label:"PIC",options:Be}],formFields:l=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:dt,value:l?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"],value:l?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:l?.period},{name:"pic",label:"PIC",type:"combobox",options:a(l?.pic),value:l?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:l?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:l?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:l?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:l?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let l=await f(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(l.ok){let m=l.data.data.map(d=>({Cabang:d.branch_name||"",Kegiatan:d.activity_type||"",Periode:d.period||"",PIC:d.pic||"","Tgl Opening":d.opening_date||"","Tgl Target":d.target_date||"","Tgl Selesai":d.completion_date||"",Status:d.status||""}));$(m,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async l=>{let d=(await f("/api/branches?all=1")).data?.data||[],p=b=>{if(!b)return null;let y=String(b||"").toLowerCase(),v=d.find(k=>String(k.full_name||"").toLowerCase()===y||String(k.code||"").toLowerCase()===y||String(k.name||"").toLowerCase()===y);return v?v.id:null},u=b=>{if(b==null||b==="")return"";if(b instanceof Date&&!isNaN(b.getTime()))return b.toISOString().slice(0,10);let y=String(b).trim();if(y===""||y==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);if(/^\d{4,5}$/.test(y)){let k=Number(y);if(k>2e4&&k<99999){let S=new Date(Date.UTC(1899,11,30)+k*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}let v=y.split(/[\/\-\.]/);if(v.length===3){let[k,S,x]=v.map(T=>T.trim());if(k.length===4&&S.length<=2&&x.length<=2)return`${k}-${S.padStart(2,"0")}-${x.padStart(2,"0")}`;if(x.length===4&&S.length<=2&&k.length<=2)return`${x}-${S.padStart(2,"0")}-${k.padStart(2,"0")}`}return y},h=l.map(b=>({branch_id:p(String(b.Cabang||"").trim()),activity_type:String(b.Kegiatan||"").trim(),period:String(b.Periode||"").trim(),pic:String(b.PIC||b.Pic||"").trim(),opening_date:u(b["Tgl Opening"]||b["Tanggal Opening"]||b["Tgl Openir"]),target_date:u(b["Tgl Target"]||b["Tanggal Target"]),completion_date:u(b["Tgl Selesai"]||b["Tanggal Selesai"]),status:String(b.Status||"").trim(),notes:String(b.Catatan||b.Keterangan||"").trim()})).filter(b=>b.activity_type&&b.period),g=await f("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:h,onDuplicate:"update"})});if(!g.ok)throw new Error(g.data?.error||"Import gagal");return g.data}}})}D();L();var pt=[],qe=[];function Ia(t,e){let i=String(t.status||"").toLowerCase();return e==="open"?i==="open":!1}async function Nt(t,e){let i=e?e.get("dash_filter"):null;pt=await P(),qe=await U();let a=r=>r&&!qe.find(s=>s.value===r)?[...qe,{value:r,label:r}]:qe,o=new Date().getFullYear(),c=Array.from({length:5},(r,s)=>String(o-s));E({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:r=>i?r.filter(s=>Ia(s,i)):r,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:r=>`<span class="badge badge-secondary">${r}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:r=>`<span title="${r}">${r?.length>50?r.slice(0,50)+"\u2026":r}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:r=>`<span title="${r||""}">${r?.length>40?r.slice(0,40)+"\u2026":r||"-"}</span>`},{key:"status",label:"Status",render:r=>N(r)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:r=>window.formatDate(r)},{key:"day_count",label:"Hari",render:r=>r??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:pt},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:c}],formFields:r=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:r?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:pt,value:r?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:r?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:r?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:r?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(r?.employee_name),value:r?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(r?.fc_specialist),value:r?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:r?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:r?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:r?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let r=await f(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let s=r.data.data.map(n=>({Tanggal:n.report_date||"",Cabang:n.branch_name||"",Kategori:n.category||"",Sumber:n.source||"",Keluhan:n.complaint||"","Nama FC":n.employee_name||"","FC Spesialis":n.fc_specialist||"",Solusi:n.solution||"","Tgl Selesai":n.completion_date||"",Status:n.status||""}));$(s,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async r=>{let n=(await f("/api/branches?all=1")).data?.data||[],l=p=>{if(!p)return null;let u=String(p||"").toLowerCase(),h=n.find(g=>String(g.full_name||"").toLowerCase()===u||String(g.code||"").toLowerCase()===u||String(g.name||"").toLowerCase()===u);return h?h.id:null},m=r.map(p=>({branch_id:l(String(p.Cabang||"").trim()),report_date:String(p.Tanggal||"").trim(),category:String(p.Kategori||"").trim(),source:String(p.Sumber||"").trim(),complaint:String(p.Keluhan||"").trim(),employee_name:String(p["Nama FC"]||"").trim(),fc_specialist:String(p["FC Spesialis"]||"").trim(),solution:String(p.Solusi||"").trim(),completion_date:String(p["Tgl Selesai"]||"").trim(),status:String(p.Status||"").trim()})).filter(p=>p.report_date&&p.complaint&&p.category),d=await f("/api/import/issues",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}}})}D();var Se=[];function Pa(t,e){let i=String(t.status||"").toLowerCase();return e==="pending"?i==="pending":!1}async function At(t,e){let i=e?e.get("dash_filter"):null;Se=await P();let a=await U(),o=["Ade","Berlin"],c=s=>s&&!a.find(n=>n.value===s)?[...a,{value:s,label:s}]:a,r=s=>s&&!o.find(n=>(typeof n=="object"?n.value:n)===s)?[...o,s]:o;E({container:t,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:s=>i?s.filter(n=>Pa(n,i)):s,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:s=>`<span title="${s||""}">${s?.length>50?s.slice(0,50)+"\u2026":s||"-"}</span>`},{key:"solution",label:"Solusi",render:s=>`<span title="${s||""}">${s?.length>40?s.slice(0,40)+"\u2026":s||"-"}</span>`},{key:"status",label:"Status",render:s=>N(s)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:s=>window.formatDate(s)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:s=>s?`<a href="${s}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"combobox",name:"branch_id",label:"Cabang",options:Se},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async s=>{let n=new URLSearchParams(s||{}).toString(),l=await f(`/api/one-on-one?limit=10000&${n}`);if(l.ok){let m=l.data.data.map(p=>({Tanggal:p.meeting_date||"",Cabang:p.branch_name||"","Nama Karyawan":p.employee_name||"",PIC:p.pic||"",Masalah:p.problem||"",Solusi:p.solution||"",Status:p.status||"","Tgl Selesai":p.completion_date||"",Dokumen:p.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(L(),Y));d(m,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let s=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(L(),Y));n(s,"Template_Import_OneOnOne")},onImport:async s=>{let n=p=>{if(!p)return null;let u=String(p||"").toLowerCase(),h=Se.find(g=>String(g.label||"").toLowerCase()===u);return h?h.value:null},l=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let u=String(p).trim();if(/^\d{4,5}$/.test(u)){let g=Number(u);if(g>2e4&&g<99999){let b=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);let h=u.split(/[\/\-\.]/);if(h.length===3){let[g,b,y]=h.map(v=>v.trim());if(g.length===4&&b.length<=2&&y.length<=2)return`${g}-${b.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&b.length<=2&&g.length<=2)return`${y}-${b.padStart(2,"0")}-${g.padStart(2,"0")}`}return u},m=s.map(p=>({meeting_date:l(p.Tanggal),employee_name:String(p["Nama Karyawan"]||"").trim(),branch_id:n(String(p.Cabang||"").trim()),pic:String(p.PIC||"").trim(),problem:String(p.Masalah||"").trim(),solution:String(p.Solusi||"").trim(),status:String(p.Status||"").trim(),completion_date:l(p["Tgl Selesai"]),document_link:String(p.Dokumen||"").trim()})).filter(p=>p.meeting_date&&p.employee_name&&p.branch_id),d=await f("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:s=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:s?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:s?.branch_id&&!Se.find(n=>n.value==s.branch_id)?[...Se,{value:s.branch_id,label:s.branch_name||s.branch_id}]:Se,createApi:{path:"/api/branches",field:"full_name"},value:s?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"combobox",required:!0,options:c(s?.employee_name),value:s?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:r(s?.pic),createApi:{path:"/api/pic",field:"name"},value:s?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:s?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:s?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:s?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:s?.document_link}]})}D();async function Ft(t){let e=await P(),i=await U(),a=["Ade","Berlin"],o=s=>s&&!i.find(n=>n.value===s)?[...i,{value:s,label:s}]:i,c=s=>s&&!a.find(n=>(typeof n=="object"?n.value:n)===s)?[...a,s]:a,r=Array.from({length:5},(s,n)=>String(new Date().getFullYear()-n));E({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:s=>window.formatDate(s)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:s=>{try{let n=JSON.parse(s);return Array.isArray(n)?n.join(", "):s||"-"}catch{return s||"-"}}},{key:"score",label:"Nilai",render:s=>s!=null?`<strong>${s}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:s=>s?`<a href="${s}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:r}],exportOptions:{moduleName:"training",onExport:async s=>{let n=new URLSearchParams(s||{}).toString(),l=await f(`/api/training?limit=10000&${n}`);if(l.ok){let m=l.data.data.map(p=>{let u=p.participants||"";try{let h=JSON.parse(u);u=Array.isArray(h)?h.join(", "):u}catch{}return{Tanggal:p.training_date||"",Batch:p.batch||"",Materi:p.subject||"",Cabang:p.branch_name||"",Trainer:p.trainer||"",Peserta:u,Nilai:p.score!==null&&p.score!==void 0?p.score:"",Dokumen:p.document_link||""}}),{downloadExcel:d}=await Promise.resolve().then(()=>(L(),Y));d(m,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let s=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(L(),Y));n(s,"Template_Import_Training")},onImport:async s=>{let n=p=>{if(!p)return null;let u=String(p||"").toLowerCase(),h=e.find(g=>String(g.label||"").toLowerCase()===u);return h?h.value:null},l=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let u=String(p).trim();if(/^\d{4,5}$/.test(u)){let g=Number(u);if(g>2e4&&g<99999){let b=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);let h=u.split(/[\/\-\.]/);if(h.length===3){let[g,b,y]=h.map(v=>v.trim());if(g.length===4&&b.length<=2&&y.length<=2)return`${g}-${b.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&b.length<=2&&g.length<=2)return`${y}-${b.padStart(2,"0")}-${g.padStart(2,"0")}`}return u},m=s.map(p=>({training_date:l(p.Tanggal),batch:String(p.Batch||"").trim(),subject:String(p.Materi||"").trim(),branch_id:n(String(p.Cabang||"").trim()),trainer:String(p.Trainer||"").trim(),participants:String(p.Peserta||"").trim(),score:p.Nilai?Number(p.Nilai):null,document_link:String(p.Dokumen||"").trim()})).filter(p=>p.training_date&&p.subject&&p.branch_id),d=await f("/api/import/training",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:s=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:s?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:s?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:s?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:s?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:c(s?.trainer),value:s?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let n=JSON.parse(s?.participants);return Array.isArray(n)?n.join(", "):s?.participants||""}catch{return s?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:s?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:s?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:s?.notes}],onBeforeSubmit:async s=>(s.participants&&(s.participants=JSON.stringify(s.participants.split(",").map(n=>n.trim()).filter(Boolean))),s)})}D();L();async function Ot(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await P(),a=await U(),o=e?e.get("dash_filter"):null;console.log("RAW",await ye()),console.log("OPTIONS",a);let c=n=>n&&!a.find(l=>l.value===n)?[...a,{value:n,label:n}]:a,r=["Agung Septiadi","Wasrikin","IQBAL AL BANNA"],s=n=>n&&!r.includes(n)?[...r,n]:r;E({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",onDataLoaded:n=>{if(o==="reliever"){let l=new Date,m=l.getFullYear(),d=String(l.getMonth()+1).padStart(2,"0");return n.filter(p=>{if(String(p.status||"").toLowerCase()!=="done")return!1;let u=p.backup_date||"";if(u.includes("/")){let h=u.split("/");if(h.length===3&&(h[2].length===4?h[2]:`20${h[2]}`)==m&&h[1].padStart(2,"0")==d)return!0}else if(u.includes("-")&&u.startsWith(`${m}-${d}`))return!0;return!1})}return n},columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:n=>ae(n)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:n=>window.formatDate(n)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:n=>n?`<span class="badge badge-info">${n}</span>`:"-"},{key:"status",label:"Status",render:n=>N(n)}],filterFields:[{type:"search",placeholder:"Cari reliefer / FC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:n?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:c(n?.original_fc_name),value:n?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:s(n?.reliever_name),value:n?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:n?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:n?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:n?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:n?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let n=await f(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let l=n.data.data.map(m=>({Cabang:m.branch_name||"","Nama Facility care":m.original_fc_name||"",Periode:m.period||"",Relifer:m.reliever_name||"","Tanggal Back Up":m.backup_date||"","Tanggal Selesai":m.completion_date||"",Keterangan:m.reason||"",Shift:m.shift||"",Status:m.status||""}));l.length===0&&l.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),$(l,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async n=>{let m=(await f("/api/branches?all=1")).data?.data||[],d=h=>{if(!h)return null;let g=String(h||"").toLowerCase(),b=m.find(y=>String(y.full_name||"").toLowerCase()===g||String(y.code||"").toLowerCase()===g||String(y.name||"").toLowerCase()===g);return b?b.id:null},p=n.map(h=>({branch_name:String(h.Cabang||"").trim(),backup_date:String(h["Tanggal Back Up"]||h["Tanggal Backup"]||"").trim(),original_fc_name:String(h["Nama Facility care"]||h["FC Digantikan"]||"").trim(),reliever_name:String(h.Relifer||h.Reliefer||"").trim(),period:String(h.Periode||"").trim(),reason:String(h.Keterangan||"").trim(),shift:String(h.Shift||"").trim(),completion_date:String(h["Tanggal Selesai"]||"").trim(),status:String(h.Status||"").trim()})).filter(h=>h.reliever_name&&h.backup_date),u=await f("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:p,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}D();L();async function Mt(t){let e=await P(),i=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));E({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>ae(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>N(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),c=await f(`/api/reports/inspection?limit=10000&${o}`);if(c.ok){let r=c.data.data.map(s=>({Cabang:s.branch_name||"",Periode:s.period||"",Tanggal:s.inspection_date||"","Point FC":s.fc_score!==null&&s.fc_score!==void 0?s.fc_score:"","Point SPV":s.spv_score!==null&&s.spv_score!==void 0?s.spv_score:"",Status:s.status||"","Link Dokumen":s.document_link||""}));$(r,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let o=n=>{if(!n)return null;let l=String(n||"").toLowerCase(),m=e.find(d=>String(d.label||"").toLowerCase()===l);return m?m.value:null},c=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let l=String(n).trim();if(l===""||l==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);if(/^\d{4,5}$/.test(l)){let d=Number(l);if(d>2e4&&d<99999){let p=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let m=l.split(/[\/\-\.]/);if(m.length===3){let[d,p,u]=m.map(h=>h.trim());if(d.length===4&&p.length<=2&&u.length<=2)return`${d}-${p.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&p.length<=2&&d.length<=2)return`${u}-${p.padStart(2,"0")}-${d.padStart(2,"0")}`}return l},r=a.map(n=>({branch_id:o(String(n.Cabang||"").trim()),period:String(n.Periode||"").trim(),inspection_date:c(n.Tanggal),fc_score:n["Point FC"]!==void 0&&n["Point FC"]!==""?Number(n["Point FC"]):null,spv_score:n["Point SPV"]!==void 0&&n["Point SPV"]!==""?Number(n["Point SPV"]):null,status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.inspection_date),s=await f("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}D();L();async function Kt(t){let e=await P(),i=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));E({container:t,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>ae(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>N(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),c=await f(`/api/reports/cleaning?limit=10000&${o}`);if(c.ok){let r=c.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||""}));$(r,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let o=n=>{if(!n)return null;let l=String(n||"").toLowerCase(),m=e.find(d=>String(d.label||"").toLowerCase()===l);return m?m.value:null},c=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let l=String(n).trim();if(l===""||l==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);if(/^\d{4,5}$/.test(l)){let d=Number(l);if(d>2e4&&d<99999){let p=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let m=l.split(/[\/\-\.]/);if(m.length===3){let[d,p,u]=m.map(h=>h.trim());if(d.length===4&&p.length<=2&&u.length<=2)return`${d}-${p.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&p.length<=2&&d.length<=2)return`${u}-${p.padStart(2,"0")}-${d.padStart(2,"0")}`}return l},r=a.map(n=>({branch_id:o(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:c(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),s=await f("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}D();L();async function Rt(t){let e=await P(),i=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));E({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge badge-warning">${a}</span>`},{key:"period",label:"Periode",render:a=>ae(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>N(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:a=>a||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:a?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),c=await f(`/api/reports/fogging?limit=10000&${o}`);if(c.ok){let r=c.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"Fogging",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||""}));$(r,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async a=>{let o=n=>{if(!n)return null;let l=String(n||"").toLowerCase(),m=e.find(d=>String(d.label||"").toLowerCase()===l);return m?m.value:null},c=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let l=String(n).trim();if(l===""||l==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);if(/^\d{4,5}$/.test(l)){let d=Number(l);if(d>2e4&&d<99999){let p=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let m=l.split(/[\/\-\.]/);if(m.length===3){let[d,p,u]=m.map(h=>h.trim());if(d.length===4&&p.length<=2&&u.length<=2)return`${d}-${p.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&p.length<=2&&d.length<=2)return`${u}-${p.padStart(2,"0")}-${d.padStart(2,"0")}`}return l},r=a.map(n=>({branch_id:o(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"Fogging").trim(),period:String(n.Periode||"").trim(),activity_date:c(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.activity_date),s=await f("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(r)});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}D();L();async function qt(t){let e=await P(),i=await U(),a=i,o=r=>r&&!i.find(s=>s.value===r)?[...i,{value:r,label:r}]:i,c=r=>r&&!a.find(s=>s.value===r)?[...a,{value:r,label:r}]:a;E({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:r=>`<span title="${r||""}">${r?.length>60?r.slice(0,60)+"\u2026":r||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:r=>window.formatDate(r)},{key:"status",label:"Status",render:r=>N(r)},{key:"notes",label:"Keterangan",render:r=>r?.length>40?r.slice(0,40)+"\u2026":r||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]}],formFields:r=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:r?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:c(r?.pic),value:r?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:r?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:r?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:r?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:r?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async r=>{let s=new URLSearchParams(r||{}).toString(),n=await f(`/api/reports/basecamp?limit=10000&${s}`);if(n.ok){let l=n.data.data.map(m=>({"Tgl Info":m.info_date||"",Cabang:m.branch_name||"",Permasalahan:m.problem||"",PIC:m.pic||"","Tgl Done":m.done_date||"",Status:m.status||"",Keterangan:m.notes||""}));$(l,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async r=>{let s=d=>{if(!d)return null;let p=String(d||"").toLowerCase(),u=e.find(h=>String(h.label||"").toLowerCase()===p);return u?u.value:null},n=d=>{if(d==null||d==="")return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let p=String(d).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let h=Number(p);if(h>2e4&&h<99999){let g=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}let u=p.split(/[\/\-\.]/);if(u.length===3){let[h,g,b]=u.map(y=>y.trim());if(h.length===4&&g.length<=2&&b.length<=2)return`${h}-${g.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&g.length<=2&&h.length<=2)return`${b}-${g.padStart(2,"0")}-${h.padStart(2,"0")}`}return p},l=r.map(d=>({info_date:n(d["Tgl Info"]||d["Tanggal Info"]),branch_id:s(String(d.Cabang||"").trim()),problem:String(d.Permasalahan||"").trim(),pic:String(d.PIC||"").trim(),done_date:n(d["Tgl Done"]||d["Tanggal Done"]),status:String(d.Status||"").trim(),notes:String(d.Keterangan||d.Catatan||"").trim()})).filter(d=>d.info_date&&d.branch_id&&d.problem),m=await f("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(l)});if(!m.ok)throw new Error(m.data?.error||"Import gagal");return m.data}}})}async function Ht(t){E({container:t,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(D(),ge)),o=await a(`/api/sop?limit=10000&${i}`);if(o.ok){let c=o.data.data.map(s=>({"Nama SOP":s.name||"",Kategori:s.category||"",Dokumen:s.document_link||"",Catatan:s.notes||s.description||""})),{downloadExcel:r}=await Promise.resolve().then(()=>(L(),Y));r(c,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(L(),Y));i(e,"Template_Import_SOP")},onImport:async e=>{let i=e.map(c=>({name:String(c["Nama SOP"]||"").trim(),category:String(c.Kategori||"").trim(),document_link:String(c.Dokumen||"").trim(),description:String(c.Catatan||"").trim()})).filter(c=>c.name),{apiFetch:a}=await Promise.resolve().then(()=>(D(),ge)),o=await a("/api/sop/import",{method:"POST",body:JSON.stringify(i)});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function jt(t){E({container:t,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(D(),ge)),o=await a(`/api/checklist?limit=10000&${i}`);if(o.ok){let c=o.data.data.map(s=>({"Nama Checklist":s.name||"",Kategori:s.category||"",Dokumen:s.document_link||"",Deskripsi:s.description||""})),{downloadExcel:r}=await Promise.resolve().then(()=>(L(),Y));r(c,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(L(),Y));i(e,"Template_Import_Checklist")},onImport:async e=>{let i=e.map(c=>({name:String(c["Nama Checklist"]||"").trim(),category:String(c.Kategori||"").trim(),document_link:String(c.Dokumen||"").trim(),description:String(c.Deskripsi||"").trim()})).filter(c=>c.name),{apiFetch:a}=await Promise.resolve().then(()=>(D(),ge)),o=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(i)});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}D();Pe();L();async function mt(t,e="forms"){if(e==="supply")return Ba(t);La(t)}function La(t){E({container:t,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}]})}async function Ba(t){let i=((await f("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));E({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,o)=>o.branch_name_ref||o.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let o=JSON.parse(a);return Array.isArray(o)?o.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let o=JSON.parse(a);return Array.isArray(o)?o.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>N(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let o=a?.tools_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}let c=a?.chemical_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:a?.branch_id&&!i.find(r=>r.value==a.branch_id)?[...i,{value:a.branch_id,label:a.branch_name||a.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:o},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:c},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),c=await f(`/api/reports/supply?limit=10000&${o}`);if(c.ok){let r=c.data.data.map(s=>{let n=s.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let l=s.chemical_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}return{Waktu:s.submitted_at||"",Pengirim:s.submitter_name||"",Cabang:s.branch_name_ref||s.branch_name||"","Alat/Barang":n||"",Chemical:l||"",Catatan:s.additional_notes||"",Status:s.status||"","Diproses Oleh":s.processed_by||""}});$(r,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let c=(await f("/api/branches?all=1")).data?.data||[],r=m=>{if(!m)return null;let d=String(m||"").toLowerCase(),p=c.find(u=>String(u.full_name||"").toLowerCase()===d||String(u.code||"").toLowerCase()===d||String(u.name||"").toLowerCase()===d);return p?p.id:null},s=m=>{if(m==null||m==="")return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let d=String(m).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let u=Number(d);if(u>2e4&&u<99999){let h=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}let p=d.split(/[\/\-\.]/);if(p.length===3){let[u,h,g]=p.map(b=>b.trim());if(u.length===4&&h.length<=2&&g.length<=2)return`${u}-${h.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&h.length<=2&&u.length<=2)return`${g}-${h.padStart(2,"0")}-${u.padStart(2,"0")}`}return d},n=a.map(m=>({submitted_at:s(m.Waktu||m.Tanggal),submitter_name:String(m.Pengirim||"").trim(),branch_id:r(String(m.Cabang||"").trim()),tools_items:String(m["Alat/Barang"]||m.Alat||"").trim(),chemical_items:String(m.Chemical||"").trim(),additional_notes:String(m.Catatan||m.Keterangan||"").trim(),status:String(m.Status||"").trim(),processed_by:String(m["Diproses Oleh"]||m.PIC||"").trim()})).filter(m=>m.submitted_at&&m.submitter_name&&m.branch_id),l=await f("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,o)=>{let c=le({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(r,s)=>{let n=r.querySelector("#supply-status").value,l=r.querySelector("#supply-processed-by").value;(await f(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:l})})).ok?(z("Status diperbarui."),s(),o()):Q("Gagal update status.")}})}}]})}D();L();async function Ut(t){let e=ue();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}E({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let a=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await f(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let a=i.data.data.map(o=>({"Nama Lengkap":o.full_name||"",Username:o.username||"",Email:o.email||"",Role:o.role||"",Status:o.is_active?"Aktif":"Nonaktif"}));$(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let a=i.map(c=>({full_name:String(c["Nama Lengkap"]||"").trim(),username:String(c.Username||"").trim(),email:String(c.Email||"").trim(),role:String(c.Role||"").trim()||"viewer",password:String(c.Password||"").trim()})).filter(c=>c.username&&c.password&&c.email&&c.full_name),o=await f("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}}})}D();L();async function Jt(t){E({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await f(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)$(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{$([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let i=e.map(o=>({code:String(o["Kode Cabang"]||"").trim(),name:String(o["Nama Pendek"]||"").trim(),full_name:String(o["Nama Lengkap"]||"").trim(),city:String(o.Kota||"").trim()})).filter(o=>o.code&&o.name),a=await f("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}D();async function Gt(t){let e=new Date,i=[];t.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),o()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),o()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(c=>c.addEventListener("change",o));async function a(){try{let c=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;i=(await f(`/api/dashboard/calendar?month=${c}`)).data?.data||[]}catch(c){console.warn("[Calendar] Failed to load events, rendering empty grid:",c),i=[]}}async function o(){let c=document.getElementById("calendar-grid");if(c){c.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let r=e.getFullYear(),s=e.getMonth(),n=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),l=document.getElementById("cal-month-label");l&&(l.textContent=n);let m=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(S=>S.value)),d=i.filter(S=>m.has(S.type)),p={};d.forEach(S=>{let x=(S.event_date||"").slice(0,10);p[x]||(p[x]=[]),p[x].push(S)});let u=new Date(r,s,1).getDay(),h=new Date(r,s+1,0).getDate(),g=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],b=new Date().toISOString().slice(0,10),y='<div class="calendar-grid">';g.forEach(S=>{y+=`<div class="cal-day-header">${S}</div>`});for(let S=0;S<u;S++)y+='<div class="cal-cell cal-cell-empty"></div>';for(let S=1;S<=h;S++){let x=`${r}-${String(s+1).padStart(2,"0")}-${String(S).padStart(2,"0")}`,T=p[x]||[],B=x===b;y+=`
          <div class="cal-cell ${B?"cal-today":""} ${T.length?"cal-has-events":""}"
               data-date="${x}" tabindex="0" role="button" aria-label="${x}">
            <div class="cal-day-num ${B?"today-num":""}">${S}</div>
            <div class="cal-events-preview">
              ${T.slice(0,3).map(C=>`
                <div class="cal-event-dot cal-color-${C.color||"gray"}" title="${He(C.title||C.type)}">
                  <span class="cal-event-dot-label">${Na(C.title||C.branch_name||C.type,18)}</span>
                </div>
              `).join("")}
              ${T.length>3?`<div class="cal-more">+${T.length-3} lagi</div>`:""}
            </div>
          </div>`}let k=(u+h)%7;if(k!==0)for(let S=0;S<7-k;S++)y+='<div class="cal-cell cal-cell-empty"></div>';y+="</div>",c.innerHTML=y,c.querySelectorAll(".cal-cell[data-date]").forEach(S=>{S.addEventListener("click",()=>{let x=S.dataset.date,T=p[x]||[];if(!T.length)return;let B=document.getElementById("cal-event-list"),C=new Date(x+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=C,document.getElementById("cal-event-items").innerHTML=T.map(I=>`
            <div class="cal-event-item cal-color-border-${I.color||"gray"}">
              <div class="cal-event-type">${Aa(I.type)}</div>
              <div class="cal-event-title">${He(I.title||"-")}</div>
              <div class="cal-event-branch">${He(I.branch_name||"")}</div>
              ${I.status?`<div class="cal-event-status">${He(I.status)}</div>`:""}
              ${I.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${I.days_remaining} hari</div>`:""}
            </div>
          `).join(""),B.style.display="block"})})}catch(r){console.error("[Calendar] Render error:",r),c&&(c.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}o()}function Na(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function He(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Aa(t){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[t]||t}D();async function Qt(t){let e=ue(),i=(e?.full_name||e?.username||"U")[0].toUpperCase(),o={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
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
  `;let c=localStorage.getItem("fm_token"),r=document.getElementById("session-info");if(c&&r)try{let s=JSON.parse(atob(c.split(".")[1])),n=new Date(s.exp*1e3);r.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{r.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async s=>{s.preventDefault();let n=document.getElementById("pwd-error"),l=document.getElementById("pwd-success"),m=document.getElementById("btn-save-pwd");n.style.display="none",l.style.display="none";let d=s.target,p=d.current_password.value,u=d.new_password.value,h=d.confirm_password.value;if(u!==h){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(u.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}m.disabled=!0,m.textContent="\u23F3 Menyimpan...";let g=await f("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:p,new_password:u})});m.disabled=!1,m.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',g.ok?(l.textContent="\u2705 Password berhasil diubah.",l.style.display="block",d.reset(),z("Password berhasil diubah.")):(n.textContent=g.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}D();var je={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function J(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let o=Number(e);if(o>2e4&&o<99999){let c=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(c.getTime())?null:c.toISOString().slice(0,10)}}let i=e.split(/[\/\-\.]/);if(i.length===3){let[o,c,r]=i.map(m=>m.trim()),s=Number(o),n=Number(c),l=Number(r);if(o.length===4&&s>1900)return`${o}-${c.padStart(2,"0")}-${r.padStart(2,"0")}`;if(r.length===4&&l>1900)return s>12?`${r}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`:n>12?`${r}-${o.padStart(2,"0")}-${c.padStart(2,"0")}`:`${r}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`;if(r.length===2&&!isNaN(l)){let m=l>=50?`19${r}`:`20${r}`;return s>12?`${m}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`:`${m}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function Vt(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Fa={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",phone:t["No. Hp"],join_date:J(t["Tanggal Masuk"]),status:t.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:J(t["Tanggal Mulai"]),end_date:J(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:J(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:J(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:J(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:J(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:J(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:J(t["Tanggal Target"]||t["Tgl Target"]),completion_date:J(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:J(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:J(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:J(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:J(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:J(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:J(t["Tanggal Back Up"]),completion_date:J(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:J(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:J(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Oa(t,e){let i=je[t];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Fa[i.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let o=[],c=[],r=[];return e.filter(n=>!Vt(n)).forEach((n,l)=>{let m=e.indexOf(n)+2,d=[];a.required.forEach(({key:u,label:h})=>{let g=n[u];if(g==null||String(g).trim()===""){let b=Object.keys(n).filter(y=>y.trim()).join(", ");d.push({column:h,originalValue:g||"",reason:`Kolom "${h}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${b.slice(0,120)}`})}});let p=a.map(n);d.length>0?c.push({row:m,data:p,raw:n,errors:d}):(o.push(n),r.push(p))}),{valid:o,errors:c,mapped:r}}function Wt(t){let e=[];return t.SheetNames.forEach(i=>{let a=je[i];if(!a)return;let o=t.Sheets[i],c=window.XLSX.utils.sheet_to_json(o,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),r=Oa(i,c),s=c.filter(n=>!Vt(n));e.push({sheetName:i,module:a.module,label:a.label,total:s.length,valid:r.mapped.length,errorCount:r.errors.length,errors:r.errors,mapped:r.mapped,skipped:!1})}),e}function zt(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu","NAMA KARYAWAN":"Budi Santoso",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02",Day:"1"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06",Day:"1","Link Document":""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}]}).forEach(([a,o])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(o),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function Yt(t){let e=window.XLSX,i=e.utils.book_new(),a=!1;return t.forEach(o=>{if(!o.errors||o.errors.length===0)return;a=!0;let c=o.errors.map(s=>({"No. Baris":s.row,"Kolom Gagal":(s.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(s.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(s.data||{}).map(([n,l])=>[n,l??""]))})),r=e.utils.json_to_sheet(c);e.utils.book_append_sheet(i,r,o.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Ma=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function Xt(t){t.innerHTML=`
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
              ${Object.entries(je).map(([g,{label:b}])=>`<span class="import-sheet-tag">\u{1F4C4} ${g} \u2192 ${b}</span>`).join("")}
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
  `;let e=null,i=null,a=0,o={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function c(g){Object.entries(o).forEach(([b,y])=>{y.style.display=b===g?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let g=document.getElementById("btn-backup-db");g.disabled=!0,g.textContent="\u23F3 Memproses Backup...";try{let b=await f("/api/import/backup");if(b.ok){let y=new Blob([JSON.stringify(b.data,null,2)],{type:"application/json"}),v=URL.createObjectURL(y),k=document.createElement("a");k.href=v,k.download=`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.json`,document.body.appendChild(k),k.click(),document.body.removeChild(k),URL.revokeObjectURL(v),z("Backup berhasil diunduh!")}else Q("Gagal memproses backup: "+(b.data?.error||"Unknown error"))}catch(b){Q("Gagal memproses backup: "+b.message)}finally{g.disabled=!1,g.textContent="\u{1F4E6} Backup Database"}});let r=document.getElementById("btn-sync-google");r&&r.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let g=r.innerHTML;r.innerHTML='<span class="spinner"></span> Menyinkronkan...',r.disabled=!0;try{let b=await f("/api/sync/google-sheets",{method:"POST"});b.ok?alert("Sinkronisasi Berhasil: "+(b.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(b.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{r.innerHTML=g,r.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{zt(),z("Template Excel berhasil didownload!")});let s=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",g=>{g.stopPropagation(),s.click()}),s.addEventListener("change",g=>{g.target.files[0]&&l(g.target.files[0])}),n.addEventListener("dragover",g=>{g.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",g=>{g.preventDefault(),n.classList.remove("drag-over");let b=g.dataTransfer.files[0];b&&b.name.match(/\.xlsx?$/i)?l(b):Q("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,s.value="",document.getElementById("file-info").style.display="none",n.style.display="",c("upload")});async function l(g){e=g,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${g.name} (${(g.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await m(g)}async function m(g){c("validating");let b=document.getElementById("validation-status"),y=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");b.textContent="Membaca file Excel...",y.style.width="20%",await Ne(200);let v=await g.arrayBuffer(),k=window.XLSX.read(v,{type:"array",cellDates:!0});b.textContent=`Memvalidasi ${k.SheetNames.length} sheet...`,y.style.width="50%",await Ne(100),i=Wt(k),y.style.width="100%",b.textContent="Validasi selesai!",await Ne(300),d()}catch(v){c("upload"),Q("Gagal memproses file: "+v.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function d(){c("preview");let g=i.filter(C=>!C.skipped).length,b=i.reduce((C,I)=>C+I.total,0),y=i.reduce((C,I)=>C+I.valid,0),v=i.reduce((C,I)=>C+I.errorCount,0),k=b>0?Math.round(y/b*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${g} sheet</span>
      <span class="badge badge-secondary">${b} baris</span>
      <span class="badge badge-success">${y} valid (${k}%)</span>
      ${v>0?`<span class="badge badge-danger">${v} error</span>`:""}
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
          ${i.map((C,I)=>`
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
                ${C.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${I}">\u{1F50D} ${C.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,S.querySelectorAll(".btn-detail-error").forEach(C=>{C.addEventListener("click",()=>{let I=i[Number(C.dataset.idx)];p(I)})});let x=document.getElementById("error-detail-section"),T=document.getElementById("error-detail-container");T.innerHTML="",x.style.display="none";let B=document.getElementById("btn-start-import");y===0?(B.disabled=!0,B.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(B.disabled=!1,v>0?(B.innerHTML=`\u{1F680} Import ${y} Data Valid (${v} dilewati)`,B.title="Baris error akan dilewati, baris valid tetap diimport"):B.innerHTML=`\u{1F680} Mulai Import ${y} Data`)}function p(g){let b=document.getElementById("error-detail-section"),y=document.getElementById("error-detail-container");b.style.display="";let v=g.errors.slice(0,100).map(k=>(Array.isArray(k.errors)?k.errors:[]).map(x=>{let T=typeof x=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${k.row}</span></td>
            <td><strong>${T?x.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${T&&x.originalValue!==void 0?x.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${T?x.reason:x}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${T&&x.aliases?`Gunakan salah satu nama kolom:<br><em>${x.aliases}</em>`:T&&x.hint?x.hint:""}
            </td>
          </tr>
        `}).join("")).join("");y.innerHTML=`
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
    `,b.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{c("upload"),document.getElementById("file-info").style.display="none",n.style.display="",e=null,s.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;Yt(i)?z("Log error berhasil didownload."):z("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let g=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";u(g)});async function u(g){c("importing"),a=Date.now();let b=[];Ma.forEach(x=>{let T=i?.find(B=>B.module===x&&B.mapped?.length>0);T&&b.push(T)});let y=document.getElementById("import-steps-list");y.innerHTML=b.map(x=>`
      <div class="import-step-item" id="step-item-${x.module}">
        <span class="step-item-icon" id="step-icon-${x.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${x.label} <span class="step-item-count">(${x.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${x.module}"></span>
      </div>
    `).join("");let v=document.getElementById("import-bar"),k=document.getElementById("import-current-status"),S={totalSheets:b.length,totalRows:b.reduce((x,T)=>x+T.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let x=0;x<b.length;x++){let T=b[x],B=document.getElementById(`step-icon-${T.module}`),C=document.getElementById(`step-status-${T.module}`);B.textContent="\u{1F504}",C.textContent="Mengimport...",k.textContent=`Mengimport ${T.label}...`,v.style.width=`${Math.round(x/b.length*100)}%`;try{let I=await f(`/api/import/${T.module}`,{method:"POST",body:JSON.stringify({rows:T.mapped,onDuplicate:g})});if(I.ok){let K=I.data;S.inserted+=K.inserted||0,S.skipped+=K.skipped||0,S.moduleResults.push({label:T.label,inserted:K.inserted||0,skipped:K.skipped||0,status:"ok"}),B.textContent="\u2705",C.innerHTML=`<span class="badge badge-success">${K.inserted||0} berhasil</span>${K.skipped>0?` <span class="badge badge-neutral">${K.skipped} skip</span>`:""}`}else S.failed++,S.moduleResults.push({label:T.label,inserted:0,skipped:0,status:"error",error:I.data?.error}),B.textContent="\u274C",C.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(I){S.failed++,S.moduleResults.push({label:T.label,inserted:0,skipped:0,status:"error",error:I.message}),B.textContent="\u274C",C.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Ne(150)}v.style.width="100%",k.textContent="Selesai!",await Ne(400),h(S)}function h(g){c("summary");let b=((Date.now()-a)/1e3).toFixed(1),y=g.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${y?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${y?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
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
        <div class="stat-value">${b}s</div>
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,i=null,s.value="",document.getElementById("file-info").style.display="none",n.style.display="",c("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Ne(t){return new Promise(e=>setTimeout(e,t))}D();var Ue=[],Zt=[];async function ea(t){Ue=await P(),Zt=await U(),E({container:t,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:Ue}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await f(`/api/sp?limit=10000&${i}`);if(a.ok){let o=a.data.data.map(r=>({"Nama Karyawan":r.employee_name||"",Divisi:r.division||"",Cabang:r.branch_name||"","Tanggal Sp":r.tanggal||"","Akhir Sp":r.akhir_sp||"","Jenis Sp":r.sp_type||"","Link Document / Foto":r.document_link||""})),{downloadExcel:c}=await Promise.resolve().then(()=>(L(),Y));c(o,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(L(),Y));i(e,"Template_Import_SP")},onImport:async e=>{let i=r=>{if(!r)return null;let s=String(r||"").toLowerCase(),n=Ue.find(l=>String(l.label||"").toLowerCase()===s);return n?n.value:null},a=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let s=String(r).trim();if(/^\d{4,5}$/.test(s)){let l=Number(s);if(l>2e4&&l<99999){let m=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);let n=s.split(/[\/\-\.]/);if(n.length===3){let[l,m,d]=n.map(p=>p.trim());if(l.length===4&&m.length<=2&&d.length<=2)return`${l}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&m.length<=2&&l.length<=2)return`${d}-${m.padStart(2,"0")}-${l.padStart(2,"0")}`}return s},o=e.map(r=>({employee_name:String(r["Nama Karyawan"]||"").trim(),division:String(r.Divisi||"").trim(),branch_id:i(String(r.Cabang||"").trim()),tanggal:a(r["Tanggal Sp"]),akhir_sp:a(r["Akhir Sp"]),sp_type:String(r["Jenis Sp"]||"").trim(),document_link:String(r["Link Document / Foto"]||"").trim()})).filter(r=>r.employee_name&&r.branch_id),c=await f("/api/import/sp",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},formFields:[{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:Zt},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"combobox",name:"branch_id",label:"Cabang",required:!0,options:Ue,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}D();var xe=[],ta=[];async function aa(t){xe=await P(),ta=await U(),E({container:t,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",options:xe},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",options:xe}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await f(`/api/mutasi?limit=10000&${i}`);if(a.ok){let o=a.data.data.map(r=>({Tanggal:r.tanggal||"","Nama Karyawan":r.employee_name||"","Cabang Asal":r.from_branch_name||"","Cabang Tujuan":r.to_branch_name||"",Status:r.status||"",Dokumen:r.document_link||""})),{downloadExcel:c}=await Promise.resolve().then(()=>(L(),Y));c(o,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(L(),Y));i(e,"Template_Import_Mutasi")},onImport:async e=>{let i=r=>{if(!r)return null;let s=String(r||"").toLowerCase(),n=xe.find(l=>String(l.label||"").toLowerCase()===s);return n?n.value:null},a=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let s=String(r).trim();if(/^\d{4,5}$/.test(s)){let l=Number(s);if(l>2e4&&l<99999){let m=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);let n=s.split(/[\/\-\.]/);if(n.length===3){let[l,m,d]=n.map(p=>p.trim());if(l.length===4&&m.length<=2&&d.length<=2)return`${l}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&m.length<=2&&l.length<=2)return`${d}-${m.padStart(2,"0")}-${l.padStart(2,"0")}`}return s},o=e.map(r=>({tanggal:a(r.Tanggal),employee_name:String(r["Nama Karyawan"]||"").trim(),from_branch_id:i(String(r["Cabang Asal"]||"").trim()),to_branch_id:i(String(r["Cabang Tujuan"]||"").trim()),status:String(r.Status||"").trim(),document_link:String(r.Dokumen||"").trim()})).filter(r=>r.tanggal&&r.employee_name&&r.from_branch_id&&r.to_branch_id),c=await f("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:ta},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:xe,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:xe,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}D();async function na(t){t.innerHTML=`
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
  `;let e=document.getElementById("btnRefreshSync"),i=document.getElementById("queueStatusFilter");e.addEventListener("click",o),i.addEventListener("change",s),document.getElementById("btnRetryAll").addEventListener("click",()=>a("retry",{allFailed:!0})),document.getElementById("btnResetStuck").addEventListener("click",()=>a("reset-stuck",{})),document.getElementById("btnPauseSync").addEventListener("click",()=>a("pause",{})),document.getElementById("btnResumeSync").addEventListener("click",()=>a("resume",{})),document.getElementById("btnReconcile").addEventListener("click",()=>a("reconcile",{module:"ALL",repairMode:!0})),document.getElementById("btnRetrySelected").addEventListener("click",()=>{let l=Array.from(document.querySelectorAll(".chk-queue:checked")).map(m=>m.value);if(l.length===0)return alert("No items selected");a("retry",{ids:l})}),document.getElementById("chkAllQueue").addEventListener("change",l=>{document.querySelectorAll(".chk-queue").forEach(m=>m.checked=l.target.checked)});async function a(l,m){if(confirm(`Are you sure you want to execute action: ${l}?`)){showLoading();try{let d=await f(`/api/sync/actions/${l}`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(m)});d.ok?(alert(d.data?.message||"Success"),o()):Q(d.error||"Action failed")}catch(d){Q(d.message)}hideLoading()}}await o();async function o(){showLoading(),await Promise.all([r(),s(),c(),n()]),hideLoading()}async function c(){try{let l=await f("/api/sync/performance");if(!l.ok)return;let{webhook:m,google_api:d,d1:p,queue:u,throughput:h}=l.data;document.getElementById("latencyHistograms").innerHTML=`
        <table class="w-full text-left">
          <thead><tr class="text-gray-500 border-b"><th>Metric</th><th>P50</th><th>P95</th><th>P99</th><th>Max</th></tr></thead>
          <tbody class="divide-y">
            <tr><td class="py-1">Webhook</td><td>${m.P50}ms</td><td>${m.P95}ms</td><td>${m.P99}ms</td><td>${m.Max}ms</td></tr>
            <tr><td class="py-1">Google API</td><td>${d.P50}ms</td><td>${d.P95}ms</td><td>${d.P99}ms</td><td>${d.Max}ms</td></tr>
            <tr><td class="py-1">D1 Execute</td><td>${p.P50}ms</td><td>${p.P95}ms</td><td>${p.P99}ms</td><td>${p.Max}ms</td></tr>
            <tr><td class="py-1">Queue Delay</td><td>${u.P50}ms</td><td>${u.P95}ms</td><td>${u.P99}ms</td><td>${u.Max}ms</td></tr>
          </tbody>
        </table>
        <div class="mt-4 pt-3 border-t text-gray-600 flex justify-between">
          <span>Throughput: <b>${h.events_per_sec}</b> ev/sec</span>
          <span><b>${h.events_per_min}</b> ev/min</span>
        </div>
      `,document.getElementById("lblSnapshotStatus").innerHTML='Checked 10 modules. Status: <b>COMPLETED</b><br><span class="text-xs text-gray-500">Run today at 02:00</span>'}catch(l){console.error(l)}}async function r(){try{let l=await f("/api/sync/health");if(!l.ok)return Q("Failed to fetch sync health");let{status:m,queue:d,circuit_breaker:p}=l.data,u=`
        <div class="bg-white p-4 rounded-lg shadow border-l-4 ${m==="HEALTHY"?"border-green-500":m==="WARNING"?"border-yellow-500":"border-red-500"}">
          <p class="text-sm text-gray-500">System Health</p>
          <p class="text-2xl font-bold ${m==="HEALTHY"?"text-green-600":m==="WARNING"?"text-yellow-600":"text-red-600"}">${m}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-blue-500">
          <p class="text-sm text-gray-500">Pending Queue</p>
          <p class="text-2xl font-bold text-gray-800">${d.pending||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-orange-500">
          <p class="text-sm text-gray-500">Failed / Retrying</p>
          <p class="text-2xl font-bold text-gray-800">${d.failed||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-red-500">
          <p class="text-sm text-gray-500">Dead Letters</p>
          <p class="text-2xl font-bold text-gray-800">${d.dead_letter||0}</p>
        </div>
      `;document.getElementById("syncOverviewCards").innerHTML=u;let h=document.getElementById("cbStateBadge"),g=document.getElementById("cbStateDesc"),b=document.getElementById("cbStatusCard");b.className="bg-white rounded-lg shadow p-6 border-l-4",p==="CLOSED"?(h.className="px-3 py-1 rounded-full text-sm font-medium bg-green-100 text-green-800",h.textContent="CLOSED",g.textContent="Traffic is flowing normally to Google Sheets.",b.classList.add("border-green-500")):p==="OPEN"?(h.className="px-3 py-1 rounded-full text-sm font-medium bg-red-100 text-red-800",h.textContent="OPEN",g.textContent="Failing fast. Traffic to Google Sheets is paused due to repeated failures.",b.classList.add("border-red-500")):p==="HALF_OPEN"?(h.className="px-3 py-1 rounded-full text-sm font-medium bg-yellow-100 text-yellow-800",h.textContent="HALF-OPEN",g.textContent="Testing recovery. Permitting limited traffic to verify stability.",b.classList.add("border-yellow-500")):h.textContent=p||"UNKNOWN"}catch(l){console.error(l)}}async function s(){try{let l=document.getElementById("queueStatusFilter").value,m=await f("/api/sync/queue?limit=15"+(l?"&status="+l:""));if(!m.ok)return;let d=document.getElementById("queueTableBody"),p=m.data?.data||m.data||[];if(p.length===0){d.innerHTML='<tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">No events found</td></tr>';return}d.innerHTML=p.map(u=>`
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
      `).join("")}catch(l){console.error(l)}}async function n(){try{let l=await f("/api/sync/metrics");if(!l.ok)return;let m=document.getElementById("metricsTableBody"),d=l.data||[];if(d.length===0){m.innerHTML='<tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">No metrics found in last 24h</td></tr>';return}m.innerHTML=d.map(p=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2 font-medium text-gray-800">${p.module}</td>
          <td class="px-4 py-2 text-gray-600">${p.total_events}</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(p.avg_webhook_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(p.avg_d1_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(p.avg_queue_wait_ms||0)} ms</td>
        </tr>
      `).join("")}catch(l){console.error(l)}}}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let e=Math.floor(Number(t)-25569);return new Date(e*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let e=t.split(/[\/\-]/);return`${e[2]}-${e[1]}-${e[0]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let i=e.split("-");if(i.length===3&&i[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],o=parseInt(i[2],10),c=a[parseInt(i[1],10)-1];return`${o} ${c} ${i[0]}`}return e};function M(t){return async e=>{if(!Ee()){be("/login");return}return t(e)}}var Ae=null;function Ka(){Ae&&clearInterval(Ae);let t=()=>{let e=new Date,i=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),o=document.getElementById("header-clock-time"),c=document.getElementById("header-clock-date");o&&(o.textContent=i),c&&(c.textContent=a)};t(),Ae=setInterval(t,1e3)}async function Ra(){try{let t=await f("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},i=(a,o)=>{let c=document.getElementById(a);c&&(c.textContent=o>0?o:"",c.style.display=o>0?"inline-flex":"none")};i("badge-issues",e.issues?.current||0),i("badge-contracts",e.expiring30?.current||0),i("badge-oo1",e.one_on_one?.current||0),i("badge-schedule",e.schedule?.current||0),i("badge-supply",e.supply?.current||0)}catch{}}var we=[];async function qa(){try{let t=await f("/api/dashboard/notifications");if(!t.ok)return;we=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=we.length>0?"block":"none",e.textContent=we.length)}catch{}}function Ha(){if(!we.length){le({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,i)=>i()});return}let t=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${we.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;le({title:`Notifikasi (${we.length})`,content:t,confirmText:"Tutup",onConfirm:(e,i)=>i()})}function ia(){let t=ue(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
                <span class="topbar-greeting-time">${(()=>{let l=new Date().getHours();return l>=4&&l<11?"Selamat Pagi":l>=11&&l<15?"Selamat Siang":l>=15&&l<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">Berlin Ariansyah</span> \u{1F44B}
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
  `;let i=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),o=document.getElementById("topbar-menu-btn"),c=document.getElementById("sidebar-close"),r=()=>{i.classList.add("open"),a.classList.add("show")},s=()=>{i.classList.remove("open"),a.classList.remove("show")};o?.addEventListener("click",r),c?.addEventListener("click",s),a?.addEventListener("click",s),document.querySelectorAll(".nav-item").forEach(l=>l.addEventListener("click",s));function n(){let l=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(p=>{let u=p.dataset.route;p.classList.toggle("active",l===u||u!=="/dashboard"&&l.startsWith(u))});let m=document.getElementById("topbar-title"),d=document.querySelector(".nav-item.active .nav-label");m&&d&&(m.textContent=d.textContent)}window.addEventListener("hashchange",n),n(),Ka(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await f("/api/auth/logout",{method:"POST"}),$e(),Ae&&clearInterval(Ae),be("/login")}),Ra(),qa(),document.getElementById("btn-notif")?.addEventListener("click",l=>{l.preventDefault(),Ha()})}async function ja(){A("/login",({main:e})=>Dt(e)),A("/dashboard",M(({main:e})=>Ct(e))),A("/calendar",M(({main:e})=>Gt(e))),A("/employees",M(({main:e,params:i})=>It(e,i))),A("/contracts",M(({main:e,params:i})=>Lt(e,i))),A("/sp",M(({main:e})=>ea(e))),A("/mutasi",M(({main:e})=>aa(e))),A("/sync-dashboard",M(({main:e})=>na(e))),A("/timeline",M(({main:e,params:i})=>Bt(e,i))),A("/issues",M(({main:e,params:i})=>Nt(e,i))),A("/one-on-one",M(({main:e,params:i})=>At(e,i))),A("/training",M(({main:e})=>Ft(e))),A("/relievers",M(({main:e,params:i})=>Ot(e,i))),A("/reports/inspection",M(({main:e})=>Mt(e))),A("/reports/cleaning",M(({main:e})=>Kt(e))),A("/reports/fogging",M(({main:e})=>Rt(e))),A("/reports/basecamp",M(({main:e})=>qt(e))),A("/reports/supply",M(({main:e})=>mt(e,"supply"))),A("/sop",M(({main:e})=>Ht(e))),A("/checklist",M(({main:e})=>jt(e))),A("/forms",M(({main:e})=>mt(e))),A("/users",M(({main:e})=>Ut(e))),A("/branches",M(({main:e})=>Jt(e))),A("/profile",M(({main:e})=>Qt(e))),A("/settings/import",M(({main:e})=>Xt(e)));let t=Ee();if(!t&&window.location.hash!=="#/login"&&be("/login"),t){let e=await f("/api/auth/me");e.ok?(De(e.data.data),ia()):($e(),be("/login"))}window.addEventListener("fm:login",()=>{ia(),be("/dashboard")}),gt()}ja();
