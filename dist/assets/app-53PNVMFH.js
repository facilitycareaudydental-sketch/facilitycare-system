var oa=Object.defineProperty;var at=(t,e)=>()=>(t&&(e=t(t=0)),e);var nt=(t,e)=>{for(var r in e)oa(t,r,{get:e[r],enumerable:!0})};var ke={};nt(ke,{API:()=>yt,CLIENT_SIDE_MAX_ROWS:()=>be,IS_DEVELOPMENT:()=>Pe,apiFetch:()=>k,clearToken:()=>Le,getToken:()=>Be,getUser:()=>he,setToken:()=>it,setUser:()=>Ne});function Be(){return localStorage.getItem("fm_token")}function it(t){localStorage.setItem("fm_token",t)}function Le(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function he(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Ne(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function k(t,e={}){let r=Be(),a={"Content-Type":"application/json",...r?{Authorization:`Bearer ${r}`}:{},...e.headers||{}};try{let l=`cb=${Date.now()}`,o=t.includes("?")?"&":"?",s=`${yt}${t}${o}${l}`,i=await fetch(s,{...e,headers:a}),n;try{let c=await i.text();try{n=JSON.parse(c)}catch{n={error:`Server Error (${i.status}): ${c.substring(0,80)}...`}}}catch{n={error:"Gagal membaca respon dari server"}}return i.status===401&&(Le(),window.location.hash="#/login"),{ok:i.ok,status:i.status,data:n}}catch(l){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${l.message})`}}}}var Pe,da,yt,be,F=at(()=>{Pe=!1,da="https://fm-operations-api.facilitycare-audydental.workers.dev",yt=da,be=1e4});var kt={};nt(kt,{confirmDialog:()=>Fe,createModal:()=>se});function se({title:t,content:e,onConfirm:r,onCancel:a,confirmText:l="Simpan",cancelText:o="Batal",size:s="md",confirmClass:i="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},c=document.createElement("div");c.className="modal-overlay",c.innerHTML=`
    <div class="modal" style="max-width:${n[s]||n.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${o}</button>
        ${r?`<button class="btn ${i} modal-confirm">${l}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&c.querySelector(".modal-body").appendChild(e);let p=()=>{c.classList.remove("show"),setTimeout(()=>c.remove(),250)};return c.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),p()}),c.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),p()}),r&&c.querySelector(".modal-confirm").addEventListener("click",()=>r(c,p)),c.addEventListener("click",u=>{u.target===c&&(a&&a(),p())}),document.body.appendChild(c),requestAnimationFrame(()=>c.classList.add("show")),{overlay:c,close:p}}function Fe(t,e,r="Konfirmasi"){return se({title:r,content:`<p>${t}</p>`,onConfirm:(a,l)=>{e(),l()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var ye=at(()=>{});var le={};nt(le,{downloadExcel:()=>B,parseExcel:()=>Oe,renderExcelButtons:()=>Re});function Oe(t){return new Promise((e,r)=>{let a=new FileReader;a.onload=l=>{try{let o=new Uint8Array(l.target.result),s=XLSX.read(o,{type:"array"}),i=s.SheetNames[0],n=s.Sheets[i];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${s.SheetNames.join(", ")}`),console.log(`Sheet Used: ${i}`);let c=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),p=c.e.r-c.s.r+1,u=c.e.c-c.s.c+1;console.log(`Total Rows (including empty): ${p}`),console.log(`Total Columns: ${u}`);let d=[];for(let h=c.s.c;h<=c.e.c;++h){let m=n[XLSX.utils.encode_cell({c:h,r:c.s.r})];m&&m.v&&d.push(m.v)}console.log(`Headers Found: ${d.join(", ")}`),console.log("---------------------------");let g=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(g,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(g,"__headers",{value:d,enumerable:!1}),e(g)}catch(o){r(o)}},a.onerror=l=>r(l),a.readAsArrayBuffer(t)})}function B(t,e){try{let r=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,r,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(r){throw console.error("Error generating Excel file:",r),r}}function Re(t){return`
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
  `}var K=at(()=>{});F();var rt={},Je=null;function J(t,e){rt[t]=e}function Se(t){window.location.hash=t}function ft(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[r,...a]=e.split("?"),l=rt[r];if(!l){for(let[s,i]of Object.entries(rt))if(s.endsWith("/*")&&r.startsWith(s.slice(0,-2))){l=i;break}}Je&&(Je(),Je=null);let o=document.getElementById("main-content");if(o&&(o.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),l){let s=new URLSearchParams(a.join("?")),i=r.split("/").filter(Boolean),n=await l({path:r,params:s,segments:i,main:o});n&&(Je=n)}else{let s=o||document.getElementById("app");s&&(s.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var Ae;function ca(){return Ae||(Ae=document.createElement("div"),Ae.id="toast-container",document.body.appendChild(Ae)),Ae}function vt(t,e="info",r=3500){let a=ca(),l=document.createElement("div");l.className=`toast toast-${e}`;let o={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};l.innerHTML=`<span class="toast-icon">${o[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(l),requestAnimationFrame(()=>l.classList.add("show")),setTimeout(()=>{l.classList.remove("show"),setTimeout(()=>l.remove(),350)},r)}var Z=t=>vt(t,"success"),G=t=>vt(t,"error");ye();F();F();ye();function Ue({columns:t,data:e,onEdit:r,onDelete:a,onView:l,actions:o=[],emptyText:s="Tidak ada data",bulkSelect:i=null}){let n=document.createElement("div");if(n.className="table-wrapper",!e||e.length===0)return n.innerHTML=`<div class="empty-state"><p>${s}</p></div>`,n;let c=document.createElement("table");c.className="data-table";let p=document.createElement("thead"),u=document.createElement("tr");if(i){let g=document.createElement("th");g.style.width="40px",g.style.textAlign="center";let h=document.createElement("input");h.type="checkbox",h.id="select-all-checkbox",h.title="Pilih semua",h.addEventListener("change",()=>{e.forEach(m=>{h.checked?i.selectedIds.add(m.id):i.selectedIds.delete(m.id)}),n.querySelectorAll(".row-checkbox").forEach(m=>m.checked=h.checked),i.onToggle()}),g.appendChild(h),u.appendChild(g)}if(t.forEach(g=>{let h=document.createElement("th");h.textContent=g.label,g.width&&(h.style.width=g.width),u.appendChild(h)}),r||a||l||o.length>0){let g=document.createElement("th");g.textContent="Aksi",g.style.width="120px",u.appendChild(g)}p.appendChild(u),c.appendChild(p);let d=document.createElement("tbody");return e.forEach(g=>{let h=document.createElement("tr");if(i){let m=document.createElement("td");m.style.textAlign="center",m.style.width="40px";let y=document.createElement("input");y.type="checkbox",y.className="row-checkbox",y.checked=i.selectedIds.has(g.id),y.addEventListener("change",()=>{if(y.checked)i.selectedIds.add(g.id);else{i.selectedIds.delete(g.id);let b=document.getElementById("select-all-checkbox");b&&(b.checked=!1)}i.onToggle()}),m.appendChild(y),h.appendChild(m)}if(t.forEach(m=>{let y=document.createElement("td");if(m.render){let b=m.render(g[m.key],g);b instanceof HTMLElement?y.appendChild(b):y.innerHTML=b||""}else y.textContent=g[m.key]!==null&&g[m.key]!==void 0&&g[m.key]!==""?g[m.key]:"";m.nowrap&&(y.style.whiteSpace="nowrap"),h.appendChild(y)}),r||a||l||o.length>0){let m=document.createElement("td");m.className="actions-cell";let y=document.createElement("div");if(y.className="btn-group",l){let b=document.createElement("button");b.className="btn btn-xs btn-ghost",b.innerHTML="\u{1F441}",b.title="Lihat",b.addEventListener("click",()=>l(g)),y.appendChild(b)}if(r){let b=document.createElement("button");b.className="btn btn-xs btn-secondary",b.innerHTML="\u270F\uFE0F",b.title="Edit",b.addEventListener("click",()=>r(g)),y.appendChild(b)}o.forEach(b=>{let x=document.createElement("button");x.className=`btn btn-xs ${b.class||"btn-ghost"}`,x.innerHTML=b.icon||b.label,x.title=b.label,x.addEventListener("click",()=>b.handler(g)),y.appendChild(x)}),m.appendChild(y),h.appendChild(m)}d.appendChild(h)}),c.appendChild(d),n.appendChild(c),n}function Ge({page:t,pages:e,total:r,limit:a,onPage:l}){if(e<=1)return null;let o=document.createElement("div");o.className="pagination";let s=document.createElement("span");s.className="pagination-info",s.textContent=`Total: ${r} data`,o.appendChild(s);let i=document.createElement("div");i.className="pagination-btns";let n=(u,d,g=!1,h=!1)=>{let m=document.createElement("button");m.className=`btn btn-sm ${h?"btn-primary":"btn-ghost"} pagination-btn`,m.textContent=u,m.disabled=g,m.addEventListener("click",()=>l(d)),i.appendChild(m)};n("\xAB",1,t===1),n("\u2039",t-1,t===1);let c=Math.max(1,t-2),p=Math.min(e,t+2);for(let u=c;u<=p;u++)n(u,u,!1,u===t);return n("\u203A",t+1,t===e),n("\xBB",e,t===e),o.appendChild(i),o}ye();function Me(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${Me(e.fields)}</div>`;let r=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",l="";switch(e.type){case"textarea":l=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${r} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let s=(e.options||[]).map(u=>{let d=typeof u=="object"?u.value:u,g=typeof u=="object"?u.label:u,h=e.value==d?"selected":"";return`<option value="${d}" ${h}>${g}</option>`}).join("");l=`<select name="${e.name}" class="form-control" ${r}><option value="">-- Pilih ${e.label||""} --</option>${s}</select>`;break;case"combobox":let i=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,n=(e.options||[]).map(u=>{let d=typeof u=="object"?u.value:u,g=typeof u=="object"?u.label||u.value||"":u||"";return(g==="undefined"||g==="[object Object]"||g==="null")&&(g=""),g?`<option value="${g}"></option>`:""}).join(""),c=e.value||"";if(e.value){let u=(e.options||[]).find(d=>(typeof d=="object"?d.value:d)==e.value);if(u){let d=typeof u=="object"?u.label||u.value||"":u||"";d&&d!=="undefined"&&d!=="[object Object]"&&d!=="null"&&(c=d)}}l=`
          <input type="text" name="${e.name}" list="${i}" class="form-control" value="${c}" placeholder="Pilih atau ketik baru..." ${r} autocomplete="off">
          <datalist id="${i}">${n}</datalist>
        `;break;case"checkbox":l=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let p=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";l=`<input type="date" name="${e.name}" class="form-control" value="${p}" ${r}>`;break;case"number":l=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${r}>`;break;case"email":l=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${r}>`;break;case"url":l=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${r}>`;break;default:l=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${r} autocomplete="off">`}let o=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${l}${o}</div>`}).join("")}function Qe(t){let e={},r=new FormData(t);for(let[a,l]of r.entries())e[a]=l===""?null:l;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function ze(t,e){e&&Object.entries(e).forEach(([r,a])=>{let l=t.querySelector(`[name="${r}"]`);l&&(l.hasAttribute("list")||(l.type==="checkbox"?l.checked=!!a:l.type==="date"&&a&&window.parseFlexibleDate?l.value=window.parseFlexibleDate(a):l.value=a??""))})}K();function L({container:t,title:e,icon:r,apiPath:a,columns:l,formFields:o,filterFields:s,defaultFilters:i={},itemLabel:n="Data",canCreate:c=!0,canEdit:p=!0,canDelete:u=!0,onBeforeSubmit:d,onAfterLoad:g,onDataLoaded:h,extraActions:m=[],initialSearch:y="",exportOptions:b=null,bulkDelete:x=!1,paginationMode:T="server"}){let S=1,w={...i};y&&(w.search=y);let _=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${r} ${e}</h1>
      <div class="page-actions">
        ${c?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
      </div>
    </div>

    ${x?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${b?Re(b.moduleName):""}

    ${s&&s.length>0?`
    <div class="filter-bar card" style="padding: 1rem;">
        ${s.map(v=>{if(v.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${w.search||""}"></div>`;if(v.type==="search-combo"){let f="dl-filter-search",$=(v.options||[]).map(P=>`<option value="${typeof P=="object"?P.label:P}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${f}" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${w.search||""}"><datalist id="${f}">${$}</datalist></div>`}return v.type==="select"?`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${v.name}" id="filter-${v.name}"><option value="">-- ${v.label} --</option>${(v.options||[]).map(f=>`<option value="${typeof f=="object"?f.value:f}" ${w[v.name]===(typeof f=="object"?f.value:f)?"selected":""}>${typeof f=="object"?f.label:f}</option>`).join("")}</select>`:""}).join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function M(){if(!document.getElementById("bulk-toolbar"))return;let f=document.getElementById("bulk-count"),$=document.getElementById("btn-bulk-delete"),P=document.getElementById("btn-bulk-cancel");f.textContent=`${_.size} item dipilih`,_.size>0?($.disabled=!1,P.disabled=!1):($.disabled=!0,P.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{_.clear(),document.querySelectorAll(".row-checkbox").forEach(f=>f.checked=!1);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1),M()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(_.size===0)return;let v=[..._],f=document.createElement("div");f.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",f.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${v.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${v.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(f),f.querySelector("#bulk-cancel-btn").addEventListener("click",()=>f.remove()),f.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let $=f.querySelector("#bulk-confirm-btn");$.disabled=!0,$.textContent="Menghapus...";let P=await k(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:v})});f.remove(),P.ok?(Z(`${v.length} ${n} berhasil dihapus.`),_.clear(),M(),O()):G(P.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),R;if(E?.addEventListener("input",v=>{clearTimeout(R),R=setTimeout(()=>{w.search=v.target.value,S=1,_.clear(),O()},400)}),s?.forEach(v=>{v.type==="select"&&document.getElementById(`filter-${v.name}`)?.addEventListener("change",f=>{w[v.name]=f.target.value,S=1,_.clear(),O()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{w={...i},E&&(E.value=""),s?.forEach(v=>{let f=document.getElementById(`filter-${v.name}`);f&&(f.value="")}),S=1,_.clear(),O()}),document.getElementById("btn-create")?.addEventListener("click",()=>ge(null)),b){document.getElementById(`btn-export-${b.moduleName}`)?.addEventListener("click",async f=>{let $=f.target,P=$.innerHTML;$.innerHTML="\u23F3 Loading...",$.disabled=!0;try{await b.onExport()}catch{G("Gagal export data")}finally{$.innerHTML=P,$.disabled=!1}}),document.getElementById(`btn-template-${b.moduleName}`)?.addEventListener("click",()=>{b.onTemplate()});let v=document.getElementById(`input-import-${b.moduleName}`);v?.addEventListener("change",async f=>{let $=f.target.files[0];if(!$)return;v.disabled=!0;let P=document.createElement("div");P.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",P.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(P);let V=P.querySelector("#import-progress-text"),U=P.querySelector("#import-progress-bar"),D=P.querySelector("#import-summary"),I=P.querySelector("#import-close-btn");I.addEventListener("click",()=>{P.remove(),O()});try{let Y=await Oe($);if(Y.length===0)throw new Error("File kosong atau format salah");let X=500,ie=0,ae=0,N=0,A=Y.length;V.textContent=`Ditemukan ${A} baris data. Memulai import...`;for(let C=0;C<A;C+=X){let z=Y.slice(C,C+X);V.textContent=`Mengimport baris ${C+1} - ${Math.min(C+X,A)} dari ${A}...`,U.style.width=`${Math.round(C/A*100)}%`;try{let j=await b.onImport(z);j?(ie+=j.inserted||j.metrics?.inserted||z.length,ae+=j.skipped||j.metrics?.updated||0):ie+=z.length}catch(j){console.error("Chunk import failed:",j),N+=z.length}}U.style.width="100%",V.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',D.style.display="block",D.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${A}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ie}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${N}</strong></div>
        `,N>0&&(D.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),I.style.display="block",v.value=""}catch(Y){V.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${Y.message}`,U.style.background="var(--danger)",U.style.width="100%",I.style.display="block",v.value=""}finally{v.disabled=!1}})}async function O(){M();let v=document.getElementById("table-container");if(!v)return;v.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let f=T==="client",$=f?1:S,P=f?be:20,V=new URLSearchParams({page:$,limit:P,...Object.fromEntries(Object.entries(w).filter(([,N])=>N))}),U=await k(`${a}?${V}`);if(!U.ok){v.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${U.data?.error||"Error"}</p></div>`;return}let D=U.data?.data||U.data||[],I=U.data?.pagination,Y=D.length,X=D;if(f){D=h(D),X=D;let N=D.length,A=20,C=Math.ceil(N/A);S>C&&C>0&&(S=C);let z=(S-1)*A,j=S*A;D=D.slice(z,j),I={page:S,limit:A,total:N,pages:C}}!1,g&&g(D);let ie=Ue({columns:l,data:D,fullData:X,onEdit:p?N=>ge(N):null,actions:m.map(N=>({...N,handler:A=>N.handler(A,O)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:x?{selectedIds:_,onToggle:M}:null});v.innerHTML="",v.appendChild(ie);let ae=document.getElementById("pagination-container");if(ae&&(ae.innerHTML="",I&&I.pages>1)){let N=Ge({page:I.page,pages:I.pages,total:I.total,limit:I.limit,onPage:A=>{S=A,O()}});N&&ae.appendChild(N)}}function fe(v){let f=typeof o=="function"?o(v):o;return Me(f)}function ge(v){let f=!!v,$=document.createElement("form");if($.noValidate=!0,$.innerHTML=fe(v),f){let V=typeof o=="function"?o(v):o;ze($,v)}let{close:P}=se({title:f?`Edit ${n}`:`Tambah ${n}`,content:$,size:"lg",confirmText:f?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(V,U)=>{if(!$.reportValidity())return;let D=V.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let I=Qe($),Y=typeof o=="function"?o(v):o,X=async A=>{for(let C of A)if(C.type==="row")await X(C.fields);else if(C.type==="combobox"&&I[C.name]){let z=I[C.name],j=(C.options||[]).find(W=>{let ne=String(typeof W=="object"?W.value:W),tt=String(typeof W=="object"?W.label:W);return ne===z||tt===z});if(j)I[C.name]=typeof j=="object"?j.value:j;else if(C.createApi){let W={};W[C.createApi.field]=z,C.createApi.extra&&Object.assign(W,C.createApi.extra);let ne=await k(C.createApi.path,{method:"POST",body:JSON.stringify(W)});if(ne.ok&&ne.data?.id)I[C.name]=ne.data.id;else if(ne.ok&&!ne.data?.id)I[C.name]=z;else throw new Error(`Gagal membuat master data: ${ne.data?.error||"Unknown error"}`)}}};try{await X(Y)}catch(A){G(A.message),D.disabled=!1,D.textContent=f?"Simpan Perubahan":`Tambah ${n}`;return}d&&(I=await d(I,v));let ie=f?"PUT":"POST",ae=f?`${a}/${v.id}`:a,N=await k(ae,{method:ie,body:JSON.stringify(I)});N.ok?(Z(f?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),U(),O()):(G(N.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=f?"Simpan Perubahan":`Tambah ${n}`)}})}function ve(v){Fe(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let f=await k(`${a}/${v.id}`,{method:"DELETE"});f.ok?(Z(`${n} berhasil dihapus.`),O()):G(f.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return O(),O}F();F();var xe=null,Ve=null;async function we(t=!1){if(xe&&!t)return console.log("Employees Raw (Cache Hit)",xe.slice(0,5)),xe;let e=await k(`/api/employees?limit=${be}&status=Aktif`);return xe=(e.data?.data||[]).map(r=>({value:r.id,label:r.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",xe.slice(0,5)),xe}async function ee(t=!1){let r=(await we(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",r.slice(0,5)),r}async function H(t=!1){return Ve&&!t||(Ve=((await k("/api/branches?all=1")).data?.data||[]).map(r=>({value:r.id,label:r.full_name}))),Ve}function q(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function lt(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function _e(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function st(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function oe(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}K();function ot(t,e){if(t.period!=="Q3")return!1;let r=String(t.status||"").toLowerCase();if(r!=="selesai"&&r!=="completed"&&r!=="done")return!1;let a=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?a.includes("inspeksi"):e==="gcdc"?a.includes("general cleaning")||a.includes("deep cleaning"):!1}F();K();function St(t,e){let r=String(t.status||"").toLowerCase();return e==="active"?r==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&r==="aktif":!1}F();K();function dt(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let l=new Date(a);l.setDate(a.getDate()+30);let o=new Date(t.end_date);return o.setHours(0,0,0,0),o>=a&&o<=l}return!1}F();K();function xt(t,e){let r=String(t.status||"").toLowerCase();return e==="open"?r==="open":!1}F();function wt(t,e){let r=String(t.status||"").toLowerCase();return e==="pending"?r==="pending":!1}var me={};function Ee(t){if(me[t]){try{me[t].destroy()}catch{}delete me[t]}}function pa(){Object.keys(me).forEach(Ee)}var pe=(t,e=0)=>{let r=Number(t);return isNaN(r)||t===null||t===void 0?e:r},Ce=(t,e="\u2014")=>{if(t==null||t==="")return e;let r=String(t).trim();return r===""||r==="[object Object]"?e:r};function Ct(t,e,r=900){if(!t)return;let a=Math.max(0,Math.round(pe(e)));if(a===0){t.textContent="0";return}let l=Date.now(),o=()=>{let s=Math.min((Date.now()-l)/r,1),i=1-Math.pow(1-s,3);t.textContent=Math.round(i*a).toLocaleString("id-ID"),s<1?requestAnimationFrame(o):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(o)}var ua={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},ma=t=>{let e=Ce(t,"\u2014");return`<span class="status-pill ${ua[e]||"pill-neutral"}">${e}</span>`};var de={family:"Inter",size:11},ue="#94A3B8",Te="#F1F5F9",ct=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],ga=()=>window.innerWidth<768;function We(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:ga()?"bottom":"top",labels:{font:de,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:de,titleFont:{...de,weight:"700"}}},scales:{x:{grid:{color:Te},ticks:{font:de,color:ue,maxRotation:0}},y:{grid:{color:Te},ticks:{font:de,color:ue},beginAtZero:!0}},...t}}var ba=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join("");function _t(t=3){return Array(t).fill(0).map((e,r)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${r<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function re(t,e,r=8e3){try{let a=new AbortController,l=setTimeout(()=>a.abort(),r),o=await k(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(l),!o||!o.ok)return e;let s=o.data;return s?s.data!==void 0?s.data??e:s:e}catch{return e}}function ya(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(a=>{let l=document.getElementById(a);l&&(l.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(a=>{let l=document.getElementById(a);if(l&&l.style.display==="none"){l.style.display="block";let o=l.parentElement;if(o&&!o.querySelector(".chart-empty")){let s=document.createElement("div");s.className="chart-empty",s.textContent="Belum ada data",l.style.display="none",o.appendChild(s)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Et({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&$t({}),["table-contracts","table-issues"].forEach(a=>{let l=document.getElementById(a);l&&l.querySelector(".skeleton")&&(l.innerHTML='<div class="chart-empty">Belum ada data</div>')});let r=document.getElementById("activity-log");r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Tt(t){pa(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
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
              <div class="chart-card-title">Jadwal Kegiatan per Bulan (<span id="jadwal-year-label">2026</span>)</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>pt(t)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async r=>{let a=r.target.value;document.getElementById("jadwal-year-label").textContent=a;let l=document.getElementById("skel-jadwal"),o=document.getElementById("chart-jadwal");l&&(l.style.display="block",l.style.position="absolute"),o&&(o.style.display="none");let s=await re(`/api/dashboard/schedule-chart?year=${a}`,{},8e3);try{Dt(s)}catch(i){console.warn("ScheduleChart render:",i),ce("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async r=>{let a=r.target.value,l=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",o=document.getElementById("skel-insp"),s=document.getElementById("chart-insp");o&&(o.style.display="block",o.style.position="absolute"),s&&(s.style.display="none");let i=await re(l,{},8e3);try{It(i)}catch(n){console.warn("InspBar render:",n),ce("skel-insp","chart-insp")}}),t._skelTimeout=setTimeout(()=>ya(),5e3),await pt(t),t._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?pt(t):clearInterval(t._dashRefresh)},6e4)}async function pt(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,r,a,l,o,s,i,n,c,p,u,d]=await Promise.all([re("/api/dashboard/kpi",{},8e3),re("/api/dashboard/issues-trend",{},8e3),re("/api/dashboard/issues-summary",{},8e3),re("/api/dashboard/stats",{},8e3),re("/api/dashboard/calendar",[],8e3),re("/api/schedule?limit=10000",{data:[]},8e3),re("/api/employees?limit=10000",{data:[]},8e3),re("/api/contracts?limit=10000",{data:[]},8e3),re("/api/issues?limit=10000",{data:[]},8e3),re("/api/one-on-one?limit=10000",{data:[]},8e3),re("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),re(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3)]),g=document.getElementById("filter-insp-month"),h=g?g.value:"",m=h?`/api/dashboard/inspection-bar?month=${h}`:"/api/dashboard/inspection-bar",y=await re(m,{},8e3);if(e){let b=Array.isArray(s?.data)?s.data:Array.isArray(s)?s:[],x=Array.isArray(i?.data)?i.data:Array.isArray(i)?i:[],T=Array.isArray(n?.data)?n.data:Array.isArray(n)?n:[],S=Array.isArray(c?.data)?c.data:Array.isArray(c)?c:[],w=Array.isArray(p?.data)?p.data:Array.isArray(p)?p:[];e.employees&&(e.employees.current=x.filter(_=>St(_,"active")).length),e.contracts&&(e.contracts.current=T.filter(_=>dt(_,"active")).length),e.expiring30&&(e.expiring30={current:T.filter(_=>dt(_,"expiring30")).length}),e.issues&&(e.issues.current=S.filter(_=>xt(_,"open")).length),e.one_on_one&&(e.one_on_one.current=w.filter(_=>wt(_,"pending")).length),e.inspection_month&&(e.inspection_month.current=b.filter(_=>ot(_,"inspeksi")).length),e.cleaning_month&&(e.cleaning_month.current=b.filter(_=>ot(_,"gcdc")).length)}try{Et(e)}catch(b){console.warn("KPI render:",b)}try{$t(e)}catch(b){console.warn("MiniStats render:",b)}try{Dt(d)}catch(b){console.warn("ScheduleChart render:",b),ce("skel-jadwal","chart-jadwal")}try{fa(Array.isArray(a?.by_category)?a.by_category:[])}catch(b){console.warn("Donut render:",b),ce("skel-donut","chart-donut")}try{va(r)}catch(b){console.warn("Trend render:",b),ce("skel-trend","chart-trend")}try{It(y)}catch(b){console.warn("InspBar render:",b),ce("skel-insp","chart-insp")}try{let b=Array.isArray(l)?l:Array.isArray(l?.recent_issues)?l.recent_issues:[];Sa(b)}catch(b){console.warn("IssuesTable render:",b)}try{let b=Array.isArray(l?.expiring_contracts)?l.expiring_contracts:[];ka(u)}catch(b){console.warn("ContractsTable render:",b)}try{xa(Array.isArray(o)?o:[])}catch(b){console.warn("Agenda render:",b)}try{wa()}catch(b){console.warn("Quick Actions render:",b)}}function Et(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let r=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=r.map(a=>{let l=pe(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${l}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${l}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{Ct(a,parseInt(a.dataset.target)||0)})}function $t(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let r=[{icon:"\u{1F4C5}",label:"Jadwal",val:t.schedule?.current,href:"#/timeline",color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F504}",label:"Reliefer",val:t.reliever_completed?.current,href:"#/relievers?dash_filter=reliever",color:"mini-teal"},{icon:"\u{1F50D}",label:"Inspeksi",val:t.inspection_month?.current,href:"#/timeline?dash_filter=inspeksi",color:"mini-blue"},{icon:"\u{1F9F9}",label:"GCDC",val:t.cleaning_month?.current,href:"#/timeline?dash_filter=gcdc",color:"mini-green"},{icon:"\u{1F4A8}",label:"Fogging",val:t.fogging_month?.current,href:"#/reports/fogging",color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=r.map(a=>`
    <a href="${a.href}" class="mini-stat ${a.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${a.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${pe(a.val)}">0</div>
        <div class="mini-stat-text">${a.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(a=>Ct(a,parseInt(a.dataset.target)||0,700))}function fa(t){ce("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),r=document.getElementById("donut-legend");if(!e||!r)return;Ee("donut");let a=(t||[]).filter(n=>pe(n.count)>0);if(!a.length){Ke(e,"Belum ada data permasalahan");return}let l=a.map(n=>`${Ce(n.category,"Lainnya")}`),o=a.map(n=>pe(n.count)),s=o.reduce((n,c)=>n+c,0);r.innerHTML=a.map((n,c)=>{let p=ct[c%ct.length],u=s>0?Math.round(n.count/s*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${p}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${u}%)</span></div>
          <div class="donut-legend-label">${l[c]}</div>
        </div>
      </div>
    `}).join("");let i={id:"centerText",beforeDraw:function(n){let c=n.width,p=n.height,u=n.ctx;u.restore();let d=(p/80).toFixed(2);u.font="bold "+d+"em Inter",u.textBaseline="middle",u.fillStyle="#1E293B";let g=s.toString(),h=Math.round((c-u.measureText(g).width)/2),m=p/2;u.fillText(g,h,m-4),u.font="600 "+(d*.35).toFixed(2)+"em Inter",u.fillStyle="#64748B";let y="Total",b=Math.round((c-u.measureText(y).width)/2);u.fillText(y,b,m+10),u.save()}};me.donut=new Chart(e,{type:"doughnut",data:{labels:l,datasets:[{data:o,backgroundColor:ct,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:de,titleFont:{...de,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutout:"75%"},plugins:[i]})}function va(t){ce("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;Ee("trend"),t=t||{};let r=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(t.labels||[]).map(s=>{if(!s||typeof s!="string")return"";try{let[i,n]=s.split("-");return(r[Number(n)-1]||n)+" "+String(i).slice(-2)}catch{return s}}),l=(t.open||[]).map(s=>pe(s)),o=(t.closed||[]).map(s=>pe(s));if(!a.length){Ke(e,"Belum ada data trend");return}me.trend=new Chart(e,{type:"line",data:{labels:a,datasets:[{label:"Open",data:l,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:o,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:We({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:ue,maxRotation:0,autoSkip:!1}},y:{grid:{color:Te},ticks:{font:{family:"Inter",size:9},color:ue},beginAtZero:!0}}})})}function Dt(t){ce("skel-jadwal","chart-jadwal");let e=document.getElementById("chart-jadwal");if(!e)return;Ee("jadwal"),t=t||{};let r=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(t).some(n=>Array.isArray(n)&&n.some(c=>c>0))){Ke(e,"Belum ada data jadwal");return}let l=t["Inspeksi Hygiene"]||Array(12).fill(0),o=t["General Cleaning"]||Array(12).fill(0),s=t["Deep Cleaning"]||Array(12).fill(0),i=t.Fogging||Array(12).fill(0);me.jadwal=new Chart(e,{type:"bar",data:{labels:r,datasets:[{label:"Inspeksi",data:l,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:o,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:s,backgroundColor:"#F59E0B"},{label:"Fogging",data:i,backgroundColor:"#EF4444"}]},options:We({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:ue,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Te},ticks:{font:{family:"Inter",size:9},color:ue},min:0}}})})}function It(t){ce("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;Ee("inspBar"),t=t||{};let r=t.labels||[],a=(t.fc||[]).map(o=>pe(o)),l=(t.spv||[]).map(o=>pe(o));if(!r.length){Ke(e,"Belum ada data inspeksi");return}me.inspBar=new Chart(e,{type:"bar",data:{labels:r,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:l,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:We({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:de,color:ue,maxRotation:45,minRotation:30}},y:{grid:{color:Te},ticks:{font:de,color:ue},min:0,max:100}}})})}function ka(t){ce("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;Ee("contractMiniBar"),t=t||{};let r={"06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(s=>{let i=s.split("-")[1];return r[i]||s}),l=(t.data||[]).map(s=>pe(s));if(!a.length){Ke(e,"Belum ada data");return}let o=e.getContext("2d");me.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:l,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:We({onClick:(s,i)=>{if(i&&i.length>0){let n=i[0].index,c=(t.labels||[])[n];c&&(window.location.hash="#/contracts?month_expiry="+c)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:de,color:ue,maxRotation:0,autoSkip:!1}},y:{grid:{color:Te,borderDash:[4,4],drawBorder:!1},ticks:{font:de,color:ue,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function Sa(t){let e=document.getElementById("table-issues");if(!e)return;let r=(t||[]).slice(0,8);if(!r.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${r.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${ma(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Ce(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Ce(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function xa(t){let e=document.getElementById("widget-agenda");if(!e)return;let r=new Date,a=`${r.getFullYear()}-${String(r.getMonth()+1).padStart(2,"0")}-${String(r.getDate()).padStart(2,"0")}`,o=(t||[]).filter(s=>(s.event_date||"").startsWith(a)).slice(0,10);if(!o.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${o.map(s=>{let i="#3B82F6",n="#EFF6FF",c="Agenda",p=(s.title||"").toLowerCase();return p.includes("inspeksi")?(i="#10B981",n="#ECFDF5",c="Inspeksi"):p.includes("cleaning")||p.includes("gcdc")?(i="#3B82F6",n="#EFF6FF",c="Cleaning"):p.includes("reliefer")?(i="#F59E0B",n="#FFFBEB",c="Reliefer"):p.includes("fogging")&&(i="#8B5CF6",n="#F5F3FF",c="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(s.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${i};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Ce(s.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Ce(s.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${i}">${c}</div>
        </div>
      `}).join("")}
    </div>
  `}function wa(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(r=>`
    <a href="${r.href}" class="action-btn">
      <div class="action-icon" style="background:${r.bg}">${r.icon}</div>
      ${r.label}
    </a>
  `).join("")}function ce(t,e){let r=document.getElementById(t),a=document.getElementById(e);if(r&&(r.style.display="none",r.style.position=""),a){a.style.display="block";let l=a.parentElement;if(l){let o=l.querySelector(".chart-empty");o&&o.remove()}}}function Ke(t,e="Belum ada data"){if(!t)return;t.style.display="none";let r=t.parentElement;if(!r)return;if(!r.querySelector(".chart-empty")){let l=document.createElement("div");l.className="chart-empty",l.textContent=e,r.appendChild(l)}}F();async function Pt(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),r=document.getElementById("login-error"),a=document.getElementById("login-btn"),l=document.getElementById("toggle-password"),o=document.getElementById("login-password");l?.addEventListener("click",()=>{let s=o.type==="text";o.type=s?"password":"text",l.style.color=s?"":"var(--primary)"}),e?.addEventListener("submit",async s=>{s.preventDefault(),r.style.display="none";let i=e.username.value.trim(),n=e.password.value;if(!i||!n){r.textContent="Username dan password wajib diisi.",r.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let c=await k("/api/auth/login",{method:"POST",body:JSON.stringify({username:i,password:n})});c.ok&&c.data.success?(it(c.data.data.token),Ne(c.data.data.user),Z("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(r.textContent=c.data.error||"Username atau password salah.",r.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{r.textContent="Gagal terhubung ke server. Periksa koneksi internet.",r.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}F();K();async function _a(){return await H()}function Ca(t,e){let r=String(t.status||"").toLowerCase();return e==="active"?r==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&r==="aktif":!1}async function Bt(t,e){let r=await _a(),a=e?e.get("dash_filter"):null;L({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:l=>a?l.filter(o=>Ca(o,a)):l,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:l=>_e(l)},{key:"phone",label:"No. HP",render:l=>l?`<a href="tel:${l}">${l}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>q(l)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:l=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:l?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:l?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:r,value:l?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:l?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:l?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:l?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let l=await k(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(l.ok){let o=l.data.data.map(s=>({"Nama Lengkap":s.full_name,Cabang:s.branch_name||"",Divisi:s.division||"","No. HP":s.phone||"","Tgl Masuk":s.join_date||"",Status:s.status||""}));B(o,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async l=>{let o=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),p=r.find(u=>String(u.label||"").toLowerCase()===c);return p?p.value:null},s=l.map(n=>({full_name:String(n["Nama Lengkap"]||"").trim(),branch_id:o(String(n.Cabang||"").trim()),division:String(n.Divisi||"").trim()||"FACILITY CARE",phone:String(n["No. HP"]||"").trim(),join_date:String(n["Tgl Masuk"]||"").trim(),status:String(n.Status||"").trim(),notes:String(n.Catatan||"").trim()})).filter(n=>n.full_name),i=await k("/api/import/employees",{method:"POST",body:JSON.stringify({rows:s,onDuplicate:"update"})});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}F();K();var mt=[],Lt=[];async function Ta(){mt=await H(),Lt=await we()}var ut=async t=>{let e=[],r=1;for(;;){let l=await(await Promise.resolve().then(()=>(F(),ke))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${r}`);if(!l.ok)break;let o=l.data?.data||l.data||[],s=Array.isArray(o)?o:[];if(e=e.concat(s),s.length<100||l.data?.pagination&&r>=l.data.pagination.pages)break;r++}return e};function Ea(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let l=new Date(a);l.setDate(a.getDate()+30);let o=new Date(t.end_date);return o.setHours(0,0,0,0),o>=a&&o<=l}return!1}async function Nt(t,e){await Ta();let r=e?e.get("dash_filter"):null;L({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",onDataLoaded:a=>r?a.filter(l=>Ea(l,r)):a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>_e(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,l)=>l.end_date&&String(l.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':lt(a)},{key:"status",label:"Status",render:a=>q(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:mt},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"month_expiry",label:"Bulan Habis",options:[{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"},{value:"2027-01",label:"Jan 2027"},{value:"2027-02",label:"Feb 2027"},{value:"2027-03",label:"Mar 2027"},{value:"2027-04",label:"Apr 2027"},{value:"2027-05",label:"Mei 2027"},{value:"2027-06",label:"Jun 2027"},{value:"2027-07",label:"Jul 2027"}]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[o,s]=await Promise.all([ut("/api/employees?status=Aktif"),ut("/api/contracts")]);if(o.length>0){let i=s.filter(u=>u.status==="Aktif"),n=new Set(i.map(u=>u.employee_id)),c=o.filter(u=>!n.has(u.id)),p=`<p style="margin-bottom:12px">Data yang terbaca: <b>${o.length}</b> Karyawan Aktif, dan <b>${i.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${c.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;c.forEach(u=>{let d=s.filter(h=>h.employee_id===u.id),g='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(d.length>0){let h=d[0];g=`Pernah ada kontrak (Status: <b style="color:#EF4444">${h.status}</b>, Selesai: ${window.formatDate(h.end_date)})`}p+=`<li style="margin-bottom:8px"><b>${u.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${u.branch_name||"-"} | ${g}</span></li>`}),p+="</ul>",Promise.resolve().then(()=>(ye(),kt)).then(u=>u.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:p,cancelText:"Tutup"}))}}catch(o){console.error(o)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let l=document.querySelector(".page-actions");l&&l.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:Lt,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:mt,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await k(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let l=a.data.data.map(o=>({"Nama Lengkap":o.employee_name,Cabang:o.branch_name||"","Div / Bagian":o.division||"","Tanggal Mulai":o.start_date||"","Tanggal Selesai":o.end_date&&String(o.end_date).startsWith("2099")?"":o.end_date||"","Sisa Kontrak":o.end_date&&String(o.end_date).startsWith("2099")?"Tetap":o.days_remaining!==null&&o.days_remaining!==void 0?`${o.days_remaining} Hari`:"",Status:o.status||""}));B(l,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[l,o]=await Promise.all([k("/api/branches?limit=10000"),ut("/api/employees")]),s=l.data?.data||[],i=o||[];console.log(`Total employee yang berhasil dimuat dari database : ${i.length}`),i.length>0&&(console.log("Contoh 5 employee pertama:"),i.slice(0,5).forEach((m,y)=>{console.log(`${y+1}. ID: ${m.id}, Name: ${m.full_name}, Status: ${m.status}`)}));let n=m=>{if(!m)return null;let y=String(m||"").replace(/\s+/g," ").toLowerCase().trim(),b=s.find(x=>String(x.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===y||String(x.code||"").replace(/\s+/g," ").toLowerCase().trim()===y||String(x.name||"").replace(/\s+/g," ").toLowerCase().trim()===y);return b?b.id:null},c=(m,y)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${y}`),console.log(`Nama dari Excel : "${m}"`),!m)return console.log("Alasan gagal mapping : Nama kosong"),null;let b=String(m||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${b}"`),console.log(`Jumlah employee di database : ${i.length}`);let x=i.find(T=>String(T.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===b);return x?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${x.id}`),x.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},p=m=>{if(!m)return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let y=String(m).trim();if(/^\d{4,5}(\.\d+)?$/.test(y)){let x=Math.floor(Number(y));if(x>2e4&&x<99999){let T=new Date(Date.UTC(1899,11,30)+x*864e5);return isNaN(T.getTime())?"":T.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);let b=y.split(/[\/\-\.]/);if(b.length===3){let[x,T,S]=b.map(w=>w.trim());if(x.length===4&&T.length<=2&&S.length<=2)return`${x}-${T.padStart(2,"0")}-${S.padStart(2,"0")}`;if(S.length===4&&T.length<=2&&x.length<=2)return`${S}-${T.padStart(2,"0")}-${x.padStart(2,"0")}`}return y},u=a.map((m,y)=>{let b=y+2,x=String(m["Nama Lengkap"]||"").trim(),T=m["Tanggal Mulai"],S=p(T);if(!S){let M=a.__worksheet,E=a.__headers||[],R=E.indexOf("Tanggal Mulai"),O="N/A",fe="N/A",ge="N/A";if(R!==-1&&M&&window.XLSX){let v=window.XLSX.utils.encode_cell({c:R,r:b-1});ge=v;let f=M[v];f?(O=f.t||"undefined",fe=f.w||"undefined"):O="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let ve="Unknown";T==null||T===""?ve="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":T instanceof Date&&isNaN(T.getTime())?ve="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":ve="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${b}`),console.log(`Employee Name : ${x}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${R})`),console.log(`Raw Cell Value : "${T}"`),console.log(`JavaScript Type : ${typeof T}`),console.log(`SheetJS Cell Type : ${O}`),console.log(`SheetJS Formatted Value : "${fe}"`),console.log(`Value After Trim : "${String(T||"").trim()}"`),console.log(`Value After Date Parser : "${S}"`),console.log(`Is Empty : ${!T}`),console.log(`Is Invalid Date : ${T instanceof Date?isNaN(T.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${ve}`),console.log(`Workbook Sheet : ${M?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${ge}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(m,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(E)),console.log(`==========================
`)}let w=c(x,b),_=null;return w||(_="Karyawan tidak ditemukan di Database"),{isValid:!!w,invalidReason:_,rowNum:b,data:{employee_id:w,branch_id:n(String(m.Cabang||"").trim()),division:String(m["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:S,end_date:p(m["Tanggal Selesai"])||"2099-12-31",status:String(m.Status||"").trim(),_rawName:x}}}),d=[],g=[];if(u.forEach(m=>{m.isValid?d.push(m.data):g.push({rowNum:m.rowNum,name:m.data._rawName,reason:m.invalidReason})}),console.log(`Split Validation - Valid: ${d.length}, Invalid: ${g.length}`),d.length===0)return{inserted:0,skipped:a.length,failed:a.length};let h=await k("/api/import/contracts",{method:"POST",body:JSON.stringify({rows:d,onDuplicate:"update"})});if(!h.ok)throw new Error(h.data?.error||"Import gagal");return h.data}}})}F();K();var gt=[],He=[];function $a(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let r of e)if(t.some(a=>a.period===r))return r;return"Q3"}async function At(t,e){gt=await H();let r=await ee();He=["BERLIN ARIANSYAH","ADE SURAHMAN"];let a=p=>p&&!He.find(u=>(typeof u=="object"?u.value:u)===p)?[...He,p]:He,l=await k(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),o=p=>{if(!p||p==="-"||String(p).trim()==="")return"";let u=String(p).split("-");return u.length===3&&u[0].length===4?`${u[2]}-${u[1]}-${u[0]}`:p},s=l.data?.data||[],i=$a(s),n=e?e.get("dash_filter"):null,c={period:i};n==="inspeksi"?c={period:"Q3",status:"Done",activity_type:"Inspeksi Hygiene"}:n==="gcdc"&&(c={period:"Q3",status:"Done"}),L({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",defaultFilters:c,onDataLoaded:p=>p.sort((u,d)=>{let g=u.opening_date?new Date(u.opening_date).getTime():0;return(d.opening_date?new Date(d.opening_date).getTime():0)-g}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:p=>st(p)},{key:"period",label:"Periode",render:p=>oe(p)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:p=>o(p)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:p=>o(p)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:p=>o(p)},{key:"status",label:"Status",render:p=>q(p)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:gt},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:He}],formFields:p=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:gt,value:p?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:p?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:p?.period},{name:"pic",label:"PIC",type:"combobox",options:a(p?.pic),value:p?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:p?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:p?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:p?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:p?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:p?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let p=await k(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(p.ok){let u=p.data.data.map(d=>({Cabang:d.branch_name||"",Kegiatan:d.activity_type||"",Periode:d.period||"",PIC:d.pic||"","Tgl Opening":d.opening_date||"","Tgl Target":d.target_date||"","Tgl Selesai":d.completion_date||"",Status:d.status||""}));B(u,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async p=>{let d=(await k("/api/branches?all=1")).data?.data||[],g=b=>{if(!b)return null;let x=String(b||"").toLowerCase(),T=d.find(S=>String(S.full_name||"").toLowerCase()===x||String(S.code||"").toLowerCase()===x||String(S.name||"").toLowerCase()===x);return T?T.id:null},h=b=>{if(b==null||b==="")return"";if(b instanceof Date&&!isNaN(b.getTime()))return b.toISOString().slice(0,10);let x=String(b).trim();if(x===""||x==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(x))return x.slice(0,10);if(/^\d{4,5}$/.test(x)){let S=Number(x);if(S>2e4&&S<99999){let w=new Date(Date.UTC(1899,11,30)+S*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}let T=x.split(/[\/\-\.]/);if(T.length===3){let[S,w,_]=T.map(M=>M.trim());if(S.length===4&&w.length<=2&&_.length<=2)return`${S}-${w.padStart(2,"0")}-${_.padStart(2,"0")}`;if(_.length===4&&w.length<=2&&S.length<=2)return`${_}-${w.padStart(2,"0")}-${S.padStart(2,"0")}`}return x},m=p.map(b=>({branch_id:g(String(b.Cabang||"").trim()),activity_type:String(b.Kegiatan||"").trim(),period:String(b.Periode||"").trim(),pic:String(b.PIC||b.Pic||"").trim(),opening_date:h(b["Tgl Opening"]||b["Tanggal Opening"]||b["Tgl Openir"]),target_date:h(b["Tgl Target"]||b["Tanggal Target"]),completion_date:h(b["Tgl Selesai"]||b["Tanggal Selesai"]),status:String(b.Status||"").trim(),notes:String(b.Catatan||b.Keterangan||"").trim()})).filter(b=>b.activity_type&&b.period),y=await k("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!y.ok)throw new Error(y.data?.error||"Import gagal");return y.data}}})}F();K();var bt=[],Ye=[];function Da(t,e){let r=String(t.status||"").toLowerCase();return e==="open"?r==="open":!1}async function Ft(t,e){let r=e?e.get("dash_filter"):null;bt=await H(),Ye=await ee();let a=s=>s&&!Ye.find(i=>i.value===s)?[...Ye,{value:s,label:s}]:Ye,l=new Date().getFullYear();L({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:s=>r?s.filter(i=>Da(i,r)):s,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:s=>`<span class="badge badge-secondary">${s}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:s=>`<span title="${s}">${s?.length>50?s.slice(0,50)+"\u2026":s}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:s=>`<span title="${s||""}">${s?.length>40?s.slice(0,40)+"\u2026":s||"-"}</span>`},{key:"status",label:"Status",render:s=>q(s)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:s=>window.formatDate(s)},{key:"day_count",label:"Hari",render:s=>s??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:bt},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:["2025","2026","2027","2028","2029","2030"]}],formFields:s=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:s?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:bt,value:s?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:s?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:s?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:s?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(s?.employee_name),value:s?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(s?.fc_specialist),value:s?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:s?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:s?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let s=await k(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(s.ok){let i=s.data.data.map(n=>({Tanggal:n.report_date||"",Cabang:n.branch_name||"",Kategori:n.category||"",Sumber:n.source||"",Keluhan:n.complaint||"","Nama FC":n.employee_name||"","FC Spesialis":n.fc_specialist||"",Solusi:n.solution||"","Tgl Selesai":n.completion_date||"",Status:n.status||""}));B(i,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async s=>{let n=(await k("/api/branches?all=1")).data?.data||[],c=d=>{if(!d)return null;let g=String(d||"").toLowerCase(),h=n.find(m=>String(m.full_name||"").toLowerCase()===g||String(m.code||"").toLowerCase()===g||String(m.name||"").toLowerCase()===g);return h?h.id:null},p=s.map(d=>({branch_id:c(String(d.Cabang||"").trim()),report_date:String(d.Tanggal||"").trim(),category:String(d.Kategori||"").trim(),source:String(d.Sumber||"").trim(),complaint:String(d.Keluhan||"").trim(),employee_name:String(d["Nama FC"]||"").trim(),fc_specialist:String(d["FC Spesialis"]||"").trim(),solution:String(d.Solusi||"").trim(),completion_date:String(d["Tgl Selesai"]||"").trim(),status:String(d.Status||"").trim()})).filter(d=>d.report_date&&d.complaint&&d.category),u=await k("/api/import/issues",{method:"POST",body:JSON.stringify({rows:p,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}F();var $e=[];function Ia(t,e){let r=String(t.status||"").toLowerCase();return e==="pending"?r==="pending":!1}async function Mt(t,e){let r=e?e.get("dash_filter"):null;$e=await H();let a=await ee(),l=["Ade","Berlin"],o=i=>i&&!a.find(n=>n.value===i)?[...a,{value:i,label:i}]:a,s=i=>i&&!l.find(n=>(typeof n=="object"?n.value:n)===i)?[...l,i]:l;L({container:t,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:i=>r?i.filter(n=>Ia(n,r)):i,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:i=>`<span title="${i||""}">${i?.length>50?i.slice(0,50)+"\u2026":i||"-"}</span>`},{key:"solution",label:"Solusi",render:i=>`<span title="${i||""}">${i?.length>40?i.slice(0,40)+"\u2026":i||"-"}</span>`},{key:"status",label:"Status",render:i=>q(i)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:i=>window.formatDate(i)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:i=>i?`<a href="${i}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"combobox",name:"branch_id",label:"Cabang",options:$e},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async i=>{let n=new URLSearchParams(i||{}).toString(),c=await k(`/api/one-on-one?limit=10000&${n}`);if(c.ok){let p=c.data.data.map(d=>({Tanggal:d.meeting_date||"",Cabang:d.branch_name||"","Nama Karyawan":d.employee_name||"",PIC:d.pic||"",Masalah:d.problem||"",Solusi:d.solution||"",Status:d.status||"","Tgl Selesai":d.completion_date||"",Dokumen:d.document_link||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(K(),le));u(p,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let i=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(K(),le));n(i,"Template_Import_OneOnOne")},onImport:async i=>{let n=d=>{if(!d)return null;let g=String(d||"").toLowerCase(),h=$e.find(m=>String(m.label||"").toLowerCase()===g);return h?h.value:null},c=d=>{if(!d)return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let g=String(d).trim();if(/^\d{4,5}$/.test(g)){let m=Number(g);if(m>2e4&&m<99999){let y=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);let h=g.split(/[\/\-\.]/);if(h.length===3){let[m,y,b]=h.map(x=>x.trim());if(m.length===4&&y.length<=2&&b.length<=2)return`${m}-${y.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&y.length<=2&&m.length<=2)return`${b}-${y.padStart(2,"0")}-${m.padStart(2,"0")}`}return g},p=i.map(d=>({meeting_date:c(d.Tanggal),employee_name:String(d["Nama Karyawan"]||"").trim(),branch_id:n(String(d.Cabang||"").trim()),pic:String(d.PIC||"").trim(),problem:String(d.Masalah||"").trim(),solution:String(d.Solusi||"").trim(),status:String(d.Status||"").trim(),completion_date:c(d["Tgl Selesai"]),document_link:String(d.Dokumen||"").trim()})).filter(d=>d.meeting_date&&d.employee_name&&d.branch_id),u=await k("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:p,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}},formFields:i=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:i?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:i?.branch_id&&!$e.find(n=>n.value==i.branch_id)?[...$e,{value:i.branch_id,label:i.branch_name||i.branch_id}]:$e,createApi:{path:"/api/branches",field:"full_name"},value:i?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"combobox",required:!0,options:o(i?.employee_name),value:i?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:s(i?.pic),createApi:{path:"/api/pic",field:"name"},value:i?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:i?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:i?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:i?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:i?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:i?.document_link}]})}F();async function Ot(t){let e=await H(),r=await ee(),a=["Ade","Berlin"],l=i=>i&&!r.find(n=>n.value===i)?[...r,{value:i,label:i}]:r,o=i=>i&&!a.find(n=>(typeof n=="object"?n.value:n)===i)?[...a,i]:a,s=Array.from({length:5},(i,n)=>String(new Date().getFullYear()-n));L({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:i=>{try{let n=JSON.parse(i);return Array.isArray(n)?n.join(", "):i||"-"}catch{return i||"-"}}},{key:"score",label:"Nilai",render:i=>i!=null?`<strong>${i}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:i=>i?`<a href="${i}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:s}],exportOptions:{moduleName:"training",onExport:async i=>{let n=new URLSearchParams(i||{}).toString(),c=await k(`/api/training?limit=10000&${n}`);if(c.ok){let p=c.data.data.map(d=>{let g=d.participants||"";try{let h=JSON.parse(g);g=Array.isArray(h)?h.join(", "):g}catch{}return{Tanggal:d.training_date||"",Batch:d.batch||"",Materi:d.subject||"",Cabang:d.branch_name||"",Trainer:d.trainer||"",Peserta:g,Nilai:d.score!==null&&d.score!==void 0?d.score:"",Dokumen:d.document_link||""}}),{downloadExcel:u}=await Promise.resolve().then(()=>(K(),le));u(p,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let i=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(K(),le));n(i,"Template_Import_Training")},onImport:async i=>{let n=d=>{if(!d)return null;let g=String(d||"").toLowerCase(),h=e.find(m=>String(m.label||"").toLowerCase()===g);return h?h.value:null},c=d=>{if(!d)return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let g=String(d).trim();if(/^\d{4,5}$/.test(g)){let m=Number(g);if(m>2e4&&m<99999){let y=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);let h=g.split(/[\/\-\.]/);if(h.length===3){let[m,y,b]=h.map(x=>x.trim());if(m.length===4&&y.length<=2&&b.length<=2)return`${m}-${y.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&y.length<=2&&m.length<=2)return`${b}-${y.padStart(2,"0")}-${m.padStart(2,"0")}`}return g},p=i.map(d=>({training_date:c(d.Tanggal),batch:String(d.Batch||"").trim(),subject:String(d.Materi||"").trim(),branch_id:n(String(d.Cabang||"").trim()),trainer:String(d.Trainer||"").trim(),participants:String(d.Peserta||"").trim(),score:d.Nilai?Number(d.Nilai):null,document_link:String(d.Dokumen||"").trim()})).filter(d=>d.training_date&&d.subject&&d.branch_id),u=await k("/api/import/training",{method:"POST",body:JSON.stringify({rows:p,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}},formFields:i=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:i?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:i?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:i?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:i?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:o(i?.trainer),value:i?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let n=JSON.parse(i?.participants);return Array.isArray(n)?n.join(", "):i?.participants||""}catch{return i?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:i?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:i?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:i?.notes}],onBeforeSubmit:async i=>(i.participants&&(i.participants=JSON.stringify(i.participants.split(",").map(n=>n.trim()).filter(Boolean))),i)})}F();ye();K();function Rt({container:t,title:e,icon:r,apiPath:a,columns:l,formFields:o,filterFields:s,defaultFilters:i={},itemLabel:n="Data",canCreate:c=!0,canEdit:p=!0,canDelete:u=!0,onBeforeSubmit:d,onAfterLoad:g,onDataLoaded:h,extraActions:m=[],initialSearch:y="",exportOptions:b=null,bulkDelete:x=!1,paginationMode:T="server"}){let S=1,w={...i};y&&(w.search=y);let _=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${r} ${e}</h1>
      <div class="page-actions">
        ${c?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
      </div>
    </div>

    ${x?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${b?Re(b.moduleName):""}

    ${s&&s.length>0?`
    <div class="filter-bar card" style="padding: 1rem;">
        ${s.map(v=>{if(v.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${w.search||""}"></div>`;if(v.type==="search-combo"){let f="dl-filter-search",$=(v.options||[]).map(P=>`<option value="${typeof P=="object"?P.label:P}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${f}" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${w.search||""}"><datalist id="${f}">${$}</datalist></div>`}return v.type==="select"?`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${v.name}" id="filter-${v.name}"><option value="">-- ${v.label} --</option>${(v.options||[]).map(f=>`<option value="${typeof f=="object"?f.value:f}" ${w[v.name]===(typeof f=="object"?f.value:f)?"selected":""}>${typeof f=="object"?f.label:f}</option>`).join("")}</select>`:""}).join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function M(){if(!document.getElementById("bulk-toolbar"))return;let f=document.getElementById("bulk-count"),$=document.getElementById("btn-bulk-delete"),P=document.getElementById("btn-bulk-cancel");f.textContent=`${_.size} item dipilih`,_.size>0?($.disabled=!1,P.disabled=!1):($.disabled=!0,P.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{_.clear(),document.querySelectorAll(".row-checkbox").forEach(f=>f.checked=!1);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1),M()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(_.size===0)return;let v=[..._],f=document.createElement("div");f.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",f.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${v.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${v.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(f),f.querySelector("#bulk-cancel-btn").addEventListener("click",()=>f.remove()),f.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let $=f.querySelector("#bulk-confirm-btn");$.disabled=!0,$.textContent="Menghapus...";let P=await k(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:v})});f.remove(),P.ok?(Z(`${v.length} ${n} berhasil dihapus.`),_.clear(),M(),O()):G(P.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),R;if(E?.addEventListener("input",v=>{clearTimeout(R),R=setTimeout(()=>{w.search=v.target.value,S=1,_.clear(),O()},400)}),s?.forEach(v=>{v.type==="select"&&document.getElementById(`filter-${v.name}`)?.addEventListener("change",f=>{w[v.name]=f.target.value,S=1,_.clear(),O()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{w={...i},E&&(E.value=""),s?.forEach(v=>{let f=document.getElementById(`filter-${v.name}`);f&&(f.value="")}),S=1,_.clear(),O()}),document.getElementById("btn-create")?.addEventListener("click",()=>ge(null)),b){document.getElementById(`btn-export-${b.moduleName}`)?.addEventListener("click",async f=>{let $=f.target,P=$.innerHTML;$.innerHTML="\u23F3 Loading...",$.disabled=!0;try{await b.onExport()}catch{G("Gagal export data")}finally{$.innerHTML=P,$.disabled=!1}}),document.getElementById(`btn-template-${b.moduleName}`)?.addEventListener("click",()=>{b.onTemplate()});let v=document.getElementById(`input-import-${b.moduleName}`);v?.addEventListener("change",async f=>{let $=f.target.files[0];if(!$)return;v.disabled=!0;let P=document.createElement("div");P.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",P.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(P);let V=P.querySelector("#import-progress-text"),U=P.querySelector("#import-progress-bar"),D=P.querySelector("#import-summary"),I=P.querySelector("#import-close-btn");I.addEventListener("click",()=>{P.remove(),O()});try{let Y=await Oe($);if(Y.length===0)throw new Error("File kosong atau format salah");let X=500,ie=0,ae=0,N=0,A=Y.length;V.textContent=`Ditemukan ${A} baris data. Memulai import...`;for(let C=0;C<A;C+=X){let z=Y.slice(C,C+X);V.textContent=`Mengimport baris ${C+1} - ${Math.min(C+X,A)} dari ${A}...`,U.style.width=`${Math.round(C/A*100)}%`;try{let j=await b.onImport(z);j?(ie+=j.inserted||j.metrics?.inserted||z.length,ae+=j.skipped||j.metrics?.updated||0):ie+=z.length}catch(j){console.error("Chunk import failed:",j),N+=z.length}}U.style.width="100%",V.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',D.style.display="block",D.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${A}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ie}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${N}</strong></div>
        `,N>0&&(D.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),I.style.display="block",v.value=""}catch(Y){V.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${Y.message}`,U.style.background="var(--danger)",U.style.width="100%",I.style.display="block",v.value=""}finally{v.disabled=!1}})}async function O(){M();let v=document.getElementById("table-container");if(!v)return;v.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let f=T==="client",$=f?1:S,P=f?be:20,V=new URLSearchParams({page:$,limit:P,...Object.fromEntries(Object.entries(w).filter(([,N])=>N))}),U=await k(`${a}?${V}`);if(!U.ok){v.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${U.data?.error||"Error"}</p></div>`;return}let D=U.data?.data||U.data||[],I=U.data?.pagination,Y=D.length,X=D;if(f){D=h(D),X=D;let N=D.length,A=20,C=Math.ceil(N/A);S>C&&C>0&&(S=C);let z=(S-1)*A,j=S*A;D=D.slice(z,j),I={page:S,limit:A,total:N,pages:C}}!1,g&&g(D);let ie=Ue({columns:l,data:D,fullData:X,onEdit:p?N=>ge(N):null,actions:m.map(N=>({...N,handler:A=>N.handler(A,O)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:x?{selectedIds:_,onToggle:M}:null});v.innerHTML="",v.appendChild(ie);let ae=document.getElementById("pagination-container");if(ae&&(ae.innerHTML="",I&&I.pages>1)){let N=Ge({page:I.page,pages:I.pages,total:I.total,limit:I.limit,onPage:A=>{S=A,O()}});N&&ae.appendChild(N)}}function fe(v){let f=typeof o=="function"?o(v):o;return Me(f)}function ge(v){let f=!!v,$=document.createElement("form");if($.noValidate=!0,$.innerHTML=fe(v),f){let V=typeof o=="function"?o(v):o;ze($,v)}let{close:P}=se({title:f?`Edit ${n}`:`Tambah ${n}`,content:$,size:"lg",confirmText:f?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(V,U)=>{if(!$.reportValidity())return;let D=V.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let I=Qe($),Y=typeof o=="function"?o(v):o,X=async A=>{for(let C of A)if(C.type==="row")await X(C.fields);else if(C.type==="combobox"&&I[C.name]){let z=I[C.name],j=(C.options||[]).find(W=>{let ne=String(typeof W=="object"?W.value:W),tt=String(typeof W=="object"?W.label:W);return ne===z||tt===z});if(j)I[C.name]=typeof j=="object"?j.value:j;else if(C.createApi){let W={};W[C.createApi.field]=z,C.createApi.extra&&Object.assign(W,C.createApi.extra);let ne=await k(C.createApi.path,{method:"POST",body:JSON.stringify(W)});if(ne.ok&&ne.data?.id)I[C.name]=ne.data.id;else if(ne.ok&&!ne.data?.id)I[C.name]=z;else throw new Error(`Gagal membuat master data: ${ne.data?.error||"Unknown error"}`)}}};try{await X(Y)}catch(A){G(A.message),D.disabled=!1,D.textContent=f?"Simpan Perubahan":`Tambah ${n}`;return}d&&(I=await d(I,v));let ie=f?"PUT":"POST",ae=f?`${a}/${v.id}`:a,N=await k(ae,{method:ie,body:JSON.stringify(I)});N.ok?(Z(f?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),U(),O()):(G(N.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=f?"Simpan Perubahan":`Tambah ${n}`)}})}function ve(v){Fe(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let f=await k(`${a}/${v.id}`,{method:"DELETE"});f.ok?(Z(`${n} berhasil dihapus.`),O()):G(f.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return O(),O}F();K();async function Kt(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let r=await H(),a=await ee(),l=e?e.get("dash_filter"):null;console.log("RAW",await we()),console.log("OPTIONS",a);let o=n=>n&&!a.find(c=>c.value===n)?[...a,{value:n,label:n}]:a,s=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],i=n=>n&&!s.includes(n)?[...s,n]:s;Rt({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",onDataLoaded:n=>{if(l==="reliever"){let c=new Date,p=c.getFullYear(),u=String(c.getMonth()+1).padStart(2,"0");return n.filter(d=>{if(String(d.status||"").toLowerCase()!=="done")return!1;let g=d.backup_date||"";if(g.includes("/")){let h=g.split("/");if(h.length===3&&(h[2].length===4?h[2]:`20${h[2]}`)==p&&h[1].padStart(2,"0")==u)return!0}else if(g.includes("-")&&g.startsWith(`${p}-${u}`))return!0;return!1})}return n},columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:n=>oe(n)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:n=>window.formatDate(n)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:n=>n?`<span class="badge badge-info">${n}</span>`:"-"},{key:"status",label:"Status",render:n=>q(n)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:s},{type:"select",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:r,value:n?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:o(n?.original_fc_name),value:n?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:i(n?.reliever_name),value:n?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:n?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:n?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:n?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:n?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let n=await k(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let c=n.data.data.map(p=>({Cabang:p.branch_name||"","Nama Facility care":p.original_fc_name||"",Periode:p.period||"",Relifer:p.reliever_name||"","Tanggal Back Up":p.backup_date||"","Tanggal Selesai":p.completion_date||"",Keterangan:p.reason||"",Shift:p.shift||"",Status:p.status||""}));c.length===0&&c.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),B(c,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async n=>{let p=(await k("/api/branches?all=1")).data?.data||[],u=h=>{if(!h)return null;let m=String(h||"").toLowerCase(),y=p.find(b=>String(b.full_name||"").toLowerCase()===m||String(b.code||"").toLowerCase()===m||String(b.name||"").toLowerCase()===m);return y?y.id:null},d=n.map(h=>({branch_name:String(h.Cabang||"").trim(),backup_date:String(h["Tanggal Back Up"]||h["Tanggal Backup"]||"").trim(),original_fc_name:String(h["Nama Facility care"]||h["FC Digantikan"]||"").trim(),reliever_name:String(h.Relifer||h.Reliefer||"").trim(),period:String(h.Periode||"").trim(),reason:String(h.Keterangan||"").trim(),shift:String(h.Shift||"").trim(),completion_date:String(h["Tanggal Selesai"]||"").trim(),status:String(h.Status||"").trim()})).filter(h=>h.reliever_name&&h.backup_date),g=await k("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:d,onDuplicate:"update"})});if(!g.ok)throw new Error(g.data?.error||"Import gagal");return g.data}}})}F();K();async function Ht(t){let e=await H(),r=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));L({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>oe(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>q(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),o=await k(`/api/reports/inspection?limit=10000&${l}`);if(o.ok){let s=o.data.data.map(i=>({Cabang:i.branch_name||"",Periode:i.period||"",Tanggal:i.inspection_date||"","Point FC":i.fc_score!==null&&i.fc_score!==void 0?i.fc_score:"","Point SPV":i.spv_score!==null&&i.spv_score!==void 0?i.spv_score:"",Status:i.status||"","Link Dokumen":i.document_link||""}));B(s,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let l=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),p=e.find(u=>String(u.label||"").toLowerCase()===c);return p?p.value:null},o=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let c=String(n).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let u=Number(c);if(u>2e4&&u<99999){let d=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let p=c.split(/[\/\-\.]/);if(p.length===3){let[u,d,g]=p.map(h=>h.trim());if(u.length===4&&d.length<=2&&g.length<=2)return`${u}-${d.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&d.length<=2&&u.length<=2)return`${g}-${d.padStart(2,"0")}-${u.padStart(2,"0")}`}return c},s=a.map(n=>({branch_id:l(String(n.Cabang||"").trim()),period:String(n.Periode||"").trim(),inspection_date:o(n.Tanggal),fc_score:n["Point FC"]!==void 0&&n["Point FC"]!==""?Number(n["Point FC"]):null,spv_score:n["Point SPV"]!==void 0&&n["Point SPV"]!==""?Number(n["Point SPV"]):null,status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.inspection_date),i=await k("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:s,onDuplicate:"update"})});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}F();K();async function jt(t){let e=await H(),r=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));L({container:t,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>oe(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>q(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),o=await k(`/api/reports/cleaning?limit=10000&${l}`);if(o.ok){let s=o.data.data.map(i=>({Cabang:i.branch_name||"",Jenis:i.activity_type||"",Periode:i.period||"",Tanggal:i.activity_date||"",Status:i.status||"","Link Dokumen":i.document_link||""}));B(s,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let l=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),p=e.find(u=>String(u.label||"").toLowerCase()===c);return p?p.value:null},o=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let c=String(n).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let u=Number(c);if(u>2e4&&u<99999){let d=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let p=c.split(/[\/\-\.]/);if(p.length===3){let[u,d,g]=p.map(h=>h.trim());if(u.length===4&&d.length<=2&&g.length<=2)return`${u}-${d.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&d.length<=2&&u.length<=2)return`${g}-${d.padStart(2,"0")}-${u.padStart(2,"0")}`}return c},s=a.map(n=>({branch_id:l(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:o(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),i=await k("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:s,onDuplicate:"update"})});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}F();K();async function qt(t){let e=await H(),r=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));L({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge badge-warning">${a}</span>`},{key:"period",label:"Periode",render:a=>oe(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>q(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:a=>a||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:a?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),o=await k(`/api/reports/fogging?limit=10000&${l}`);if(o.ok){let s=o.data.data.map(i=>({Cabang:i.branch_name||"",Jenis:i.activity_type||"Fogging",Periode:i.period||"",Tanggal:i.activity_date||"",Status:i.status||"","Link Dokumen":i.document_link||""}));B(s,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async a=>{let l=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),p=e.find(u=>String(u.label||"").toLowerCase()===c);return p?p.value:null},o=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let c=String(n).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let u=Number(c);if(u>2e4&&u<99999){let d=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let p=c.split(/[\/\-\.]/);if(p.length===3){let[u,d,g]=p.map(h=>h.trim());if(u.length===4&&d.length<=2&&g.length<=2)return`${u}-${d.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&d.length<=2&&u.length<=2)return`${g}-${d.padStart(2,"0")}-${u.padStart(2,"0")}`}return c},s=a.map(n=>({branch_id:l(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"Fogging").trim(),period:String(n.Periode||"").trim(),activity_date:o(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.activity_date),i=await k("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(s)});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}F();K();async function Jt(t){let e=await H(),r=await ee(),a=r,l=Array.from({length:4},(i,n)=>String(new Date().getFullYear()-n)),o=i=>i&&!r.find(n=>n.value===i)?[...r,{value:i,label:i}]:r,s=i=>i&&!a.find(n=>n.value===i)?[...a,{value:i,label:i}]:a;L({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:i=>window.formatDate(i)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:i=>`<span title="${i||""}">${i?.length>60?i.slice(0,60)+"\u2026":i||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:i=>window.formatDate(i)},{key:"status",label:"Status",render:i=>q(i)},{key:"notes",label:"Keterangan",render:i=>i?.length>40?i.slice(0,40)+"\u2026":i||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade","Mizwar"]},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:l}],formFields:i=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:i?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:s(i?.pic),value:i?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:i?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:i?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:i?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:i?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:i?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async i=>{let n=new URLSearchParams(i||{}).toString(),c=await k(`/api/reports/basecamp?limit=10000&${n}`);if(c.ok){let p=c.data.data.map(u=>({"Tgl Info":u.info_date||"",Cabang:u.branch_name||"",Permasalahan:u.problem||"",PIC:u.pic||"","Tgl Done":u.done_date||"",Status:u.status||"",Keterangan:u.notes||""}));B(p,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async i=>{let n=d=>{if(!d)return null;let g=String(d||"").toLowerCase(),h=e.find(m=>String(m.label||"").toLowerCase()===g);return h?h.value:null},c=d=>{if(d==null||d==="")return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let g=String(d).trim();if(g===""||g==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);if(/^\d{4,5}$/.test(g)){let m=Number(g);if(m>2e4&&m<99999){let y=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}let h=g.split(/[\/\-\.]/);if(h.length===3){let[m,y,b]=h.map(x=>x.trim());if(m.length===4&&y.length<=2&&b.length<=2)return`${m}-${y.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&y.length<=2&&m.length<=2)return`${b}-${y.padStart(2,"0")}-${m.padStart(2,"0")}`}return g},p=i.map(d=>({info_date:c(d["Tgl Info"]||d["Tanggal Info"]),branch_id:n(String(d.Cabang||"").trim()),problem:String(d.Permasalahan||"").trim(),pic:String(d.PIC||"").trim(),done_date:c(d["Tgl Done"]||d["Tanggal Done"]),status:String(d.Status||"").trim(),notes:String(d.Keterangan||d.Catatan||"").trim()})).filter(d=>d.info_date&&d.branch_id&&d.problem),u=await k("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(p)});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}async function Ut(t){L({container:t,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(F(),ke)),l=await a(`/api/sop?limit=10000&${r}`);if(l.ok){let o=l.data.data.map(i=>({"Nama SOP":i.name||"",Kategori:i.category||"",Dokumen:i.document_link||"",Catatan:i.notes||i.description||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(K(),le));s(o,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:r}=await Promise.resolve().then(()=>(K(),le));r(e,"Template_Import_SOP")},onImport:async e=>{let r=e.map(o=>({name:String(o["Nama SOP"]||"").trim(),category:String(o.Kategori||"").trim(),document_link:String(o.Dokumen||"").trim(),description:String(o.Catatan||"").trim()})).filter(o=>o.name),{apiFetch:a}=await Promise.resolve().then(()=>(F(),ke)),l=await a("/api/sop/import",{method:"POST",body:JSON.stringify(r)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function Gt(t){L({container:t,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(F(),ke)),l=await a(`/api/checklist?limit=10000&${r}`);if(l.ok){let o=l.data.data.map(i=>({"Nama Checklist":i.name||"",Kategori:i.category||"",Dokumen:i.document_link||"",Deskripsi:i.description||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(K(),le));s(o,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:r}=await Promise.resolve().then(()=>(K(),le));r(e,"Template_Import_Checklist")},onImport:async e=>{let r=e.map(o=>({name:String(o["Nama Checklist"]||"").trim(),category:String(o.Kategori||"").trim(),document_link:String(o.Dokumen||"").trim(),description:String(o.Deskripsi||"").trim()})).filter(o=>o.name),{apiFetch:a}=await Promise.resolve().then(()=>(F(),ke)),l=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(r)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}F();ye();K();async function ht(t,e="forms"){if(e==="supply")return Ba(t);Pa(t)}function Pa(t){L({container:t,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await k(`/api/forms?limit=10000&${r}`);a.data?.data?B(a.data.data,"Data_Master_Form"):G("Gagal export data master form")},onImport:async e=>{let r=await k("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!r.ok)throw new Error(r.data?.error||"Import failed");return r.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function Ba(t){let r=((await k("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));L({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,l)=>l.branch_name_ref||l.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>q(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let l=a?.tools_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}let o=a?.chemical_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:a?.branch_id&&!r.find(s=>s.value==a.branch_id)?[...r,{value:a.branch_id,label:a.branch_name||a.branch_id}]:r,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:l},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:o},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),o=await k(`/api/reports/supply?limit=10000&${l}`);if(o.ok){let s=o.data.data.map(i=>{let n=i.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let c=i.chemical_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}return{Waktu:i.submitted_at||"",Pengirim:i.submitter_name||"",Cabang:i.branch_name_ref||i.branch_name||"","Alat/Barang":n||"",Chemical:c||"",Catatan:i.additional_notes||"",Status:i.status||"","Diproses Oleh":i.processed_by||""}});B(s,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let o=(await k("/api/branches?all=1")).data?.data||[],s=p=>{if(!p)return null;let u=String(p||"").toLowerCase(),d=o.find(g=>String(g.full_name||"").toLowerCase()===u||String(g.code||"").toLowerCase()===u||String(g.name||"").toLowerCase()===u);return d?d.id:null},i=p=>{if(p==null||p==="")return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let u=String(p).trim();if(u===""||u==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);if(/^\d{4,5}$/.test(u)){let g=Number(u);if(g>2e4&&g<99999){let h=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}let d=u.split(/[\/\-\.]/);if(d.length===3){let[g,h,m]=d.map(y=>y.trim());if(g.length===4&&h.length<=2&&m.length<=2)return`${g}-${h.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&h.length<=2&&g.length<=2)return`${m}-${h.padStart(2,"0")}-${g.padStart(2,"0")}`}return u},n=a.map(p=>({submitted_at:i(p.Waktu||p.Tanggal),submitter_name:String(p.Pengirim||"").trim(),branch_id:s(String(p.Cabang||"").trim()),tools_items:String(p["Alat/Barang"]||p.Alat||"").trim(),chemical_items:String(p.Chemical||"").trim(),additional_notes:String(p.Catatan||p.Keterangan||"").trim(),status:String(p.Status||"").trim(),processed_by:String(p["Diproses Oleh"]||p.PIC||"").trim()})).filter(p=>p.submitted_at&&p.submitter_name&&p.branch_id),c=await k("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,l)=>{let o=se({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(s,i)=>{let n=s.querySelector("#supply-status").value,c=s.querySelector("#supply-processed-by").value;(await k(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:c})})).ok?(Z("Status diperbarui."),i(),l()):G("Gagal update status.")}})}}]})}F();K();async function Qt(t){let e=he();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}L({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:r=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[r]||"badge-neutral"}">${r}</span>`},{key:"is_active",label:"Status",render:r=>r?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:r=>r?new Date(r).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:r=>{let a=!!r;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:r?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:r?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:r?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:r?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?r?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let r=await k(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let a=r.data.data.map(l=>({"Nama Lengkap":l.full_name||"",Username:l.username||"",Email:l.email||"",Role:l.role||"",Status:l.is_active?"Aktif":"Nonaktif"}));B(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async r=>{let a=r.map(o=>({full_name:String(o["Nama Lengkap"]||"").trim(),username:String(o.Username||"").trim(),email:String(o.Email||"").trim(),role:String(o.Role||"").trim()||"viewer",password:String(o.Password||"").trim()})).filter(o=>o.username&&o.password&&o.email&&o.full_name),l=await k("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}F();K();async function zt(t){L({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await k(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)B(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let r=e.map(l=>({code:String(l["Kode Cabang"]||"").trim(),name:String(l["Nama Pendek"]||"").trim(),full_name:String(l["Nama Lengkap"]||"").trim(),city:String(l.Kota||"").trim()})).filter(l=>l.code&&l.name),a=await k("/api/branches/import",{method:"POST",body:JSON.stringify(r)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}F();async function Vt(t){let e=new Date,r=[];t.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),l()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),l()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(o=>o.addEventListener("change",l));async function a(){try{let o=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;r=(await k(`/api/dashboard/calendar?month=${o}`)).data?.data||[]}catch(o){console.warn("[Calendar] Failed to load events, rendering empty grid:",o),r=[]}}async function l(){let o=document.getElementById("calendar-grid");if(o){o.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let s=e.getFullYear(),i=e.getMonth(),n=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),c=document.getElementById("cal-month-label");c&&(c.textContent=n);let p=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(S=>S.value)),u=r.filter(S=>p.has(S.type)),d={};u.forEach(S=>{let w=(S.event_date||"").slice(0,10);d[w]||(d[w]=[]),d[w].push(S)});let g=new Date(s,i,1).getDay(),h=new Date(s,i+1,0).getDate(),m=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],y=new Date().toISOString().slice(0,10),b='<div class="calendar-grid">';m.forEach(S=>{b+=`<div class="cal-day-header">${S}</div>`});for(let S=0;S<g;S++)b+='<div class="cal-cell cal-cell-empty"></div>';for(let S=1;S<=h;S++){let w=`${s}-${String(i+1).padStart(2,"0")}-${String(S).padStart(2,"0")}`,_=d[w]||[],M=w===y;b+=`
          <div class="cal-cell ${M?"cal-today":""} ${_.length?"cal-has-events":""}"
               data-date="${w}" tabindex="0" role="button" aria-label="${w}">
            <div class="cal-day-num ${M?"today-num":""}">${S}</div>
            <div class="cal-events-preview">
              ${_.slice(0,3).map(E=>`
                <div class="cal-event-dot cal-color-${E.color||"gray"}" title="${Xe(E.title||E.type)}">
                  <span class="cal-event-dot-label">${La(E.title||E.branch_name||E.type,18)}</span>
                </div>
              `).join("")}
              ${_.length>3?`<div class="cal-more">+${_.length-3} lagi</div>`:""}
            </div>
          </div>`}let T=(g+h)%7;if(T!==0)for(let S=0;S<7-T;S++)b+='<div class="cal-cell cal-cell-empty"></div>';b+="</div>",o.innerHTML=b,o.querySelectorAll(".cal-cell[data-date]").forEach(S=>{S.addEventListener("click",()=>{let w=S.dataset.date,_=d[w]||[];if(!_.length)return;let M=document.getElementById("cal-event-list"),E=new Date(w+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=E,document.getElementById("cal-event-items").innerHTML=_.map(R=>`
            <div class="cal-event-item cal-color-border-${R.color||"gray"}">
              <div class="cal-event-type">${Na(R.type)}</div>
              <div class="cal-event-title">${Xe(R.title||"-")}</div>
              <div class="cal-event-branch">${Xe(R.branch_name||"")}</div>
              ${R.status?`<div class="cal-event-status">${Xe(R.status)}</div>`:""}
              ${R.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${R.days_remaining} hari</div>`:""}
            </div>
          `).join(""),M.style.display="block"})})}catch(s){console.error("[Calendar] Render error:",s),o&&(o.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}l()}function La(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function Xe(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Na(t){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[t]||t}F();async function Wt(t){let e=he(),r=(e?.full_name||e?.username||"U")[0].toUpperCase(),l={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
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
  `;let o=localStorage.getItem("fm_token"),s=document.getElementById("session-info");if(o&&s)try{let i=JSON.parse(atob(o.split(".")[1])),n=new Date(i.exp*1e3);s.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{s.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async i=>{i.preventDefault();let n=document.getElementById("pwd-error"),c=document.getElementById("pwd-success"),p=document.getElementById("btn-save-pwd");n.style.display="none",c.style.display="none";let u=i.target,d=u.current_password.value,g=u.new_password.value,h=u.confirm_password.value;if(g!==h){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(g.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}p.disabled=!0,p.textContent="\u23F3 Menyimpan...";let m=await k("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:d,new_password:g})});p.disabled=!1,p.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',m.ok?(c.textContent="\u2705 Password berhasil diubah.",c.style.display="block",u.reset(),Z("Password berhasil diubah.")):(n.textContent=m.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}F();var Ze={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function te(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let l=Number(e);if(l>2e4&&l<99999){let o=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(o.getTime())?null:o.toISOString().slice(0,10)}}let r=e.split(/[\/\-\.]/);if(r.length===3){let[l,o,s]=r.map(p=>p.trim()),i=Number(l),n=Number(o),c=Number(s);if(l.length===4&&i>1900)return`${l}-${o.padStart(2,"0")}-${s.padStart(2,"0")}`;if(s.length===4&&c>1900)return i>12?`${s}-${o.padStart(2,"0")}-${l.padStart(2,"0")}`:n>12?`${s}-${l.padStart(2,"0")}-${o.padStart(2,"0")}`:`${s}-${o.padStart(2,"0")}-${l.padStart(2,"0")}`;if(s.length===2&&!isNaN(c)){let p=c>=50?`19${s}`:`20${s}`;return i>12?`${p}-${o.padStart(2,"0")}-${l.padStart(2,"0")}`:`${p}-${o.padStart(2,"0")}-${l.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function Yt(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Aa={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",phone:t["No. Hp"],join_date:te(t["Tanggal Masuk"]),status:t.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:te(t["Tanggal Mulai"]),end_date:te(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:te(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:te(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:te(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:te(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:te(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:te(t["Tanggal Target"]||t["Tgl Target"]),completion_date:te(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:te(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:te(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:te(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:te(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:te(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:te(t["Tanggal Back Up"]),completion_date:te(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:te(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:te(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Fa(t,e){let r=Ze[t];if(!r)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Aa[r.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let l=[],o=[],s=[];return e.filter(n=>!Yt(n)).forEach((n,c)=>{let p=e.indexOf(n)+2,u=[];a.required.forEach(({key:g,label:h})=>{let m=n[g];if(m==null||String(m).trim()===""){let y=Object.keys(n).filter(b=>b.trim()).join(", ");u.push({column:h,originalValue:m||"",reason:`Kolom "${h}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${y.slice(0,120)}`})}});let d=a.map(n);u.length>0?o.push({row:p,data:d,raw:n,errors:u}):(l.push(n),s.push(d))}),{valid:l,errors:o,mapped:s}}function Xt(t){let e=[];return t.SheetNames.forEach(r=>{let a=Ze[r];if(!a)return;let l=t.Sheets[r],o=window.XLSX.utils.sheet_to_json(l,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),s=Fa(r,o),i=o.filter(n=>!Yt(n));e.push({sheetName:r,module:a.module,label:a.label,total:i.length,valid:s.mapped.length,errorCount:s.errors.length,errors:s.errors,mapped:s.mapped,skipped:!1})}),e}function Zt(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,l])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(l),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function ea(t){let e=window.XLSX,r=e.utils.book_new(),a=!1;return t.forEach(l=>{if(!l.errors||l.errors.length===0)return;a=!0;let o=l.errors.map(i=>({"No. Baris":i.row,"Kolom Gagal":(i.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(i.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(i.data||{}).map(([n,c])=>[n,c??""]))})),s=e.utils.json_to_sheet(o);e.utils.book_append_sheet(r,s,l.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(r,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Ma=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function ta(t){t.innerHTML=`
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
              ${Object.entries(Ze).map(([m,{label:y}])=>`<span class="import-sheet-tag">\u{1F4C4} ${m} \u2192 ${y}</span>`).join("")}
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
  `;let e=null,r=null,a=0,l={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function o(m){Object.entries(l).forEach(([y,b])=>{b.style.display=y===m?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let m=document.getElementById("btn-backup-db");m.disabled=!0,m.textContent="\u23F3 Memproses Backup...";try{let y=await k("/api/import/backup");if(y.ok){if(!window.XLSX){G("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let b=window.XLSX,x=b.utils.book_new();Object.entries(y.data.database).forEach(([T,S])=>{let w=S.length>0?S:[{}],_=b.utils.json_to_sheet(w);b.utils.book_append_sheet(x,_,T.substring(0,31))}),b.writeFile(x,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),Z("Backup berhasil diunduh!")}else G("Gagal memproses backup: "+(y.data?.error||"Unknown error"))}catch(y){G("Gagal memproses backup: "+y.message)}finally{m.disabled=!1,m.textContent="\u{1F4E6} Backup Database"}});let s=document.getElementById("btn-sync-google");s&&s.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let m=s.innerHTML;s.innerHTML='<span class="spinner"></span> Menyinkronkan...',s.disabled=!0;try{let y=await k("/api/sync/google-sheets",{method:"POST"});y.ok?alert("Sinkronisasi Berhasil: "+(y.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(y.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{s.innerHTML=m,s.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Zt(),Z("Template Excel berhasil didownload!")});let i=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",m=>{m.stopPropagation(),i.click()}),i.addEventListener("change",m=>{m.target.files[0]&&c(m.target.files[0])}),n.addEventListener("dragover",m=>{m.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",m=>{m.preventDefault(),n.classList.remove("drag-over");let y=m.dataTransfer.files[0];y&&y.name.match(/\.xlsx?$/i)?c(y):G("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,i.value="",document.getElementById("file-info").style.display="none",n.style.display="",o("upload")});async function c(m){e=m,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${m.name} (${(m.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await p(m)}async function p(m){o("validating");let y=document.getElementById("validation-status"),b=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");y.textContent="Membaca file Excel...",b.style.width="20%",await je(200);let x=await m.arrayBuffer(),T=window.XLSX.read(x,{type:"array",cellDates:!0});y.textContent=`Memvalidasi ${T.SheetNames.length} sheet...`,b.style.width="50%",await je(100),r=Xt(T),b.style.width="100%",y.textContent="Validasi selesai!",await je(300),u()}catch(x){o("upload"),G("Gagal memproses file: "+x.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function u(){o("preview");let m=r.filter(E=>!E.skipped).length,y=r.reduce((E,R)=>E+R.total,0),b=r.reduce((E,R)=>E+R.valid,0),x=r.reduce((E,R)=>E+R.errorCount,0),T=y>0?Math.round(b/y*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${m} sheet</span>
      <span class="badge badge-secondary">${y} baris</span>
      <span class="badge badge-success">${b} valid (${T}%)</span>
      ${x>0?`<span class="badge badge-danger">${x} error</span>`:""}
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
          ${r.map((E,R)=>`
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
                ${E.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${R}">\u{1F50D} ${E.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,S.querySelectorAll(".btn-detail-error").forEach(E=>{E.addEventListener("click",()=>{let R=r[Number(E.dataset.idx)];d(R)})});let w=document.getElementById("error-detail-section"),_=document.getElementById("error-detail-container");_.innerHTML="",w.style.display="none";let M=document.getElementById("btn-start-import");b===0?(M.disabled=!0,M.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(M.disabled=!1,x>0?(M.innerHTML=`\u{1F680} Import ${b} Data Valid (${x} dilewati)`,M.title="Baris error akan dilewati, baris valid tetap diimport"):M.innerHTML=`\u{1F680} Mulai Import ${b} Data`)}function d(m){let y=document.getElementById("error-detail-section"),b=document.getElementById("error-detail-container");y.style.display="";let x=m.errors.slice(0,100).map(T=>(Array.isArray(T.errors)?T.errors:[]).map(w=>{let _=typeof w=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${T.row}</span></td>
            <td><strong>${_?w.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${_&&w.originalValue!==void 0?w.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${_?w.reason:w}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${_&&w.aliases?`Gunakan salah satu nama kolom:<br><em>${w.aliases}</em>`:_&&w.hint?w.hint:""}
            </td>
          </tr>
        `}).join("")).join("");b.innerHTML=`
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
            <tbody>${x||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${m.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,y.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{o("upload"),document.getElementById("file-info").style.display="none",n.style.display="",e=null,i.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!r)return;ea(r)?Z("Log error berhasil didownload."):Z("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let m=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";g(m)});async function g(m){o("importing"),a=Date.now();let y=[];Ma.forEach(w=>{let _=r?.find(M=>M.module===w&&M.mapped?.length>0);_&&y.push(_)});let b=document.getElementById("import-steps-list");b.innerHTML=y.map(w=>`
      <div class="import-step-item" id="step-item-${w.module}">
        <span class="step-item-icon" id="step-icon-${w.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${w.label} <span class="step-item-count">(${w.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${w.module}"></span>
      </div>
    `).join("");let x=document.getElementById("import-bar"),T=document.getElementById("import-current-status"),S={totalSheets:y.length,totalRows:y.reduce((w,_)=>w+_.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let w=0;w<y.length;w++){let _=y[w],M=document.getElementById(`step-icon-${_.module}`),E=document.getElementById(`step-status-${_.module}`);M.textContent="\u{1F504}",E.textContent="Mengimport...",T.textContent=`Mengimport ${_.label}...`,x.style.width=`${Math.round(w/y.length*100)}%`;try{let R=await k(`/api/import/${_.module}`,{method:"POST",body:JSON.stringify({rows:_.mapped,onDuplicate:m})});if(R.ok){let O=R.data;S.inserted+=O.inserted||0,S.skipped+=O.skipped||0,S.moduleResults.push({label:_.label,inserted:O.inserted||0,skipped:O.skipped||0,status:"ok"}),M.textContent="\u2705",E.innerHTML=`<span class="badge badge-success">${O.inserted||0} berhasil</span>${O.skipped>0?` <span class="badge badge-neutral">${O.skipped} skip</span>`:""}`}else S.failed++,S.moduleResults.push({label:_.label,inserted:0,skipped:0,status:"error",error:R.data?.error}),M.textContent="\u274C",E.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(R){S.failed++,S.moduleResults.push({label:_.label,inserted:0,skipped:0,status:"error",error:R.message}),M.textContent="\u274C",E.innerHTML='<span class="badge badge-danger">Gagal</span>'}await je(150)}x.style.width="100%",T.textContent="Selesai!",await je(400),h(S)}function h(m){o("summary");let y=((Date.now()-a)/1e3).toFixed(1),b=m.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${b?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${b?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
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
        <div class="stat-value">${y}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${m.moduleResults.map(x=>`
            <tr>
              <td>${x.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${x.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${x.skipped}</span></td>
              <td style="text-align:center">
                ${x.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${x.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,r=null,i.value="",document.getElementById("file-info").style.display="none",n.style.display="",o("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function je(t){return new Promise(e=>setTimeout(e,t))}F();var et=[],aa=[];async function na(t){et=await H(),aa=await ee(),L({container:t,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:et}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await k(`/api/sp?limit=10000&${r}`);if(a.ok){let l=a.data.data.map(s=>({"Nama Karyawan":s.employee_name||"",Divisi:s.division||"",Cabang:s.branch_name||"","Tanggal Sp":s.tanggal||"","Akhir Sp":s.akhir_sp||"","Jenis Sp":s.sp_type||"","Link Document / Foto":s.document_link||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(K(),le));o(l,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(K(),le));r(e,"Template_Import_SP")},onImport:async e=>{let r=s=>{if(!s)return null;let i=String(s||"").toLowerCase(),n=et.find(c=>String(c.label||"").toLowerCase()===i);return n?n.value:null},a=s=>{if(!s)return"";if(s instanceof Date&&!isNaN(s.getTime()))return s.toISOString().slice(0,10);let i=String(s).trim();if(/^\d{4,5}$/.test(i)){let c=Number(i);if(c>2e4&&c<99999){let p=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(i))return i.slice(0,10);let n=i.split(/[\/\-\.]/);if(n.length===3){let[c,p,u]=n.map(d=>d.trim());if(c.length===4&&p.length<=2&&u.length<=2)return`${c}-${p.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&p.length<=2&&c.length<=2)return`${u}-${p.padStart(2,"0")}-${c.padStart(2,"0")}`}return i},l=e.map(s=>({employee_name:String(s["Nama Karyawan"]||"").trim(),division:String(s.Divisi||"").trim(),branch_id:r(String(s.Cabang||"").trim()),tanggal:a(s["Tanggal Sp"]),akhir_sp:a(s["Akhir Sp"]),sp_type:String(s["Jenis Sp"]||"").trim(),document_link:String(s["Link Document / Foto"]||"").trim()})).filter(s=>s.employee_name&&s.branch_id),o=await k("/api/import/sp",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}},formFields:[{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:aa},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"combobox",name:"branch_id",label:"Cabang",required:!0,options:et,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}F();var De=[],ia=[];async function ra(t){De=await H(),ia=await ee(),L({container:t,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",options:De},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",options:De}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await k(`/api/mutasi?limit=10000&${r}`);if(a.ok){let l=a.data.data.map(s=>({Tanggal:s.tanggal||"","Nama Karyawan":s.employee_name||"","Cabang Asal":s.from_branch_name||"","Cabang Tujuan":s.to_branch_name||"",Status:s.status||"",Dokumen:s.document_link||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(K(),le));o(l,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(K(),le));r(e,"Template_Import_Mutasi")},onImport:async e=>{let r=s=>{if(!s)return null;let i=String(s||"").toLowerCase(),n=De.find(c=>String(c.label||"").toLowerCase()===i);return n?n.value:null},a=s=>{if(!s)return"";if(s instanceof Date&&!isNaN(s.getTime()))return s.toISOString().slice(0,10);let i=String(s).trim();if(/^\d{4,5}$/.test(i)){let c=Number(i);if(c>2e4&&c<99999){let p=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(i))return i.slice(0,10);let n=i.split(/[\/\-\.]/);if(n.length===3){let[c,p,u]=n.map(d=>d.trim());if(c.length===4&&p.length<=2&&u.length<=2)return`${c}-${p.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&p.length<=2&&c.length<=2)return`${u}-${p.padStart(2,"0")}-${c.padStart(2,"0")}`}return i},l=e.map(s=>({tanggal:a(s.Tanggal),employee_name:String(s["Nama Karyawan"]||"").trim(),from_branch_id:r(String(s["Cabang Asal"]||"").trim()),to_branch_id:r(String(s["Cabang Tujuan"]||"").trim()),status:String(s.Status||"").trim(),document_link:String(s.Dokumen||"").trim()})).filter(s=>s.tanggal&&s.employee_name&&s.from_branch_id&&s.to_branch_id),o=await k("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:ia},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:De,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:De,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}F();async function la(t){t.innerHTML=`
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
  `;let e=document.getElementById("btnRefreshSync"),r=document.getElementById("queueStatusFilter");e.addEventListener("click",l),r.addEventListener("change",i),document.getElementById("btnRetryAll").addEventListener("click",()=>a("retry",{allFailed:!0})),document.getElementById("btnResetStuck").addEventListener("click",()=>a("reset-stuck",{})),document.getElementById("btnPauseSync").addEventListener("click",()=>a("pause",{})),document.getElementById("btnResumeSync").addEventListener("click",()=>a("resume",{})),document.getElementById("btnReconcile").addEventListener("click",()=>a("reconcile",{module:"ALL",repairMode:!0})),document.getElementById("btnRetrySelected").addEventListener("click",()=>{let c=Array.from(document.querySelectorAll(".chk-queue:checked")).map(p=>p.value);if(c.length===0)return alert("No items selected");a("retry",{ids:c})}),document.getElementById("chkAllQueue").addEventListener("change",c=>{document.querySelectorAll(".chk-queue").forEach(p=>p.checked=c.target.checked)});async function a(c,p){if(confirm(`Are you sure you want to execute action: ${c}?`)){showLoading();try{let u=await k(`/api/sync/actions/${c}`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(p)});u.ok?(alert(u.data?.message||"Success"),l()):G(u.error||"Action failed")}catch(u){G(u.message)}hideLoading()}}await l();async function l(){showLoading(),await Promise.all([s(),i(),o(),n()]),hideLoading()}async function o(){try{let c=await k("/api/sync/performance");if(!c.ok)return;let{webhook:p,google_api:u,d1:d,queue:g,throughput:h}=c.data;document.getElementById("latencyHistograms").innerHTML=`
        <table class="w-full text-left">
          <thead><tr class="text-gray-500 border-b"><th>Metric</th><th>P50</th><th>P95</th><th>P99</th><th>Max</th></tr></thead>
          <tbody class="divide-y">
            <tr><td class="py-1">Webhook</td><td>${p.P50}ms</td><td>${p.P95}ms</td><td>${p.P99}ms</td><td>${p.Max}ms</td></tr>
            <tr><td class="py-1">Google API</td><td>${u.P50}ms</td><td>${u.P95}ms</td><td>${u.P99}ms</td><td>${u.Max}ms</td></tr>
            <tr><td class="py-1">D1 Execute</td><td>${d.P50}ms</td><td>${d.P95}ms</td><td>${d.P99}ms</td><td>${d.Max}ms</td></tr>
            <tr><td class="py-1">Queue Delay</td><td>${g.P50}ms</td><td>${g.P95}ms</td><td>${g.P99}ms</td><td>${g.Max}ms</td></tr>
          </tbody>
        </table>
        <div class="mt-4 pt-3 border-t text-gray-600 flex justify-between">
          <span>Throughput: <b>${h.events_per_sec}</b> ev/sec</span>
          <span><b>${h.events_per_min}</b> ev/min</span>
        </div>
      `,document.getElementById("lblSnapshotStatus").innerHTML='Checked 10 modules. Status: <b>COMPLETED</b><br><span class="text-xs text-gray-500">Run today at 02:00</span>'}catch(c){console.error(c)}}async function s(){try{let c=await k("/api/sync/health");if(!c.ok)return G("Failed to fetch sync health");let{status:p,queue:u,circuit_breaker:d}=c.data,g=`
        <div class="bg-white p-4 rounded-lg shadow border-l-4 ${p==="HEALTHY"?"border-green-500":p==="WARNING"?"border-yellow-500":"border-red-500"}">
          <p class="text-sm text-gray-500">System Health</p>
          <p class="text-2xl font-bold ${p==="HEALTHY"?"text-green-600":p==="WARNING"?"text-yellow-600":"text-red-600"}">${p}</p>
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
      `;document.getElementById("syncOverviewCards").innerHTML=g;let h=document.getElementById("cbStateBadge"),m=document.getElementById("cbStateDesc"),y=document.getElementById("cbStatusCard");y.className="bg-white rounded-lg shadow p-6 border-l-4",d==="CLOSED"?(h.className="px-3 py-1 rounded-full text-sm font-medium bg-green-100 text-green-800",h.textContent="CLOSED",m.textContent="Traffic is flowing normally to Google Sheets.",y.classList.add("border-green-500")):d==="OPEN"?(h.className="px-3 py-1 rounded-full text-sm font-medium bg-red-100 text-red-800",h.textContent="OPEN",m.textContent="Failing fast. Traffic to Google Sheets is paused due to repeated failures.",y.classList.add("border-red-500")):d==="HALF_OPEN"?(h.className="px-3 py-1 rounded-full text-sm font-medium bg-yellow-100 text-yellow-800",h.textContent="HALF-OPEN",m.textContent="Testing recovery. Permitting limited traffic to verify stability.",y.classList.add("border-yellow-500")):h.textContent=d||"UNKNOWN"}catch(c){console.error(c)}}async function i(){try{let c=document.getElementById("queueStatusFilter").value,p=await k("/api/sync/queue?limit=15"+(c?"&status="+c:""));if(!p.ok)return;let u=document.getElementById("queueTableBody"),d=p.data?.data||p.data||[];if(d.length===0){u.innerHTML='<tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">No events found</td></tr>';return}u.innerHTML=d.map(g=>`
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
      `).join("")}catch(c){console.error(c)}}async function n(){try{let c=await k("/api/sync/metrics");if(!c.ok)return;let p=document.getElementById("metricsTableBody"),u=c.data||[];if(u.length===0){p.innerHTML='<tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">No metrics found in last 24h</td></tr>';return}p.innerHTML=u.map(d=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2 font-medium text-gray-800">${d.module}</td>
          <td class="px-4 py-2 text-gray-600">${d.total_events}</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(d.avg_webhook_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(d.avg_d1_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(d.avg_queue_wait_ms||0)} ms</td>
        </tr>
      `).join("")}catch(c){console.error(c)}}}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let e=Math.floor(Number(t)-25569);return new Date(e*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let e=t.split(/[\/\-]/);return`${e[2]}-${e[1]}-${e[0]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let r=e.split("-");if(r.length===3&&r[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],l=parseInt(r[2],10),o=a[parseInt(r[1],10)-1];return`${l} ${o} ${r[0]}`}return e};function Q(t){return async e=>{if(!Be()){Se("/login");return}return t(e)}}var qe=null;function Oa(){qe&&clearInterval(qe);let t=()=>{let e=new Date,r=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),l=document.getElementById("header-clock-time"),o=document.getElementById("header-clock-date");l&&(l.textContent=r),o&&(o.textContent=a)};t(),qe=setInterval(t,1e3)}async function Ra(){try{let t=await k("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},r=(a,l)=>{let o=document.getElementById(a);o&&(o.textContent=l>0?l:"",o.style.display=l>0?"inline-flex":"none")};r("badge-issues",e.issues?.current||0),r("badge-contracts",e.expiring30?.current||0),r("badge-oo1",e.one_on_one?.current||0),r("badge-schedule",e.schedule?.current||0),r("badge-supply",e.supply?.current||0)}catch{}}var Ie=[];async function Ka(){try{let t=await k("/api/dashboard/notifications");if(!t.ok)return;Ie=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Ie.length>0?"block":"none",e.textContent=Ie.length)}catch{}}function Ha(){if(!Ie.length){se({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,r)=>r()});return}let t=`
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
  `;se({title:`Notifikasi (${Ie.length})`,content:t,confirmText:"Tutup",onConfirm:(e,r)=>r()})}function sa(){let t=he(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
  `;let r=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),l=document.getElementById("topbar-menu-btn"),o=document.getElementById("sidebar-close"),s=()=>{r.classList.add("open"),a.classList.add("show")},i=()=>{r.classList.remove("open"),a.classList.remove("show")};l?.addEventListener("click",s),o?.addEventListener("click",i),a?.addEventListener("click",i),document.querySelectorAll(".nav-item").forEach(c=>c.addEventListener("click",i));function n(){let c=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(d=>{let g=d.dataset.route;d.classList.toggle("active",c===g||g!=="/dashboard"&&c.startsWith(g))});let p=document.getElementById("topbar-title"),u=document.querySelector(".nav-item.active .nav-label");p&&u&&(p.textContent=u.textContent)}window.addEventListener("hashchange",n),n(),Oa(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await k("/api/auth/logout",{method:"POST"}),Le(),qe&&clearInterval(qe),Se("/login")}),Ra(),Ka(),document.getElementById("btn-notif")?.addEventListener("click",c=>{c.preventDefault(),Ha()})}async function ja(){J("/login",({main:e})=>Pt(e)),J("/dashboard",Q(({main:e})=>Tt(e))),J("/calendar",Q(({main:e})=>Vt(e))),J("/employees",Q(({main:e,params:r})=>Bt(e,r))),J("/contracts",Q(({main:e,params:r})=>Nt(e,r))),J("/sp",Q(({main:e})=>na(e))),J("/mutasi",Q(({main:e})=>ra(e))),J("/sync-dashboard",Q(({main:e})=>la(e))),J("/timeline",Q(({main:e,params:r})=>At(e,r))),J("/issues",Q(({main:e,params:r})=>Ft(e,r))),J("/one-on-one",Q(({main:e,params:r})=>Mt(e,r))),J("/training",Q(({main:e})=>Ot(e))),J("/relievers",Q(({main:e,params:r})=>Kt(e,r))),J("/reports/inspection",Q(({main:e})=>Ht(e))),J("/reports/cleaning",Q(({main:e})=>jt(e))),J("/reports/fogging",Q(({main:e})=>qt(e))),J("/reports/basecamp",Q(({main:e})=>Jt(e))),J("/reports/supply",Q(({main:e})=>ht(e,"supply"))),J("/sop",Q(({main:e})=>Ut(e))),J("/checklist",Q(({main:e})=>Gt(e))),J("/forms",Q(({main:e})=>ht(e))),J("/users",Q(({main:e})=>Qt(e))),J("/branches",Q(({main:e})=>zt(e))),J("/profile",Q(({main:e})=>Wt(e))),J("/settings/import",Q(({main:e})=>ta(e)));let t=Be();if(!t&&window.location.hash!=="#/login"&&Se("/login"),t){let e=await k("/api/auth/me");e.ok?(Ne(e.data.data),sa()):(Le(),Se("/login"))}window.addEventListener("fm:login",()=>{sa(),Se("/dashboard")}),ft()}ja();
