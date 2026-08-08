var oa=Object.defineProperty;var at=(t,e)=>()=>(t&&(e=t(t=0)),e);var nt=(t,e)=>{for(var s in e)oa(t,s,{get:e[s],enumerable:!0})};var ke={};nt(ke,{API:()=>yt,CLIENT_SIDE_MAX_ROWS:()=>be,IS_DEVELOPMENT:()=>Pe,apiFetch:()=>k,clearToken:()=>Le,getToken:()=>Be,getUser:()=>he,setToken:()=>it,setUser:()=>Ne});function Be(){return localStorage.getItem("fm_token")}function it(t){localStorage.setItem("fm_token",t)}function Le(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function he(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Ne(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function k(t,e={}){let s=Be(),a={"Content-Type":"application/json",...s?{Authorization:`Bearer ${s}`}:{},...e.headers||{}};try{let l=`cb=${Date.now()}`,d=t.includes("?")?"&":"?",i=`${yt}${t}${d}${l}`,r=await fetch(i,{...e,headers:a}),n;try{let o=await r.text();try{n=JSON.parse(o)}catch{n={error:`Server Error (${r.status}): ${o.substring(0,80)}...`}}}catch{n={error:"Gagal membaca respon dari server"}}return r.status===401&&(Le(),window.location.hash="#/login"),{ok:r.ok,status:r.status,data:n}}catch(l){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${l.message})`}}}}var Pe,da,yt,be,F=at(()=>{Pe=!1,da="https://fm-operations-api.facilitycare-audydental.workers.dev",yt=da,be=1e4});var kt={};nt(kt,{confirmDialog:()=>Fe,createModal:()=>le});function le({title:t,content:e,onConfirm:s,onCancel:a,confirmText:l="Simpan",cancelText:d="Batal",size:i="md",confirmClass:r="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},o=document.createElement("div");o.className="modal-overlay",o.innerHTML=`
    <div class="modal" style="max-width:${n[i]||n.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${d}</button>
        ${s?`<button class="btn ${r} modal-confirm">${l}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&o.querySelector(".modal-body").appendChild(e);let m=()=>{o.classList.remove("show"),setTimeout(()=>o.remove(),250)};return o.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),m()}),o.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),m()}),s&&o.querySelector(".modal-confirm").addEventListener("click",()=>s(o,m)),o.addEventListener("click",c=>{c.target===o&&(a&&a(),m())}),document.body.appendChild(o),requestAnimationFrame(()=>o.classList.add("show")),{overlay:o,close:m}}function Fe(t,e,s="Konfirmasi"){return le({title:s,content:`<p>${t}</p>`,onConfirm:(a,l)=>{e(),l()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var ye=at(()=>{});var se={};nt(se,{downloadExcel:()=>B,parseExcel:()=>Oe,renderExcelButtons:()=>Re});function Oe(t){return new Promise((e,s)=>{let a=new FileReader;a.onload=l=>{try{let d=new Uint8Array(l.target.result),i=XLSX.read(d,{type:"array"}),r=i.SheetNames[0],n=i.Sheets[r];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${i.SheetNames.join(", ")}`),console.log(`Sheet Used: ${r}`);let o=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),m=o.e.r-o.s.r+1,c=o.e.c-o.s.c+1;console.log(`Total Rows (including empty): ${m}`),console.log(`Total Columns: ${c}`);let p=[];for(let y=o.s.c;y<=o.e.c;++y){let u=n[XLSX.utils.encode_cell({c:y,r:o.s.r})];u&&u.v&&p.push(u.v)}console.log(`Headers Found: ${p.join(", ")}`),console.log("---------------------------");let g=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(g,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(g,"__headers",{value:p,enumerable:!1}),e(g)}catch(d){s(d)}},a.onerror=l=>s(l),a.readAsArrayBuffer(t)})}function B(t,e){try{let s=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,s,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(s){throw console.error("Error generating Excel file:",s),s}}function Re(t){return`
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
  `}var K=at(()=>{});F();var rt={},Je=null;function J(t,e){rt[t]=e}function Se(t){window.location.hash=t}function ft(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[s,...a]=e.split("?"),l=rt[s];if(!l){for(let[i,r]of Object.entries(rt))if(i.endsWith("/*")&&s.startsWith(i.slice(0,-2))){l=r;break}}Je&&(Je(),Je=null);let d=document.getElementById("main-content");if(d&&(d.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),l){let i=new URLSearchParams(a.join("?")),r=s.split("/").filter(Boolean),n=await l({path:s,params:i,segments:r,main:d});n&&(Je=n)}else{let i=d||document.getElementById("app");i&&(i.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var Ae;function ca(){return Ae||(Ae=document.createElement("div"),Ae.id="toast-container",document.body.appendChild(Ae)),Ae}function vt(t,e="info",s=3500){let a=ca(),l=document.createElement("div");l.className=`toast toast-${e}`;let d={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};l.innerHTML=`<span class="toast-icon">${d[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(l),requestAnimationFrame(()=>l.classList.add("show")),setTimeout(()=>{l.classList.remove("show"),setTimeout(()=>l.remove(),350)},s)}var Z=t=>vt(t,"success"),G=t=>vt(t,"error");ye();F();F();ye();function Ue({columns:t,data:e,onEdit:s,onDelete:a,onView:l,actions:d=[],emptyText:i="Tidak ada data",bulkSelect:r=null}){let n=document.createElement("div");if(n.className="table-wrapper",!e||e.length===0)return n.innerHTML=`<div class="empty-state"><p>${i}</p></div>`,n;let o=document.createElement("table");o.className="data-table";let m=document.createElement("thead"),c=document.createElement("tr");if(r){let g=document.createElement("th");g.style.width="40px",g.style.textAlign="center";let y=document.createElement("input");y.type="checkbox",y.id="select-all-checkbox",y.title="Pilih semua",y.addEventListener("change",()=>{e.forEach(u=>{y.checked?r.selectedIds.add(u.id):r.selectedIds.delete(u.id)}),n.querySelectorAll(".row-checkbox").forEach(u=>u.checked=y.checked),r.onToggle()}),g.appendChild(y),c.appendChild(g)}if(t.forEach(g=>{let y=document.createElement("th");y.textContent=g.label,g.width&&(y.style.width=g.width),c.appendChild(y)}),s||a||l||d.length>0){let g=document.createElement("th");g.textContent="Aksi",g.style.width="120px",c.appendChild(g)}m.appendChild(c),o.appendChild(m);let p=document.createElement("tbody");return e.forEach(g=>{let y=document.createElement("tr");if(r){let u=document.createElement("td");u.style.textAlign="center",u.style.width="40px";let b=document.createElement("input");b.type="checkbox",b.className="row-checkbox",b.checked=r.selectedIds.has(g.id),b.addEventListener("change",()=>{if(b.checked)r.selectedIds.add(g.id);else{r.selectedIds.delete(g.id);let h=document.getElementById("select-all-checkbox");h&&(h.checked=!1)}r.onToggle()}),u.appendChild(b),y.appendChild(u)}if(t.forEach(u=>{let b=document.createElement("td");if(u.render){let h=u.render(g[u.key],g);h instanceof HTMLElement?b.appendChild(h):b.innerHTML=h||""}else b.textContent=g[u.key]!==null&&g[u.key]!==void 0&&g[u.key]!==""?g[u.key]:"";u.nowrap&&(b.style.whiteSpace="nowrap"),y.appendChild(b)}),s||a||l||d.length>0){let u=document.createElement("td");u.className="actions-cell";let b=document.createElement("div");if(b.className="btn-group",l){let h=document.createElement("button");h.className="btn btn-xs btn-ghost",h.innerHTML="\u{1F441}",h.title="Lihat",h.addEventListener("click",()=>l(g)),b.appendChild(h)}if(s){let h=document.createElement("button");h.className="btn btn-xs btn-secondary",h.innerHTML="\u270F\uFE0F",h.title="Edit",h.addEventListener("click",()=>s(g)),b.appendChild(h)}d.forEach(h=>{let _=document.createElement("button");_.className=`btn btn-xs ${h.class||"btn-ghost"}`,_.innerHTML=h.icon||h.label,_.title=h.label,_.addEventListener("click",()=>h.handler(g)),b.appendChild(_)}),u.appendChild(b),y.appendChild(u)}p.appendChild(y)}),o.appendChild(p),n.appendChild(o),n}function Ge({page:t,pages:e,total:s,limit:a,onPage:l}){if(e<=1)return null;let d=document.createElement("div");d.className="pagination";let i=document.createElement("span");i.className="pagination-info",i.textContent=`Total: ${s} data`,d.appendChild(i);let r=document.createElement("div");r.className="pagination-btns";let n=(c,p,g=!1,y=!1)=>{let u=document.createElement("button");u.className=`btn btn-sm ${y?"btn-primary":"btn-ghost"} pagination-btn`,u.textContent=c,u.disabled=g,u.addEventListener("click",()=>l(p)),r.appendChild(u)};n("\xAB",1,t===1),n("\u2039",t-1,t===1);let o=Math.max(1,t-2),m=Math.min(e,t+2);for(let c=o;c<=m;c++)n(c,c,!1,c===t);return n("\u203A",t+1,t===e),n("\xBB",e,t===e),d.appendChild(r),d}ye();function Me(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${Me(e.fields)}</div>`;let s=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",l="";switch(e.type){case"textarea":l=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${s} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let i=(e.options||[]).map(c=>{let p=typeof c=="object"?c.value:c,g=typeof c=="object"?c.label:c,y=e.value==p?"selected":"";return`<option value="${p}" ${y}>${g}</option>`}).join("");l=`<select name="${e.name}" class="form-control" ${s}><option value="">-- Pilih ${e.label||""} --</option>${i}</select>`;break;case"combobox":let r=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,n=(e.options||[]).map(c=>{let p=typeof c=="object"?c.value:c,g=typeof c=="object"?c.label||c.value||"":c||"";return(g==="undefined"||g==="[object Object]"||g==="null")&&(g=""),g?`<option value="${g}"></option>`:""}).join(""),o=e.value||"";if(e.value){let c=(e.options||[]).find(p=>(typeof p=="object"?p.value:p)==e.value);if(c){let p=typeof c=="object"?c.label||c.value||"":c||"";p&&p!=="undefined"&&p!=="[object Object]"&&p!=="null"&&(o=p)}}l=`
          <input type="text" name="${e.name}" list="${r}" class="form-control" value="${o}" placeholder="Pilih atau ketik baru..." ${s} autocomplete="off">
          <datalist id="${r}">${n}</datalist>
        `;break;case"checkbox":l=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let m=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";l=`<input type="date" name="${e.name}" class="form-control" value="${m}" ${s}>`;break;case"number":l=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${s}>`;break;case"email":l=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${s}>`;break;case"url":l=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${s}>`;break;default:l=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${s} autocomplete="off">`}let d=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${l}${d}</div>`}).join("")}function Qe(t){let e={},s=new FormData(t);for(let[a,l]of s.entries())e[a]=l===""?null:l;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function ze(t,e){e&&Object.entries(e).forEach(([s,a])=>{let l=t.querySelector(`[name="${s}"]`);l&&(l.hasAttribute("list")||(l.type==="checkbox"?l.checked=!!a:l.type==="date"&&a&&window.parseFlexibleDate?l.value=window.parseFlexibleDate(a):l.value=a??""))})}K();function L({container:t,title:e,icon:s,apiPath:a,columns:l,formFields:d,filterFields:i,defaultFilters:r={},itemLabel:n="Data",canCreate:o=!0,canEdit:m=!0,canDelete:c=!0,onBeforeSubmit:p,onAfterLoad:g,onDataLoaded:y,extraActions:u=[],initialSearch:b="",exportOptions:h=null,bulkDelete:_=!1,paginationMode:C="server"}){let S=1,x={...r};b&&(x.search=b);let w=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${s} ${e}</h1>
      <div class="page-actions">
        ${o?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
      </div>
    </div>

    ${_?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${h?Re(h.moduleName):""}

    ${i&&i.length>0?`
    <div class="filter-bar card" style="padding: 1rem;">
        ${i.map(v=>{if(v.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}"></div>`;if(v.type==="search-combo"){let f="dl-filter-search",$=(v.options||[]).map(P=>`<option value="${typeof P=="object"?P.label:P}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${f}" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}"><datalist id="${f}">${$}</datalist></div>`}return v.type==="select"?`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${v.name}" id="filter-${v.name}"><option value="">-- ${v.label} --</option>${(v.options||[]).map(f=>`<option value="${typeof f=="object"?f.value:f}" ${x[v.name]===(typeof f=="object"?f.value:f)?"selected":""}>${typeof f=="object"?f.label:f}</option>`).join("")}</select>`:""}).join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function R(){if(!document.getElementById("bulk-toolbar"))return;let f=document.getElementById("bulk-count"),$=document.getElementById("btn-bulk-delete"),P=document.getElementById("btn-bulk-cancel");f.textContent=`${w.size} item dipilih`,w.size>0?($.disabled=!1,P.disabled=!1):($.disabled=!0,P.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{w.clear(),document.querySelectorAll(".row-checkbox").forEach(f=>f.checked=!1);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1),R()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(w.size===0)return;let v=[...w],f=document.createElement("div");f.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",f.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${v.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${v.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(f),f.querySelector("#bulk-cancel-btn").addEventListener("click",()=>f.remove()),f.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let $=f.querySelector("#bulk-confirm-btn");$.disabled=!0,$.textContent="Menghapus...";let P=await k(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:v})});f.remove(),P.ok?(Z(`${v.length} ${n} berhasil dihapus.`),w.clear(),R(),M()):G(P.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),O;if(E?.addEventListener("input",v=>{clearTimeout(O),O=setTimeout(()=>{x.search=v.target.value,S=1,w.clear(),M()},400)}),i?.forEach(v=>{v.type==="select"&&document.getElementById(`filter-${v.name}`)?.addEventListener("change",f=>{x[v.name]=f.target.value,S=1,w.clear(),M()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{x={...r},E&&(E.value=""),i?.forEach(v=>{let f=document.getElementById(`filter-${v.name}`);f&&(f.value="")}),S=1,w.clear(),M()}),document.getElementById("btn-create")?.addEventListener("click",()=>ge(null)),h){document.getElementById(`btn-export-${h.moduleName}`)?.addEventListener("click",async f=>{let $=f.target,P=$.innerHTML;$.innerHTML="\u23F3 Loading...",$.disabled=!0;try{await h.onExport()}catch{G("Gagal export data")}finally{$.innerHTML=P,$.disabled=!1}}),document.getElementById(`btn-template-${h.moduleName}`)?.addEventListener("click",()=>{h.onTemplate()});let v=document.getElementById(`input-import-${h.moduleName}`);v?.addEventListener("change",async f=>{let $=f.target.files[0];if(!$)return;v.disabled=!0;let P=document.createElement("div");P.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",P.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(P);let V=P.querySelector("#import-progress-text"),U=P.querySelector("#import-progress-bar"),D=P.querySelector("#import-summary"),I=P.querySelector("#import-close-btn");I.addEventListener("click",()=>{P.remove(),M()});try{let Y=await Oe($);if(Y.length===0)throw new Error("File kosong atau format salah");let X=500,ie=0,ae=0,N=0,A=Y.length;V.textContent=`Ditemukan ${A} baris data. Memulai import...`;for(let T=0;T<A;T+=X){let z=Y.slice(T,T+X);V.textContent=`Mengimport baris ${T+1} - ${Math.min(T+X,A)} dari ${A}...`,U.style.width=`${Math.round(T/A*100)}%`;try{let j=await h.onImport(z);j?(ie+=j.inserted||j.metrics?.inserted||z.length,ae+=j.skipped||j.metrics?.updated||0):ie+=z.length}catch(j){console.error("Chunk import failed:",j),N+=z.length}}U.style.width="100%",V.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',D.style.display="block",D.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${A}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ie}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${N}</strong></div>
        `,N>0&&(D.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),I.style.display="block",v.value=""}catch(Y){V.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${Y.message}`,U.style.background="var(--danger)",U.style.width="100%",I.style.display="block",v.value=""}finally{v.disabled=!1}})}async function M(){R();let v=document.getElementById("table-container");if(!v)return;v.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let f=C==="client",$=f?1:S,P=f?be:20,V=new URLSearchParams({page:$,limit:P,...Object.fromEntries(Object.entries(x).filter(([,N])=>N))}),U=await k(`${a}?${V}`);if(!U.ok){v.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${U.data?.error||"Error"}</p></div>`;return}let D=U.data?.data||U.data||[],I=U.data?.pagination,Y=D.length,X=D;if(f){D=y(D),X=D;let N=D.length,A=20,T=Math.ceil(N/A);S>T&&T>0&&(S=T);let z=(S-1)*A,j=S*A;D=D.slice(z,j),I={page:S,limit:A,total:N,pages:T}}!1,g&&g(D);let ie=Ue({columns:l,data:D,fullData:X,onEdit:m?N=>ge(N):null,actions:u.map(N=>({...N,handler:A=>N.handler(A,M)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:_?{selectedIds:w,onToggle:R}:null});v.innerHTML="",v.appendChild(ie);let ae=document.getElementById("pagination-container");if(ae&&(ae.innerHTML="",I&&I.pages>1)){let N=Ge({page:I.page,pages:I.pages,total:I.total,limit:I.limit,onPage:A=>{S=A,M()}});N&&ae.appendChild(N)}}function fe(v){let f=typeof d=="function"?d(v):d;return Me(f)}function ge(v){let f=!!v,$=document.createElement("form");if($.noValidate=!0,$.innerHTML=fe(v),f){let V=typeof d=="function"?d(v):d;ze($,v)}let{close:P}=le({title:f?`Edit ${n}`:`Tambah ${n}`,content:$,size:"lg",confirmText:f?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(V,U)=>{if(!$.reportValidity())return;let D=V.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let I=Qe($),Y=typeof d=="function"?d(v):d,X=async A=>{for(let T of A)if(T.type==="row")await X(T.fields);else if(T.type==="combobox"&&I[T.name]){let z=I[T.name],j=(T.options||[]).find(W=>{let ne=String(typeof W=="object"?W.value:W),tt=String(typeof W=="object"?W.label:W);return ne===z||tt===z});if(j)I[T.name]=typeof j=="object"?j.value:j;else if(T.createApi){let W={};W[T.createApi.field]=z,T.createApi.extra&&Object.assign(W,T.createApi.extra);let ne=await k(T.createApi.path,{method:"POST",body:JSON.stringify(W)});if(ne.ok&&ne.data?.id)I[T.name]=ne.data.id;else if(ne.ok&&!ne.data?.id)I[T.name]=z;else throw new Error(`Gagal membuat master data: ${ne.data?.error||"Unknown error"}`)}}};try{await X(Y)}catch(A){G(A.message),D.disabled=!1,D.textContent=f?"Simpan Perubahan":`Tambah ${n}`;return}p&&(I=await p(I,v));let ie=f?"PUT":"POST",ae=f?`${a}/${v.id}`:a,N=await k(ae,{method:ie,body:JSON.stringify(I)});N.ok?(Z(f?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),U(),M()):(G(N.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=f?"Simpan Perubahan":`Tambah ${n}`)}})}function ve(v){Fe(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let f=await k(`${a}/${v.id}`,{method:"DELETE"});f.ok?(Z(`${n} berhasil dihapus.`),M()):G(f.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return M(),M}F();F();var xe=null,Ve=null;async function we(t=!1){if(xe&&!t)return console.log("Employees Raw (Cache Hit)",xe.slice(0,5)),xe;let e=await k(`/api/employees?limit=${be}&status=Aktif`);return xe=(e.data?.data||[]).map(s=>({value:s.id,label:s.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",xe.slice(0,5)),xe}async function ee(t=!1){let s=(await we(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",s.slice(0,5)),s}async function H(t=!1){return Ve&&!t||(Ve=((await k("/api/branches?all=1")).data?.data||[]).map(s=>({value:s.id,label:s.full_name}))),Ve}function q(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function st(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function _e(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function lt(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function oe(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}K();function ot(t,e){if(t.period!=="Q3")return!1;let s=String(t.status||"").toLowerCase();if(s!=="selesai"&&s!=="completed"&&s!=="done")return!1;let a=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?a.includes("inspeksi"):e==="gcdc"?a.includes("general cleaning")||a.includes("deep cleaning"):!1}F();K();function St(t,e){let s=String(t.status||"").toLowerCase();return e==="active"?s==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&s==="aktif":!1}F();K();function dt(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let l=new Date(a);l.setDate(a.getDate()+30);let d=new Date(t.end_date);return d.setHours(0,0,0,0),d>=a&&d<=l}return!1}F();K();function xt(t,e){let s=String(t.status||"").toLowerCase();return e==="open"?s==="open":!1}F();function wt(t,e){let s=String(t.status||"").toLowerCase();return e==="pending"?s==="pending":!1}var ue={};function Ee(t){if(ue[t]){try{ue[t].destroy()}catch{}delete ue[t]}}function pa(){Object.keys(ue).forEach(Ee)}var pe=(t,e=0)=>{let s=Number(t);return isNaN(s)||t===null||t===void 0?e:s},Ce=(t,e="\u2014")=>{if(t==null||t==="")return e;let s=String(t).trim();return s===""||s==="[object Object]"?e:s};function Ct(t,e,s=900){if(!t)return;let a=Math.max(0,Math.round(pe(e)));if(a===0){t.textContent="0";return}let l=Date.now(),d=()=>{let i=Math.min((Date.now()-l)/s,1),r=1-Math.pow(1-i,3);t.textContent=Math.round(r*a).toLocaleString("id-ID"),i<1?requestAnimationFrame(d):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(d)}var ma={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},ua=t=>{let e=Ce(t,"\u2014");return`<span class="status-pill ${ma[e]||"pill-neutral"}">${e}</span>`};var de={family:"Inter",size:11},me="#94A3B8",Te="#F1F5F9",ct=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],ga=()=>window.innerWidth<768;function We(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:ga()?"bottom":"top",labels:{font:de,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:de,titleFont:{...de,weight:"700"}}},scales:{x:{grid:{color:Te},ticks:{font:de,color:me,maxRotation:0}},y:{grid:{color:Te},ticks:{font:de,color:me},beginAtZero:!0}},...t}}var ba=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join("");function _t(t=3){return Array(t).fill(0).map((e,s)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${s<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function re(t,e,s=8e3){try{let a=new AbortController,l=setTimeout(()=>a.abort(),s),d=await k(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(l),!d||!d.ok)return e;let i=d.data;return i?i.data!==void 0?i.data??e:i:e}catch{return e}}function ya(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(a=>{let l=document.getElementById(a);l&&(l.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(a=>{let l=document.getElementById(a);if(l&&l.style.display==="none"){l.style.display="block";let d=l.parentElement;if(d&&!d.querySelector(".chart-empty")){let i=document.createElement("div");i.className="chart-empty",i.textContent="Belum ada data",l.style.display="none",d.appendChild(i)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Et({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&$t({}),["table-contracts","table-issues"].forEach(a=>{let l=document.getElementById(a);l&&l.querySelector(".skeleton")&&(l.innerHTML='<div class="chart-empty">Belum ada data</div>')});let s=document.getElementById("activity-log");s&&s.querySelector(".skeleton")&&(s.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Tt(t){pa(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
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
              <div class="chart-canvas-wrap" style="width:130px;height:130px;position:relative">
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>pt(t)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async s=>{let a=s.target.value;document.getElementById("jadwal-year-label").textContent=a;let l=document.getElementById("skel-jadwal"),d=document.getElementById("chart-jadwal");l&&(l.style.display="block",l.style.position="absolute"),d&&(d.style.display="none");let i=await re(`/api/dashboard/schedule-chart?year=${a}`,{},8e3);try{Dt(i)}catch(r){console.warn("ScheduleChart render:",r),ce("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async s=>{let a=s.target.value,l=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",d=document.getElementById("skel-insp"),i=document.getElementById("chart-insp");d&&(d.style.display="block",d.style.position="absolute"),i&&(i.style.display="none");let r=await re(l,{},8e3);try{It(r)}catch(n){console.warn("InspBar render:",n),ce("skel-insp","chart-insp")}}),t._skelTimeout=setTimeout(()=>ya(),5e3),await pt(t),t._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?pt(t):clearInterval(t._dashRefresh)},6e4)}async function pt(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,s,a,l,d,i,r,n,o,m,c,p]=await Promise.all([re("/api/dashboard/kpi",{},8e3),re("/api/dashboard/issues-trend",{},8e3),re("/api/dashboard/issues-summary",{},8e3),re("/api/dashboard/stats",{},8e3),re("/api/dashboard/calendar",[],8e3),re("/api/schedule?limit=10000",{data:[]},8e3),re("/api/employees?limit=10000",{data:[]},8e3),re("/api/contracts?limit=10000",{data:[]},8e3),re("/api/issues?limit=10000",{data:[]},8e3),re("/api/one-on-one?limit=10000",{data:[]},8e3),re("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),re(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3)]),g=document.getElementById("filter-insp-month"),y=g?g.value:"",u=y?`/api/dashboard/inspection-bar?month=${y}`:"/api/dashboard/inspection-bar",b=await re(u,{},8e3);if(e){let h=Array.isArray(i?.data)?i.data:Array.isArray(i)?i:[],_=Array.isArray(r?.data)?r.data:Array.isArray(r)?r:[],C=Array.isArray(n?.data)?n.data:Array.isArray(n)?n:[],S=Array.isArray(o?.data)?o.data:Array.isArray(o)?o:[],x=Array.isArray(m?.data)?m.data:Array.isArray(m)?m:[];e.employees&&(e.employees.current=_.filter(w=>St(w,"active")).length),e.contracts&&(e.contracts.current=C.filter(w=>dt(w,"active")).length),e.expiring30&&(e.expiring30={current:C.filter(w=>dt(w,"expiring30")).length}),e.issues&&(e.issues.current=S.filter(w=>xt(w,"open")).length),e.one_on_one&&(e.one_on_one.current=x.filter(w=>wt(w,"pending")).length),e.inspection_month&&(e.inspection_month.current=h.filter(w=>ot(w,"inspeksi")).length),e.cleaning_month&&(e.cleaning_month.current=h.filter(w=>ot(w,"gcdc")).length)}try{Et(e)}catch(h){console.warn("KPI render:",h)}try{$t(e)}catch(h){console.warn("MiniStats render:",h)}try{Dt(p)}catch(h){console.warn("ScheduleChart render:",h),ce("skel-jadwal","chart-jadwal")}try{fa(Array.isArray(a?.by_category)?a.by_category:[])}catch(h){console.warn("Donut render:",h),ce("skel-donut","chart-donut")}try{va(s)}catch(h){console.warn("Trend render:",h),ce("skel-trend","chart-trend")}try{It(b)}catch(h){console.warn("InspBar render:",h),ce("skel-insp","chart-insp")}try{let h=Array.isArray(l)?l:Array.isArray(l?.recent_issues)?l.recent_issues:[];Sa(h)}catch(h){console.warn("IssuesTable render:",h)}try{let h=Array.isArray(l?.expiring_contracts)?l.expiring_contracts:[];ka(c)}catch(h){console.warn("ContractsTable render:",h)}try{xa(Array.isArray(d)?d:[])}catch(h){console.warn("Agenda render:",h)}try{wa()}catch(h){console.warn("Quick Actions render:",h)}}function Et(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let s=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=s.map(a=>{let l=pe(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${l}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${l}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{Ct(a,parseInt(a.dataset.target)||0)})}function $t(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let s=[{icon:"\u{1F4C5}",label:"Jadwal",val:t.schedule?.current,href:"#/timeline",color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F504}",label:"Reliefer",val:t.reliever_completed?.current,href:"#/relievers?dash_filter=reliever",color:"mini-teal"},{icon:"\u{1F50D}",label:"Inspeksi",val:t.inspection_month?.current,href:"#/timeline?dash_filter=inspeksi",color:"mini-blue"},{icon:"\u{1F9F9}",label:"GCDC",val:t.cleaning_month?.current,href:"#/timeline?dash_filter=gcdc",color:"mini-green"},{icon:"\u{1F4A8}",label:"Fogging",val:t.fogging_month?.current,href:"#/reports/fogging",color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=s.map(a=>`
    <a href="${a.href}" class="mini-stat ${a.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${a.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${pe(a.val)}">0</div>
        <div class="mini-stat-text">${a.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(a=>Ct(a,parseInt(a.dataset.target)||0,700))}function fa(t){ce("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),s=document.getElementById("donut-legend");if(!e||!s)return;Ee("donut");let a=(t||[]).filter(n=>pe(n.count)>0);if(!a.length){Ke(e,"Belum ada data permasalahan");return}let l=a.map(n=>`${Ce(n.category,"Lainnya")}`),d=a.map(n=>pe(n.count)),i=d.reduce((n,o)=>n+o,0);s.innerHTML=a.map((n,o)=>{let m=ct[o%ct.length],c=i>0?Math.round(n.count/i*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${m}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${c}%)</span></div>
          <div class="donut-legend-label">${l[o]}</div>
        </div>
      </div>
    `}).join("");let r={id:"centerText",beforeDraw:function(n){let o=n.width,m=n.height,c=n.ctx;c.restore();let p=(m/80).toFixed(2);c.font="bold "+p+"em Inter",c.textBaseline="middle",c.fillStyle="#1E293B";let g=i.toString(),y=Math.round((o-c.measureText(g).width)/2),u=m/2;c.fillText(g,y,u-8),c.font="600 "+(p*.35).toFixed(2)+"em Inter",c.fillStyle="#64748B";let b="Total",h=Math.round((o-c.measureText(b).width)/2);c.fillText(b,h,u+14),c.save()}};ue.donut=new Chart(e,{type:"doughnut",data:{labels:l,datasets:[{data:d,backgroundColor:ct,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:de,titleFont:{...de,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutout:"75%"},plugins:[r]})}function va(t){ce("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;Ee("trend"),t=t||{};let s=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(t.labels||[]).map(i=>{if(!i||typeof i!="string")return"";try{let[r,n]=i.split("-");return(s[Number(n)-1]||n)+" "+String(r).slice(-2)}catch{return i}}),l=(t.open||[]).map(i=>pe(i)),d=(t.closed||[]).map(i=>pe(i));if(!a.length){Ke(e,"Belum ada data trend");return}ue.trend=new Chart(e,{type:"line",data:{labels:a,datasets:[{label:"Open",data:l,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:d,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:We({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:me,maxRotation:0,autoSkip:!1}},y:{grid:{color:Te},ticks:{font:{family:"Inter",size:9},color:me},beginAtZero:!0}}})})}function Dt(t){ce("skel-jadwal","chart-jadwal");let e=document.getElementById("chart-jadwal");if(!e)return;Ee("jadwal"),t=t||{};let s=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(t).some(n=>Array.isArray(n)&&n.some(o=>o>0))){Ke(e,"Belum ada data jadwal");return}let l=t["Inspeksi Hygiene"]||Array(12).fill(0),d=t["General Cleaning"]||Array(12).fill(0),i=t["Deep Cleaning"]||Array(12).fill(0),r=t.Fogging||Array(12).fill(0);ue.jadwal=new Chart(e,{type:"bar",data:{labels:s,datasets:[{label:"Inspeksi",data:l,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:d,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:i,backgroundColor:"#F59E0B"},{label:"Fogging",data:r,backgroundColor:"#EF4444"}]},options:We({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.7,categoryPercentage:.8}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:me,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Te},ticks:{font:{family:"Inter",size:9},color:me},min:0}}})})}function It(t){ce("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;Ee("inspBar"),t=t||{};let s=t.labels||[],a=(t.fc||[]).map(d=>pe(d)),l=(t.spv||[]).map(d=>pe(d));if(!s.length){Ke(e,"Belum ada data inspeksi");return}ue.inspBar=new Chart(e,{type:"bar",data:{labels:s,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:l,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:We({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:de,color:me,maxRotation:45,minRotation:30}},y:{grid:{color:Te},ticks:{font:de,color:me},min:0,max:100}}})})}function ka(t){ce("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;Ee("contractMiniBar"),t=t||{};let s={"06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(i=>{let r=i.split("-")[1];return s[r]||i}),l=(t.data||[]).map(i=>pe(i));if(!a.length){Ke(e,"Belum ada data");return}let d=e.getContext("2d");ue.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:l,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:We({onClick:(i,r)=>{if(r&&r.length>0){let n=r[0].index,o=(t.labels||[])[n];o&&(window.location.hash="#/contracts?month_expiry="+o)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:de,color:me,maxRotation:0,autoSkip:!1}},y:{grid:{color:Te,borderDash:[4,4],drawBorder:!1},ticks:{font:de,color:me,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function Sa(t){let e=document.getElementById("table-issues");if(!e)return;let s=(t||[]).slice(0,8);if(!s.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${s.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${ua(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Ce(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Ce(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function xa(t){let e=document.getElementById("widget-agenda");if(!e)return;let s=new Date,a=`${s.getFullYear()}-${String(s.getMonth()+1).padStart(2,"0")}-${String(s.getDate()).padStart(2,"0")}`,d=(t||[]).filter(i=>(i.event_date||"").startsWith(a)).slice(0,10);if(!d.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${d.map(i=>{let r="#3B82F6",n="#EFF6FF",o="Agenda",m=(i.title||"").toLowerCase();return m.includes("inspeksi")?(r="#10B981",n="#ECFDF5",o="Inspeksi"):m.includes("cleaning")||m.includes("gcdc")?(r="#3B82F6",n="#EFF6FF",o="Cleaning"):m.includes("reliefer")?(r="#F59E0B",n="#FFFBEB",o="Reliefer"):m.includes("fogging")&&(r="#8B5CF6",n="#F5F3FF",o="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(i.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${r};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Ce(i.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Ce(i.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${r}">${o}</div>
        </div>
      `}).join("")}
    </div>
  `}function wa(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(s=>`
    <a href="${s.href}" class="action-btn">
      <div class="action-icon" style="background:${s.bg}">${s.icon}</div>
      ${s.label}
    </a>
  `).join("")}function ce(t,e){let s=document.getElementById(t),a=document.getElementById(e);if(s&&(s.style.display="none",s.style.position=""),a){a.style.display="block";let l=a.parentElement;if(l){let d=l.querySelector(".chart-empty");d&&d.remove()}}}function Ke(t,e="Belum ada data"){if(!t)return;t.style.display="none";let s=t.parentElement;if(!s)return;if(!s.querySelector(".chart-empty")){let l=document.createElement("div");l.className="chart-empty",l.textContent=e,s.appendChild(l)}}F();async function Pt(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),s=document.getElementById("login-error"),a=document.getElementById("login-btn"),l=document.getElementById("toggle-password"),d=document.getElementById("login-password");l?.addEventListener("click",()=>{let i=d.type==="text";d.type=i?"password":"text",l.style.color=i?"":"var(--primary)"}),e?.addEventListener("submit",async i=>{i.preventDefault(),s.style.display="none";let r=e.username.value.trim(),n=e.password.value;if(!r||!n){s.textContent="Username dan password wajib diisi.",s.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let o=await k("/api/auth/login",{method:"POST",body:JSON.stringify({username:r,password:n})});o.ok&&o.data.success?(it(o.data.data.token),Ne(o.data.data.user),Z("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(s.textContent=o.data.error||"Username atau password salah.",s.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{s.textContent="Gagal terhubung ke server. Periksa koneksi internet.",s.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}F();K();async function _a(){return await H()}function Ca(t,e){let s=String(t.status||"").toLowerCase();return e==="active"?s==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&s==="aktif":!1}async function Bt(t,e){let s=await _a(),a=e?e.get("dash_filter"):null;L({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:l=>a?l.filter(d=>Ca(d,a)):l,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:l=>_e(l)},{key:"phone",label:"No. HP",render:l=>l?`<a href="tel:${l}">${l}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>q(l)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:s},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:l=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:l?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:l?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:s,value:l?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:l?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:l?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:l?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let l=await k(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(l.ok){let d=l.data.data.map(i=>({"Nama Lengkap":i.full_name,Cabang:i.branch_name||"",Divisi:i.division||"","No. HP":i.phone||"","Tgl Masuk":i.join_date||"",Status:i.status||""}));B(d,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async l=>{let d=n=>{if(!n)return null;let o=String(n||"").toLowerCase(),m=s.find(c=>String(c.label||"").toLowerCase()===o);return m?m.value:null},i=l.map(n=>({full_name:String(n["Nama Lengkap"]||"").trim(),branch_id:d(String(n.Cabang||"").trim()),division:String(n.Divisi||"").trim()||"FACILITY CARE",phone:String(n["No. HP"]||"").trim(),join_date:String(n["Tgl Masuk"]||"").trim(),status:String(n.Status||"").trim(),notes:String(n.Catatan||"").trim()})).filter(n=>n.full_name),r=await k("/api/import/employees",{method:"POST",body:JSON.stringify({rows:i,onDuplicate:"update"})});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}F();K();var ut=[],Lt=[];async function Ta(){ut=await H(),Lt=await we()}var mt=async t=>{let e=[],s=1;for(;;){let l=await(await Promise.resolve().then(()=>(F(),ke))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${s}`);if(!l.ok)break;let d=l.data?.data||l.data||[],i=Array.isArray(d)?d:[];if(e=e.concat(i),i.length<100||l.data?.pagination&&s>=l.data.pagination.pages)break;s++}return e};function Ea(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let l=new Date(a);l.setDate(a.getDate()+30);let d=new Date(t.end_date);return d.setHours(0,0,0,0),d>=a&&d<=l}return!1}async function Nt(t,e){await Ta();let s=e?e.get("dash_filter"):null;L({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",onDataLoaded:a=>s?a.filter(l=>Ea(l,s)):a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>_e(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,l)=>l.end_date&&String(l.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':st(a)},{key:"status",label:"Status",render:a=>q(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:ut},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"month_expiry",label:"Bulan Habis",options:[{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"},{value:"2027-01",label:"Jan 2027"},{value:"2027-02",label:"Feb 2027"},{value:"2027-03",label:"Mar 2027"},{value:"2027-04",label:"Apr 2027"},{value:"2027-05",label:"Mei 2027"},{value:"2027-06",label:"Jun 2027"},{value:"2027-07",label:"Jul 2027"}]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[d,i]=await Promise.all([mt("/api/employees?status=Aktif"),mt("/api/contracts")]);if(d.length>0){let r=i.filter(c=>c.status==="Aktif"),n=new Set(r.map(c=>c.employee_id)),o=d.filter(c=>!n.has(c.id)),m=`<p style="margin-bottom:12px">Data yang terbaca: <b>${d.length}</b> Karyawan Aktif, dan <b>${r.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${o.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;o.forEach(c=>{let p=i.filter(y=>y.employee_id===c.id),g='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(p.length>0){let y=p[0];g=`Pernah ada kontrak (Status: <b style="color:#EF4444">${y.status}</b>, Selesai: ${window.formatDate(y.end_date)})`}m+=`<li style="margin-bottom:8px"><b>${c.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${c.branch_name||"-"} | ${g}</span></li>`}),m+="</ul>",Promise.resolve().then(()=>(ye(),kt)).then(c=>c.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:m,cancelText:"Tutup"}))}}catch(d){console.error(d)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let l=document.querySelector(".page-actions");l&&l.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:Lt,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:ut,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await k(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let l=a.data.data.map(d=>({"Nama Lengkap":d.employee_name,Cabang:d.branch_name||"","Div / Bagian":d.division||"","Tanggal Mulai":d.start_date||"","Tanggal Selesai":d.end_date&&String(d.end_date).startsWith("2099")?"":d.end_date||"","Sisa Kontrak":d.end_date&&String(d.end_date).startsWith("2099")?"Tetap":d.days_remaining!==null&&d.days_remaining!==void 0?`${d.days_remaining} Hari`:"",Status:d.status||""}));B(l,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[l,d]=await Promise.all([k("/api/branches?limit=10000"),mt("/api/employees")]),i=l.data?.data||[],r=d||[];console.log(`Total employee yang berhasil dimuat dari database : ${r.length}`),r.length>0&&(console.log("Contoh 5 employee pertama:"),r.slice(0,5).forEach((u,b)=>{console.log(`${b+1}. ID: ${u.id}, Name: ${u.full_name}, Status: ${u.status}`)}));let n=u=>{if(!u)return null;let b=String(u||"").replace(/\s+/g," ").toLowerCase().trim(),h=i.find(_=>String(_.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===b||String(_.code||"").replace(/\s+/g," ").toLowerCase().trim()===b||String(_.name||"").replace(/\s+/g," ").toLowerCase().trim()===b);return h?h.id:null},o=(u,b)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${b}`),console.log(`Nama dari Excel : "${u}"`),!u)return console.log("Alasan gagal mapping : Nama kosong"),null;let h=String(u||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${h}"`),console.log(`Jumlah employee di database : ${r.length}`);let _=r.find(C=>String(C.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===h);return _?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${_.id}`),_.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},m=u=>{if(!u)return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let b=String(u).trim();if(/^\d{4,5}(\.\d+)?$/.test(b)){let _=Math.floor(Number(b));if(_>2e4&&_<99999){let C=new Date(Date.UTC(1899,11,30)+_*864e5);return isNaN(C.getTime())?"":C.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(b))return b.slice(0,10);let h=b.split(/[\/\-\.]/);if(h.length===3){let[_,C,S]=h.map(x=>x.trim());if(_.length===4&&C.length<=2&&S.length<=2)return`${_}-${C.padStart(2,"0")}-${S.padStart(2,"0")}`;if(S.length===4&&C.length<=2&&_.length<=2)return`${S}-${C.padStart(2,"0")}-${_.padStart(2,"0")}`}return b},c=a.map((u,b)=>{let h=b+2,_=String(u["Nama Lengkap"]||"").trim(),C=u["Tanggal Mulai"],S=m(C);if(!S){let R=a.__worksheet,E=a.__headers||[],O=E.indexOf("Tanggal Mulai"),M="N/A",fe="N/A",ge="N/A";if(O!==-1&&R&&window.XLSX){let v=window.XLSX.utils.encode_cell({c:O,r:h-1});ge=v;let f=R[v];f?(M=f.t||"undefined",fe=f.w||"undefined"):M="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let ve="Unknown";C==null||C===""?ve="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":C instanceof Date&&isNaN(C.getTime())?ve="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":ve="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${h}`),console.log(`Employee Name : ${_}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${O})`),console.log(`Raw Cell Value : "${C}"`),console.log(`JavaScript Type : ${typeof C}`),console.log(`SheetJS Cell Type : ${M}`),console.log(`SheetJS Formatted Value : "${fe}"`),console.log(`Value After Trim : "${String(C||"").trim()}"`),console.log(`Value After Date Parser : "${S}"`),console.log(`Is Empty : ${!C}`),console.log(`Is Invalid Date : ${C instanceof Date?isNaN(C.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${ve}`),console.log(`Workbook Sheet : ${R?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${ge}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(u,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(E)),console.log(`==========================
`)}let x=o(_,h),w=null;return x||(w="Karyawan tidak ditemukan di Database"),{isValid:!!x,invalidReason:w,rowNum:h,data:{employee_id:x,branch_id:n(String(u.Cabang||"").trim()),division:String(u["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:S,end_date:m(u["Tanggal Selesai"])||"2099-12-31",status:String(u.Status||"").trim(),_rawName:_}}}),p=[],g=[];if(c.forEach(u=>{u.isValid?p.push(u.data):g.push({rowNum:u.rowNum,name:u.data._rawName,reason:u.invalidReason})}),console.log(`Split Validation - Valid: ${p.length}, Invalid: ${g.length}`),p.length===0)return{inserted:0,skipped:a.length,failed:a.length};let y=await k("/api/import/contracts",{method:"POST",body:JSON.stringify({rows:p,onDuplicate:"update"})});if(!y.ok)throw new Error(y.data?.error||"Import gagal");return y.data}}})}F();K();var gt=[],He=[];function $a(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let s of e)if(t.some(a=>a.period===s))return s;return"Q3"}function Da(t,e){if(t.period!=="Q3")return!1;let s=String(t.status||"").toLowerCase();if(s!=="selesai"&&s!=="completed"&&s!=="done")return!1;let a=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?a.includes("inspeksi"):e==="gcdc"?a.includes("general cleaning")||a.includes("deep cleaning"):!1}async function At(t,e){gt=await H();let s=await ee();He=["BERLIN ARIANSYAH","ADE SURAHMAN"];let a=o=>o&&!He.find(m=>(typeof m=="object"?m.value:m)===o)?[...He,o]:He,l=await k(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),d=o=>{if(!o||o==="-"||String(o).trim()==="")return"";let m=String(o).split("-");return m.length===3&&m[0].length===4?`${m[2]}-${m[1]}-${m[0]}`:o},i=l.data?.data||[],r=$a(i),n=e?e.get("dash_filter"):null;L({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",defaultFilters:n?{period:"Q3"}:{period:r},onDataLoaded:o=>(n&&(o=o.filter(m=>Da(m,n))),o.sort((m,c)=>{let p=m.opening_date?new Date(m.opening_date).getTime():0;return(c.opening_date?new Date(c.opening_date).getTime():0)-p})),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:o=>lt(o)},{key:"period",label:"Periode",render:o=>oe(o)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:o=>d(o)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:o=>d(o)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:o=>d(o)},{key:"status",label:"Status",render:o=>q(o)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:gt},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:He}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:gt,value:o?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:o?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:o?.period},{name:"pic",label:"PIC",type:"combobox",options:a(o?.pic),value:o?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:o?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:o?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:o?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:o?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let o=await k(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(o.ok){let m=o.data.data.map(c=>({Cabang:c.branch_name||"",Kegiatan:c.activity_type||"",Periode:c.period||"",PIC:c.pic||"","Tgl Opening":c.opening_date||"","Tgl Target":c.target_date||"","Tgl Selesai":c.completion_date||"",Status:c.status||""}));B(m,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async o=>{let c=(await k("/api/branches?all=1")).data?.data||[],p=b=>{if(!b)return null;let h=String(b||"").toLowerCase(),_=c.find(C=>String(C.full_name||"").toLowerCase()===h||String(C.code||"").toLowerCase()===h||String(C.name||"").toLowerCase()===h);return _?_.id:null},g=b=>{if(b==null||b==="")return"";if(b instanceof Date&&!isNaN(b.getTime()))return b.toISOString().slice(0,10);let h=String(b).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let C=Number(h);if(C>2e4&&C<99999){let S=new Date(Date.UTC(1899,11,30)+C*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}let _=h.split(/[\/\-\.]/);if(_.length===3){let[C,S,x]=_.map(w=>w.trim());if(C.length===4&&S.length<=2&&x.length<=2)return`${C}-${S.padStart(2,"0")}-${x.padStart(2,"0")}`;if(x.length===4&&S.length<=2&&C.length<=2)return`${x}-${S.padStart(2,"0")}-${C.padStart(2,"0")}`}return h},y=o.map(b=>({branch_id:p(String(b.Cabang||"").trim()),activity_type:String(b.Kegiatan||"").trim(),period:String(b.Periode||"").trim(),pic:String(b.PIC||b.Pic||"").trim(),opening_date:g(b["Tgl Opening"]||b["Tanggal Opening"]||b["Tgl Openir"]),target_date:g(b["Tgl Target"]||b["Tanggal Target"]),completion_date:g(b["Tgl Selesai"]||b["Tanggal Selesai"]),status:String(b.Status||"").trim(),notes:String(b.Catatan||b.Keterangan||"").trim()})).filter(b=>b.activity_type&&b.period),u=await k("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:y,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}F();K();var bt=[],Ye=[];function Ia(t,e){let s=String(t.status||"").toLowerCase();return e==="open"?s==="open":!1}async function Ft(t,e){let s=e?e.get("dash_filter"):null;bt=await H(),Ye=await ee();let a=i=>i&&!Ye.find(r=>r.value===i)?[...Ye,{value:i,label:i}]:Ye,l=new Date().getFullYear();L({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:i=>s?i.filter(r=>Ia(r,s)):i,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:i=>`<span class="badge badge-secondary">${i}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:i=>`<span title="${i}">${i?.length>50?i.slice(0,50)+"\u2026":i}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:i=>`<span title="${i||""}">${i?.length>40?i.slice(0,40)+"\u2026":i||"-"}</span>`},{key:"status",label:"Status",render:i=>q(i)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:i=>window.formatDate(i)},{key:"day_count",label:"Hari",render:i=>i??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:bt},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:["2025","2026","2027","2028","2029","2030"]}],formFields:i=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:i?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:bt,value:i?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:i?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:i?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:i?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(i?.employee_name),value:i?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(i?.fc_specialist),value:i?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:i?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:i?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:i?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let i=await k(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let r=i.data.data.map(n=>({Tanggal:n.report_date||"",Cabang:n.branch_name||"",Kategori:n.category||"",Sumber:n.source||"",Keluhan:n.complaint||"","Nama FC":n.employee_name||"","FC Spesialis":n.fc_specialist||"",Solusi:n.solution||"","Tgl Selesai":n.completion_date||"",Status:n.status||""}));B(r,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async i=>{let n=(await k("/api/branches?all=1")).data?.data||[],o=p=>{if(!p)return null;let g=String(p||"").toLowerCase(),y=n.find(u=>String(u.full_name||"").toLowerCase()===g||String(u.code||"").toLowerCase()===g||String(u.name||"").toLowerCase()===g);return y?y.id:null},m=i.map(p=>({branch_id:o(String(p.Cabang||"").trim()),report_date:String(p.Tanggal||"").trim(),category:String(p.Kategori||"").trim(),source:String(p.Sumber||"").trim(),complaint:String(p.Keluhan||"").trim(),employee_name:String(p["Nama FC"]||"").trim(),fc_specialist:String(p["FC Spesialis"]||"").trim(),solution:String(p.Solusi||"").trim(),completion_date:String(p["Tgl Selesai"]||"").trim(),status:String(p.Status||"").trim()})).filter(p=>p.report_date&&p.complaint&&p.category),c=await k("/api/import/issues",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}}})}F();var $e=[];function Pa(t,e){let s=String(t.status||"").toLowerCase();return e==="pending"?s==="pending":!1}async function Mt(t,e){let s=e?e.get("dash_filter"):null;$e=await H();let a=await ee(),l=["Ade","Berlin"],d=r=>r&&!a.find(n=>n.value===r)?[...a,{value:r,label:r}]:a,i=r=>r&&!l.find(n=>(typeof n=="object"?n.value:n)===r)?[...l,r]:l;L({container:t,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:r=>s?r.filter(n=>Pa(n,s)):r,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:r=>`<span title="${r||""}">${r?.length>50?r.slice(0,50)+"\u2026":r||"-"}</span>`},{key:"solution",label:"Solusi",render:r=>`<span title="${r||""}">${r?.length>40?r.slice(0,40)+"\u2026":r||"-"}</span>`},{key:"status",label:"Status",render:r=>q(r)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:r=>window.formatDate(r)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:r=>r?`<a href="${r}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"combobox",name:"branch_id",label:"Cabang",options:$e},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async r=>{let n=new URLSearchParams(r||{}).toString(),o=await k(`/api/one-on-one?limit=10000&${n}`);if(o.ok){let m=o.data.data.map(p=>({Tanggal:p.meeting_date||"",Cabang:p.branch_name||"","Nama Karyawan":p.employee_name||"",PIC:p.pic||"",Masalah:p.problem||"",Solusi:p.solution||"",Status:p.status||"","Tgl Selesai":p.completion_date||"",Dokumen:p.document_link||""})),{downloadExcel:c}=await Promise.resolve().then(()=>(K(),se));c(m,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let r=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(K(),se));n(r,"Template_Import_OneOnOne")},onImport:async r=>{let n=p=>{if(!p)return null;let g=String(p||"").toLowerCase(),y=$e.find(u=>String(u.label||"").toLowerCase()===g);return y?y.value:null},o=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let g=String(p).trim();if(/^\d{4,5}$/.test(g)){let u=Number(g);if(u>2e4&&u<99999){let b=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);let y=g.split(/[\/\-\.]/);if(y.length===3){let[u,b,h]=y.map(_=>_.trim());if(u.length===4&&b.length<=2&&h.length<=2)return`${u}-${b.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&b.length<=2&&u.length<=2)return`${h}-${b.padStart(2,"0")}-${u.padStart(2,"0")}`}return g},m=r.map(p=>({meeting_date:o(p.Tanggal),employee_name:String(p["Nama Karyawan"]||"").trim(),branch_id:n(String(p.Cabang||"").trim()),pic:String(p.PIC||"").trim(),problem:String(p.Masalah||"").trim(),solution:String(p.Solusi||"").trim(),status:String(p.Status||"").trim(),completion_date:o(p["Tgl Selesai"]),document_link:String(p.Dokumen||"").trim()})).filter(p=>p.meeting_date&&p.employee_name&&p.branch_id),c=await k("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},formFields:r=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:r?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:r?.branch_id&&!$e.find(n=>n.value==r.branch_id)?[...$e,{value:r.branch_id,label:r.branch_name||r.branch_id}]:$e,createApi:{path:"/api/branches",field:"full_name"},value:r?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"combobox",required:!0,options:d(r?.employee_name),value:r?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:i(r?.pic),createApi:{path:"/api/pic",field:"name"},value:r?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:r?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:r?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:r?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:r?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:r?.document_link}]})}F();async function Ot(t){let e=await H(),s=await ee(),a=["Ade","Berlin"],l=r=>r&&!s.find(n=>n.value===r)?[...s,{value:r,label:r}]:s,d=r=>r&&!a.find(n=>(typeof n=="object"?n.value:n)===r)?[...a,r]:a,i=Array.from({length:5},(r,n)=>String(new Date().getFullYear()-n));L({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:r=>{try{let n=JSON.parse(r);return Array.isArray(n)?n.join(", "):r||"-"}catch{return r||"-"}}},{key:"score",label:"Nilai",render:r=>r!=null?`<strong>${r}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:r=>r?`<a href="${r}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:i}],exportOptions:{moduleName:"training",onExport:async r=>{let n=new URLSearchParams(r||{}).toString(),o=await k(`/api/training?limit=10000&${n}`);if(o.ok){let m=o.data.data.map(p=>{let g=p.participants||"";try{let y=JSON.parse(g);g=Array.isArray(y)?y.join(", "):g}catch{}return{Tanggal:p.training_date||"",Batch:p.batch||"",Materi:p.subject||"",Cabang:p.branch_name||"",Trainer:p.trainer||"",Peserta:g,Nilai:p.score!==null&&p.score!==void 0?p.score:"",Dokumen:p.document_link||""}}),{downloadExcel:c}=await Promise.resolve().then(()=>(K(),se));c(m,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let r=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(K(),se));n(r,"Template_Import_Training")},onImport:async r=>{let n=p=>{if(!p)return null;let g=String(p||"").toLowerCase(),y=e.find(u=>String(u.label||"").toLowerCase()===g);return y?y.value:null},o=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let g=String(p).trim();if(/^\d{4,5}$/.test(g)){let u=Number(g);if(u>2e4&&u<99999){let b=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);let y=g.split(/[\/\-\.]/);if(y.length===3){let[u,b,h]=y.map(_=>_.trim());if(u.length===4&&b.length<=2&&h.length<=2)return`${u}-${b.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&b.length<=2&&u.length<=2)return`${h}-${b.padStart(2,"0")}-${u.padStart(2,"0")}`}return g},m=r.map(p=>({training_date:o(p.Tanggal),batch:String(p.Batch||"").trim(),subject:String(p.Materi||"").trim(),branch_id:n(String(p.Cabang||"").trim()),trainer:String(p.Trainer||"").trim(),participants:String(p.Peserta||"").trim(),score:p.Nilai?Number(p.Nilai):null,document_link:String(p.Dokumen||"").trim()})).filter(p=>p.training_date&&p.subject&&p.branch_id),c=await k("/api/import/training",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},formFields:r=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:r?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:r?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:r?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:r?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:d(r?.trainer),value:r?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let n=JSON.parse(r?.participants);return Array.isArray(n)?n.join(", "):r?.participants||""}catch{return r?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:r?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:r?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],onBeforeSubmit:async r=>(r.participants&&(r.participants=JSON.stringify(r.participants.split(",").map(n=>n.trim()).filter(Boolean))),r)})}F();ye();K();function Rt({container:t,title:e,icon:s,apiPath:a,columns:l,formFields:d,filterFields:i,defaultFilters:r={},itemLabel:n="Data",canCreate:o=!0,canEdit:m=!0,canDelete:c=!0,onBeforeSubmit:p,onAfterLoad:g,onDataLoaded:y,extraActions:u=[],initialSearch:b="",exportOptions:h=null,bulkDelete:_=!1,paginationMode:C="server"}){let S=1,x={...r};b&&(x.search=b);let w=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${s} ${e}</h1>
      <div class="page-actions">
        ${o?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
      </div>
    </div>

    ${_?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${h?Re(h.moduleName):""}

    ${i&&i.length>0?`
    <div class="filter-bar card" style="padding: 1rem;">
        ${i.map(v=>{if(v.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}"></div>`;if(v.type==="search-combo"){let f="dl-filter-search",$=(v.options||[]).map(P=>`<option value="${typeof P=="object"?P.label:P}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${f}" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${x.search||""}"><datalist id="${f}">${$}</datalist></div>`}return v.type==="select"?`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${v.name}" id="filter-${v.name}"><option value="">-- ${v.label} --</option>${(v.options||[]).map(f=>`<option value="${typeof f=="object"?f.value:f}" ${x[v.name]===(typeof f=="object"?f.value:f)?"selected":""}>${typeof f=="object"?f.label:f}</option>`).join("")}</select>`:""}).join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function R(){if(!document.getElementById("bulk-toolbar"))return;let f=document.getElementById("bulk-count"),$=document.getElementById("btn-bulk-delete"),P=document.getElementById("btn-bulk-cancel");f.textContent=`${w.size} item dipilih`,w.size>0?($.disabled=!1,P.disabled=!1):($.disabled=!0,P.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{w.clear(),document.querySelectorAll(".row-checkbox").forEach(f=>f.checked=!1);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1),R()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(w.size===0)return;let v=[...w],f=document.createElement("div");f.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",f.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${v.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${v.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(f),f.querySelector("#bulk-cancel-btn").addEventListener("click",()=>f.remove()),f.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let $=f.querySelector("#bulk-confirm-btn");$.disabled=!0,$.textContent="Menghapus...";let P=await k(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:v})});f.remove(),P.ok?(Z(`${v.length} ${n} berhasil dihapus.`),w.clear(),R(),M()):G(P.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),O;if(E?.addEventListener("input",v=>{clearTimeout(O),O=setTimeout(()=>{x.search=v.target.value,S=1,w.clear(),M()},400)}),i?.forEach(v=>{v.type==="select"&&document.getElementById(`filter-${v.name}`)?.addEventListener("change",f=>{x[v.name]=f.target.value,S=1,w.clear(),M()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{x={...r},E&&(E.value=""),i?.forEach(v=>{let f=document.getElementById(`filter-${v.name}`);f&&(f.value="")}),S=1,w.clear(),M()}),document.getElementById("btn-create")?.addEventListener("click",()=>ge(null)),h){document.getElementById(`btn-export-${h.moduleName}`)?.addEventListener("click",async f=>{let $=f.target,P=$.innerHTML;$.innerHTML="\u23F3 Loading...",$.disabled=!0;try{await h.onExport()}catch{G("Gagal export data")}finally{$.innerHTML=P,$.disabled=!1}}),document.getElementById(`btn-template-${h.moduleName}`)?.addEventListener("click",()=>{h.onTemplate()});let v=document.getElementById(`input-import-${h.moduleName}`);v?.addEventListener("change",async f=>{let $=f.target.files[0];if(!$)return;v.disabled=!0;let P=document.createElement("div");P.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",P.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(P);let V=P.querySelector("#import-progress-text"),U=P.querySelector("#import-progress-bar"),D=P.querySelector("#import-summary"),I=P.querySelector("#import-close-btn");I.addEventListener("click",()=>{P.remove(),M()});try{let Y=await Oe($);if(Y.length===0)throw new Error("File kosong atau format salah");let X=500,ie=0,ae=0,N=0,A=Y.length;V.textContent=`Ditemukan ${A} baris data. Memulai import...`;for(let T=0;T<A;T+=X){let z=Y.slice(T,T+X);V.textContent=`Mengimport baris ${T+1} - ${Math.min(T+X,A)} dari ${A}...`,U.style.width=`${Math.round(T/A*100)}%`;try{let j=await h.onImport(z);j?(ie+=j.inserted||j.metrics?.inserted||z.length,ae+=j.skipped||j.metrics?.updated||0):ie+=z.length}catch(j){console.error("Chunk import failed:",j),N+=z.length}}U.style.width="100%",V.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',D.style.display="block",D.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${A}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ie}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${N}</strong></div>
        `,N>0&&(D.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),I.style.display="block",v.value=""}catch(Y){V.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${Y.message}`,U.style.background="var(--danger)",U.style.width="100%",I.style.display="block",v.value=""}finally{v.disabled=!1}})}async function M(){R();let v=document.getElementById("table-container");if(!v)return;v.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let f=C==="client",$=f?1:S,P=f?be:20,V=new URLSearchParams({page:$,limit:P,...Object.fromEntries(Object.entries(x).filter(([,N])=>N))}),U=await k(`${a}?${V}`);if(!U.ok){v.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${U.data?.error||"Error"}</p></div>`;return}let D=U.data?.data||U.data||[],I=U.data?.pagination,Y=D.length,X=D;if(f){D=y(D),X=D;let N=D.length,A=20,T=Math.ceil(N/A);S>T&&T>0&&(S=T);let z=(S-1)*A,j=S*A;D=D.slice(z,j),I={page:S,limit:A,total:N,pages:T}}!1,g&&g(D);let ie=Ue({columns:l,data:D,fullData:X,onEdit:m?N=>ge(N):null,actions:u.map(N=>({...N,handler:A=>N.handler(A,M)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:_?{selectedIds:w,onToggle:R}:null});v.innerHTML="",v.appendChild(ie);let ae=document.getElementById("pagination-container");if(ae&&(ae.innerHTML="",I&&I.pages>1)){let N=Ge({page:I.page,pages:I.pages,total:I.total,limit:I.limit,onPage:A=>{S=A,M()}});N&&ae.appendChild(N)}}function fe(v){let f=typeof d=="function"?d(v):d;return Me(f)}function ge(v){let f=!!v,$=document.createElement("form");if($.noValidate=!0,$.innerHTML=fe(v),f){let V=typeof d=="function"?d(v):d;ze($,v)}let{close:P}=le({title:f?`Edit ${n}`:`Tambah ${n}`,content:$,size:"lg",confirmText:f?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(V,U)=>{if(!$.reportValidity())return;let D=V.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let I=Qe($),Y=typeof d=="function"?d(v):d,X=async A=>{for(let T of A)if(T.type==="row")await X(T.fields);else if(T.type==="combobox"&&I[T.name]){let z=I[T.name],j=(T.options||[]).find(W=>{let ne=String(typeof W=="object"?W.value:W),tt=String(typeof W=="object"?W.label:W);return ne===z||tt===z});if(j)I[T.name]=typeof j=="object"?j.value:j;else if(T.createApi){let W={};W[T.createApi.field]=z,T.createApi.extra&&Object.assign(W,T.createApi.extra);let ne=await k(T.createApi.path,{method:"POST",body:JSON.stringify(W)});if(ne.ok&&ne.data?.id)I[T.name]=ne.data.id;else if(ne.ok&&!ne.data?.id)I[T.name]=z;else throw new Error(`Gagal membuat master data: ${ne.data?.error||"Unknown error"}`)}}};try{await X(Y)}catch(A){G(A.message),D.disabled=!1,D.textContent=f?"Simpan Perubahan":`Tambah ${n}`;return}p&&(I=await p(I,v));let ie=f?"PUT":"POST",ae=f?`${a}/${v.id}`:a,N=await k(ae,{method:ie,body:JSON.stringify(I)});N.ok?(Z(f?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),U(),M()):(G(N.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=f?"Simpan Perubahan":`Tambah ${n}`)}})}function ve(v){Fe(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let f=await k(`${a}/${v.id}`,{method:"DELETE"});f.ok?(Z(`${n} berhasil dihapus.`),M()):G(f.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return M(),M}F();K();async function Kt(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let s=await H(),a=await ee(),l=e?e.get("dash_filter"):null;console.log("RAW",await we()),console.log("OPTIONS",a);let d=n=>n&&!a.find(o=>o.value===n)?[...a,{value:n,label:n}]:a,i=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],r=n=>n&&!i.includes(n)?[...i,n]:i;Rt({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",onDataLoaded:n=>{if(l==="reliever"){let o=new Date,m=o.getFullYear(),c=String(o.getMonth()+1).padStart(2,"0");return n.filter(p=>{if(String(p.status||"").toLowerCase()!=="done")return!1;let g=p.backup_date||"";if(g.includes("/")){let y=g.split("/");if(y.length===3&&(y[2].length===4?y[2]:`20${y[2]}`)==m&&y[1].padStart(2,"0")==c)return!0}else if(g.includes("-")&&g.startsWith(`${m}-${c}`))return!0;return!1})}return n},columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:n=>oe(n)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:n=>window.formatDate(n)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:n=>n?`<span class="badge badge-info">${n}</span>`:"-"},{key:"status",label:"Status",render:n=>q(n)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:i},{type:"select",name:"branch_id",label:"Cabang",options:s},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:s,value:n?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:d(n?.original_fc_name),value:n?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:r(n?.reliever_name),value:n?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:n?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:n?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:n?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:n?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let n=await k(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let o=n.data.data.map(m=>({Cabang:m.branch_name||"","Nama Facility care":m.original_fc_name||"",Periode:m.period||"",Relifer:m.reliever_name||"","Tanggal Back Up":m.backup_date||"","Tanggal Selesai":m.completion_date||"",Keterangan:m.reason||"",Shift:m.shift||"",Status:m.status||""}));o.length===0&&o.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),B(o,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async n=>{let m=(await k("/api/branches?all=1")).data?.data||[],c=y=>{if(!y)return null;let u=String(y||"").toLowerCase(),b=m.find(h=>String(h.full_name||"").toLowerCase()===u||String(h.code||"").toLowerCase()===u||String(h.name||"").toLowerCase()===u);return b?b.id:null},p=n.map(y=>({branch_name:String(y.Cabang||"").trim(),backup_date:String(y["Tanggal Back Up"]||y["Tanggal Backup"]||"").trim(),original_fc_name:String(y["Nama Facility care"]||y["FC Digantikan"]||"").trim(),reliever_name:String(y.Relifer||y.Reliefer||"").trim(),period:String(y.Periode||"").trim(),reason:String(y.Keterangan||"").trim(),shift:String(y.Shift||"").trim(),completion_date:String(y["Tanggal Selesai"]||"").trim(),status:String(y.Status||"").trim()})).filter(y=>y.reliever_name&&y.backup_date),g=await k("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:p,onDuplicate:"update"})});if(!g.ok)throw new Error(g.data?.error||"Import gagal");return g.data}}})}F();K();async function Ht(t){let e=await H(),s=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));L({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>oe(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>q(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:s}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),d=await k(`/api/reports/inspection?limit=10000&${l}`);if(d.ok){let i=d.data.data.map(r=>({Cabang:r.branch_name||"",Periode:r.period||"",Tanggal:r.inspection_date||"","Point FC":r.fc_score!==null&&r.fc_score!==void 0?r.fc_score:"","Point SPV":r.spv_score!==null&&r.spv_score!==void 0?r.spv_score:"",Status:r.status||"","Link Dokumen":r.document_link||""}));B(i,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let l=n=>{if(!n)return null;let o=String(n||"").toLowerCase(),m=e.find(c=>String(c.label||"").toLowerCase()===o);return m?m.value:null},d=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let o=String(n).trim();if(o===""||o==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(o))return o.slice(0,10);if(/^\d{4,5}$/.test(o)){let c=Number(o);if(c>2e4&&c<99999){let p=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let m=o.split(/[\/\-\.]/);if(m.length===3){let[c,p,g]=m.map(y=>y.trim());if(c.length===4&&p.length<=2&&g.length<=2)return`${c}-${p.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&p.length<=2&&c.length<=2)return`${g}-${p.padStart(2,"0")}-${c.padStart(2,"0")}`}return o},i=a.map(n=>({branch_id:l(String(n.Cabang||"").trim()),period:String(n.Periode||"").trim(),inspection_date:d(n.Tanggal),fc_score:n["Point FC"]!==void 0&&n["Point FC"]!==""?Number(n["Point FC"]):null,spv_score:n["Point SPV"]!==void 0&&n["Point SPV"]!==""?Number(n["Point SPV"]):null,status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.inspection_date),r=await k("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:i,onDuplicate:"update"})});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}F();K();async function jt(t){let e=await H(),s=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));L({container:t,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>oe(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>q(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:s}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),d=await k(`/api/reports/cleaning?limit=10000&${l}`);if(d.ok){let i=d.data.data.map(r=>({Cabang:r.branch_name||"",Jenis:r.activity_type||"",Periode:r.period||"",Tanggal:r.activity_date||"",Status:r.status||"","Link Dokumen":r.document_link||""}));B(i,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let l=n=>{if(!n)return null;let o=String(n||"").toLowerCase(),m=e.find(c=>String(c.label||"").toLowerCase()===o);return m?m.value:null},d=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let o=String(n).trim();if(o===""||o==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(o))return o.slice(0,10);if(/^\d{4,5}$/.test(o)){let c=Number(o);if(c>2e4&&c<99999){let p=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let m=o.split(/[\/\-\.]/);if(m.length===3){let[c,p,g]=m.map(y=>y.trim());if(c.length===4&&p.length<=2&&g.length<=2)return`${c}-${p.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&p.length<=2&&c.length<=2)return`${g}-${p.padStart(2,"0")}-${c.padStart(2,"0")}`}return o},i=a.map(n=>({branch_id:l(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:d(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),r=await k("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:i,onDuplicate:"update"})});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}F();K();async function qt(t){let e=await H(),s=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));L({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge badge-warning">${a}</span>`},{key:"period",label:"Periode",render:a=>oe(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>q(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:a=>a||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:s}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:a?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),d=await k(`/api/reports/fogging?limit=10000&${l}`);if(d.ok){let i=d.data.data.map(r=>({Cabang:r.branch_name||"",Jenis:r.activity_type||"Fogging",Periode:r.period||"",Tanggal:r.activity_date||"",Status:r.status||"","Link Dokumen":r.document_link||""}));B(i,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async a=>{let l=n=>{if(!n)return null;let o=String(n||"").toLowerCase(),m=e.find(c=>String(c.label||"").toLowerCase()===o);return m?m.value:null},d=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let o=String(n).trim();if(o===""||o==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(o))return o.slice(0,10);if(/^\d{4,5}$/.test(o)){let c=Number(o);if(c>2e4&&c<99999){let p=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let m=o.split(/[\/\-\.]/);if(m.length===3){let[c,p,g]=m.map(y=>y.trim());if(c.length===4&&p.length<=2&&g.length<=2)return`${c}-${p.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&p.length<=2&&c.length<=2)return`${g}-${p.padStart(2,"0")}-${c.padStart(2,"0")}`}return o},i=a.map(n=>({branch_id:l(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"Fogging").trim(),period:String(n.Periode||"").trim(),activity_date:d(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.activity_date),r=await k("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(i)});if(!r.ok)throw new Error(r.data?.error||"Import gagal");return r.data}}})}F();K();async function Jt(t){let e=await H(),s=await ee(),a=s,l=i=>i&&!s.find(r=>r.value===i)?[...s,{value:i,label:i}]:s,d=i=>i&&!a.find(r=>r.value===i)?[...a,{value:i,label:i}]:a;L({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:i=>window.formatDate(i)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:i=>`<span title="${i||""}">${i?.length>60?i.slice(0,60)+"\u2026":i||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:i=>window.formatDate(i)},{key:"status",label:"Status",render:i=>q(i)},{key:"notes",label:"Keterangan",render:i=>i?.length>40?i.slice(0,40)+"\u2026":i||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]}],formFields:i=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:i?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:d(i?.pic),value:i?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:i?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:i?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:i?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:i?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:i?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async i=>{let r=new URLSearchParams(i||{}).toString(),n=await k(`/api/reports/basecamp?limit=10000&${r}`);if(n.ok){let o=n.data.data.map(m=>({"Tgl Info":m.info_date||"",Cabang:m.branch_name||"",Permasalahan:m.problem||"",PIC:m.pic||"","Tgl Done":m.done_date||"",Status:m.status||"",Keterangan:m.notes||""}));B(o,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async i=>{let r=c=>{if(!c)return null;let p=String(c||"").toLowerCase(),g=e.find(y=>String(y.label||"").toLowerCase()===p);return g?g.value:null},n=c=>{if(c==null||c==="")return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let p=String(c).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let y=Number(p);if(y>2e4&&y<99999){let u=new Date(Date.UTC(1899,11,30)+y*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}let g=p.split(/[\/\-\.]/);if(g.length===3){let[y,u,b]=g.map(h=>h.trim());if(y.length===4&&u.length<=2&&b.length<=2)return`${y}-${u.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&u.length<=2&&y.length<=2)return`${b}-${u.padStart(2,"0")}-${y.padStart(2,"0")}`}return p},o=i.map(c=>({info_date:n(c["Tgl Info"]||c["Tanggal Info"]),branch_id:r(String(c.Cabang||"").trim()),problem:String(c.Permasalahan||"").trim(),pic:String(c.PIC||"").trim(),done_date:n(c["Tgl Done"]||c["Tanggal Done"]),status:String(c.Status||"").trim(),notes:String(c.Keterangan||c.Catatan||"").trim()})).filter(c=>c.info_date&&c.branch_id&&c.problem),m=await k("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(o)});if(!m.ok)throw new Error(m.data?.error||"Import gagal");return m.data}}})}async function Ut(t){L({container:t,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let s=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(F(),ke)),l=await a(`/api/sop?limit=10000&${s}`);if(l.ok){let d=l.data.data.map(r=>({"Nama SOP":r.name||"",Kategori:r.category||"",Dokumen:r.document_link||"",Catatan:r.notes||r.description||""})),{downloadExcel:i}=await Promise.resolve().then(()=>(K(),se));i(d,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:s}=await Promise.resolve().then(()=>(K(),se));s(e,"Template_Import_SOP")},onImport:async e=>{let s=e.map(d=>({name:String(d["Nama SOP"]||"").trim(),category:String(d.Kategori||"").trim(),document_link:String(d.Dokumen||"").trim(),description:String(d.Catatan||"").trim()})).filter(d=>d.name),{apiFetch:a}=await Promise.resolve().then(()=>(F(),ke)),l=await a("/api/sop/import",{method:"POST",body:JSON.stringify(s)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function Gt(t){L({container:t,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let s=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(F(),ke)),l=await a(`/api/checklist?limit=10000&${s}`);if(l.ok){let d=l.data.data.map(r=>({"Nama Checklist":r.name||"",Kategori:r.category||"",Dokumen:r.document_link||"",Deskripsi:r.description||""})),{downloadExcel:i}=await Promise.resolve().then(()=>(K(),se));i(d,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:s}=await Promise.resolve().then(()=>(K(),se));s(e,"Template_Import_Checklist")},onImport:async e=>{let s=e.map(d=>({name:String(d["Nama Checklist"]||"").trim(),category:String(d.Kategori||"").trim(),document_link:String(d.Dokumen||"").trim(),description:String(d.Deskripsi||"").trim()})).filter(d=>d.name),{apiFetch:a}=await Promise.resolve().then(()=>(F(),ke)),l=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(s)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}F();ye();K();async function ht(t,e="forms"){if(e==="supply")return La(t);Ba(t)}function Ba(t){L({container:t,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let s=new URLSearchParams(e||{}).toString(),a=await k(`/api/forms?limit=10000&${s}`);a.data?.data?B(a.data.data,"Data_Master_Form"):G("Gagal export data master form")},onImport:async e=>{let s=await k("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!s.ok)throw new Error(s.data?.error||"Import failed");return s.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function La(t){let s=((await k("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));L({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,l)=>l.branch_name_ref||l.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>q(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let l=a?.tools_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}let d=a?.chemical_items;try{d=Array.isArray(JSON.parse(d))?JSON.parse(d).join(", "):d}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:a?.branch_id&&!s.find(i=>i.value==a.branch_id)?[...s,{value:a.branch_id,label:a.branch_name||a.branch_id}]:s,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:l},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:d},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),d=await k(`/api/reports/supply?limit=10000&${l}`);if(d.ok){let i=d.data.data.map(r=>{let n=r.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let o=r.chemical_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}return{Waktu:r.submitted_at||"",Pengirim:r.submitter_name||"",Cabang:r.branch_name_ref||r.branch_name||"","Alat/Barang":n||"",Chemical:o||"",Catatan:r.additional_notes||"",Status:r.status||"","Diproses Oleh":r.processed_by||""}});B(i,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let d=(await k("/api/branches?all=1")).data?.data||[],i=m=>{if(!m)return null;let c=String(m||"").toLowerCase(),p=d.find(g=>String(g.full_name||"").toLowerCase()===c||String(g.code||"").toLowerCase()===c||String(g.name||"").toLowerCase()===c);return p?p.id:null},r=m=>{if(m==null||m==="")return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let c=String(m).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let g=Number(c);if(g>2e4&&g<99999){let y=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}let p=c.split(/[\/\-\.]/);if(p.length===3){let[g,y,u]=p.map(b=>b.trim());if(g.length===4&&y.length<=2&&u.length<=2)return`${g}-${y.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&y.length<=2&&g.length<=2)return`${u}-${y.padStart(2,"0")}-${g.padStart(2,"0")}`}return c},n=a.map(m=>({submitted_at:r(m.Waktu||m.Tanggal),submitter_name:String(m.Pengirim||"").trim(),branch_id:i(String(m.Cabang||"").trim()),tools_items:String(m["Alat/Barang"]||m.Alat||"").trim(),chemical_items:String(m.Chemical||"").trim(),additional_notes:String(m.Catatan||m.Keterangan||"").trim(),status:String(m.Status||"").trim(),processed_by:String(m["Diproses Oleh"]||m.PIC||"").trim()})).filter(m=>m.submitted_at&&m.submitter_name&&m.branch_id),o=await k("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,l)=>{let d=le({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(i,r)=>{let n=i.querySelector("#supply-status").value,o=i.querySelector("#supply-processed-by").value;(await k(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:o})})).ok?(Z("Status diperbarui."),r(),l()):G("Gagal update status.")}})}}]})}F();K();async function Qt(t){let e=he();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}L({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:s=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[s]||"badge-neutral"}">${s}</span>`},{key:"is_active",label:"Status",render:s=>s?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:s=>s?new Date(s).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:s=>{let a=!!s;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:s?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:s?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:s?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:s?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?s?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let s=await k(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(s.ok){let a=s.data.data.map(l=>({"Nama Lengkap":l.full_name||"",Username:l.username||"",Email:l.email||"",Role:l.role||"",Status:l.is_active?"Aktif":"Nonaktif"}));B(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async s=>{let a=s.map(d=>({full_name:String(d["Nama Lengkap"]||"").trim(),username:String(d.Username||"").trim(),email:String(d.Email||"").trim(),role:String(d.Role||"").trim()||"viewer",password:String(d.Password||"").trim()})).filter(d=>d.username&&d.password&&d.email&&d.full_name),l=await k("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}F();K();async function zt(t){L({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await k(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)B(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{B([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let s=e.map(l=>({code:String(l["Kode Cabang"]||"").trim(),name:String(l["Nama Pendek"]||"").trim(),full_name:String(l["Nama Lengkap"]||"").trim(),city:String(l.Kota||"").trim()})).filter(l=>l.code&&l.name),a=await k("/api/branches/import",{method:"POST",body:JSON.stringify(s)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}F();async function Vt(t){let e=new Date,s=[];t.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),l()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),l()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(d=>d.addEventListener("change",l));async function a(){try{let d=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;s=(await k(`/api/dashboard/calendar?month=${d}`)).data?.data||[]}catch(d){console.warn("[Calendar] Failed to load events, rendering empty grid:",d),s=[]}}async function l(){let d=document.getElementById("calendar-grid");if(d){d.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let i=e.getFullYear(),r=e.getMonth(),n=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),o=document.getElementById("cal-month-label");o&&(o.textContent=n);let m=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(S=>S.value)),c=s.filter(S=>m.has(S.type)),p={};c.forEach(S=>{let x=(S.event_date||"").slice(0,10);p[x]||(p[x]=[]),p[x].push(S)});let g=new Date(i,r,1).getDay(),y=new Date(i,r+1,0).getDate(),u=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],b=new Date().toISOString().slice(0,10),h='<div class="calendar-grid">';u.forEach(S=>{h+=`<div class="cal-day-header">${S}</div>`});for(let S=0;S<g;S++)h+='<div class="cal-cell cal-cell-empty"></div>';for(let S=1;S<=y;S++){let x=`${i}-${String(r+1).padStart(2,"0")}-${String(S).padStart(2,"0")}`,w=p[x]||[],R=x===b;h+=`
          <div class="cal-cell ${R?"cal-today":""} ${w.length?"cal-has-events":""}"
               data-date="${x}" tabindex="0" role="button" aria-label="${x}">
            <div class="cal-day-num ${R?"today-num":""}">${S}</div>
            <div class="cal-events-preview">
              ${w.slice(0,3).map(E=>`
                <div class="cal-event-dot cal-color-${E.color||"gray"}" title="${Xe(E.title||E.type)}">
                  <span class="cal-event-dot-label">${Na(E.title||E.branch_name||E.type,18)}</span>
                </div>
              `).join("")}
              ${w.length>3?`<div class="cal-more">+${w.length-3} lagi</div>`:""}
            </div>
          </div>`}let C=(g+y)%7;if(C!==0)for(let S=0;S<7-C;S++)h+='<div class="cal-cell cal-cell-empty"></div>';h+="</div>",d.innerHTML=h,d.querySelectorAll(".cal-cell[data-date]").forEach(S=>{S.addEventListener("click",()=>{let x=S.dataset.date,w=p[x]||[];if(!w.length)return;let R=document.getElementById("cal-event-list"),E=new Date(x+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=E,document.getElementById("cal-event-items").innerHTML=w.map(O=>`
            <div class="cal-event-item cal-color-border-${O.color||"gray"}">
              <div class="cal-event-type">${Aa(O.type)}</div>
              <div class="cal-event-title">${Xe(O.title||"-")}</div>
              <div class="cal-event-branch">${Xe(O.branch_name||"")}</div>
              ${O.status?`<div class="cal-event-status">${Xe(O.status)}</div>`:""}
              ${O.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${O.days_remaining} hari</div>`:""}
            </div>
          `).join(""),R.style.display="block"})})}catch(i){console.error("[Calendar] Render error:",i),d&&(d.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}l()}function Na(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function Xe(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Aa(t){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[t]||t}F();async function Wt(t){let e=he(),s=(e?.full_name||e?.username||"U")[0].toUpperCase(),l={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${l},${l}99)">
            ${s}
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
  `;let d=localStorage.getItem("fm_token"),i=document.getElementById("session-info");if(d&&i)try{let r=JSON.parse(atob(d.split(".")[1])),n=new Date(r.exp*1e3);i.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{i.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async r=>{r.preventDefault();let n=document.getElementById("pwd-error"),o=document.getElementById("pwd-success"),m=document.getElementById("btn-save-pwd");n.style.display="none",o.style.display="none";let c=r.target,p=c.current_password.value,g=c.new_password.value,y=c.confirm_password.value;if(g!==y){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(g.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}m.disabled=!0,m.textContent="\u23F3 Menyimpan...";let u=await k("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:p,new_password:g})});m.disabled=!1,m.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',u.ok?(o.textContent="\u2705 Password berhasil diubah.",o.style.display="block",c.reset(),Z("Password berhasil diubah.")):(n.textContent=u.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}F();var Ze={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function te(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let l=Number(e);if(l>2e4&&l<99999){let d=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(d.getTime())?null:d.toISOString().slice(0,10)}}let s=e.split(/[\/\-\.]/);if(s.length===3){let[l,d,i]=s.map(m=>m.trim()),r=Number(l),n=Number(d),o=Number(i);if(l.length===4&&r>1900)return`${l}-${d.padStart(2,"0")}-${i.padStart(2,"0")}`;if(i.length===4&&o>1900)return r>12?`${i}-${d.padStart(2,"0")}-${l.padStart(2,"0")}`:n>12?`${i}-${l.padStart(2,"0")}-${d.padStart(2,"0")}`:`${i}-${d.padStart(2,"0")}-${l.padStart(2,"0")}`;if(i.length===2&&!isNaN(o)){let m=o>=50?`19${i}`:`20${i}`;return r>12?`${m}-${d.padStart(2,"0")}-${l.padStart(2,"0")}`:`${m}-${d.padStart(2,"0")}-${l.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function Yt(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Fa={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",phone:t["No. Hp"],join_date:te(t["Tanggal Masuk"]),status:t.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:te(t["Tanggal Mulai"]),end_date:te(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:te(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:te(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:te(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:te(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:te(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:te(t["Tanggal Target"]||t["Tgl Target"]),completion_date:te(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:te(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:te(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:te(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:te(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:te(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:te(t["Tanggal Back Up"]),completion_date:te(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:te(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:te(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Ma(t,e){let s=Ze[t];if(!s)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Fa[s.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let l=[],d=[],i=[];return e.filter(n=>!Yt(n)).forEach((n,o)=>{let m=e.indexOf(n)+2,c=[];a.required.forEach(({key:g,label:y})=>{let u=n[g];if(u==null||String(u).trim()===""){let b=Object.keys(n).filter(h=>h.trim()).join(", ");c.push({column:y,originalValue:u||"",reason:`Kolom "${y}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${b.slice(0,120)}`})}});let p=a.map(n);c.length>0?d.push({row:m,data:p,raw:n,errors:c}):(l.push(n),i.push(p))}),{valid:l,errors:d,mapped:i}}function Xt(t){let e=[];return t.SheetNames.forEach(s=>{let a=Ze[s];if(!a)return;let l=t.Sheets[s],d=window.XLSX.utils.sheet_to_json(l,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),i=Ma(s,d),r=d.filter(n=>!Yt(n));e.push({sheetName:s,module:a.module,label:a.label,total:r.length,valid:i.mapped.length,errorCount:i.errors.length,errors:i.errors,mapped:i.mapped,skipped:!1})}),e}function Zt(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,l])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(l),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function ea(t){let e=window.XLSX,s=e.utils.book_new(),a=!1;return t.forEach(l=>{if(!l.errors||l.errors.length===0)return;a=!0;let d=l.errors.map(r=>({"No. Baris":r.row,"Kolom Gagal":(r.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(r.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(r.data||{}).map(([n,o])=>[n,o??""]))})),i=e.utils.json_to_sheet(d);e.utils.book_append_sheet(s,i,l.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(s,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Oa=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function ta(t){t.innerHTML=`
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
              ${Object.entries(Ze).map(([u,{label:b}])=>`<span class="import-sheet-tag">\u{1F4C4} ${u} \u2192 ${b}</span>`).join("")}
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
  `;let e=null,s=null,a=0,l={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function d(u){Object.entries(l).forEach(([b,h])=>{h.style.display=b===u?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let u=document.getElementById("btn-backup-db");u.disabled=!0,u.textContent="\u23F3 Memproses Backup...";try{let b=await k("/api/import/backup");if(b.ok){if(!window.XLSX){G("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let h=window.XLSX,_=h.utils.book_new();Object.entries(b.data.database).forEach(([C,S])=>{let x=S.length>0?S:[{}],w=h.utils.json_to_sheet(x);h.utils.book_append_sheet(_,w,C.substring(0,31))}),h.writeFile(_,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),Z("Backup berhasil diunduh!")}else G("Gagal memproses backup: "+(b.data?.error||"Unknown error"))}catch(b){G("Gagal memproses backup: "+b.message)}finally{u.disabled=!1,u.textContent="\u{1F4E6} Backup Database"}});let i=document.getElementById("btn-sync-google");i&&i.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let u=i.innerHTML;i.innerHTML='<span class="spinner"></span> Menyinkronkan...',i.disabled=!0;try{let b=await k("/api/sync/google-sheets",{method:"POST"});b.ok?alert("Sinkronisasi Berhasil: "+(b.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(b.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{i.innerHTML=u,i.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Zt(),Z("Template Excel berhasil didownload!")});let r=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",u=>{u.stopPropagation(),r.click()}),r.addEventListener("change",u=>{u.target.files[0]&&o(u.target.files[0])}),n.addEventListener("dragover",u=>{u.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",u=>{u.preventDefault(),n.classList.remove("drag-over");let b=u.dataTransfer.files[0];b&&b.name.match(/\.xlsx?$/i)?o(b):G("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,r.value="",document.getElementById("file-info").style.display="none",n.style.display="",d("upload")});async function o(u){e=u,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${u.name} (${(u.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await m(u)}async function m(u){d("validating");let b=document.getElementById("validation-status"),h=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");b.textContent="Membaca file Excel...",h.style.width="20%",await je(200);let _=await u.arrayBuffer(),C=window.XLSX.read(_,{type:"array",cellDates:!0});b.textContent=`Memvalidasi ${C.SheetNames.length} sheet...`,h.style.width="50%",await je(100),s=Xt(C),h.style.width="100%",b.textContent="Validasi selesai!",await je(300),c()}catch(_){d("upload"),G("Gagal memproses file: "+_.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function c(){d("preview");let u=s.filter(E=>!E.skipped).length,b=s.reduce((E,O)=>E+O.total,0),h=s.reduce((E,O)=>E+O.valid,0),_=s.reduce((E,O)=>E+O.errorCount,0),C=b>0?Math.round(h/b*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${u} sheet</span>
      <span class="badge badge-secondary">${b} baris</span>
      <span class="badge badge-success">${h} valid (${C}%)</span>
      ${_>0?`<span class="badge badge-danger">${_} error</span>`:""}
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
          ${s.map((E,O)=>`
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
                ${E.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${O}">\u{1F50D} ${E.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,S.querySelectorAll(".btn-detail-error").forEach(E=>{E.addEventListener("click",()=>{let O=s[Number(E.dataset.idx)];p(O)})});let x=document.getElementById("error-detail-section"),w=document.getElementById("error-detail-container");w.innerHTML="",x.style.display="none";let R=document.getElementById("btn-start-import");h===0?(R.disabled=!0,R.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(R.disabled=!1,_>0?(R.innerHTML=`\u{1F680} Import ${h} Data Valid (${_} dilewati)`,R.title="Baris error akan dilewati, baris valid tetap diimport"):R.innerHTML=`\u{1F680} Mulai Import ${h} Data`)}function p(u){let b=document.getElementById("error-detail-section"),h=document.getElementById("error-detail-container");b.style.display="";let _=u.errors.slice(0,100).map(C=>(Array.isArray(C.errors)?C.errors:[]).map(x=>{let w=typeof x=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${C.row}</span></td>
            <td><strong>${w?x.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${w&&x.originalValue!==void 0?x.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${w?x.reason:x}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${w&&x.aliases?`Gunakan salah satu nama kolom:<br><em>${x.aliases}</em>`:w&&x.hint?x.hint:""}
            </td>
          </tr>
        `}).join("")).join("");h.innerHTML=`
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
            <tbody>${_||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${u.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,b.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{d("upload"),document.getElementById("file-info").style.display="none",n.style.display="",e=null,r.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!s)return;ea(s)?Z("Log error berhasil didownload."):Z("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let u=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";g(u)});async function g(u){d("importing"),a=Date.now();let b=[];Oa.forEach(x=>{let w=s?.find(R=>R.module===x&&R.mapped?.length>0);w&&b.push(w)});let h=document.getElementById("import-steps-list");h.innerHTML=b.map(x=>`
      <div class="import-step-item" id="step-item-${x.module}">
        <span class="step-item-icon" id="step-icon-${x.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${x.label} <span class="step-item-count">(${x.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${x.module}"></span>
      </div>
    `).join("");let _=document.getElementById("import-bar"),C=document.getElementById("import-current-status"),S={totalSheets:b.length,totalRows:b.reduce((x,w)=>x+w.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let x=0;x<b.length;x++){let w=b[x],R=document.getElementById(`step-icon-${w.module}`),E=document.getElementById(`step-status-${w.module}`);R.textContent="\u{1F504}",E.textContent="Mengimport...",C.textContent=`Mengimport ${w.label}...`,_.style.width=`${Math.round(x/b.length*100)}%`;try{let O=await k(`/api/import/${w.module}`,{method:"POST",body:JSON.stringify({rows:w.mapped,onDuplicate:u})});if(O.ok){let M=O.data;S.inserted+=M.inserted||0,S.skipped+=M.skipped||0,S.moduleResults.push({label:w.label,inserted:M.inserted||0,skipped:M.skipped||0,status:"ok"}),R.textContent="\u2705",E.innerHTML=`<span class="badge badge-success">${M.inserted||0} berhasil</span>${M.skipped>0?` <span class="badge badge-neutral">${M.skipped} skip</span>`:""}`}else S.failed++,S.moduleResults.push({label:w.label,inserted:0,skipped:0,status:"error",error:O.data?.error}),R.textContent="\u274C",E.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(O){S.failed++,S.moduleResults.push({label:w.label,inserted:0,skipped:0,status:"error",error:O.message}),R.textContent="\u274C",E.innerHTML='<span class="badge badge-danger">Gagal</span>'}await je(150)}_.style.width="100%",C.textContent="Selesai!",await je(400),y(S)}function y(u){d("summary");let b=((Date.now()-a)/1e3).toFixed(1),h=u.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${h?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${h?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
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
        <div class="stat-value">${b}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${u.moduleResults.map(_=>`
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,s=null,r.value="",document.getElementById("file-info").style.display="none",n.style.display="",d("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function je(t){return new Promise(e=>setTimeout(e,t))}F();var et=[],aa=[];async function na(t){et=await H(),aa=await ee(),L({container:t,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:et}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let s=new URLSearchParams(e||{}).toString(),a=await k(`/api/sp?limit=10000&${s}`);if(a.ok){let l=a.data.data.map(i=>({"Nama Karyawan":i.employee_name||"",Divisi:i.division||"",Cabang:i.branch_name||"","Tanggal Sp":i.tanggal||"","Akhir Sp":i.akhir_sp||"","Jenis Sp":i.sp_type||"","Link Document / Foto":i.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(K(),se));d(l,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:s}=await Promise.resolve().then(()=>(K(),se));s(e,"Template_Import_SP")},onImport:async e=>{let s=i=>{if(!i)return null;let r=String(i||"").toLowerCase(),n=et.find(o=>String(o.label||"").toLowerCase()===r);return n?n.value:null},a=i=>{if(!i)return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let r=String(i).trim();if(/^\d{4,5}$/.test(r)){let o=Number(r);if(o>2e4&&o<99999){let m=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(r))return r.slice(0,10);let n=r.split(/[\/\-\.]/);if(n.length===3){let[o,m,c]=n.map(p=>p.trim());if(o.length===4&&m.length<=2&&c.length<=2)return`${o}-${m.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&m.length<=2&&o.length<=2)return`${c}-${m.padStart(2,"0")}-${o.padStart(2,"0")}`}return r},l=e.map(i=>({employee_name:String(i["Nama Karyawan"]||"").trim(),division:String(i.Divisi||"").trim(),branch_id:s(String(i.Cabang||"").trim()),tanggal:a(i["Tanggal Sp"]),akhir_sp:a(i["Akhir Sp"]),sp_type:String(i["Jenis Sp"]||"").trim(),document_link:String(i["Link Document / Foto"]||"").trim()})).filter(i=>i.employee_name&&i.branch_id),d=await k("/api/import/sp",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:[{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:aa},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"combobox",name:"branch_id",label:"Cabang",required:!0,options:et,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}F();var De=[],ia=[];async function ra(t){De=await H(),ia=await ee(),L({container:t,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",options:De},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",options:De}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let s=new URLSearchParams(e||{}).toString(),a=await k(`/api/mutasi?limit=10000&${s}`);if(a.ok){let l=a.data.data.map(i=>({Tanggal:i.tanggal||"","Nama Karyawan":i.employee_name||"","Cabang Asal":i.from_branch_name||"","Cabang Tujuan":i.to_branch_name||"",Status:i.status||"",Dokumen:i.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(K(),se));d(l,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:s}=await Promise.resolve().then(()=>(K(),se));s(e,"Template_Import_Mutasi")},onImport:async e=>{let s=i=>{if(!i)return null;let r=String(i||"").toLowerCase(),n=De.find(o=>String(o.label||"").toLowerCase()===r);return n?n.value:null},a=i=>{if(!i)return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let r=String(i).trim();if(/^\d{4,5}$/.test(r)){let o=Number(r);if(o>2e4&&o<99999){let m=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(r))return r.slice(0,10);let n=r.split(/[\/\-\.]/);if(n.length===3){let[o,m,c]=n.map(p=>p.trim());if(o.length===4&&m.length<=2&&c.length<=2)return`${o}-${m.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&m.length<=2&&o.length<=2)return`${c}-${m.padStart(2,"0")}-${o.padStart(2,"0")}`}return r},l=e.map(i=>({tanggal:a(i.Tanggal),employee_name:String(i["Nama Karyawan"]||"").trim(),from_branch_id:s(String(i["Cabang Asal"]||"").trim()),to_branch_id:s(String(i["Cabang Tujuan"]||"").trim()),status:String(i.Status||"").trim(),document_link:String(i.Dokumen||"").trim()})).filter(i=>i.tanggal&&i.employee_name&&i.from_branch_id&&i.to_branch_id),d=await k("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:ia},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:De,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:De,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}F();async function sa(t){t.innerHTML=`
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
  `;let e=document.getElementById("btnRefreshSync"),s=document.getElementById("queueStatusFilter");e.addEventListener("click",l),s.addEventListener("change",r),document.getElementById("btnRetryAll").addEventListener("click",()=>a("retry",{allFailed:!0})),document.getElementById("btnResetStuck").addEventListener("click",()=>a("reset-stuck",{})),document.getElementById("btnPauseSync").addEventListener("click",()=>a("pause",{})),document.getElementById("btnResumeSync").addEventListener("click",()=>a("resume",{})),document.getElementById("btnReconcile").addEventListener("click",()=>a("reconcile",{module:"ALL",repairMode:!0})),document.getElementById("btnRetrySelected").addEventListener("click",()=>{let o=Array.from(document.querySelectorAll(".chk-queue:checked")).map(m=>m.value);if(o.length===0)return alert("No items selected");a("retry",{ids:o})}),document.getElementById("chkAllQueue").addEventListener("change",o=>{document.querySelectorAll(".chk-queue").forEach(m=>m.checked=o.target.checked)});async function a(o,m){if(confirm(`Are you sure you want to execute action: ${o}?`)){showLoading();try{let c=await k(`/api/sync/actions/${o}`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(m)});c.ok?(alert(c.data?.message||"Success"),l()):G(c.error||"Action failed")}catch(c){G(c.message)}hideLoading()}}await l();async function l(){showLoading(),await Promise.all([i(),r(),d(),n()]),hideLoading()}async function d(){try{let o=await k("/api/sync/performance");if(!o.ok)return;let{webhook:m,google_api:c,d1:p,queue:g,throughput:y}=o.data;document.getElementById("latencyHistograms").innerHTML=`
        <table class="w-full text-left">
          <thead><tr class="text-gray-500 border-b"><th>Metric</th><th>P50</th><th>P95</th><th>P99</th><th>Max</th></tr></thead>
          <tbody class="divide-y">
            <tr><td class="py-1">Webhook</td><td>${m.P50}ms</td><td>${m.P95}ms</td><td>${m.P99}ms</td><td>${m.Max}ms</td></tr>
            <tr><td class="py-1">Google API</td><td>${c.P50}ms</td><td>${c.P95}ms</td><td>${c.P99}ms</td><td>${c.Max}ms</td></tr>
            <tr><td class="py-1">D1 Execute</td><td>${p.P50}ms</td><td>${p.P95}ms</td><td>${p.P99}ms</td><td>${p.Max}ms</td></tr>
            <tr><td class="py-1">Queue Delay</td><td>${g.P50}ms</td><td>${g.P95}ms</td><td>${g.P99}ms</td><td>${g.Max}ms</td></tr>
          </tbody>
        </table>
        <div class="mt-4 pt-3 border-t text-gray-600 flex justify-between">
          <span>Throughput: <b>${y.events_per_sec}</b> ev/sec</span>
          <span><b>${y.events_per_min}</b> ev/min</span>
        </div>
      `,document.getElementById("lblSnapshotStatus").innerHTML='Checked 10 modules. Status: <b>COMPLETED</b><br><span class="text-xs text-gray-500">Run today at 02:00</span>'}catch(o){console.error(o)}}async function i(){try{let o=await k("/api/sync/health");if(!o.ok)return G("Failed to fetch sync health");let{status:m,queue:c,circuit_breaker:p}=o.data,g=`
        <div class="bg-white p-4 rounded-lg shadow border-l-4 ${m==="HEALTHY"?"border-green-500":m==="WARNING"?"border-yellow-500":"border-red-500"}">
          <p class="text-sm text-gray-500">System Health</p>
          <p class="text-2xl font-bold ${m==="HEALTHY"?"text-green-600":m==="WARNING"?"text-yellow-600":"text-red-600"}">${m}</p>
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
      `;document.getElementById("syncOverviewCards").innerHTML=g;let y=document.getElementById("cbStateBadge"),u=document.getElementById("cbStateDesc"),b=document.getElementById("cbStatusCard");b.className="bg-white rounded-lg shadow p-6 border-l-4",p==="CLOSED"?(y.className="px-3 py-1 rounded-full text-sm font-medium bg-green-100 text-green-800",y.textContent="CLOSED",u.textContent="Traffic is flowing normally to Google Sheets.",b.classList.add("border-green-500")):p==="OPEN"?(y.className="px-3 py-1 rounded-full text-sm font-medium bg-red-100 text-red-800",y.textContent="OPEN",u.textContent="Failing fast. Traffic to Google Sheets is paused due to repeated failures.",b.classList.add("border-red-500")):p==="HALF_OPEN"?(y.className="px-3 py-1 rounded-full text-sm font-medium bg-yellow-100 text-yellow-800",y.textContent="HALF-OPEN",u.textContent="Testing recovery. Permitting limited traffic to verify stability.",b.classList.add("border-yellow-500")):y.textContent=p||"UNKNOWN"}catch(o){console.error(o)}}async function r(){try{let o=document.getElementById("queueStatusFilter").value,m=await k("/api/sync/queue?limit=15"+(o?"&status="+o:""));if(!m.ok)return;let c=document.getElementById("queueTableBody"),p=m.data?.data||m.data||[];if(p.length===0){c.innerHTML='<tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">No events found</td></tr>';return}c.innerHTML=p.map(g=>`
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
      `).join("")}catch(o){console.error(o)}}async function n(){try{let o=await k("/api/sync/metrics");if(!o.ok)return;let m=document.getElementById("metricsTableBody"),c=o.data||[];if(c.length===0){m.innerHTML='<tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">No metrics found in last 24h</td></tr>';return}m.innerHTML=c.map(p=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2 font-medium text-gray-800">${p.module}</td>
          <td class="px-4 py-2 text-gray-600">${p.total_events}</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(p.avg_webhook_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(p.avg_d1_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(p.avg_queue_wait_ms||0)} ms</td>
        </tr>
      `).join("")}catch(o){console.error(o)}}}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let e=Math.floor(Number(t)-25569);return new Date(e*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let e=t.split(/[\/\-]/);return`${e[2]}-${e[1]}-${e[0]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let s=e.split("-");if(s.length===3&&s[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],l=parseInt(s[2],10),d=a[parseInt(s[1],10)-1];return`${l} ${d} ${s[0]}`}return e};function Q(t){return async e=>{if(!Be()){Se("/login");return}return t(e)}}var qe=null;function Ra(){qe&&clearInterval(qe);let t=()=>{let e=new Date,s=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),l=document.getElementById("header-clock-time"),d=document.getElementById("header-clock-date");l&&(l.textContent=s),d&&(d.textContent=a)};t(),qe=setInterval(t,1e3)}async function Ka(){try{let t=await k("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},s=(a,l)=>{let d=document.getElementById(a);d&&(d.textContent=l>0?l:"",d.style.display=l>0?"inline-flex":"none")};s("badge-issues",e.issues?.current||0),s("badge-contracts",e.expiring30?.current||0),s("badge-oo1",e.one_on_one?.current||0),s("badge-schedule",e.schedule?.current||0),s("badge-supply",e.supply?.current||0)}catch{}}var Ie=[];async function Ha(){try{let t=await k("/api/dashboard/notifications");if(!t.ok)return;Ie=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Ie.length>0?"block":"none",e.textContent=Ie.length)}catch{}}function ja(){if(!Ie.length){le({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,s)=>s()});return}let t=`
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
  `;le({title:`Notifikasi (${Ie.length})`,content:t,confirmText:"Tutup",onConfirm:(e,s)=>s()})}function la(){let t=he(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
                <span class="topbar-greeting-time">${(()=>{let o=new Date().getHours();return o>=4&&o<11?"Selamat Pagi":o>=11&&o<15?"Selamat Siang":o>=15&&o<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">Berlin Ariansyah</span> \u{1F44B}
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
  `;let s=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),l=document.getElementById("topbar-menu-btn"),d=document.getElementById("sidebar-close"),i=()=>{s.classList.add("open"),a.classList.add("show")},r=()=>{s.classList.remove("open"),a.classList.remove("show")};l?.addEventListener("click",i),d?.addEventListener("click",r),a?.addEventListener("click",r),document.querySelectorAll(".nav-item").forEach(o=>o.addEventListener("click",r));function n(){let o=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(p=>{let g=p.dataset.route;p.classList.toggle("active",o===g||g!=="/dashboard"&&o.startsWith(g))});let m=document.getElementById("topbar-title"),c=document.querySelector(".nav-item.active .nav-label");m&&c&&(m.textContent=c.textContent)}window.addEventListener("hashchange",n),n(),Ra(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await k("/api/auth/logout",{method:"POST"}),Le(),qe&&clearInterval(qe),Se("/login")}),Ka(),Ha(),document.getElementById("btn-notif")?.addEventListener("click",o=>{o.preventDefault(),ja()})}async function qa(){J("/login",({main:e})=>Pt(e)),J("/dashboard",Q(({main:e})=>Tt(e))),J("/calendar",Q(({main:e})=>Vt(e))),J("/employees",Q(({main:e,params:s})=>Bt(e,s))),J("/contracts",Q(({main:e,params:s})=>Nt(e,s))),J("/sp",Q(({main:e})=>na(e))),J("/mutasi",Q(({main:e})=>ra(e))),J("/sync-dashboard",Q(({main:e})=>sa(e))),J("/timeline",Q(({main:e,params:s})=>At(e,s))),J("/issues",Q(({main:e,params:s})=>Ft(e,s))),J("/one-on-one",Q(({main:e,params:s})=>Mt(e,s))),J("/training",Q(({main:e})=>Ot(e))),J("/relievers",Q(({main:e,params:s})=>Kt(e,s))),J("/reports/inspection",Q(({main:e})=>Ht(e))),J("/reports/cleaning",Q(({main:e})=>jt(e))),J("/reports/fogging",Q(({main:e})=>qt(e))),J("/reports/basecamp",Q(({main:e})=>Jt(e))),J("/reports/supply",Q(({main:e})=>ht(e,"supply"))),J("/sop",Q(({main:e})=>Ut(e))),J("/checklist",Q(({main:e})=>Gt(e))),J("/forms",Q(({main:e})=>ht(e))),J("/users",Q(({main:e})=>Qt(e))),J("/branches",Q(({main:e})=>zt(e))),J("/profile",Q(({main:e})=>Wt(e))),J("/settings/import",Q(({main:e})=>ta(e)));let t=Be();if(!t&&window.location.hash!=="#/login"&&Se("/login"),t){let e=await k("/api/auth/me");e.ok?(Ne(e.data.data),la()):(Le(),Se("/login"))}window.addEventListener("fm:login",()=>{la(),Se("/dashboard")}),ft()}qa();
