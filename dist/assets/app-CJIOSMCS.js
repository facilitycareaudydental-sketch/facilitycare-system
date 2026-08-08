var ca=Object.defineProperty;var nt=(t,e)=>()=>(t&&(e=t(t=0)),e);var it=(t,e)=>{for(var i in e)ca(t,i,{get:e[i],enumerable:!0})};var ke={};it(ke,{API:()=>vt,CLIENT_SIDE_MAX_ROWS:()=>fe,IS_DEVELOPMENT:()=>Pe,apiFetch:()=>k,clearToken:()=>Be,getToken:()=>Le,getUser:()=>se,setToken:()=>rt,setUser:()=>Ae});function Le(){return localStorage.getItem("fm_token")}function rt(t){localStorage.setItem("fm_token",t)}function Be(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function se(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Ae(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function k(t,e={}){let i=Le(),a={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...e.headers||{}};try{let s=`cb=${Date.now()}`,o=t.includes("?")?"&":"?",l=`${vt}${t}${o}${s}`,n=await fetch(l,{...e,headers:a}),r;try{let d=await n.text();try{r=JSON.parse(d)}catch{r={error:`Server Error (${n.status}): ${d.substring(0,80)}...`}}}catch{r={error:"Gagal membaca respon dari server"}}return n.status===401&&(Be(),window.location.hash="#/login"),{ok:n.ok,status:n.status,data:r}}catch(s){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${s.message})`}}}}var Pe,pa,vt,fe,O=nt(()=>{Pe=!1,pa="https://fm-operations-api.facilitycare-audydental.workers.dev",vt=pa,fe=1e4});var st={};it(st,{confirmDialog:()=>Fe,createModal:()=>de});function de({title:t,content:e,onConfirm:i,onCancel:a,confirmText:s="Simpan",cancelText:o="Batal",size:l="md",confirmClass:n="btn-primary"}){let r={sm:"400px",md:"560px",lg:"720px",xl:"900px"},d=document.createElement("div");d.className="modal-overlay",d.innerHTML=`
    <div class="modal" style="max-width:${r[l]||r.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${o}</button>
        ${i?`<button class="btn ${n} modal-confirm">${s}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&d.querySelector(".modal-body").appendChild(e);let g=()=>{d.classList.remove("show"),setTimeout(()=>d.remove(),250)};return d.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),g()}),d.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),g()}),i&&d.querySelector(".modal-confirm").addEventListener("click",()=>i(d,g)),d.addEventListener("click",c=>{c.target===d&&(a&&a(),g())}),document.body.appendChild(d),requestAnimationFrame(()=>d.classList.add("show")),{overlay:d,close:g}}function Fe(t,e,i="Konfirmasi"){return de({title:i,content:`<p>${t}</p>`,onConfirm:(a,s)=>{e(),s()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var ye=nt(()=>{});var oe={};it(oe,{downloadExcel:()=>A,parseExcel:()=>Oe,renderExcelButtons:()=>Re});function Oe(t){return new Promise((e,i)=>{let a=new FileReader;a.onload=s=>{try{let o=new Uint8Array(s.target.result),l=XLSX.read(o,{type:"array"}),n=l.SheetNames[0],r=l.Sheets[n];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${l.SheetNames.join(", ")}`),console.log(`Sheet Used: ${n}`);let d=XLSX.utils.decode_range(r["!ref"]||"A1:A1"),g=d.e.r-d.s.r+1,c=d.e.c-d.s.c+1;console.log(`Total Rows (including empty): ${g}`),console.log(`Total Columns: ${c}`);let p=[];for(let h=d.s.c;h<=d.e.c;++h){let m=r[XLSX.utils.encode_cell({c:h,r:d.s.r})];m&&m.v&&p.push(m.v)}console.log(`Headers Found: ${p.join(", ")}`),console.log("---------------------------");let u=XLSX.utils.sheet_to_json(r,{defval:""});Object.defineProperty(u,"__worksheet",{value:r,enumerable:!1}),Object.defineProperty(u,"__headers",{value:p,enumerable:!1}),e(u)}catch(o){i(o)}},a.onerror=s=>i(s),a.readAsArrayBuffer(t)})}function A(t,e){try{let i=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,i,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function Re(t){return`
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
  `}var K=nt(()=>{});O();var lt={},Ue=null;function q(t,e){lt[t]=e}function Se(t){window.location.hash=t}function kt(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[i,...a]=e.split("?"),s=lt[i];if(!s){for(let[l,n]of Object.entries(lt))if(l.endsWith("/*")&&i.startsWith(l.slice(0,-2))){s=n;break}}Ue&&(Ue(),Ue=null);let o=document.getElementById("main-content");if(o&&(o.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),s){let l=new URLSearchParams(a.join("?")),n=i.split("/").filter(Boolean),r=await s({path:i,params:l,segments:n,main:o});r&&(Ue=r)}else{let l=o||document.getElementById("app");l&&(l.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var Ne;function ua(){return Ne||(Ne=document.createElement("div"),Ne.id="toast-container",document.body.appendChild(Ne)),Ne}function St(t,e="info",i=3500){let a=ua(),s=document.createElement("div");s.className=`toast toast-${e}`;let o={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};s.innerHTML=`<span class="toast-icon">${o[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(s),requestAnimationFrame(()=>s.classList.add("show")),setTimeout(()=>{s.classList.remove("show"),setTimeout(()=>s.remove(),350)},i)}var ee=t=>St(t,"success"),Q=t=>St(t,"error");ye();O();O();ye();function Ge({columns:t,data:e,onEdit:i,onDelete:a,onView:s,actions:o=[],emptyText:l="Tidak ada data",bulkSelect:n=null}){let r=document.createElement("div");if(r.className="table-wrapper",!e||e.length===0)return r.innerHTML=`<div class="empty-state"><p>${l}</p></div>`,r;let d=document.createElement("table");d.className="data-table";let g=document.createElement("thead"),c=document.createElement("tr");if(n){let u=document.createElement("th");u.style.width="40px",u.style.textAlign="center";let h=document.createElement("input");h.type="checkbox",h.id="select-all-checkbox",h.title="Pilih semua",h.addEventListener("change",()=>{e.forEach(m=>{h.checked?n.selectedIds.add(m.id):n.selectedIds.delete(m.id)}),r.querySelectorAll(".row-checkbox").forEach(m=>m.checked=h.checked),n.onToggle()}),u.appendChild(h),c.appendChild(u)}if(t.forEach(u=>{let h=document.createElement("th");h.textContent=u.label,u.width&&(h.style.width=u.width),c.appendChild(h)}),i||a||s||o.length>0){let u=document.createElement("th");u.textContent="Aksi",u.style.width="120px",c.appendChild(u)}g.appendChild(c),d.appendChild(g);let p=document.createElement("tbody");return e.forEach(u=>{let h=document.createElement("tr");if(n){let m=document.createElement("td");m.style.textAlign="center",m.style.width="40px";let b=document.createElement("input");b.type="checkbox",b.className="row-checkbox",b.checked=n.selectedIds.has(u.id),b.addEventListener("change",()=>{if(b.checked)n.selectedIds.add(u.id);else{n.selectedIds.delete(u.id);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1)}n.onToggle()}),m.appendChild(b),h.appendChild(m)}if(t.forEach(m=>{let b=document.createElement("td");if(m.render){let y=m.render(u[m.key],u);y instanceof HTMLElement?b.appendChild(y):b.innerHTML=y||""}else b.textContent=u[m.key]!==null&&u[m.key]!==void 0&&u[m.key]!==""?u[m.key]:"";m.nowrap&&(b.style.whiteSpace="nowrap"),h.appendChild(b)}),i||a||s||o.length>0){let m=document.createElement("td");m.className="actions-cell";let b=document.createElement("div");if(b.className="btn-group",s){let y=document.createElement("button");y.className="btn btn-xs btn-ghost",y.innerHTML="\u{1F441}",y.title="Lihat",y.addEventListener("click",()=>s(u)),b.appendChild(y)}if(i){let y=document.createElement("button");y.className="btn btn-xs btn-secondary",y.innerHTML="\u270F\uFE0F",y.title="Edit",y.addEventListener("click",()=>i(u)),b.appendChild(y)}o.forEach(y=>{let _=document.createElement("button");_.className=`btn btn-xs ${y.class||"btn-ghost"}`,_.innerHTML=y.icon||y.label,_.title=y.label,_.addEventListener("click",()=>y.handler(u)),b.appendChild(_)}),m.appendChild(b),h.appendChild(m)}p.appendChild(h)}),d.appendChild(p),r.appendChild(d),r}function Qe({page:t,pages:e,total:i,limit:a,onPage:s}){if(e<=1)return null;let o=document.createElement("div");o.className="pagination";let l=document.createElement("span");l.className="pagination-info",l.textContent=`Total: ${i} data`,o.appendChild(l);let n=document.createElement("div");n.className="pagination-btns";let r=(c,p,u=!1,h=!1)=>{let m=document.createElement("button");m.className=`btn btn-sm ${h?"btn-primary":"btn-ghost"} pagination-btn`,m.textContent=c,m.disabled=u,m.addEventListener("click",()=>s(p)),n.appendChild(m)};r("\xAB",1,t===1),r("\u2039",t-1,t===1);let d=Math.max(1,t-2),g=Math.min(e,t+2);for(let c=d;c<=g;c++)r(c,c,!1,c===t);return r("\u203A",t+1,t===e),r("\xBB",e,t===e),o.appendChild(n),o}ye();function Me(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${Me(e.fields)}</div>`;let i=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",s="";switch(e.type){case"textarea":s=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${i} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let l=(e.options||[]).map(c=>{let p=typeof c=="object"?c.value:c,u=typeof c=="object"?c.label:c,h=e.value==p?"selected":"";return`<option value="${p}" ${h}>${u}</option>`}).join("");s=`<select name="${e.name}" class="form-control" ${i}><option value="">-- Pilih ${e.label||""} --</option>${l}</select>`;break;case"combobox":let n=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,r=(e.options||[]).map(c=>{let p=typeof c=="object"?c.value:c,u=typeof c=="object"?c.label||c.value||"":c||"";return(u==="undefined"||u==="[object Object]"||u==="null")&&(u=""),u?`<option value="${u}"></option>`:""}).join(""),d=e.value||"";if(e.value){let c=(e.options||[]).find(p=>(typeof p=="object"?p.value:p)==e.value);if(c){let p=typeof c=="object"?c.label||c.value||"":c||"";p&&p!=="undefined"&&p!=="[object Object]"&&p!=="null"&&(d=p)}}s=`
          <input type="text" name="${e.name}" list="${n}" class="form-control" value="${d}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${n}">${r}</datalist>
        `;break;case"checkbox":s=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let g=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";s=`<input type="date" name="${e.name}" class="form-control" value="${g}" ${i}>`;break;case"number":s=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${i}>`;break;case"email":s=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i}>`;break;case"url":s=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${i}>`;break;default:s=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i} autocomplete="off">`}let o=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${s}${o}</div>`}).join("")}function ze(t){let e={},i=new FormData(t);for(let[a,s]of i.entries())e[a]=s===""?null:s;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function We(t,e){e&&Object.entries(e).forEach(([i,a])=>{let s=t.querySelector(`[name="${i}"]`);s&&(s.hasAttribute("list")||(s.type==="checkbox"?s.checked=!!a:s.type==="date"&&a&&window.parseFlexibleDate?s.value=window.parseFlexibleDate(a):s.value=a??""))})}K();function N({container:t,title:e,icon:i,apiPath:a,columns:s,formFields:o,filterFields:l,defaultFilters:n={},itemLabel:r="Data",canCreate:d=!0,canEdit:g=!0,canDelete:c=!0,onBeforeSubmit:p,onAfterLoad:u,onDataLoaded:h,extraActions:m=[],initialSearch:b="",exportOptions:y=null,bulkDelete:_=!1,paginationMode:x="server"}){let w=se();w&&typeof w=="object"&&w.role==="viewer"&&(d=!1,g=!1,c=!1,_=!1,y=null);let S=1,T={...n};b&&(T.search=b);let $=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions">
        ${d?`<button class="btn btn-primary" id="btn-create">+ Tambah ${r}</button>`:""}
      </div>
    </div>

    ${_?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${y?Re(y.moduleName):""}

    ${l&&l.length>0?`
    <div class="filter-bar card" style="padding: 1rem; display: flex; gap: 0.5rem; flex-wrap: wrap;">
        ${l.map(v=>{if(v.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${T.search||""}"></div>`;if(v.type==="search-combo"){let f="dl-filter-search",D=(v.options||[]).map(L=>`<option value="${typeof L=="object"?L.label:L}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${f}" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${T.search||""}"><datalist id="${f}">${D}</datalist></div>`}return v.type==="select"?`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${v.name}" id="filter-${v.name}"><option value="">Pilih ${v.label}</option>${(v.options||[]).map(f=>`<option value="${typeof f=="object"?f.value:f}" ${T[v.name]===(typeof f=="object"?f.value:f)?"selected":""}>${typeof f=="object"?f.label:f}</option>`).join("")}</select>`:""}).join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function C(){if(!document.getElementById("bulk-toolbar"))return;let f=document.getElementById("bulk-count"),D=document.getElementById("btn-bulk-delete"),L=document.getElementById("btn-bulk-cancel");f.textContent=`${$.size} item dipilih`,$.size>0?(D.disabled=!1,L.disabled=!1):(D.disabled=!0,L.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{$.clear(),document.querySelectorAll(".row-checkbox").forEach(f=>f.checked=!1);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1),C()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if($.size===0)return;let v=[...$],f=document.createElement("div");f.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",f.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${v.length} ${r}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${v.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(f),f.querySelector("#bulk-cancel-btn").addEventListener("click",()=>f.remove()),f.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let D=f.querySelector("#bulk-confirm-btn");D.disabled=!0,D.textContent="Menghapus...";let L=await k(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:v})});f.remove(),L.ok?(ee(`${v.length} ${r} berhasil dihapus.`),$.clear(),C(),U()):Q(L.data?.error||"Gagal menghapus data.")})});let B=document.getElementById("filter-search"),R;if(B?.addEventListener("input",v=>{clearTimeout(R),R=setTimeout(()=>{T.search=v.target.value,S=1,$.clear(),U()},400)}),l?.forEach(v=>{v.type==="select"&&document.getElementById(`filter-${v.name}`)?.addEventListener("change",f=>{T[v.name]=f.target.value,S=1,$.clear(),U()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{T={...n},B&&(B.value=""),l?.forEach(v=>{let f=document.getElementById(`filter-${v.name}`);f&&(f.value="")}),S=1,$.clear(),U()}),document.getElementById("btn-create")?.addEventListener("click",()=>me(null)),y){document.getElementById(`btn-export-${y.moduleName}`)?.addEventListener("click",async f=>{let D=f.target,L=D.innerHTML;D.innerHTML="\u23F3 Loading...",D.disabled=!0;try{await y.onExport()}catch{Q("Gagal export data")}finally{D.innerHTML=L,D.disabled=!1}}),document.getElementById(`btn-template-${y.moduleName}`)?.addEventListener("click",()=>{y.onTemplate()});let v=document.getElementById(`input-import-${y.moduleName}`);v?.addEventListener("change",async f=>{let D=f.target.files[0];if(!D)return;v.disabled=!0;let L=document.createElement("div");L.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",L.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(L);let V=L.querySelector("#import-progress-text"),G=L.querySelector("#import-progress-bar"),I=L.querySelector("#import-summary"),P=L.querySelector("#import-close-btn");P.addEventListener("click",()=>{L.remove(),U()});try{let X=await Oe(D);if(X.length===0)throw new Error("File kosong atau format salah");let Z=500,le=0,ne=0,F=0,M=X.length;V.textContent=`Ditemukan ${M} baris data. Memulai import...`;for(let E=0;E<M;E+=Z){let W=X.slice(E,E+Z);V.textContent=`Mengimport baris ${E+1} - ${Math.min(E+Z,M)} dari ${M}...`,G.style.width=`${Math.round(E/M*100)}%`;try{let j=await y.onImport(W);j?(le+=j.inserted||j.metrics?.inserted||W.length,ne+=j.skipped||j.metrics?.updated||0):le+=W.length}catch(j){console.error("Chunk import failed:",j),F+=W.length}}G.style.width="100%",V.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',I.style.display="block",I.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${M}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${le}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ne}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${F}</strong></div>
        `,F>0&&(I.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),P.style.display="block",v.value=""}catch(X){V.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${X.message}`,G.style.background="var(--danger)",G.style.width="100%",P.style.display="block",v.value=""}finally{v.disabled=!1}})}async function U(){C();let v=document.getElementById("table-container");if(!v)return;v.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let f=x==="client",D=f?1:S,L=f?fe:20,V=new URLSearchParams({page:D,limit:L,...Object.fromEntries(Object.entries(T).filter(([,F])=>F))}),G=await k(`${a}?${V}`);if(!G.ok){v.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${G.data?.error||"Error"}</p></div>`;return}let I=G.data?.data||G.data||[],P=G.data?.pagination,X=I.length,Z=I;if(f){I=h(I),Z=I;let F=I.length,M=20,E=Math.ceil(F/M);S>E&&E>0&&(S=E);let W=(S-1)*M,j=S*M;I=I.slice(W,j),P={page:S,limit:M,total:F,pages:E}}!1,u&&u(I);let le=Ge({columns:s,data:I,fullData:Z,onEdit:g?F=>me(F):null,actions:m.map(F=>({...F,handler:M=>F.handler(M,U)})),emptyText:`Tidak ada ${String(r||"").toLowerCase()}`,bulkSelect:_?{selectedIds:$,onToggle:C}:null});v.innerHTML="",v.appendChild(le);let ne=document.getElementById("pagination-container");if(ne&&(ne.innerHTML="",P&&P.pages>1)){let F=Qe({page:P.page,pages:P.pages,total:P.total,limit:P.limit,onPage:M=>{S=M,U()}});F&&ne.appendChild(F)}}function ve(v){let f=typeof o=="function"?o(v):o;return Me(f)}function me(v){let f=!!v,D=document.createElement("form");if(D.noValidate=!0,D.innerHTML=ve(v),f){let V=typeof o=="function"?o(v):o;We(D,v)}let{close:L}=de({title:f?`Edit ${r}`:`Tambah ${r}`,content:D,size:"lg",confirmText:f?"Simpan Perubahan":`Tambah ${r}`,onConfirm:async(V,G)=>{if(!D.reportValidity())return;let I=V.querySelector(".modal-confirm");I.disabled=!0,I.textContent="Menyimpan...";let P=ze(D),X=typeof o=="function"?o(v):o,Z=async M=>{for(let E of M)if(E.type==="row")await Z(E.fields);else if(E.type==="combobox"&&P[E.name]){let W=P[E.name],j=(E.options||[]).find(Y=>{let ie=String(typeof Y=="object"?Y.value:Y),at=String(typeof Y=="object"?Y.label:Y);return ie===W||at===W});if(j)P[E.name]=typeof j=="object"?j.value:j;else if(E.createApi){let Y={};Y[E.createApi.field]=W,E.createApi.extra&&Object.assign(Y,E.createApi.extra);let ie=await k(E.createApi.path,{method:"POST",body:JSON.stringify(Y)});if(ie.ok&&ie.data?.id)P[E.name]=ie.data.id;else if(ie.ok&&!ie.data?.id)P[E.name]=W;else throw new Error(`Gagal membuat master data: ${ie.data?.error||"Unknown error"}`)}}};try{await Z(X)}catch(M){Q(M.message),I.disabled=!1,I.textContent=f?"Simpan Perubahan":`Tambah ${r}`;return}p&&(P=await p(P,v));let le=f?"PUT":"POST",ne=f?`${a}/${v.id}`:a,F=await k(ne,{method:le,body:JSON.stringify(P)});F.ok?(ee(f?`${r} berhasil diperbarui.`:`${r} berhasil ditambahkan.`),G(),U()):(Q(F.data?.error||"Gagal menyimpan data."),I.disabled=!1,I.textContent=f?"Simpan Perubahan":`Tambah ${r}`)}})}function Je(v){Fe(`Hapus ${r} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let f=await k(`${a}/${v.id}`,{method:"DELETE"});f.ok?(ee(`${r} berhasil dihapus.`),U()):Q(f.data?.error||"Gagal menghapus.")},`Hapus ${r}`)}return U(),U}O();O();var we=null,Ve=null;async function xe(t=!1){if(we&&!t)return console.log("Employees Raw (Cache Hit)",we.slice(0,5)),we;let e=await k(`/api/employees?limit=${fe}&status=Aktif`);return we=(e.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",we.slice(0,5)),we}async function te(t=!1){let i=(await xe(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",i.slice(0,5)),i}async function H(t=!1){return Ve&&!t||(Ve=((await k("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),Ve}function J(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function ot(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function _e(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function dt(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function ce(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}K();function ct(t,e){let i=new Date,a=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}`;if(!(t.target_date||t.opening_date||"").startsWith(a))return!1;let o=String(t.status||"").toLowerCase();if(o!=="selesai"&&o!=="completed"&&o!=="done")return!1;let l=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?l.includes("inspeksi"):e==="gcdc"?l.includes("general cleaning")||l.includes("deep cleaning"):!1}O();K();function wt(t,e){let i=String(t.status||"").toLowerCase();return e==="active"?i==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&i==="aktif":!1}O();K();function pt(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let s=new Date(a);s.setDate(a.getDate()+30);let o=new Date(t.end_date);return o.setHours(0,0,0,0),o>=a&&o<=s}return!1}O();K();function xt(t,e){let i=String(t.status||"").toLowerCase();return e==="open"?i==="open":!1}O();function _t(t,e){let i=String(t.status||"").toLowerCase();return e==="pending"?i==="pending":!1}var he={};function $e(t){if(he[t]){try{he[t].destroy()}catch{}delete he[t]}}function ma(){Object.keys(he).forEach($e)}var ge=(t,e=0)=>{let i=Number(t);return isNaN(i)||t===null||t===void 0?e:i},Ce=(t,e="\u2014")=>{if(t==null||t==="")return e;let i=String(t).trim();return i===""||i==="[object Object]"?e:i};function Tt(t,e,i=900){if(!t)return;let a=Math.max(0,Math.round(ge(e)));if(a===0){t.textContent="0";return}let s=Date.now(),o=()=>{let l=Math.min((Date.now()-s)/i,1),n=1-Math.pow(1-l,3);t.textContent=Math.round(n*a).toLocaleString("id-ID"),l<1?requestAnimationFrame(o):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(o)}var ga={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},ba=t=>{let e=Ce(t,"\u2014");return`<span class="status-pill ${ga[e]||"pill-neutral"}">${e}</span>`};var pe={family:"Inter",size:11},be="#94A3B8",Te="#F1F5F9",ut=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],ha=()=>window.innerWidth<768;function Ye(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:ha()?"bottom":"top",labels:{font:pe,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:pe,titleFont:{...pe,weight:"700"}}},scales:{x:{grid:{color:Te},ticks:{font:pe,color:be,maxRotation:0}},y:{grid:{color:Te},ticks:{font:pe,color:be},beginAtZero:!0}},...t}}var ya=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),fa=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Ct(t=3){return Array(t).fill(0).map((e,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function re(t,e,i=8e3){try{let a=new AbortController,s=setTimeout(()=>a.abort(),i),o=await k(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(s),!o||!o.ok)return e;let l=o.data;return l?l.data!==void 0?l.data??e:l:e}catch{return e}}function va(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-jadwal"].forEach(a=>{let s=document.getElementById(a);s&&(s.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-jadwal"].forEach(a=>{let s=document.getElementById(a);if(s&&s.style.display==="none"){s.style.display="block";let o=s.parentElement;if(o&&!o.querySelector(".chart-empty")){let l=document.createElement("div");l.className="chart-empty",l.textContent="Belum ada data",s.style.display="none",o.appendChild(l)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Et({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Dt({}),["table-contracts","table-issues"].forEach(a=>{let s=document.getElementById(a);s&&s.querySelector(".skeleton")&&(s.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function $t(t){ma(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">


      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${ya()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${fa()}</div>

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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${Ct(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${Ct(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>mt(t)),document.getElementById("filter-jadwal-year")?.addEventListener("change",async i=>{let a=i.target.value,s=document.getElementById("jadwal-year-label");s&&(s.textContent=a);let o=document.getElementById("skel-jadwal"),l=document.getElementById("chart-jadwal");o&&(o.style.display="block",o.style.position="absolute"),l&&(l.style.display="none");let n=await re(`/api/dashboard/schedule-chart?year=${a}`,{},8e3);try{It(n)}catch(r){console.warn("ScheduleChart render:",r),ue("skel-jadwal","chart-jadwal")}}),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let a=i.target.value,s=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",o=document.getElementById("skel-insp"),l=document.getElementById("chart-insp");o&&(o.style.display="block",o.style.position="absolute"),l&&(l.style.display="none");let n=await re(s,{},8e3);try{Pt(n)}catch(r){console.warn("InspBar render:",r),ue("skel-insp","chart-insp")}}),t._skelTimeout=setTimeout(()=>va(),5e3),await mt(t),t._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?mt(t):clearInterval(t._dashRefresh)},6e4)}async function mt(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,i,a,s,o,l,n,r,d,g,c,p,u,h]=await Promise.all([re("/api/dashboard/kpi",{},8e3),re("/api/dashboard/issues-trend",{},8e3),re("/api/dashboard/issues-summary",{},8e3),re("/api/dashboard/stats",{},8e3),re("/api/dashboard/calendar",[],8e3),re("/api/schedule?limit=10000",{data:[]},8e3),re("/api/employees?limit=10000",{data:[]},8e3),re("/api/contracts?limit=10000",{data:[]},8e3),re("/api/issues?limit=10000",{data:[]},8e3),re("/api/one-on-one?limit=10000",{data:[]},8e3),re("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),re(`/api/dashboard/schedule-chart?year=${document.getElementById("filter-jadwal-year")?.value||new Date().getFullYear()}`,{},8e3),re("/api/relievers?limit=10000",{data:[]},8e3),re("/api/reports/fogging?limit=10000",{data:[]},8e3)]),m=document.getElementById("filter-insp-month"),b=m?m.value:"",y=b?`/api/dashboard/inspection-bar?month=${b}`:"/api/dashboard/inspection-bar",_=await re(y,{},8e3);if(e){let x=Array.isArray(l?.data)?l.data:Array.isArray(l)?l:[];window.dashboardSchedules=x;let w=Array.isArray(n?.data)?n.data:Array.isArray(n)?n:[],S=Array.isArray(r?.data)?r.data:Array.isArray(r)?r:[],T=Array.isArray(d?.data)?d.data:Array.isArray(d)?d:[],$=Array.isArray(g?.data)?g.data:Array.isArray(g)?g:[],C=Array.isArray(u?.data)?u.data:Array.isArray(u)?u:[];window.dashboardRelievers=C;let B=Array.isArray(h?.data)?h.data:Array.isArray(h)?h:[];window.dashboardFogging=B,e.employees&&(e.employees.current=w.filter(R=>wt(R,"active")).length),e.contracts&&(e.contracts.current=S.filter(R=>pt(R,"active")).length),e.expiring30&&(e.expiring30={current:S.filter(R=>pt(R,"expiring30")).length}),e.issues&&(e.issues.current=T.filter(R=>xt(R,"open")).length),e.one_on_one&&(e.one_on_one.current=$.filter(R=>_t(R,"pending")).length),e.inspection_month&&(e.inspection_month.current=x.filter(R=>ct(R,"inspeksi")).length),e.cleaning_month&&(e.cleaning_month.current=x.filter(R=>ct(R,"gcdc")).length)}try{Et(e)}catch(x){console.warn("KPI render:",x)}try{Dt(e)}catch(x){console.warn("MiniStats render:",x)}try{It(p)}catch(x){console.warn("ScheduleChart render:",x),ue("skel-jadwal","chart-jadwal")}try{ka(Array.isArray(a?.by_category)?a.by_category:[])}catch(x){console.warn("Donut render:",x),ue("skel-donut","chart-donut")}try{Sa(i)}catch(x){console.warn("Trend render:",x),ue("skel-trend","chart-trend")}try{Pt(_)}catch(x){console.warn("InspBar render:",x),ue("skel-insp","chart-insp")}try{let x=Array.isArray(s)?s:Array.isArray(s?.recent_issues)?s.recent_issues:[];xa(x)}catch(x){console.warn("IssuesTable render:",x)}try{let x=Array.isArray(s?.expiring_contracts)?s.expiring_contracts:[];wa(c)}catch(x){console.warn("ContractsTable render:",x)}try{_a(Array.isArray(o)?o:[])}catch(x){console.warn("Agenda render:",x)}try{Ca()}catch(x){console.warn("Quick Actions render:",x)}}function Et(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let i=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=i.map(a=>{let s=ge(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${s}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${s}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{Tt(a,parseInt(a.dataset.target)||0)})}function Dt(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let i=`Q${Math.ceil((new Date().getMonth()+1)/3)}`,a=new Date().getFullYear(),s=String(new Date().getMonth()+1).padStart(2,"0"),o=`${a}-${s}`,l=c=>`
    <select id="${c}" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.preventDefault(); event.stopPropagation();">
      <option value="${a}-01" ${o===`${a}-01`?"selected":""}>Jan</option>
      <option value="${a}-02" ${o===`${a}-02`?"selected":""}>Feb</option>
      <option value="${a}-03" ${o===`${a}-03`?"selected":""}>Mar</option>
      <option value="${a}-04" ${o===`${a}-04`?"selected":""}>Apr</option>
      <option value="${a}-05" ${o===`${a}-05`?"selected":""}>Mei</option>
      <option value="${a}-06" ${o===`${a}-06`?"selected":""}>Jun</option>
      <option value="${a}-07" ${o===`${a}-07`?"selected":""}>Jul</option>
      <option value="${a}-08" ${o===`${a}-08`?"selected":""}>Agu</option>
      <option value="${a}-09" ${o===`${a}-09`?"selected":""}>Sep</option>
      <option value="${a}-10" ${o===`${a}-10`?"selected":""}>Okt</option>
      <option value="${a}-11" ${o===`${a}-11`?"selected":""}>Nov</option>
      <option value="${a}-12" ${o===`${a}-12`?"selected":""}>Des</option>
    </select>
  `,n=[{id:"mini-jadwal",icon:"\u{1F4C5}",label:"Jadwal",dropdown:`
        <select id="dash-jadwal-period" style="padding:0; font-size:1rem; line-height:1; border-radius:4px; background:transparent; border:none; color:var(--text-1); font-weight:700; cursor:pointer; outline:none;" onclick="event.preventDefault(); event.stopPropagation();">
          <option value="Q1" ${i==="Q1"?"selected":""}>Q1</option>
          <option value="Q2" ${i==="Q2"?"selected":""}>Q2</option>
          <option value="Q3" ${i==="Q3"?"selected":""}>Q3</option>
          <option value="Q4" ${i==="Q4"?"selected":""}>Q4</option>
        </select>
      `,val:t.schedule?.current,href:`#/timeline?dash_filter=period_${i.toLowerCase()}`,color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{id:"mini-reliefer",icon:"\u{1F504}",label:"Report Reliefer",dropdown:l("dash-reliefer-month"),val:t.reliever_completed?.current,href:`#/relievers?dash_filter=reliever&month=${o}`,color:"mini-teal"},{id:"mini-inspeksi",icon:"\u{1F50D}",label:"Report Inspeksi",dropdown:l("dash-inspeksi-month"),val:t.inspection_month?.current,href:`#/timeline?dash_filter=inspeksi&month=${o}`,color:"mini-blue"},{id:"mini-gcdc",icon:"\u{1F9F9}",label:"Report GCDC",dropdown:l("dash-gcdc-month"),val:t.cleaning_month?.current,href:`#/timeline?dash_filter=gcdc&month=${o}`,color:"mini-green"},{id:"mini-fogging",icon:"\u{1F4A8}",label:"Report Fogging",dropdown:l("dash-fogging-month"),val:t.fogging_month?.current,href:`#/reports/fogging?dash_filter=fogging&month=${o}`,color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=n.map(c=>`
    <a href="${c.href}" class="mini-stat ${c.color}" style="text-decoration:none" id="${c.id||""}">
      <div class="mini-stat-icon">${c.icon}</div>
      <div class="mini-stat-body" style="flex:1; min-width:0; overflow:visible;">
        <div style="display:flex; align-items:baseline; gap:3px;">
          <div class="mini-stat-value" data-target="${ge(c.val)}">0</div>
          ${c.dropdown?c.dropdown:""}
        </div>
        <div class="mini-stat-text">${c.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(c=>Tt(c,parseInt(c.dataset.target)||0,700));let r=document.getElementById("dash-jadwal-period");r&&r.addEventListener("change",c=>{let p=c.target.value,u=(window.dashboardSchedules||[]).filter(b=>b.period===p).length,h=document.querySelector("#mini-jadwal .mini-stat-value");h&&(h.dataset.target=u,h.textContent=u);let m=document.getElementById("mini-jadwal");m&&(m.href=`#/timeline?dash_filter=period_${p.toLowerCase()}`)});let d=(c,p,u,h,m)=>{let b=document.getElementById(c);b&&b.addEventListener("change",y=>{let _=y.target.value,x=(u||[]).filter(T=>h(T,_)).length,w=document.querySelector(`#${p} .mini-stat-value`);w&&(w.dataset.target=x,w.textContent=x);let S=document.getElementById(p);S&&(S.href=`${m}&month=${_}`)})},g=c=>{let p=String(c.status||"").toLowerCase();return p==="done"||p==="selesai"||p==="completed"};d("dash-reliefer-month","mini-reliefer",window.dashboardRelievers,(c,p)=>window.parseFlexibleDate(c.backup_date).startsWith(p)&&g(c),"#/relievers?dash_filter=reliever"),d("dash-inspeksi-month","mini-inspeksi",window.dashboardSchedules,(c,p)=>c.activity_type==="Inspeksi Hygiene"&&g(c)&&window.parseFlexibleDate(c.opening_date||c.target_date).startsWith(p),"#/timeline?dash_filter=inspeksi"),d("dash-gcdc-month","mini-gcdc",window.dashboardSchedules,(c,p)=>(c.activity_type==="General Cleaning"||c.activity_type==="Deep Cleaning")&&g(c)&&window.parseFlexibleDate(c.opening_date||c.target_date).startsWith(p),"#/timeline?dash_filter=gcdc"),d("dash-fogging-month","mini-fogging",window.dashboardFogging,(c,p)=>g(c)&&window.parseFlexibleDate(c.activity_date).startsWith(p),"#/reports/fogging?dash_filter=fogging")}function ka(t){ue("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!e||!i)return;$e("donut");let a=(t||[]).filter(r=>ge(r.count)>0);if(!a.length){Ke(e,"Belum ada data permasalahan");return}let s=a.map(r=>`${Ce(r.category,"Lainnya")}`),o=a.map(r=>ge(r.count)),l=o.reduce((r,d)=>r+d,0);i.innerHTML=a.map((r,d)=>{let g=ut[d%ut.length],c=l>0?Math.round(r.count/l*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${g}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${r.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${c}%)</span></div>
          <div class="donut-legend-label">${s[d]}</div>
        </div>
      </div>
    `}).join("");let n={id:"centerText",beforeDraw:function(r){let d=r.width,g=r.height,c=r.ctx;c.restore();let p=(g/80).toFixed(2);c.font="bold "+p+"em Inter",c.textBaseline="middle",c.fillStyle="#1E293B";let u=l.toString(),h=Math.round((d-c.measureText(u).width)/2),m=g/2;c.fillText(u,h,m-4),c.font="600 "+(p*.35).toFixed(2)+"em Inter",c.fillStyle="#64748B";let b="Total",y=Math.round((d-c.measureText(b).width)/2);c.fillText(b,y,m+10),c.save()}};he.donut=new Chart(e,{type:"doughnut",data:{labels:s,datasets:[{data:o,backgroundColor:ut,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:pe,titleFont:{...pe,weight:"700"},callbacks:{label:r=>` ${r.label}: ${r.parsed} kasus`}}},cutout:"75%"},plugins:[n]})}function Sa(t){ue("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;$e("trend"),t=t||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],a=(t.labels||[]).map(l=>{if(!l||typeof l!="string")return"";try{let[n,r]=l.split("-");return(i[Number(r)-1]||r)+" "+String(n).slice(-2)}catch{return l}}),s=(t.open||[]).map(l=>ge(l)),o=(t.closed||[]).map(l=>ge(l));if(!a.length){Ke(e,"Belum ada data trend");return}he.trend=new Chart(e,{type:"line",data:{labels:a,datasets:[{label:"Open",data:s,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:o,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.1)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:Ye({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:be,maxRotation:0,autoSkip:!1}},y:{grid:{color:Te},ticks:{font:{family:"Inter",size:9},color:be},beginAtZero:!0}}})})}function It(t){ue("skel-jadwal","chart-jadwal");let e=document.getElementById("chart-jadwal");if(!e)return;$e("jadwal"),t=t||{};let i=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"];if(!Object.values(t).some(r=>Array.isArray(r)&&r.some(d=>d>0))){Ke(e,"Belum ada data jadwal");return}let s=t["Inspeksi Hygiene"]||Array(12).fill(0),o=t["General Cleaning"]||Array(12).fill(0),l=t["Deep Cleaning"]||Array(12).fill(0),n=t.Fogging||Array(12).fill(0);he.jadwal=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Inspeksi",data:s,backgroundColor:"#3B82F6"},{label:"General Cleaning",data:o,backgroundColor:"#10B981"},{label:"Deep Cleaning",data:l,backgroundColor:"#F59E0B"},{label:"Fogging",data:n,backgroundColor:"#EF4444"}]},options:Ye({plugins:{legend:{display:!1}},datasets:{bar:{barPercentage:.85,categoryPercentage:.9}},scales:{x:{stacked:!0,grid:{display:!1},ticks:{font:{family:"Inter",size:9},color:be,maxRotation:0,autoSkip:!1}},y:{stacked:!0,grid:{color:Te},ticks:{font:{family:"Inter",size:9},color:be},min:0}}})})}function Pt(t){ue("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;$e("inspBar"),t=t||{};let i=t.labels||[],a=(t.fc||[]).map(o=>ge(o)),s=(t.spv||[]).map(o=>ge(o));if(!i.length){Ke(e,"Belum ada data inspeksi");return}he.inspBar=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:s,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:Ye({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:pe,color:be,maxRotation:45,minRotation:30}},y:{grid:{color:Te},ticks:{font:pe,color:be},min:0,max:100}}})})}function wa(t){ue("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;$e("contractMiniBar"),t=t||{};let i={"06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(l=>{let n=l.split("-")[1];return i[n]||l}),s=(t.data||[]).map(l=>ge(l));if(!a.length){Ke(e,"Belum ada data");return}let o=e.getContext("2d");he.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:s,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:Ye({onClick:(l,n)=>{if(n&&n.length>0){let r=n[0].index,d=(t.labels||[])[r];d&&(window.location.hash="#/contracts?month_expiry="+d)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:pe,color:be,maxRotation:0,autoSkip:!1}},y:{grid:{color:Te,borderDash:[4,4],drawBorder:!1},ticks:{font:pe,color:be,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function xa(t){let e=document.getElementById("table-issues");if(!e)return;let i=(t||[]).slice(0,8);if(!i.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${i.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${ba(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Ce(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Ce(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function _a(t){let e=document.getElementById("widget-agenda");if(!e)return;let i=new Date,a=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,o=(t||[]).filter(l=>(l.event_date||"").startsWith(a)).slice(0,10);if(!o.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${o.map(l=>{let n="#3B82F6",r="#EFF6FF",d="Agenda",g=(l.title||"").toLowerCase();return g.includes("inspeksi")?(n="#10B981",r="#ECFDF5",d="Inspeksi"):g.includes("cleaning")||g.includes("gcdc")?(n="#3B82F6",r="#EFF6FF",d="Cleaning"):g.includes("reliefer")?(n="#F59E0B",r="#FFFBEB",d="Reliefer"):g.includes("fogging")&&(n="#8B5CF6",r="#F5F3FF",d="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(l.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${n};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Ce(l.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Ce(l.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${r};color:${n}">${d}</div>
        </div>
      `}).join("")}
    </div>
  `}function Ca(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function ue(t,e){let i=document.getElementById(t),a=document.getElementById(e);if(i&&(i.style.display="none",i.style.position=""),a){a.style.display="block";let s=a.parentElement;if(s){let o=s.querySelector(".chart-empty");o&&o.remove()}}}function Ke(t,e="Belum ada data"){if(!t)return;t.style.display="none";let i=t.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let s=document.createElement("div");s.className="chart-empty",s.textContent=e,i.appendChild(s)}}O();async function Lt(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),i=document.getElementById("login-error"),a=document.getElementById("login-btn"),s=document.getElementById("toggle-password"),o=document.getElementById("login-password");s?.addEventListener("click",()=>{let l=o.type==="text";o.type=l?"password":"text",s.style.color=l?"":"var(--primary)"}),e?.addEventListener("submit",async l=>{l.preventDefault(),i.style.display="none";let n=e.username.value.trim(),r=e.password.value;if(!n||!r){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let d=await k("/api/auth/login",{method:"POST",body:JSON.stringify({username:n,password:r})});d.ok&&d.data.success?(rt(d.data.data.token),Ae(d.data.data.user),ee("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=d.data.error||"Username atau password salah.",i.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}O();K();async function Ta(){return await H()}function $a(t,e){let i=String(t.status||"").toLowerCase();return e==="active"?i==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&i==="aktif":!1}async function Bt(t,e){let i=await Ta(),a=e?e.get("dash_filter"):null;N({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:s=>a?s.filter(o=>$a(o,a)):s,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:s=>_e(s)},{key:"phone",label:"No. HP",render:s=>s?`<a href="tel:${s}">${s}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:s=>window.formatDate(s)},{key:"status",label:"Status",render:s=>J(s)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:s=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:s?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:s?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:i,value:s?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:s?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:s?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:s?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:s?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let s=await k(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(s.ok){let o=s.data.data.map(l=>({"Nama Lengkap":l.full_name,Cabang:l.branch_name||"",Divisi:l.division||"","No. HP":l.phone||"","Tgl Masuk":l.join_date||"",Status:l.status||""}));A(o,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{A([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async s=>{let o=r=>{if(!r)return null;let d=String(r||"").toLowerCase(),g=i.find(c=>String(c.label||"").toLowerCase()===d);return g?g.value:null},l=s.map(r=>({full_name:String(r["Nama Lengkap"]||"").trim(),branch_id:o(String(r.Cabang||"").trim()),division:String(r.Divisi||"").trim()||"FACILITY CARE",phone:String(r["No. HP"]||"").trim(),join_date:String(r["Tgl Masuk"]||"").trim(),status:String(r.Status||"").trim(),notes:String(r.Catatan||"").trim()})).filter(r=>r.full_name),n=await k("/api/import/employees",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}O();K();var bt=[],At=[];async function Ea(){bt=await H(),At=await xe()}var gt=async t=>{let e=[],i=1;for(;;){let s=await(await Promise.resolve().then(()=>(O(),ke))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${i}`);if(!s.ok)break;let o=s.data?.data||s.data||[],l=Array.isArray(o)?o:[];if(e=e.concat(l),l.length<100||s.data?.pagination&&i>=s.data.pagination.pages)break;i++}return e};function Da(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let s=new Date(a);s.setDate(a.getDate()+30);let o=new Date(t.end_date);return o.setHours(0,0,0,0),o>=a&&o<=s}return!1}async function Nt(t,e){await Ea();let i=e?e.get("dash_filter"):null;N({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",onDataLoaded:a=>i?a.filter(s=>Da(s,i)):a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>_e(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,s)=>s.end_date&&String(s.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':ot(a)},{key:"status",label:"Status",render:a=>J(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:bt},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"month_expiry",label:"Bulan Habis",options:[{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"},{value:"2027-01",label:"Jan 2027"},{value:"2027-02",label:"Feb 2027"},{value:"2027-03",label:"Mar 2027"},{value:"2027-04",label:"Apr 2027"},{value:"2027-05",label:"Mei 2027"},{value:"2027-06",label:"Jun 2027"},{value:"2027-07",label:"Jul 2027"}]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Detail Kontrak",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[o,l]=await Promise.all([gt("/api/employees?status=Aktif"),gt("/api/contracts")]);if(o.length>0){let n=l.filter(u=>u.status==="Aktif"&&(u.days_remaining==null||u.days_remaining>=0||String(u.end_date).startsWith("2099"))),r=new Set(n.map(u=>u.employee_id)),d=o.filter(u=>!r.has(u.id)),g=[],c=[];d.forEach(u=>{let h=l.filter(m=>m.employee_id===u.id);h.length===0?g.push(u):c.push({emp:u,lastContract:h[0]})});let p=`<p style="margin-bottom:12px">Data yang terbaca: <b>${o.length}</b> Karyawan Aktif, dan <b>${n.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:6px">Terdapat <b>${g.length}</b> karyawan aktif tanpa kontrak.</p>
              <p style="margin-bottom:12px">Terdapat <b>${c.length}</b> karyawan aktif yang memiliki masa kontrak Expired.</p>
              <ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;g.forEach(u=>{p+=`<li style="margin-bottom:8px"><b>${u.full_name}</b> <br><span style="font-size:0.85em;color:#F59E0B">Cabang: ${u.branch_name||"-"} | Belum pernah di-input kontrak</span></li>`}),c.forEach(u=>{let h=u.emp,m=u.lastContract,b=m.status==="Aktif"&&m.days_remaining<0?"Aktif (Masa Habis)":m.status;p+=`<li style="margin-bottom:8px"><b>${h.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${h.branch_name||"-"} | Status Terakhir: <b style="color:#EF4444">${b}</b>, Tgl Berakhir: ${window.formatDate(m.end_date)}</span></li>`}),p+="</ul>",Promise.resolve().then(()=>(ye(),st)).then(u=>u.createModal({title:"Detail Karyawan Tanpa Kontrak Aktif",content:p,cancelText:"Tutup"}))}}catch(o){console.error(o)}a.innerHTML="\u{1F50D} Detail Kontrak",a.disabled=!1};let s=document.querySelector(".page-actions");s&&s.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:At,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:bt,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await k(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let s=a.data.data.map(o=>({"Nama Lengkap":o.employee_name,Cabang:o.branch_name||"","Div / Bagian":o.division||"","Tanggal Mulai":o.start_date||"","Tanggal Selesai":o.end_date&&String(o.end_date).startsWith("2099")?"":o.end_date||"","Sisa Kontrak":o.end_date&&String(o.end_date).startsWith("2099")?"Tetap":o.days_remaining!==null&&o.days_remaining!==void 0?`${o.days_remaining} Hari`:"",Status:o.status||""}));A(s,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{A([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[s,o]=await Promise.all([k("/api/branches?limit=10000"),gt("/api/employees")]),l=s.data?.data||[],n=o||[];console.log(`Total employee yang berhasil dimuat dari database : ${n.length}`),n.length>0&&(console.log("Contoh 5 employee pertama:"),n.slice(0,5).forEach((m,b)=>{console.log(`${b+1}. ID: ${m.id}, Name: ${m.full_name}, Status: ${m.status}`)}));let r=m=>{if(!m)return null;let b=String(m||"").replace(/\s+/g," ").toLowerCase().trim(),y=l.find(_=>String(_.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===b||String(_.code||"").replace(/\s+/g," ").toLowerCase().trim()===b||String(_.name||"").replace(/\s+/g," ").toLowerCase().trim()===b);return y?y.id:null},d=(m,b)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${b}`),console.log(`Nama dari Excel : "${m}"`),!m)return console.log("Alasan gagal mapping : Nama kosong"),null;let y=String(m||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${y}"`),console.log(`Jumlah employee di database : ${n.length}`);let _=n.find(x=>String(x.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===y);return _?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${_.id}`),_.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},g=m=>{if(!m)return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let b=String(m).trim();if(/^\d{4,5}(\.\d+)?$/.test(b)){let _=Math.floor(Number(b));if(_>2e4&&_<99999){let x=new Date(Date.UTC(1899,11,30)+_*864e5);return isNaN(x.getTime())?"":x.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(b))return b.slice(0,10);let y=b.split(/[\/\-\.]/);if(y.length===3){let[_,x,w]=y.map(S=>S.trim());if(_.length===4&&x.length<=2&&w.length<=2)return`${_}-${x.padStart(2,"0")}-${w.padStart(2,"0")}`;if(w.length===4&&x.length<=2&&_.length<=2)return`${w}-${x.padStart(2,"0")}-${_.padStart(2,"0")}`}return b},c=a.map((m,b)=>{let y=b+2,_=String(m["Nama Lengkap"]||"").trim(),x=m["Tanggal Mulai"],w=g(x);if(!w){let $=a.__worksheet,C=a.__headers||[],B=C.indexOf("Tanggal Mulai"),R="N/A",U="N/A",ve="N/A";if(B!==-1&&$&&window.XLSX){let Je=window.XLSX.utils.encode_cell({c:B,r:y-1});ve=Je;let v=$[Je];v?(R=v.t||"undefined",U=v.w||"undefined"):R="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let me="Unknown";x==null||x===""?me="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":x instanceof Date&&isNaN(x.getTime())?me="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":me="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${y}`),console.log(`Employee Name : ${_}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${B})`),console.log(`Raw Cell Value : "${x}"`),console.log(`JavaScript Type : ${typeof x}`),console.log(`SheetJS Cell Type : ${R}`),console.log(`SheetJS Formatted Value : "${U}"`),console.log(`Value After Trim : "${String(x||"").trim()}"`),console.log(`Value After Date Parser : "${w}"`),console.log(`Is Empty : ${!x}`),console.log(`Is Invalid Date : ${x instanceof Date?isNaN(x.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${me}`),console.log(`Workbook Sheet : ${$?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${ve}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(m,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(C)),console.log(`==========================
`)}let S=d(_,y),T=null;return S||(T="Karyawan tidak ditemukan di Database"),{isValid:!!S,invalidReason:T,rowNum:y,data:{employee_id:S,branch_id:r(String(m.Cabang||"").trim()),division:String(m["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:w,end_date:g(m["Tanggal Selesai"])||"2099-12-31",status:String(m.Status||"").trim(),_rawName:_}}}),p=[],u=[];if(c.forEach(m=>{m.isValid?p.push(m.data):u.push({rowNum:m.rowNum,name:m.data._rawName,reason:m.invalidReason})}),console.log(`Split Validation - Valid: ${p.length}, Invalid: ${u.length}`),p.length===0)return{inserted:0,skipped:a.length,failed:a.length};let h=await k("/api/import/contracts",{method:"POST",body:JSON.stringify({rows:p,onDuplicate:"update"})});if(!h.ok)throw new Error(h.data?.error||"Import gagal");return h.data}}})}O();K();var ht=[],He=[];function Ia(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let i of e)if(t.some(a=>a.period===i))return i;return"Q3"}async function Ft(t,e){ht=await H();let i=await te();He=["Berlin Ariansyah","Ade Surahman"];let a=u=>u&&!He.find(h=>String(typeof h=="object"?h.value:h).toLowerCase()===String(u).toLowerCase())?[...He,u]:He,s=await k(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),o=u=>{if(!u||u==="-"||String(u).trim()==="")return"";let h=String(u).split("-");return h.length===3&&h[0].length===4?`${h[2]}-${h[1]}-${h[0]}`:u},l=s.data?.data||[],n=Ia(l),r=e?e.get("dash_filter"):null,d=new Date,g=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`,c={},p=e&&e.get("month")?e.get("month"):null;r==="inspeksi"?c={status:"Done",activity_type:"Inspeksi Hygiene",month:p}:r==="gcdc"?c={status:"Done",activity_type:"GCDC",month:p}:r&&r.startsWith("period_")&&(c={period:r.replace("period_","").toUpperCase()}),N({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",defaultFilters:c,onDataLoaded:u=>u.sort((h,m)=>{let b=h.opening_date?new Date(h.opening_date).getTime():0;return(m.opening_date?new Date(m.opening_date).getTime():0)-b}),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:u=>dt(u)},{key:"period",label:"Periode",render:u=>ce(u)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:u=>o(u)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:u=>o(u)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:u=>o(u)},{key:"status",label:"Status",render:u=>J(u)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:ht},{type:"select",name:"activity_type",label:"Kegiatan",options:[{value:"Inspeksi Hygiene",label:"Inspeksi Hygiene"},{value:"General Cleaning",label:"General Cleaning"},{value:"Deep Cleaning",label:"Deep Cleaning"},{value:"Fogging",label:"Fogging"},{value:"GCDC",label:"GCDC (GC & DC)"}]},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:He}],formFields:u=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:ht,value:u?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:u?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:u?.period},{name:"pic",label:"PIC",type:"combobox",options:a(u?.pic),value:u?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:u?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:u?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:u?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:u?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:u?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let u=await k(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(u.ok){let h=u.data.data.map(m=>({Cabang:m.branch_name||"",Kegiatan:m.activity_type||"",Periode:m.period||"",PIC:m.pic||"","Tgl Opening":m.opening_date||"","Tgl Target":m.target_date||"","Tgl Selesai":m.completion_date||"",Status:m.status||""}));A(h,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{A([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async u=>{let m=(await k("/api/branches?all=1")).data?.data||[],b=w=>{if(!w)return null;let S=String(w||"").toLowerCase(),T=m.find($=>String($.full_name||"").toLowerCase()===S||String($.code||"").toLowerCase()===S||String($.name||"").toLowerCase()===S);return T?T.id:null},y=w=>{if(w==null||w==="")return"";if(w instanceof Date&&!isNaN(w.getTime()))return w.toISOString().slice(0,10);let S=String(w).trim();if(S===""||S==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(S))return S.slice(0,10);if(/^\d{4,5}$/.test(S)){let $=Number(S);if($>2e4&&$<99999){let C=new Date(Date.UTC(1899,11,30)+$*864e5);return isNaN(C.getTime())?"":C.toISOString().slice(0,10)}}let T=S.split(/[\/\-\.]/);if(T.length===3){let[$,C,B]=T.map(R=>R.trim());if($.length===4&&C.length<=2&&B.length<=2)return`${$}-${C.padStart(2,"0")}-${B.padStart(2,"0")}`;if(B.length===4&&C.length<=2&&$.length<=2)return`${B}-${C.padStart(2,"0")}-${$.padStart(2,"0")}`}return S},_=u.map(w=>({branch_id:b(String(w.Cabang||"").trim()),activity_type:String(w.Kegiatan||"").trim(),period:String(w.Periode||"").trim(),pic:String(w.PIC||w.Pic||"").trim(),opening_date:y(w["Tgl Opening"]||w["Tanggal Opening"]||w["Tgl Openir"]),target_date:y(w["Tgl Target"]||w["Tanggal Target"]),completion_date:y(w["Tgl Selesai"]||w["Tanggal Selesai"]),status:String(w.Status||"").trim(),notes:String(w.Catatan||w.Keterangan||"").trim()})).filter(w=>w.activity_type&&w.period),x=await k("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:_,onDuplicate:"update"})});if(!x.ok)throw new Error(x.data?.error||"Import gagal");return x.data}}})}O();K();var yt=[],Xe=[];function Pa(t,e){let i=String(t.status||"").toLowerCase();return e==="open"?i==="open":!1}async function Mt(t,e){let i=e?e.get("dash_filter"):null;yt=await H(),Xe=await te();let a=l=>l&&!Xe.find(n=>n.value===l)?[...Xe,{value:l,label:l}]:Xe,s=new Date().getFullYear();N({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:l=>i?l.filter(n=>Pa(n,i)):l,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:l=>`<span class="badge badge-secondary">${l}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:l=>`<span title="${l}">${l?.length>50?l.slice(0,50)+"\u2026":l}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:l=>`<span title="${l||""}">${l?.length>40?l.slice(0,40)+"\u2026":l||"-"}</span>`},{key:"status",label:"Status",render:l=>J(l)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:l=>window.formatDate(l)},{key:"day_count",label:"Hari",render:l=>l??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:yt},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:["2025","2026","2027","2028","2029","2030"]}],formFields:l=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:l?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:yt,value:l?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:l?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:l?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:l?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(l?.employee_name),value:l?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(l?.fc_specialist),value:l?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:l?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:l?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:l?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let l=await k(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(l.ok){let n=l.data.data.map(r=>({Tanggal:r.report_date||"",Cabang:r.branch_name||"",Kategori:r.category||"",Sumber:r.source||"",Keluhan:r.complaint||"","Nama FC":r.employee_name||"","FC Spesialis":r.fc_specialist||"",Solusi:r.solution||"","Tgl Selesai":r.completion_date||"",Status:r.status||""}));A(n,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{A([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async l=>{let r=(await k("/api/branches?all=1")).data?.data||[],d=p=>{if(!p)return null;let u=String(p||"").toLowerCase(),h=r.find(m=>String(m.full_name||"").toLowerCase()===u||String(m.code||"").toLowerCase()===u||String(m.name||"").toLowerCase()===u);return h?h.id:null},g=l.map(p=>({branch_id:d(String(p.Cabang||"").trim()),report_date:String(p.Tanggal||"").trim(),category:String(p.Kategori||"").trim(),source:String(p.Sumber||"").trim(),complaint:String(p.Keluhan||"").trim(),employee_name:String(p["Nama FC"]||"").trim(),fc_specialist:String(p["FC Spesialis"]||"").trim(),solution:String(p.Solusi||"").trim(),completion_date:String(p["Tgl Selesai"]||"").trim(),status:String(p.Status||"").trim()})).filter(p=>p.report_date&&p.complaint&&p.category),c=await k("/api/import/issues",{method:"POST",body:JSON.stringify({rows:g,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}}})}O();var Ee=[];function La(t,e){let i=String(t.status||"").toLowerCase();return e==="pending"?i==="pending":!1}async function Ot(t,e){let i=e?e.get("dash_filter"):null;Ee=await H();let a=await te(),s=["Ade","Berlin"],o=n=>n&&!a.find(r=>r.value===n)?[...a,{value:n,label:n}]:a,l=n=>n&&!s.find(r=>(typeof r=="object"?r.value:r)===n)?[...s,n]:s;N({container:t,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:n=>i?n.filter(r=>La(r,i)):n,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:n=>`<span title="${n||""}">${n?.length>50?n.slice(0,50)+"\u2026":n||"-"}</span>`},{key:"solution",label:"Solusi",render:n=>`<span title="${n||""}">${n?.length>40?n.slice(0,40)+"\u2026":n||"-"}</span>`},{key:"status",label:"Status",render:n=>J(n)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"combobox",name:"branch_id",label:"Cabang",options:Ee},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async n=>{let r=new URLSearchParams(n||{}).toString(),d=await k(`/api/one-on-one?limit=10000&${r}`);if(d.ok){let g=d.data.data.map(p=>({Tanggal:p.meeting_date||"",Cabang:p.branch_name||"","Nama Karyawan":p.employee_name||"",PIC:p.pic||"",Masalah:p.problem||"",Solusi:p.solution||"",Status:p.status||"","Tgl Selesai":p.completion_date||"",Dokumen:p.document_link||""})),{downloadExcel:c}=await Promise.resolve().then(()=>(K(),oe));c(g,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(K(),oe));r(n,"Template_Import_OneOnOne")},onImport:async n=>{let r=p=>{if(!p)return null;let u=String(p||"").toLowerCase(),h=Ee.find(m=>String(m.label||"").toLowerCase()===u);return h?h.value:null},d=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let u=String(p).trim();if(/^\d{4,5}$/.test(u)){let m=Number(u);if(m>2e4&&m<99999){let b=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);let h=u.split(/[\/\-\.]/);if(h.length===3){let[m,b,y]=h.map(_=>_.trim());if(m.length===4&&b.length<=2&&y.length<=2)return`${m}-${b.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&b.length<=2&&m.length<=2)return`${y}-${b.padStart(2,"0")}-${m.padStart(2,"0")}`}return u},g=n.map(p=>({meeting_date:d(p.Tanggal),employee_name:String(p["Nama Karyawan"]||"").trim(),branch_id:r(String(p.Cabang||"").trim()),pic:String(p.PIC||"").trim(),problem:String(p.Masalah||"").trim(),solution:String(p.Solusi||"").trim(),status:String(p.Status||"").trim(),completion_date:d(p["Tgl Selesai"]),document_link:String(p.Dokumen||"").trim()})).filter(p=>p.meeting_date&&p.employee_name&&p.branch_id),c=await k("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:g,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},formFields:n=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:n?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:n?.branch_id&&!Ee.find(r=>r.value==n.branch_id)?[...Ee,{value:n.branch_id,label:n.branch_name||n.branch_id}]:Ee,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"combobox",required:!0,options:o(n?.employee_name),value:n?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:l(n?.pic),createApi:{path:"/api/pic",field:"name"},value:n?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:n?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:n?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:n?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:n?.document_link}]})}O();async function Rt(t){let e=await H(),i=await te(),a=["Ade","Berlin"],s=n=>n&&!i.find(r=>r.value===n)?[...i,{value:n,label:n}]:i,o=n=>n&&!a.find(r=>(typeof r=="object"?r.value:r)===n)?[...a,n]:a,l=Array.from({length:5},(n,r)=>String(new Date().getFullYear()-r));N({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:n=>window.formatDate(n)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:n=>{try{let r=JSON.parse(n);return Array.isArray(r)?r.join(", "):n||"-"}catch{return n||"-"}}},{key:"score",label:"Nilai",render:n=>n!=null?`<strong>${n}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:l}],exportOptions:{moduleName:"training",onExport:async n=>{let r=new URLSearchParams(n||{}).toString(),d=await k(`/api/training?limit=10000&${r}`);if(d.ok){let g=d.data.data.map(p=>{let u=p.participants||"";try{let h=JSON.parse(u);u=Array.isArray(h)?h.join(", "):u}catch{}return{Tanggal:p.training_date||"",Batch:p.batch||"",Materi:p.subject||"",Cabang:p.branch_name||"",Trainer:p.trainer||"",Peserta:u,Nilai:p.score!==null&&p.score!==void 0?p.score:"",Dokumen:p.document_link||""}}),{downloadExcel:c}=await Promise.resolve().then(()=>(K(),oe));c(g,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(K(),oe));r(n,"Template_Import_Training")},onImport:async n=>{let r=p=>{if(!p)return null;let u=String(p||"").toLowerCase(),h=e.find(m=>String(m.label||"").toLowerCase()===u);return h?h.value:null},d=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let u=String(p).trim();if(/^\d{4,5}$/.test(u)){let m=Number(u);if(m>2e4&&m<99999){let b=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);let h=u.split(/[\/\-\.]/);if(h.length===3){let[m,b,y]=h.map(_=>_.trim());if(m.length===4&&b.length<=2&&y.length<=2)return`${m}-${b.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&b.length<=2&&m.length<=2)return`${y}-${b.padStart(2,"0")}-${m.padStart(2,"0")}`}return u},g=n.map(p=>({training_date:d(p.Tanggal),batch:String(p.Batch||"").trim(),subject:String(p.Materi||"").trim(),branch_id:r(String(p.Cabang||"").trim()),trainer:String(p.Trainer||"").trim(),participants:String(p.Peserta||"").trim(),score:p.Nilai?Number(p.Nilai):null,document_link:String(p.Dokumen||"").trim()})).filter(p=>p.training_date&&p.subject&&p.branch_id),c=await k("/api/import/training",{method:"POST",body:JSON.stringify({rows:g,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},formFields:n=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:n?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:n?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:n?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:n?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:o(n?.trainer),value:n?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let r=JSON.parse(n?.participants);return Array.isArray(r)?r.join(", "):n?.participants||""}catch{return n?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:n?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:n?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:n?.notes}],onBeforeSubmit:async n=>(n.participants&&(n.participants=JSON.stringify(n.participants.split(",").map(r=>r.trim()).filter(Boolean))),n)})}O();ye();K();function Kt({container:t,title:e,icon:i,apiPath:a,columns:s,formFields:o,filterFields:l,defaultFilters:n={},itemLabel:r="Data",canCreate:d=!0,canEdit:g=!0,canDelete:c=!0,onBeforeSubmit:p,onAfterLoad:u,onDataLoaded:h,extraActions:m=[],initialSearch:b="",exportOptions:y=null,bulkDelete:_=!1,paginationMode:x="server"}){let w=se();w&&typeof w=="object"&&w.role==="viewer"&&(d=!1,g=!1,c=!1,_=!1,y=null);let S=1,T={...n};b&&(T.search=b);let $=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions">
        ${d?`<button class="btn btn-primary" id="btn-create">+ Tambah ${r}</button>`:""}
      </div>
    </div>

    ${_?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${y?Re(y.moduleName):""}

    ${l&&l.length>0?`
    <div class="filter-bar card" style="padding: 1rem; display: flex; gap: 0.5rem; flex-wrap: wrap;">
        ${l.map(v=>{if(v.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${T.search||""}"></div>`;if(v.type==="search-combo"){let f="dl-filter-search",D=(v.options||[]).map(L=>`<option value="${typeof L=="object"?L.label:L}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${f}" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${T.search||""}"><datalist id="${f}">${D}</datalist></div>`}return v.type==="select"?`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${v.name}" id="filter-${v.name}"><option value="">Pilih ${v.label}</option>${(v.options||[]).map(f=>`<option value="${typeof f=="object"?f.value:f}" ${T[v.name]===(typeof f=="object"?f.value:f)?"selected":""}>${typeof f=="object"?f.label:f}</option>`).join("")}</select>`:""}).join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function C(){if(!document.getElementById("bulk-toolbar"))return;let f=document.getElementById("bulk-count"),D=document.getElementById("btn-bulk-delete"),L=document.getElementById("btn-bulk-cancel");f.textContent=`${$.size} item dipilih`,$.size>0?(D.disabled=!1,L.disabled=!1):(D.disabled=!0,L.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{$.clear(),document.querySelectorAll(".row-checkbox").forEach(f=>f.checked=!1);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1),C()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if($.size===0)return;let v=[...$],f=document.createElement("div");f.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",f.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${v.length} ${r}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${v.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(f),f.querySelector("#bulk-cancel-btn").addEventListener("click",()=>f.remove()),f.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let D=f.querySelector("#bulk-confirm-btn");D.disabled=!0,D.textContent="Menghapus...";let L=await k(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:v})});f.remove(),L.ok?(ee(`${v.length} ${r} berhasil dihapus.`),$.clear(),C(),U()):Q(L.data?.error||"Gagal menghapus data.")})});let B=document.getElementById("filter-search"),R;if(B?.addEventListener("input",v=>{clearTimeout(R),R=setTimeout(()=>{T.search=v.target.value,S=1,$.clear(),U()},400)}),l?.forEach(v=>{v.type==="select"&&document.getElementById(`filter-${v.name}`)?.addEventListener("change",f=>{T[v.name]=f.target.value,S=1,$.clear(),U()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{T={...n},B&&(B.value=""),l?.forEach(v=>{let f=document.getElementById(`filter-${v.name}`);f&&(f.value="")}),S=1,$.clear(),U()}),document.getElementById("btn-create")?.addEventListener("click",()=>me(null)),y){document.getElementById(`btn-export-${y.moduleName}`)?.addEventListener("click",async f=>{let D=f.target,L=D.innerHTML;D.innerHTML="\u23F3 Loading...",D.disabled=!0;try{await y.onExport()}catch{Q("Gagal export data")}finally{D.innerHTML=L,D.disabled=!1}}),document.getElementById(`btn-template-${y.moduleName}`)?.addEventListener("click",()=>{y.onTemplate()});let v=document.getElementById(`input-import-${y.moduleName}`);v?.addEventListener("change",async f=>{let D=f.target.files[0];if(!D)return;v.disabled=!0;let L=document.createElement("div");L.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",L.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(L);let V=L.querySelector("#import-progress-text"),G=L.querySelector("#import-progress-bar"),I=L.querySelector("#import-summary"),P=L.querySelector("#import-close-btn");P.addEventListener("click",()=>{L.remove(),U()});try{let X=await Oe(D);if(X.length===0)throw new Error("File kosong atau format salah");let Z=500,le=0,ne=0,F=0,M=X.length;V.textContent=`Ditemukan ${M} baris data. Memulai import...`;for(let E=0;E<M;E+=Z){let W=X.slice(E,E+Z);V.textContent=`Mengimport baris ${E+1} - ${Math.min(E+Z,M)} dari ${M}...`,G.style.width=`${Math.round(E/M*100)}%`;try{let j=await y.onImport(W);j?(le+=j.inserted||j.metrics?.inserted||W.length,ne+=j.skipped||j.metrics?.updated||0):le+=W.length}catch(j){console.error("Chunk import failed:",j),F+=W.length}}G.style.width="100%",V.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',I.style.display="block",I.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${M}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${le}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ne}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${F}</strong></div>
        `,F>0&&(I.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),P.style.display="block",v.value=""}catch(X){V.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${X.message}`,G.style.background="var(--danger)",G.style.width="100%",P.style.display="block",v.value=""}finally{v.disabled=!1}})}async function U(){C();let v=document.getElementById("table-container");if(!v)return;v.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let f=x==="client",D=f?1:S,L=f?fe:20,V=new URLSearchParams({page:D,limit:L,...Object.fromEntries(Object.entries(T).filter(([,F])=>F))}),G=await k(`${a}?${V}`);if(!G.ok){v.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${G.data?.error||"Error"}</p></div>`;return}let I=G.data?.data||G.data||[],P=G.data?.pagination,X=I.length,Z=I;if(f){I=h(I),Z=I;let F=I.length,M=20,E=Math.ceil(F/M);S>E&&E>0&&(S=E);let W=(S-1)*M,j=S*M;I=I.slice(W,j),P={page:S,limit:M,total:F,pages:E}}!1,u&&u(I);let le=Ge({columns:s,data:I,fullData:Z,onEdit:g?F=>me(F):null,actions:m.map(F=>({...F,handler:M=>F.handler(M,U)})),emptyText:`Tidak ada ${String(r||"").toLowerCase()}`,bulkSelect:_?{selectedIds:$,onToggle:C}:null});v.innerHTML="",v.appendChild(le);let ne=document.getElementById("pagination-container");if(ne&&(ne.innerHTML="",P&&P.pages>1)){let F=Qe({page:P.page,pages:P.pages,total:P.total,limit:P.limit,onPage:M=>{S=M,U()}});F&&ne.appendChild(F)}}function ve(v){let f=typeof o=="function"?o(v):o;return Me(f)}function me(v){let f=!!v,D=document.createElement("form");if(D.noValidate=!0,D.innerHTML=ve(v),f){let V=typeof o=="function"?o(v):o;We(D,v)}let{close:L}=de({title:f?`Edit ${r}`:`Tambah ${r}`,content:D,size:"lg",confirmText:f?"Simpan Perubahan":`Tambah ${r}`,onConfirm:async(V,G)=>{if(!D.reportValidity())return;let I=V.querySelector(".modal-confirm");I.disabled=!0,I.textContent="Menyimpan...";let P=ze(D),X=typeof o=="function"?o(v):o,Z=async M=>{for(let E of M)if(E.type==="row")await Z(E.fields);else if(E.type==="combobox"&&P[E.name]){let W=P[E.name],j=(E.options||[]).find(Y=>{let ie=String(typeof Y=="object"?Y.value:Y),at=String(typeof Y=="object"?Y.label:Y);return ie===W||at===W});if(j)P[E.name]=typeof j=="object"?j.value:j;else if(E.createApi){let Y={};Y[E.createApi.field]=W,E.createApi.extra&&Object.assign(Y,E.createApi.extra);let ie=await k(E.createApi.path,{method:"POST",body:JSON.stringify(Y)});if(ie.ok&&ie.data?.id)P[E.name]=ie.data.id;else if(ie.ok&&!ie.data?.id)P[E.name]=W;else throw new Error(`Gagal membuat master data: ${ie.data?.error||"Unknown error"}`)}}};try{await Z(X)}catch(M){Q(M.message),I.disabled=!1,I.textContent=f?"Simpan Perubahan":`Tambah ${r}`;return}p&&(P=await p(P,v));let le=f?"PUT":"POST",ne=f?`${a}/${v.id}`:a,F=await k(ne,{method:le,body:JSON.stringify(P)});F.ok?(ee(f?`${r} berhasil diperbarui.`:`${r} berhasil ditambahkan.`),G(),U()):(Q(F.data?.error||"Gagal menyimpan data."),I.disabled=!1,I.textContent=f?"Simpan Perubahan":`Tambah ${r}`)}})}function Je(v){Fe(`Hapus ${r} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let f=await k(`${a}/${v.id}`,{method:"DELETE"});f.ok?(ee(`${r} berhasil dihapus.`),U()):Q(f.data?.error||"Gagal menghapus.")},`Hapus ${r}`)}return U(),U}O();K();async function Ht(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await H(),a=await te(),s=e?e.get("dash_filter"):null,o={};if(s==="reliever"){let d=new Date,g=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;o={status:"Done",month:e&&e.get("month")?e.get("month"):g}}console.log("RAW",await xe()),console.log("OPTIONS",a);let l=d=>d&&!a.find(g=>g.value===d)?[...a,{value:d,label:d}]:a,n=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],r=d=>d&&!n.includes(d)?[...n,d]:n;Kt({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",defaultFilters:o,onDataLoaded:d=>d.sort((g,c)=>{let p=g.backup_date?new Date(g.backup_date).getTime():0;return(c.backup_date?new Date(c.backup_date).getTime():0)-p}),columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:d=>ce(d)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:d=>window.formatDate(d)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:d=>window.formatDate(d)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:d=>d?`<span class="badge badge-info">${d}</span>`:"-"},{key:"status",label:"Status",render:d=>J(d)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:n},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:d=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:d?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:d?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:l(d?.original_fc_name),value:d?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:r(d?.reliever_name),value:d?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:d?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:d?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:d?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:d?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:d?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let d=await k(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(d.ok){let g=d.data.data.map(c=>({Cabang:c.branch_name||"","Nama Facility care":c.original_fc_name||"",Periode:c.period||"",Relifer:c.reliever_name||"","Tanggal Back Up":c.backup_date||"","Tanggal Selesai":c.completion_date||"",Keterangan:c.reason||"",Shift:c.shift||"",Status:c.status||""}));g.length===0&&g.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),A(g,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{A([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async d=>{let c=(await k("/api/branches?all=1")).data?.data||[],p=m=>{if(!m)return null;let b=String(m||"").toLowerCase(),y=c.find(_=>String(_.full_name||"").toLowerCase()===b||String(_.code||"").toLowerCase()===b||String(_.name||"").toLowerCase()===b);return y?y.id:null},u=d.map(m=>({branch_name:String(m.Cabang||"").trim(),backup_date:String(m["Tanggal Back Up"]||m["Tanggal Backup"]||"").trim(),original_fc_name:String(m["Nama Facility care"]||m["FC Digantikan"]||"").trim(),reliever_name:String(m.Relifer||m.Reliefer||"").trim(),period:String(m.Periode||"").trim(),reason:String(m.Keterangan||"").trim(),shift:String(m.Shift||"").trim(),completion_date:String(m["Tanggal Selesai"]||"").trim(),status:String(m.Status||"").trim()})).filter(m=>m.reliever_name&&m.backup_date),h=await k("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:u,onDuplicate:"update"})});if(!h.ok)throw new Error(h.data?.error||"Import gagal");return h.data}}})}O();K();async function jt(t){let e=await H(),i=Array.from({length:4},(a,s)=>String(new Date().getFullYear()-s));N({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>ce(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>J(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let s=new URLSearchParams(a||{}).toString(),o=await k(`/api/reports/inspection?limit=10000&${s}`);if(o.ok){let l=o.data.data.map(n=>({Cabang:n.branch_name||"",Periode:n.period||"",Tanggal:n.inspection_date||"","Point FC":n.fc_score!==null&&n.fc_score!==void 0?n.fc_score:"","Point SPV":n.spv_score!==null&&n.spv_score!==void 0?n.spv_score:"",Status:n.status||"","Link Dokumen":n.document_link||""}));A(l,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{A([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let s=r=>{if(!r)return null;let d=String(r||"").toLowerCase(),g=e.find(c=>String(c.label||"").toLowerCase()===d);return g?g.value:null},o=r=>{if(r==null||r==="")return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let d=String(r).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let c=Number(d);if(c>2e4&&c<99999){let p=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let g=d.split(/[\/\-\.]/);if(g.length===3){let[c,p,u]=g.map(h=>h.trim());if(c.length===4&&p.length<=2&&u.length<=2)return`${c}-${p.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&p.length<=2&&c.length<=2)return`${u}-${p.padStart(2,"0")}-${c.padStart(2,"0")}`}return d},l=a.map(r=>({branch_id:s(String(r.Cabang||"").trim()),period:String(r.Periode||"").trim(),inspection_date:o(r.Tanggal),fc_score:r["Point FC"]!==void 0&&r["Point FC"]!==""?Number(r["Point FC"]):null,spv_score:r["Point SPV"]!==void 0&&r["Point SPV"]!==""?Number(r["Point SPV"]):null,status:String(r.Status||"").trim(),document_link:String(r["Link Dokumen"]||"").trim(),notes:String(r.Catatan||r.Keterangan||"").trim()})).filter(r=>r.branch_id&&r.period&&r.inspection_date),n=await k("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}O();K();async function qt(t){let e=await H(),i=Array.from({length:4},(a,s)=>String(new Date().getFullYear()-s));N({container:t,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>ce(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>J(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let s=new URLSearchParams(a||{}).toString(),o=await k(`/api/reports/cleaning?limit=10000&${s}`);if(o.ok){let l=o.data.data.map(n=>({Cabang:n.branch_name||"",Jenis:n.activity_type||"",Periode:n.period||"",Tanggal:n.activity_date||"",Status:n.status||"","Link Dokumen":n.document_link||""}));A(l,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{A([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let s=r=>{if(!r)return null;let d=String(r||"").toLowerCase(),g=e.find(c=>String(c.label||"").toLowerCase()===d);return g?g.value:null},o=r=>{if(r==null||r==="")return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let d=String(r).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let c=Number(d);if(c>2e4&&c<99999){let p=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let g=d.split(/[\/\-\.]/);if(g.length===3){let[c,p,u]=g.map(h=>h.trim());if(c.length===4&&p.length<=2&&u.length<=2)return`${c}-${p.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&p.length<=2&&c.length<=2)return`${u}-${p.padStart(2,"0")}-${c.padStart(2,"0")}`}return d},l=a.map(r=>({branch_id:s(String(r.Cabang||"").trim()),activity_type:String(r.Jenis||r.Kegiatan||"").trim(),period:String(r.Periode||"").trim(),activity_date:o(r.Tanggal),status:String(r.Status||"").trim(),document_link:String(r["Link Dokumen"]||"").trim(),notes:String(r.Catatan||r.Keterangan||"").trim()})).filter(r=>r.branch_id&&r.activity_type&&r.period&&r.activity_date),n=await k("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!n.ok)throw new Error(n.data?.error||"Import gagal");return n.data}}})}O();K();async function Jt(t,e){let i=await H(),a=Array.from({length:4},(l,n)=>String(new Date().getFullYear()-n)),s=e?e.get("dash_filter"):null,o={};if(s==="fogging"){let l=new Date,n=String(l.getMonth()+1).padStart(2,"0"),r=String(l.getFullYear()),d=e?e.get("month"):null;d&&d.length===7&&(r=d.split("-")[0],n=d.split("-")[1]),o={status:"Done",month:n,year:r}}N({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,defaultFilters:o,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:l=>`<span class="badge badge-warning">${l}</span>`},{key:"period",label:"Periode",render:l=>ce(l)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>J(l)},{key:"document_link",label:"Dokumen",render:l=>l?`<a href="${l}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:l=>l||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:a}],formFields:l=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:l?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:l?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:l?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:l?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:l?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async l=>{let n=new URLSearchParams(l||{}).toString(),r=await k(`/api/reports/fogging?limit=10000&${n}`);if(r.ok){let d=r.data.data.map(g=>({Cabang:g.branch_name||"",Jenis:g.activity_type||"Fogging",Periode:g.period||"",Tanggal:g.activity_date||"",Status:g.status||"","Link Dokumen":g.document_link||""}));A(d,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{A([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async l=>{let n=c=>{if(!c)return null;let p=String(c||"").toLowerCase(),u=i.find(h=>String(h.label||"").toLowerCase()===p);return u?u.value:null},r=c=>{if(c==null||c==="")return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let p=String(c).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let h=Number(p);if(h>2e4&&h<99999){let m=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}let u=p.split(/[\/\-\.]/);if(u.length===3){let[h,m,b]=u.map(y=>y.trim());if(h.length===4&&m.length<=2&&b.length<=2)return`${h}-${m.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&m.length<=2&&h.length<=2)return`${b}-${m.padStart(2,"0")}-${h.padStart(2,"0")}`}return p},d=l.map(c=>({branch_id:n(String(c.Cabang||"").trim()),activity_type:String(c.Jenis||c.Kegiatan||"Fogging").trim(),period:String(c.Periode||"").trim(),activity_date:r(c.Tanggal),status:String(c.Status||"").trim(),document_link:String(c["Link Dokumen"]||"").trim(),notes:String(c.Catatan||c.Keterangan||"").trim()})).filter(c=>c.branch_id&&c.period&&c.activity_date),g=await k("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(d)});if(!g.ok)throw new Error(g.data?.error||"Import gagal");return g.data}}})}O();K();async function Ut(t){let e=await H(),i=await te(),a=i,s=Array.from({length:4},(n,r)=>String(new Date().getFullYear()-r)),o=n=>n&&!i.find(r=>r.value===n)?[...i,{value:n,label:n}]:i,l=n=>n&&!a.find(r=>r.value===n)?[...a,{value:n,label:n}]:a;N({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:n=>window.formatDate(n)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:n=>`<span title="${n||""}">${n?.length>60?n.slice(0,60)+"\u2026":n||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:n=>window.formatDate(n)},{key:"status",label:"Status",render:n=>J(n)},{key:"notes",label:"Keterangan",render:n=>n?.length>40?n.slice(0,40)+"\u2026":n||"-"}],filterFields:[{type:"select",name:"pic",label:"PIC",options:["Berlin","Ade","Mizwar"]},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"month",label:"Bulan",options:[{value:"01",label:"Jan"},{value:"02",label:"Feb"},{value:"03",label:"Mar"},{value:"04",label:"Apr"},{value:"05",label:"Mei"},{value:"06",label:"Jun"},{value:"07",label:"Jul"},{value:"08",label:"Agu"},{value:"09",label:"Sep"},{value:"10",label:"Okt"},{value:"11",label:"Nov"},{value:"12",label:"Des"}]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:s}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:n?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:l(n?.pic),value:n?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:n?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:n?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:n?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:n?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:n?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async n=>{let r=new URLSearchParams(n||{}).toString(),d=await k(`/api/reports/basecamp?limit=10000&${r}`);if(d.ok){let g=d.data.data.map(c=>({"Tgl Info":c.info_date||"",Cabang:c.branch_name||"",Permasalahan:c.problem||"",PIC:c.pic||"","Tgl Done":c.done_date||"",Status:c.status||"",Keterangan:c.notes||""}));A(g,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{A([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async n=>{let r=p=>{if(!p)return null;let u=String(p||"").toLowerCase(),h=e.find(m=>String(m.label||"").toLowerCase()===u);return h?h.value:null},d=p=>{if(p==null||p==="")return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let u=String(p).trim();if(u===""||u==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);if(/^\d{4,5}$/.test(u)){let m=Number(u);if(m>2e4&&m<99999){let b=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}let h=u.split(/[\/\-\.]/);if(h.length===3){let[m,b,y]=h.map(_=>_.trim());if(m.length===4&&b.length<=2&&y.length<=2)return`${m}-${b.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&b.length<=2&&m.length<=2)return`${y}-${b.padStart(2,"0")}-${m.padStart(2,"0")}`}return u},g=n.map(p=>({info_date:d(p["Tgl Info"]||p["Tanggal Info"]),branch_id:r(String(p.Cabang||"").trim()),problem:String(p.Permasalahan||"").trim(),pic:String(p.PIC||"").trim(),done_date:d(p["Tgl Done"]||p["Tanggal Done"]),status:String(p.Status||"").trim(),notes:String(p.Keterangan||p.Catatan||"").trim()})).filter(p=>p.info_date&&p.branch_id&&p.problem),c=await k("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(g)});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}}})}async function Gt(t){N({container:t,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(O(),ke)),s=await a(`/api/sop?limit=10000&${i}`);if(s.ok){let o=s.data.data.map(n=>({"Nama SOP":n.name||"",Kategori:n.category||"",Dokumen:n.document_link||"",Catatan:n.notes||n.description||""})),{downloadExcel:l}=await Promise.resolve().then(()=>(K(),oe));l(o,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(K(),oe));i(e,"Template_Import_SOP")},onImport:async e=>{let i=e.map(o=>({name:String(o["Nama SOP"]||"").trim(),category:String(o.Kategori||"").trim(),document_link:String(o.Dokumen||"").trim(),description:String(o.Catatan||"").trim()})).filter(o=>o.name),{apiFetch:a}=await Promise.resolve().then(()=>(O(),ke)),s=await a("/api/sop/import",{method:"POST",body:JSON.stringify(i)});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function Qt(t){N({container:t,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(O(),ke)),s=await a(`/api/checklist?limit=10000&${i}`);if(s.ok){let o=s.data.data.map(n=>({"Nama Checklist":n.name||"",Kategori:n.category||"",Dokumen:n.document_link||"",Deskripsi:n.description||""})),{downloadExcel:l}=await Promise.resolve().then(()=>(K(),oe));l(o,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(K(),oe));i(e,"Template_Import_Checklist")},onImport:async e=>{let i=e.map(o=>({name:String(o["Nama Checklist"]||"").trim(),category:String(o.Kategori||"").trim(),document_link:String(o.Dokumen||"").trim(),description:String(o.Deskripsi||"").trim()})).filter(o=>o.name),{apiFetch:a}=await Promise.resolve().then(()=>(O(),ke)),s=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(i)});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}O();ye();K();async function ft(t,e="forms"){if(e==="supply")return Aa(t);Ba(t)}function Ba(t){N({container:t,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await k(`/api/forms?limit=10000&${i}`);a.data?.data?A(a.data.data,"Data_Master_Form"):Q("Gagal export data master form")},onImport:async e=>{let i=await k("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!i.ok)throw new Error(i.data?.error||"Import failed");return i.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function Aa(t){let i=((await k("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));N({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,s)=>s.branch_name_ref||s.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let s=JSON.parse(a);return Array.isArray(s)?s.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let s=JSON.parse(a);return Array.isArray(s)?s.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>J(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let s=a?.tools_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}let o=a?.chemical_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:a?.branch_id&&!i.find(l=>l.value==a.branch_id)?[...i,{value:a.branch_id,label:a.branch_name||a.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:s},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:o},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let s=new URLSearchParams(a||{}).toString(),o=await k(`/api/reports/supply?limit=10000&${s}`);if(o.ok){let l=o.data.data.map(n=>{let r=n.tools_items;try{r=Array.isArray(JSON.parse(r))?JSON.parse(r).join(", "):r}catch{}let d=n.chemical_items;try{d=Array.isArray(JSON.parse(d))?JSON.parse(d).join(", "):d}catch{}return{Waktu:n.submitted_at||"",Pengirim:n.submitter_name||"",Cabang:n.branch_name_ref||n.branch_name||"","Alat/Barang":r||"",Chemical:d||"",Catatan:n.additional_notes||"",Status:n.status||"","Diproses Oleh":n.processed_by||""}});A(l,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{A([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let o=(await k("/api/branches?all=1")).data?.data||[],l=g=>{if(!g)return null;let c=String(g||"").toLowerCase(),p=o.find(u=>String(u.full_name||"").toLowerCase()===c||String(u.code||"").toLowerCase()===c||String(u.name||"").toLowerCase()===c);return p?p.id:null},n=g=>{if(g==null||g==="")return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let c=String(g).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let u=Number(c);if(u>2e4&&u<99999){let h=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}let p=c.split(/[\/\-\.]/);if(p.length===3){let[u,h,m]=p.map(b=>b.trim());if(u.length===4&&h.length<=2&&m.length<=2)return`${u}-${h.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&h.length<=2&&u.length<=2)return`${m}-${h.padStart(2,"0")}-${u.padStart(2,"0")}`}return c},r=a.map(g=>({submitted_at:n(g.Waktu||g.Tanggal),submitter_name:String(g.Pengirim||"").trim(),branch_id:l(String(g.Cabang||"").trim()),tools_items:String(g["Alat/Barang"]||g.Alat||"").trim(),chemical_items:String(g.Chemical||"").trim(),additional_notes:String(g.Catatan||g.Keterangan||"").trim(),status:String(g.Status||"").trim(),processed_by:String(g["Diproses Oleh"]||g.PIC||"").trim()})).filter(g=>g.submitted_at&&g.submitter_name&&g.branch_id),d=await k("/api/reports/supply/import",{method:"POST",body:JSON.stringify(r)});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,s)=>{let o=de({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(l,n)=>{let r=l.querySelector("#supply-status").value,d=l.querySelector("#supply-processed-by").value;(await k(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:r,processed_by:d})})).ok?(ee("Status diperbarui."),n(),s()):Q("Gagal update status.")}})}}]})}O();async function zt(t){let e=se();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}N({container:t,title:"Riwayat Aktivitas",icon:"\u{1F575}\uFE0F\u200D\u2642\uFE0F",apiPath:"/api/audit-logs",itemLabel:"Log",canCreate:!1,canEdit:!1,canDelete:!1,bulkDelete:!1,exportOptions:null,columns:[{key:"created_at",label:"Waktu",nowrap:!0,render:i=>new Date(i).toLocaleString("id-ID",{dateStyle:"short",timeStyle:"medium"})},{key:"user_name",label:"Pengguna",render:(i,a)=>`<strong>${i||"Sistem"}</strong><br><small class="text-muted" style="text-transform:capitalize">${a.user_role||""}</small>`},{key:"action",label:"Aksi",render:i=>`<span class="badge ${{CREATE:"badge-success",UPDATE:"badge-info",DELETE:"badge-danger"}[i]||"badge-neutral"}">${i}</span>`},{key:"module",label:"Modul",render:i=>`<span style="text-transform:capitalize">${(i||"").replace("_"," ")}</span>`},{key:"target_id",label:"ID Target"},{key:"id",label:"Detail",render:(i,a)=>`<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${i}')">Lihat Detail</button>`}],filterFields:[{type:"search",placeholder:"Cari pengguna, modul..."},{type:"select",name:"action",options:[{value:"",label:"Semua Aksi"},{value:"CREATE",label:"Tambah (CREATE)"},{value:"UPDATE",label:"Ubah (UPDATE)"},{value:"DELETE",label:"Hapus (DELETE)"}]},{type:"select",name:"module",options:[{value:"",label:"Semua Modul"},{value:"employees",label:"Karyawan"},{value:"schedule",label:"Jadwal"},{value:"issues",label:"Permasalahan"},{value:"relievers",label:"Reliefer"},{value:"contracts",label:"Kontrak"}]}]}),window.viewAuditDetail=async i=>{try{let o=((await(await fetch(`/api/audit-logs?search=${i}`,{headers:{Authorization:`Bearer ${localStorage.getItem("fm_token")}`}})).json()).data||[]).find(d=>String(d.id)===String(i));if(!o)return alert("Data tidak ditemukan");let l=d=>{if(!d)return"Tidak ada data";try{return JSON.stringify(JSON.parse(d),null,2)}catch{return d}},n=`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${l(o.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${l(o.new_data)}</pre>
           </div>
         </div>
       `,{createModal:r}=await Promise.resolve().then(()=>(ye(),st));r({title:`Detail Audit Log #${i}`,content:n,width:"800px",hideFooter:!0})}catch{alert("Gagal mengambil detail")}}}O();K();async function Wt(t){let e=se();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}N({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let a=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await k(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let a=i.data.data.map(s=>({"Nama Lengkap":s.full_name||"",Username:s.username||"",Email:s.email||"",Role:s.role||"",Status:s.is_active?"Aktif":"Nonaktif"}));A(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{A([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let a=i.map(o=>({full_name:String(o["Nama Lengkap"]||"").trim(),username:String(o.Username||"").trim(),email:String(o.Email||"").trim(),role:String(o.Role||"").trim()||"viewer",password:String(o.Password||"").trim()})).filter(o=>o.username&&o.password&&o.email&&o.full_name),s=await k("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}O();K();async function Vt(t){N({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await k(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)A(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{A([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let i=e.map(s=>({code:String(s["Kode Cabang"]||"").trim(),name:String(s["Nama Pendek"]||"").trim(),full_name:String(s["Nama Lengkap"]||"").trim(),city:String(s.Kota||"").trim()})).filter(s=>s.code&&s.name),a=await k("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}O();async function Yt(t){let e=new Date,i=[];t.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),s()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),s()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(o=>o.addEventListener("change",s));async function a(){try{let o=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;i=(await k(`/api/dashboard/calendar?month=${o}`)).data?.data||[]}catch(o){console.warn("[Calendar] Failed to load events, rendering empty grid:",o),i=[]}}async function s(){let o=document.getElementById("calendar-grid");if(o){o.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await a();try{let l=e.getFullYear(),n=e.getMonth(),r=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),d=document.getElementById("cal-month-label");d&&(d.textContent=r);let g=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(w=>w.value)),c=i.filter(w=>g.has(w.type)),p={};c.forEach(w=>{let S=(w.event_date||"").slice(0,10);p[S]||(p[S]=[]),p[S].push(w)});let u=new Date(l,n,1).getDay(),h=new Date(l,n+1,0).getDate(),m=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],b=new Date().toISOString().slice(0,10),y='<div class="calendar-grid">';m.forEach(w=>{y+=`<div class="cal-day-header">${w}</div>`});for(let w=0;w<u;w++)y+='<div class="cal-cell cal-cell-empty"></div>';for(let w=1;w<=h;w++){let S=`${l}-${String(n+1).padStart(2,"0")}-${String(w).padStart(2,"0")}`,T=p[S]||[],$=S===b;y+=`
          <div class="cal-cell ${$?"cal-today":""} ${T.length?"cal-has-events":""}"
               data-date="${S}" tabindex="0" role="button" aria-label="${S}">
            <div class="cal-day-num ${$?"today-num":""}">${w}</div>
            <div class="cal-events-preview">
              ${T.slice(0,3).map(C=>`
                <div class="cal-event-dot cal-color-${C.color||"gray"}" title="${Ze(C.title||C.type)}">
                  <span class="cal-event-dot-label">${Na(C.title||C.branch_name||C.type,18)}</span>
                </div>
              `).join("")}
              ${T.length>3?`<div class="cal-more">+${T.length-3} lagi</div>`:""}
            </div>
          </div>`}let x=(u+h)%7;if(x!==0)for(let w=0;w<7-x;w++)y+='<div class="cal-cell cal-cell-empty"></div>';y+="</div>",o.innerHTML=y,o.querySelectorAll(".cal-cell[data-date]").forEach(w=>{w.addEventListener("click",()=>{let S=w.dataset.date,T=p[S]||[];if(!T.length)return;let $=document.getElementById("cal-event-list"),C=new Date(S+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=C,document.getElementById("cal-event-items").innerHTML=T.map(B=>`
            <div class="cal-event-item cal-color-border-${B.color||"gray"}">
              <div class="cal-event-type">${Fa(B.type)}</div>
              <div class="cal-event-title">${Ze(B.title||"-")}</div>
              <div class="cal-event-branch">${Ze(B.branch_name||"")}</div>
              ${B.status?`<div class="cal-event-status">${Ze(B.status)}</div>`:""}
              ${B.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${B.days_remaining} hari</div>`:""}
            </div>
          `).join(""),$.style.display="block"})})}catch(l){console.error("[Calendar] Render error:",l),o&&(o.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}s()}function Na(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function Ze(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Fa(t){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[t]||t}O();async function Xt(t){let e=se(),i=(e?.full_name||e?.username||"U")[0].toUpperCase(),s={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${s},${s}99)">
            ${i}
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
  `;let o=localStorage.getItem("fm_token"),l=document.getElementById("session-info");if(o&&l)try{let n=JSON.parse(atob(o.split(".")[1])),r=new Date(n.exp*1e3);l.textContent=`Berakhir: ${r.toLocaleString("id-ID")}`}catch{l.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async n=>{n.preventDefault();let r=document.getElementById("pwd-error"),d=document.getElementById("pwd-success"),g=document.getElementById("btn-save-pwd");r.style.display="none",d.style.display="none";let c=n.target,p=c.current_password.value,u=c.new_password.value,h=c.confirm_password.value;if(u!==h){r.textContent="\u274C Konfirmasi password tidak cocok.",r.style.display="block";return}if(u.length<6){r.textContent="\u274C Password baru minimal 6 karakter.",r.style.display="block";return}g.disabled=!0,g.textContent="\u23F3 Menyimpan...";let m=await k("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:p,new_password:u})});g.disabled=!1,g.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',m.ok?(d.textContent="\u2705 Password berhasil diubah.",d.style.display="block",c.reset(),ee("Password berhasil diubah.")):(r.textContent=m.data?.error||"Gagal mengubah password.",r.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}O();var et={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function ae(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let s=Number(e);if(s>2e4&&s<99999){let o=new Date(Date.UTC(1899,11,30)+s*864e5);return isNaN(o.getTime())?null:o.toISOString().slice(0,10)}}let i=e.split(/[\/\-\.]/);if(i.length===3){let[s,o,l]=i.map(g=>g.trim()),n=Number(s),r=Number(o),d=Number(l);if(s.length===4&&n>1900)return`${s}-${o.padStart(2,"0")}-${l.padStart(2,"0")}`;if(l.length===4&&d>1900)return n>12?`${l}-${o.padStart(2,"0")}-${s.padStart(2,"0")}`:r>12?`${l}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`:`${l}-${o.padStart(2,"0")}-${s.padStart(2,"0")}`;if(l.length===2&&!isNaN(d)){let g=d>=50?`19${l}`:`20${l}`;return n>12?`${g}-${o.padStart(2,"0")}-${s.padStart(2,"0")}`:`${g}-${o.padStart(2,"0")}-${s.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function Zt(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Ma={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",phone:t["No. Hp"],join_date:ae(t["Tanggal Masuk"]),status:t.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:ae(t["Tanggal Mulai"]),end_date:ae(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:ae(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:ae(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:ae(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:ae(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:ae(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:ae(t["Tanggal Target"]||t["Tgl Target"]),completion_date:ae(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:ae(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:ae(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:ae(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:ae(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:ae(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:ae(t["Tanggal Back Up"]),completion_date:ae(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:ae(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:ae(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Oa(t,e){let i=et[t];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Ma[i.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let s=[],o=[],l=[];return e.filter(r=>!Zt(r)).forEach((r,d)=>{let g=e.indexOf(r)+2,c=[];a.required.forEach(({key:u,label:h})=>{let m=r[u];if(m==null||String(m).trim()===""){let b=Object.keys(r).filter(y=>y.trim()).join(", ");c.push({column:h,originalValue:m||"",reason:`Kolom "${h}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${b.slice(0,120)}`})}});let p=a.map(r);c.length>0?o.push({row:g,data:p,raw:r,errors:c}):(s.push(r),l.push(p))}),{valid:s,errors:o,mapped:l}}function ea(t){let e=[];return t.SheetNames.forEach(i=>{let a=et[i];if(!a)return;let s=t.Sheets[i],o=window.XLSX.utils.sheet_to_json(s,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),l=Oa(i,o),n=o.filter(r=>!Zt(r));e.push({sheetName:i,module:a.module,label:a.label,total:n.length,valid:l.mapped.length,errorCount:l.errors.length,errors:l.errors,mapped:l.mapped,skipped:!1})}),e}function ta(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,s])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(s),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function aa(t){let e=window.XLSX,i=e.utils.book_new(),a=!1;return t.forEach(s=>{if(!s.errors||s.errors.length===0)return;a=!0;let o=s.errors.map(n=>({"No. Baris":n.row,"Kolom Gagal":(n.errors||[]).map(r=>r.column||r).join("; "),"Alasan Error":(n.errors||[]).map(r=>r.reason||r).join("; "),...Object.fromEntries(Object.entries(n.data||{}).map(([r,d])=>[r,d??""]))})),l=e.utils.json_to_sheet(o);e.utils.book_append_sheet(i,l,s.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Ra=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function na(t){t.innerHTML=`
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
              ${Object.entries(et).map(([m,{label:b}])=>`<span class="import-sheet-tag">\u{1F4C4} ${m} \u2192 ${b}</span>`).join("")}
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
  `;let e=null,i=null,a=0,s={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function o(m){Object.entries(s).forEach(([b,y])=>{y.style.display=b===m?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let m=document.getElementById("btn-backup-db");m.disabled=!0,m.textContent="\u23F3 Memproses Backup...";try{let b=await k("/api/import/backup");if(b.ok){if(!window.XLSX){Q("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let y=window.XLSX,_=y.utils.book_new();Object.entries(b.data.database).forEach(([x,w])=>{let S=w.length>0?w:[{}],T=y.utils.json_to_sheet(S);y.utils.book_append_sheet(_,T,x.substring(0,31))}),y.writeFile(_,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),ee("Backup berhasil diunduh!")}else Q("Gagal memproses backup: "+(b.data?.error||"Unknown error"))}catch(b){Q("Gagal memproses backup: "+b.message)}finally{m.disabled=!1,m.textContent="\u{1F4E6} Backup Database"}});let l=document.getElementById("btn-sync-google");l&&l.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let m=l.innerHTML;l.innerHTML='<span class="spinner"></span> Menyinkronkan...',l.disabled=!0;try{let b=await k("/api/sync/google-sheets",{method:"POST"});b.ok?alert("Sinkronisasi Berhasil: "+(b.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(b.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{l.innerHTML=m,l.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{ta(),ee("Template Excel berhasil didownload!")});let n=document.getElementById("file-input"),r=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",m=>{m.stopPropagation(),n.click()}),n.addEventListener("change",m=>{m.target.files[0]&&d(m.target.files[0])}),r.addEventListener("dragover",m=>{m.preventDefault(),r.classList.add("drag-over")}),r.addEventListener("dragleave",()=>r.classList.remove("drag-over")),r.addEventListener("drop",m=>{m.preventDefault(),r.classList.remove("drag-over");let b=m.dataTransfer.files[0];b&&b.name.match(/\.xlsx?$/i)?d(b):Q("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,n.value="",document.getElementById("file-info").style.display="none",r.style.display="",o("upload")});async function d(m){e=m,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${m.name} (${(m.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",r.style.display="none",await g(m)}async function g(m){o("validating");let b=document.getElementById("validation-status"),y=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");b.textContent="Membaca file Excel...",y.style.width="20%",await je(200);let _=await m.arrayBuffer(),x=window.XLSX.read(_,{type:"array",cellDates:!0});b.textContent=`Memvalidasi ${x.SheetNames.length} sheet...`,y.style.width="50%",await je(100),i=ea(x),y.style.width="100%",b.textContent="Validasi selesai!",await je(300),c()}catch(_){o("upload"),Q("Gagal memproses file: "+_.message),document.getElementById("file-info").style.display="flex",r.style.display="none"}}function c(){o("preview");let m=i.filter(C=>!C.skipped).length,b=i.reduce((C,B)=>C+B.total,0),y=i.reduce((C,B)=>C+B.valid,0),_=i.reduce((C,B)=>C+B.errorCount,0),x=b>0?Math.round(y/b*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${m} sheet</span>
      <span class="badge badge-secondary">${b} baris</span>
      <span class="badge badge-success">${y} valid (${x}%)</span>
      ${_>0?`<span class="badge badge-danger">${_} error</span>`:""}
    `;let w=document.getElementById("preview-table-container");w.innerHTML=`
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
          ${i.map((C,B)=>`
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
                ${C.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${B}">\u{1F50D} ${C.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,w.querySelectorAll(".btn-detail-error").forEach(C=>{C.addEventListener("click",()=>{let B=i[Number(C.dataset.idx)];p(B)})});let S=document.getElementById("error-detail-section"),T=document.getElementById("error-detail-container");T.innerHTML="",S.style.display="none";let $=document.getElementById("btn-start-import");y===0?($.disabled=!0,$.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):($.disabled=!1,_>0?($.innerHTML=`\u{1F680} Import ${y} Data Valid (${_} dilewati)`,$.title="Baris error akan dilewati, baris valid tetap diimport"):$.innerHTML=`\u{1F680} Mulai Import ${y} Data`)}function p(m){let b=document.getElementById("error-detail-section"),y=document.getElementById("error-detail-container");b.style.display="";let _=m.errors.slice(0,100).map(x=>(Array.isArray(x.errors)?x.errors:[]).map(S=>{let T=typeof S=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${x.row}</span></td>
            <td><strong>${T?S.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${T&&S.originalValue!==void 0?S.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${T?S.reason:S}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${T&&S.aliases?`Gunakan salah satu nama kolom:<br><em>${S.aliases}</em>`:T&&S.hint?S.hint:""}
            </td>
          </tr>
        `}).join("")).join("");y.innerHTML=`
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
    `,b.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{o("upload"),document.getElementById("file-info").style.display="none",r.style.display="",e=null,n.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;aa(i)?ee("Log error berhasil didownload."):ee("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let m=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";u(m)});async function u(m){o("importing"),a=Date.now();let b=[];Ra.forEach(S=>{let T=i?.find($=>$.module===S&&$.mapped?.length>0);T&&b.push(T)});let y=document.getElementById("import-steps-list");y.innerHTML=b.map(S=>`
      <div class="import-step-item" id="step-item-${S.module}">
        <span class="step-item-icon" id="step-icon-${S.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${S.label} <span class="step-item-count">(${S.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${S.module}"></span>
      </div>
    `).join("");let _=document.getElementById("import-bar"),x=document.getElementById("import-current-status"),w={totalSheets:b.length,totalRows:b.reduce((S,T)=>S+T.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let S=0;S<b.length;S++){let T=b[S],$=document.getElementById(`step-icon-${T.module}`),C=document.getElementById(`step-status-${T.module}`);$.textContent="\u{1F504}",C.textContent="Mengimport...",x.textContent=`Mengimport ${T.label}...`,_.style.width=`${Math.round(S/b.length*100)}%`;try{let B=await k(`/api/import/${T.module}`,{method:"POST",body:JSON.stringify({rows:T.mapped,onDuplicate:m})});if(B.ok){let R=B.data;w.inserted+=R.inserted||0,w.skipped+=R.skipped||0,w.moduleResults.push({label:T.label,inserted:R.inserted||0,skipped:R.skipped||0,status:"ok"}),$.textContent="\u2705",C.innerHTML=`<span class="badge badge-success">${R.inserted||0} berhasil</span>${R.skipped>0?` <span class="badge badge-neutral">${R.skipped} skip</span>`:""}`}else w.failed++,w.moduleResults.push({label:T.label,inserted:0,skipped:0,status:"error",error:B.data?.error}),$.textContent="\u274C",C.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(B){w.failed++,w.moduleResults.push({label:T.label,inserted:0,skipped:0,status:"error",error:B.message}),$.textContent="\u274C",C.innerHTML='<span class="badge badge-danger">Gagal</span>'}await je(150)}_.style.width="100%",x.textContent="Selesai!",await je(400),h(w)}function h(m){o("summary");let b=((Date.now()-a)/1e3).toFixed(1),y=m.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${y?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${y?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,i=null,n.value="",document.getElementById("file-info").style.display="none",r.style.display="",o("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function je(t){return new Promise(e=>setTimeout(e,t))}O();var tt=[],ia=[];async function ra(t){tt=await H(),ia=await te(),N({container:t,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:tt}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await k(`/api/sp?limit=10000&${i}`);if(a.ok){let s=a.data.data.map(l=>({"Nama Karyawan":l.employee_name||"",Divisi:l.division||"",Cabang:l.branch_name||"","Tanggal Sp":l.tanggal||"","Akhir Sp":l.akhir_sp||"","Jenis Sp":l.sp_type||"","Link Document / Foto":l.document_link||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(K(),oe));o(s,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(K(),oe));i(e,"Template_Import_SP")},onImport:async e=>{let i=l=>{if(!l)return null;let n=String(l||"").toLowerCase(),r=tt.find(d=>String(d.label||"").toLowerCase()===n);return r?r.value:null},a=l=>{if(!l)return"";if(l instanceof Date&&!isNaN(l.getTime()))return l.toISOString().slice(0,10);let n=String(l).trim();if(/^\d{4,5}$/.test(n)){let d=Number(n);if(d>2e4&&d<99999){let g=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(n))return n.slice(0,10);let r=n.split(/[\/\-\.]/);if(r.length===3){let[d,g,c]=r.map(p=>p.trim());if(d.length===4&&g.length<=2&&c.length<=2)return`${d}-${g.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&g.length<=2&&d.length<=2)return`${c}-${g.padStart(2,"0")}-${d.padStart(2,"0")}`}return n},s=e.map(l=>({employee_name:String(l["Nama Karyawan"]||"").trim(),division:String(l.Divisi||"").trim(),branch_id:i(String(l.Cabang||"").trim()),tanggal:a(l["Tanggal Sp"]),akhir_sp:a(l["Akhir Sp"]),sp_type:String(l["Jenis Sp"]||"").trim(),document_link:String(l["Link Document / Foto"]||"").trim()})).filter(l=>l.employee_name&&l.branch_id),o=await k("/api/import/sp",{method:"POST",body:JSON.stringify({rows:s,onDuplicate:"update"})});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}},formFields:[{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:ia},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"combobox",name:"branch_id",label:"Cabang",required:!0,options:tt,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}O();var De=[],la=[];async function sa(t){De=await H(),la=await te(),N({container:t,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",options:De},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",options:De}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await k(`/api/mutasi?limit=10000&${i}`);if(a.ok){let s=a.data.data.map(l=>({Tanggal:l.tanggal||"","Nama Karyawan":l.employee_name||"","Cabang Asal":l.from_branch_name||"","Cabang Tujuan":l.to_branch_name||"",Status:l.status||"",Dokumen:l.document_link||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(K(),oe));o(s,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(K(),oe));i(e,"Template_Import_Mutasi")},onImport:async e=>{let i=l=>{if(!l)return null;let n=String(l||"").toLowerCase(),r=De.find(d=>String(d.label||"").toLowerCase()===n);return r?r.value:null},a=l=>{if(!l)return"";if(l instanceof Date&&!isNaN(l.getTime()))return l.toISOString().slice(0,10);let n=String(l).trim();if(/^\d{4,5}$/.test(n)){let d=Number(n);if(d>2e4&&d<99999){let g=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(n))return n.slice(0,10);let r=n.split(/[\/\-\.]/);if(r.length===3){let[d,g,c]=r.map(p=>p.trim());if(d.length===4&&g.length<=2&&c.length<=2)return`${d}-${g.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&g.length<=2&&d.length<=2)return`${c}-${g.padStart(2,"0")}-${d.padStart(2,"0")}`}return n},s=e.map(l=>({tanggal:a(l.Tanggal),employee_name:String(l["Nama Karyawan"]||"").trim(),from_branch_id:i(String(l["Cabang Asal"]||"").trim()),to_branch_id:i(String(l["Cabang Tujuan"]||"").trim()),status:String(l.Status||"").trim(),document_link:String(l.Dokumen||"").trim()})).filter(l=>l.tanggal&&l.employee_name&&l.from_branch_id&&l.to_branch_id),o=await k("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:s,onDuplicate:"update"})});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:la},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:De,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:De,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}O();async function oa(t){t.innerHTML=`
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
  `;let e=document.getElementById("btnRefreshSync"),i=document.getElementById("queueStatusFilter");e.addEventListener("click",s),i.addEventListener("change",n),document.getElementById("btnRetryAll").addEventListener("click",()=>a("retry",{allFailed:!0})),document.getElementById("btnResetStuck").addEventListener("click",()=>a("reset-stuck",{})),document.getElementById("btnPauseSync").addEventListener("click",()=>a("pause",{})),document.getElementById("btnResumeSync").addEventListener("click",()=>a("resume",{})),document.getElementById("btnReconcile").addEventListener("click",()=>a("reconcile",{module:"ALL",repairMode:!0})),document.getElementById("btnRetrySelected").addEventListener("click",()=>{let d=Array.from(document.querySelectorAll(".chk-queue:checked")).map(g=>g.value);if(d.length===0)return alert("No items selected");a("retry",{ids:d})}),document.getElementById("chkAllQueue").addEventListener("change",d=>{document.querySelectorAll(".chk-queue").forEach(g=>g.checked=d.target.checked)});async function a(d,g){if(confirm(`Are you sure you want to execute action: ${d}?`)){showLoading();try{let c=await k(`/api/sync/actions/${d}`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(g)});c.ok?(alert(c.data?.message||"Success"),s()):Q(c.error||"Action failed")}catch(c){Q(c.message)}hideLoading()}}await s();async function s(){showLoading(),await Promise.all([l(),n(),o(),r()]),hideLoading()}async function o(){try{let d=await k("/api/sync/performance");if(!d.ok)return;let{webhook:g,google_api:c,d1:p,queue:u,throughput:h}=d.data;document.getElementById("latencyHistograms").innerHTML=`
        <table class="w-full text-left">
          <thead><tr class="text-gray-500 border-b"><th>Metric</th><th>P50</th><th>P95</th><th>P99</th><th>Max</th></tr></thead>
          <tbody class="divide-y">
            <tr><td class="py-1">Webhook</td><td>${g.P50}ms</td><td>${g.P95}ms</td><td>${g.P99}ms</td><td>${g.Max}ms</td></tr>
            <tr><td class="py-1">Google API</td><td>${c.P50}ms</td><td>${c.P95}ms</td><td>${c.P99}ms</td><td>${c.Max}ms</td></tr>
            <tr><td class="py-1">D1 Execute</td><td>${p.P50}ms</td><td>${p.P95}ms</td><td>${p.P99}ms</td><td>${p.Max}ms</td></tr>
            <tr><td class="py-1">Queue Delay</td><td>${u.P50}ms</td><td>${u.P95}ms</td><td>${u.P99}ms</td><td>${u.Max}ms</td></tr>
          </tbody>
        </table>
        <div class="mt-4 pt-3 border-t text-gray-600 flex justify-between">
          <span>Throughput: <b>${h.events_per_sec}</b> ev/sec</span>
          <span><b>${h.events_per_min}</b> ev/min</span>
        </div>
      `,document.getElementById("lblSnapshotStatus").innerHTML='Checked 10 modules. Status: <b>COMPLETED</b><br><span class="text-xs text-gray-500">Run today at 02:00</span>'}catch(d){console.error(d)}}async function l(){try{let d=await k("/api/sync/health");if(!d.ok)return Q("Failed to fetch sync health");let{status:g,queue:c,circuit_breaker:p}=d.data,u=`
        <div class="bg-white p-4 rounded-lg shadow border-l-4 ${g==="HEALTHY"?"border-green-500":g==="WARNING"?"border-yellow-500":"border-red-500"}">
          <p class="text-sm text-gray-500">System Health</p>
          <p class="text-2xl font-bold ${g==="HEALTHY"?"text-green-600":g==="WARNING"?"text-yellow-600":"text-red-600"}">${g}</p>
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
      `;document.getElementById("syncOverviewCards").innerHTML=u;let h=document.getElementById("cbStateBadge"),m=document.getElementById("cbStateDesc"),b=document.getElementById("cbStatusCard");b.className="bg-white rounded-lg shadow p-6 border-l-4",p==="CLOSED"?(h.className="px-3 py-1 rounded-full text-sm font-medium bg-green-100 text-green-800",h.textContent="CLOSED",m.textContent="Traffic is flowing normally to Google Sheets.",b.classList.add("border-green-500")):p==="OPEN"?(h.className="px-3 py-1 rounded-full text-sm font-medium bg-red-100 text-red-800",h.textContent="OPEN",m.textContent="Failing fast. Traffic to Google Sheets is paused due to repeated failures.",b.classList.add("border-red-500")):p==="HALF_OPEN"?(h.className="px-3 py-1 rounded-full text-sm font-medium bg-yellow-100 text-yellow-800",h.textContent="HALF-OPEN",m.textContent="Testing recovery. Permitting limited traffic to verify stability.",b.classList.add("border-yellow-500")):h.textContent=p||"UNKNOWN"}catch(d){console.error(d)}}async function n(){try{let d=document.getElementById("queueStatusFilter").value,g=await k("/api/sync/queue?limit=15"+(d?"&status="+d:""));if(!g.ok)return;let c=document.getElementById("queueTableBody"),p=g.data?.data||g.data||[];if(p.length===0){c.innerHTML='<tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">No events found</td></tr>';return}c.innerHTML=p.map(u=>`
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
      `).join("")}catch(d){console.error(d)}}async function r(){try{let d=await k("/api/sync/metrics");if(!d.ok)return;let g=document.getElementById("metricsTableBody"),c=d.data||[];if(c.length===0){g.innerHTML='<tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">No metrics found in last 24h</td></tr>';return}g.innerHTML=c.map(p=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2 font-medium text-gray-800">${p.module}</td>
          <td class="px-4 py-2 text-gray-600">${p.total_events}</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(p.avg_webhook_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(p.avg_d1_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(p.avg_queue_wait_ms||0)} ms</td>
        </tr>
      `).join("")}catch(d){console.error(d)}}}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let e=Math.floor(Number(t)-25569);return new Date(e*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let e=t.split(/[\/\-]/);return`${e[2]}-${e[1]}-${e[0]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let i=e.split("-");if(i.length===3&&i[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],s=parseInt(i[2],10),o=a[parseInt(i[1],10)-1];return`${s} ${o} ${i[0]}`}return e};function z(t){return async e=>{if(!Le()){Se("/login");return}return t(e)}}var qe=null;function Ka(){qe&&clearInterval(qe);let t=()=>{let e=new Date,i=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),s=document.getElementById("header-clock-time"),o=document.getElementById("header-clock-date");s&&(s.textContent=i),o&&(o.textContent=a)};t(),qe=setInterval(t,1e3)}async function Ha(){try{let t=await k("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},i=(a,s)=>{let o=document.getElementById(a);o&&(o.textContent=s>0?s:"",o.style.display=s>0?"inline-flex":"none")};i("badge-issues",e.issues?.current||0),i("badge-contracts",e.expiring30?.current||0),i("badge-oo1",e.one_on_one?.current||0),i("badge-schedule",e.schedule?.current||0),i("badge-supply",e.supply?.current||0)}catch{}}var Ie=[];async function ja(){try{let t=await k("/api/dashboard/notifications");if(!t.ok)return;Ie=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Ie.length>0?"block":"none",e.textContent=Ie.length)}catch{}}function qa(){if(!Ie.length){de({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,i)=>i()});return}let t=`
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
  `;de({title:`Notifikasi (${Ie.length})`,content:t,confirmText:"Tutup",onConfirm:(e,i)=>i()})}function da(){let t=se(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
  `;let i=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),s=document.getElementById("topbar-menu-btn"),o=document.getElementById("sidebar-close"),l=()=>{i.classList.add("open"),a.classList.add("show")},n=()=>{i.classList.remove("open"),a.classList.remove("show")};s?.addEventListener("click",l),o?.addEventListener("click",n),a?.addEventListener("click",n),document.querySelectorAll(".nav-item").forEach(d=>d.addEventListener("click",n));function r(){let d=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(p=>{let u=p.dataset.route;p.classList.toggle("active",d===u||u!=="/dashboard"&&d.startsWith(u))});let g=document.getElementById("topbar-title"),c=document.querySelector(".nav-item.active .nav-label");g&&c&&(g.textContent=c.textContent)}window.addEventListener("hashchange",r),r(),Ka(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await k("/api/auth/logout",{method:"POST"}),Be(),qe&&clearInterval(qe),Se("/login")}),Ha(),ja(),document.getElementById("btn-notif")?.addEventListener("click",d=>{d.preventDefault(),qa()})}async function Ja(){q("/login",({main:e})=>Lt(e)),q("/dashboard",z(({main:e})=>$t(e))),q("/calendar",z(({main:e})=>Yt(e))),q("/employees",z(({main:e,params:i})=>Bt(e,i))),q("/contracts",z(({main:e,params:i})=>Nt(e,i))),q("/sp",z(({main:e})=>ra(e))),q("/mutasi",z(({main:e})=>sa(e))),q("/sync-dashboard",z(({main:e})=>oa(e))),q("/timeline",z(({main:e,params:i})=>Ft(e,i))),q("/issues",z(({main:e,params:i})=>Mt(e,i))),q("/one-on-one",z(({main:e,params:i})=>Ot(e,i))),q("/training",z(({main:e})=>Rt(e))),q("/relievers",z(({main:e,params:i})=>Ht(e,i))),q("/reports/inspection",z(({main:e})=>jt(e))),q("/reports/cleaning",z(({main:e})=>qt(e))),q("/reports/fogging",z(({main:e})=>Jt(e))),q("/reports/basecamp",z(({main:e})=>Ut(e))),q("/reports/supply",z(({main:e})=>ft(e,"supply"))),q("/sop",z(({main:e})=>Gt(e))),q("/checklist",z(({main:e})=>Qt(e))),q("/forms",z(({main:e})=>ft(e))),q("/users",z(({main:e})=>Wt(e))),q("/branches",z(({main:e})=>Vt(e))),q("/profile",z(({main:e})=>Xt(e))),q("/settings/import",z(({main:e})=>na(e))),q("/audit-logs",z(({main:e})=>zt(e)));let t=Le();if(!t&&window.location.hash!=="#/login"&&Se("/login"),t){let e=await k("/api/auth/me");e.ok?(Ae(e.data.data),da()):(Be(),Se("/login"))}window.addEventListener("fm:login",()=>{da(),Se("/dashboard")}),kt()}Ja();
