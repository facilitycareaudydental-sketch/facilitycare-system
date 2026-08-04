var da=Object.defineProperty;var nt=(t,e)=>()=>(t&&(e=t(t=0)),e);var it=(t,e)=>{for(var r in e)da(t,r,{get:e[r],enumerable:!0})};var ke={};it(ke,{API:()=>ft,CLIENT_SIDE_MAX_ROWS:()=>ye,IS_DEVELOPMENT:()=>Be,apiFetch:()=>k,clearToken:()=>Ae,getToken:()=>Le,getUser:()=>fe,setToken:()=>rt,setUser:()=>Ne});function Le(){return localStorage.getItem("fm_token")}function rt(t){localStorage.setItem("fm_token",t)}function Ae(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function fe(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Ne(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function k(t,e={}){let r=Le(),a={"Content-Type":"application/json",...r?{Authorization:`Bearer ${r}`}:{},...e.headers||{}};try{let o=`cb=${Date.now()}`,c=t.includes("?")?"&":"?",s=`${ft}${t}${c}${o}`,i=await fetch(s,{...e,headers:a}),n;try{let p=await i.text();try{n=JSON.parse(p)}catch{n={error:`Server Error (${i.status}): ${p.substring(0,80)}...`}}}catch{n={error:"Gagal membaca respon dari server"}}return i.status===401&&(Ae(),window.location.hash="#/login"),{ok:i.ok,status:i.status,data:n}}catch(o){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${o.message})`}}}}var Be,pa,ft,ye,O=nt(()=>{Be=!1,pa="https://fm-operations-api.facilitycare-audydental.workers.dev",ft=pa,ye=1e4});var xt={};it(xt,{confirmDialog:()=>Me,createModal:()=>ce});function ce({title:t,content:e,onConfirm:r,onCancel:a,confirmText:o="Simpan",cancelText:c="Batal",size:s="md",confirmClass:i="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},p=document.createElement("div");p.className="modal-overlay",p.innerHTML=`
    <div class="modal" style="max-width:${n[s]||n.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${c}</button>
        ${r?`<button class="btn ${i} modal-confirm">${o}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&p.querySelector(".modal-body").appendChild(e);let m=()=>{p.classList.remove("show"),setTimeout(()=>p.remove(),250)};return p.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),m()}),p.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),m()}),r&&p.querySelector(".modal-confirm").addEventListener("click",()=>r(p,m)),p.addEventListener("click",l=>{l.target===p&&(a&&a(),m())}),document.body.appendChild(p),requestAnimationFrame(()=>p.classList.add("show")),{overlay:p,close:m}}function Me(t,e,r="Konfirmasi"){return ce({title:r,content:`<p>${t}</p>`,onConfirm:(a,o)=>{e(),o()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var ve=nt(()=>{});var oe={};it(oe,{downloadExcel:()=>L,parseExcel:()=>Re,renderExcelButtons:()=>Ke});function Re(t){return new Promise((e,r)=>{let a=new FileReader;a.onload=o=>{try{let c=new Uint8Array(o.target.result),s=XLSX.read(c,{type:"array"}),i=s.SheetNames[0],n=s.Sheets[i];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${s.SheetNames.join(", ")}`),console.log(`Sheet Used: ${i}`);let p=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),m=p.e.r-p.s.r+1,l=p.e.c-p.s.c+1;console.log(`Total Rows (including empty): ${m}`),console.log(`Total Columns: ${l}`);let d=[];for(let b=p.s.c;b<=p.e.c;++b){let g=n[XLSX.utils.encode_cell({c:b,r:p.s.r})];g&&g.v&&d.push(g.v)}console.log(`Headers Found: ${d.join(", ")}`),console.log("---------------------------");let u=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(u,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(u,"__headers",{value:d,enumerable:!1}),e(u)}catch(c){r(c)}},a.onerror=o=>r(o),a.readAsArrayBuffer(t)})}function L(t,e){try{let r=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,r,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(r){throw console.error("Error generating Excel file:",r),r}}function Ke(t){return`
    <div class="excel-actions excel-btns" style="display:flex; margin-bottom:1rem;">
      <button class="btn btn-secondary btn-sm" id="btn-export-${t}">
        \u{1F4E5} Export Excel
      </button>
      <button class="btn btn-secondary btn-sm" id="btn-template-${t}">
        \u{1F4C4} Download Template
      </button>
      <label class="btn btn-primary btn-sm" style="cursor:pointer; margin:0;" id="label-import-${t}">
        <span class="import-text">\u{1F4E4} Import Excel</span>
        <input type="file" id="input-import-${t}" accept=".xlsx, .xls, .csv" style="display:none;">
      </label>
    </div>
  `}var R=nt(()=>{});O();var st={},Ue=null;function J(t,e){st[t]=e}function xe(t){window.location.hash=t}function vt(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[r,...a]=e.split("?"),o=st[r];if(!o){for(let[s,i]of Object.entries(st))if(s.endsWith("/*")&&r.startsWith(s.slice(0,-2))){o=i;break}}Ue&&(Ue(),Ue=null);let c=document.getElementById("main-content");if(c&&(c.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),o){let s=new URLSearchParams(a.join("?")),i=r.split("/").filter(Boolean),n=await o({path:r,params:s,segments:i,main:c});n&&(Ue=n)}else{let s=c||document.getElementById("app");s&&(s.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var Fe;function ua(){return Fe||(Fe=document.createElement("div"),Fe.id="toast-container",document.body.appendChild(Fe)),Fe}function kt(t,e="info",r=3500){let a=ua(),o=document.createElement("div");o.className=`toast toast-${e}`;let c={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};o.innerHTML=`<span class="toast-icon">${c[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(o),requestAnimationFrame(()=>o.classList.add("show")),setTimeout(()=>{o.classList.remove("show"),setTimeout(()=>o.remove(),350)},r)}var Z=t=>kt(t,"success"),G=t=>kt(t,"error");ve();O();O();ve();function Ge({columns:t,data:e,onEdit:r,onDelete:a,onView:o,actions:c=[],emptyText:s="Tidak ada data",bulkSelect:i=null}){let n=document.createElement("div");if(n.className="table-wrapper",!e||e.length===0)return n.innerHTML=`<div class="empty-state"><p>${s}</p></div>`,n;let p=document.createElement("table");p.className="data-table";let m=document.createElement("thead"),l=document.createElement("tr");if(i){let u=document.createElement("th");u.style.width="40px",u.style.textAlign="center";let b=document.createElement("input");b.type="checkbox",b.id="select-all-checkbox",b.title="Pilih semua",b.addEventListener("change",()=>{e.forEach(g=>{b.checked?i.selectedIds.add(g.id):i.selectedIds.delete(g.id)}),n.querySelectorAll(".row-checkbox").forEach(g=>g.checked=b.checked),i.onToggle()}),u.appendChild(b),l.appendChild(u)}if(t.forEach(u=>{let b=document.createElement("th");b.textContent=u.label,u.width&&(b.style.width=u.width),l.appendChild(b)}),r||a||o||c.length>0){let u=document.createElement("th");u.textContent="Aksi",u.style.width="120px",l.appendChild(u)}m.appendChild(l),p.appendChild(m);let d=document.createElement("tbody");return e.forEach(u=>{let b=document.createElement("tr");if(i){let g=document.createElement("td");g.style.textAlign="center",g.style.width="40px";let h=document.createElement("input");h.type="checkbox",h.className="row-checkbox",h.checked=i.selectedIds.has(u.id),h.addEventListener("change",()=>{if(h.checked)i.selectedIds.add(u.id);else{i.selectedIds.delete(u.id);let f=document.getElementById("select-all-checkbox");f&&(f.checked=!1)}i.onToggle()}),g.appendChild(h),b.appendChild(g)}if(t.forEach(g=>{let h=document.createElement("td");if(g.render){let f=g.render(u[g.key],u);f instanceof HTMLElement?h.appendChild(f):h.innerHTML=f||""}else h.textContent=u[g.key]!==null&&u[g.key]!==void 0&&u[g.key]!==""?u[g.key]:"";g.nowrap&&(h.style.whiteSpace="nowrap"),b.appendChild(h)}),r||a||o||c.length>0){let g=document.createElement("td");g.className="actions-cell";let h=document.createElement("div");if(h.className="btn-group",o){let f=document.createElement("button");f.className="btn btn-xs btn-ghost",f.innerHTML="\u{1F441}",f.title="Lihat",f.addEventListener("click",()=>o(u)),h.appendChild(f)}if(r){let f=document.createElement("button");f.className="btn btn-xs btn-secondary",f.innerHTML="\u270F\uFE0F",f.title="Edit",f.addEventListener("click",()=>r(u)),h.appendChild(f)}c.forEach(f=>{let x=document.createElement("button");x.className=`btn btn-xs ${f.class||"btn-ghost"}`,x.innerHTML=f.icon||f.label,x.title=f.label,x.addEventListener("click",()=>f.handler(u)),h.appendChild(x)}),g.appendChild(h),b.appendChild(g)}d.appendChild(b)}),p.appendChild(d),n.appendChild(p),n}function ze({page:t,pages:e,total:r,limit:a,onPage:o}){if(e<=1)return null;let c=document.createElement("div");c.className="pagination";let s=document.createElement("span");s.className="pagination-info",s.textContent=`Total: ${r} data`,c.appendChild(s);let i=document.createElement("div");i.className="pagination-btns";let n=(l,d,u=!1,b=!1)=>{let g=document.createElement("button");g.className=`btn btn-sm ${b?"btn-primary":"btn-ghost"} pagination-btn`,g.textContent=l,g.disabled=u,g.addEventListener("click",()=>o(d)),i.appendChild(g)};n("\xAB",1,t===1),n("\u2039",t-1,t===1);let p=Math.max(1,t-2),m=Math.min(e,t+2);for(let l=p;l<=m;l++)n(l,l,!1,l===t);return n("\u203A",t+1,t===e),n("\xBB",e,t===e),c.appendChild(i),c}ve();function Oe(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${Oe(e.fields)}</div>`;let r=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",o="";switch(e.type){case"textarea":o=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${r} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let s=(e.options||[]).map(l=>{let d=typeof l=="object"?l.value:l,u=typeof l=="object"?l.label:l,b=e.value==d?"selected":"";return`<option value="${d}" ${b}>${u}</option>`}).join("");o=`<select name="${e.name}" class="form-control" ${r}><option value="">-- Pilih ${e.label||""} --</option>${s}</select>`;break;case"combobox":let i=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,n=(e.options||[]).map(l=>{let d=typeof l=="object"?l.value:l,u=typeof l=="object"?l.label||l.value||"":l||"";return(u==="undefined"||u==="[object Object]"||u==="null")&&(u=""),u?`<option value="${u}"></option>`:""}).join(""),p=e.value||"";if(e.value){let l=(e.options||[]).find(d=>(typeof d=="object"?d.value:d)==e.value);if(l){let d=typeof l=="object"?l.label||l.value||"":l||"";d&&d!=="undefined"&&d!=="[object Object]"&&d!=="null"&&(p=d)}}o=`
          <input type="text" name="${e.name}" list="${i}" class="form-control" value="${p}" placeholder="Pilih atau ketik baru..." ${r} autocomplete="off">
          <datalist id="${i}">${n}</datalist>
        `;break;case"checkbox":o=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let m=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";o=`<input type="date" name="${e.name}" class="form-control" value="${m}" ${r}>`;break;case"number":o=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${r}>`;break;case"email":o=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${r}>`;break;case"url":o=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${r}>`;break;default:o=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${r} autocomplete="off">`}let c=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${o}${c}</div>`}).join("")}function Qe(t){let e={},r=new FormData(t);for(let[a,o]of r.entries())e[a]=o===""?null:o;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function We(t,e){e&&Object.entries(e).forEach(([r,a])=>{let o=t.querySelector(`[name="${r}"]`);o&&(o.hasAttribute("list")||(o.type==="checkbox"?o.checked=!!a:o.type==="date"&&a&&window.parseFlexibleDate?o.value=window.parseFlexibleDate(a):o.value=a??""))})}R();function A({container:t,title:e,icon:r,apiPath:a,columns:o,formFields:c,filterFields:s,defaultFilters:i={},itemLabel:n="Data",canCreate:p=!0,canEdit:m=!0,canDelete:l=!0,onBeforeSubmit:d,onAfterLoad:u,onDataLoaded:b,extraActions:g=[],initialSearch:h="",exportOptions:f=null,bulkDelete:x=!1,paginationMode:_="server"}){let C=1,S={...i};h&&(S.search=h);let w=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${r} ${e}</h1>
      <div class="page-actions">
        ${p?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
      </div>
    </div>

    ${x?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="align-items:center; background:var(--bg-card); border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${f?Ke(f.moduleName):""}

    ${s&&s.length>0?`
    <div class="filter-bar card" style="padding: 1rem; display:flex; flex-wrap:wrap; gap:1rem; align-items:center;">
        ${s.map(v=>{if(v.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${S.search||""}"></div>`;if(v.type==="search-combo"){let y="dl-filter-search",$=(v.options||[]).map(P=>`<option value="${typeof P=="object"?P.label:P}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${y}" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${S.search||""}"><datalist id="${y}">${$}</datalist></div>`}return v.type==="select"?`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${v.name}" id="filter-${v.name}"><option value="">-- ${v.label} --</option>${(v.options||[]).map(y=>`<option value="${typeof y=="object"?y.value:y}" ${S[v.name]===(typeof y=="object"?y.value:y)?"selected":""}>${typeof y=="object"?y.label:y}</option>`).join("")}</select>`:""}).join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function B(){if(!document.getElementById("bulk-toolbar"))return;let y=document.getElementById("bulk-count"),$=document.getElementById("btn-bulk-delete"),P=document.getElementById("btn-bulk-cancel");y.textContent=`${w.size} item dipilih`,w.size>0?($.disabled=!1,P.disabled=!1):($.disabled=!0,P.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{w.clear(),document.querySelectorAll(".row-checkbox").forEach(y=>y.checked=!1);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1),B()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(w.size===0)return;let v=[...w],y=document.createElement("div");y.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",y.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${v.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${v.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(y),y.querySelector("#bulk-cancel-btn").addEventListener("click",()=>y.remove()),y.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let $=y.querySelector("#bulk-confirm-btn");$.disabled=!0,$.textContent="Menghapus...";let P=await k(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:v})});y.remove(),P.ok?(Z(`${v.length} ${n} berhasil dihapus.`),w.clear(),B(),M()):G(P.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),q;if(E?.addEventListener("input",v=>{clearTimeout(q),q=setTimeout(()=>{S.search=v.target.value,C=1,w.clear(),M()},400)}),s?.forEach(v=>{v.type==="select"&&document.getElementById(`filter-${v.name}`)?.addEventListener("change",y=>{S[v.name]=y.target.value,C=1,w.clear(),M()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{S={...i},E&&(E.value=""),s?.forEach(v=>{let y=document.getElementById(`filter-${v.name}`);y&&(y.value="")}),C=1,w.clear(),M()}),document.getElementById("btn-create")?.addEventListener("click",()=>se(null)),f){document.getElementById(`btn-export-${f.moduleName}`)?.addEventListener("click",async y=>{let $=y.target,P=$.innerHTML;$.innerHTML="\u23F3 Loading...",$.disabled=!0;try{await f.onExport()}catch{G("Gagal export data")}finally{$.innerHTML=P,$.disabled=!1}}),document.getElementById(`btn-template-${f.moduleName}`)?.addEventListener("click",()=>{f.onTemplate()});let v=document.getElementById(`input-import-${f.moduleName}`);v?.addEventListener("change",async y=>{let $=y.target.files[0];if(!$)return;v.disabled=!0;let P=document.createElement("div");P.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",P.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(P);let W=P.querySelector("#import-progress-text"),U=P.querySelector("#import-progress-bar"),D=P.querySelector("#import-summary"),I=P.querySelector("#import-close-btn");I.addEventListener("click",()=>{P.remove(),M()});try{let V=await Re($);if(V.length===0)throw new Error("File kosong atau format salah");let X=500,ie=0,ae=0,N=0,F=V.length;W.textContent=`Ditemukan ${F} baris data. Memulai import...`;for(let T=0;T<F;T+=X){let Q=V.slice(T,T+X);W.textContent=`Mengimport baris ${T+1} - ${Math.min(T+X,F)} dari ${F}...`,U.style.width=`${Math.round(T/F*100)}%`;try{let H=await f.onImport(Q);H?(ie+=H.inserted||H.metrics?.inserted||Q.length,ae+=H.skipped||H.metrics?.updated||0):ie+=Q.length}catch(H){console.error("Chunk import failed:",H),N+=Q.length}}U.style.width="100%",W.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',D.style.display="block",D.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${F}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ie}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${N}</strong></div>
        `,N>0&&(D.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),I.style.display="block",v.value=""}catch(V){W.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${V.message}`,U.style.background="var(--danger)",U.style.width="100%",I.style.display="block",v.value=""}finally{v.disabled=!1}})}async function M(){B();let v=document.getElementById("table-container");if(!v)return;v.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let y=_==="client",$=y?1:C,P=y?ye:20,W=new URLSearchParams({page:$,limit:P,...Object.fromEntries(Object.entries(S).filter(([,N])=>N))}),U=await k(`${a}?${W}`);if(!U.ok){v.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${U.data?.error||"Error"}</p></div>`;return}let D=U.data?.data||U.data||[],I=U.data?.pagination,V=D.length,X=D;if(y){D=b(D),X=D;let N=D.length,F=20,T=Math.ceil(N/F);C>T&&T>0&&(C=T);let Q=(C-1)*F,H=C*F;D=D.slice(Q,H),I={page:C,limit:F,total:N,pages:T}}!1,u&&u(D);let ie=Ge({columns:o,data:D,fullData:X,onEdit:m?N=>se(N):null,actions:g.map(N=>({...N,handler:F=>N.handler(F,M)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:x?{selectedIds:w,onToggle:B}:null});v.innerHTML="",v.appendChild(ie);let ae=document.getElementById("pagination-container");if(ae&&(ae.innerHTML="",I&&I.pages>1)){let N=ze({page:I.page,pages:I.pages,total:I.total,limit:I.limit,onPage:F=>{C=F,M()}});N&&ae.appendChild(N)}}function ge(v){let y=typeof c=="function"?c(v):c;return Oe(y)}function se(v){let y=!!v,$=document.createElement("form");if($.noValidate=!0,$.innerHTML=ge(v),y){let W=typeof c=="function"?c(v):c;We($,v)}let{close:P}=ce({title:y?`Edit ${n}`:`Tambah ${n}`,content:$,size:"lg",confirmText:y?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(W,U)=>{if(!$.reportValidity())return;let D=W.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let I=Qe($),V=typeof c=="function"?c(v):c,X=async F=>{for(let T of F)if(T.type==="row")await X(T.fields);else if(T.type==="combobox"&&I[T.name]){let Q=I[T.name],H=(T.options||[]).find(Y=>{let ne=String(typeof Y=="object"?Y.value:Y),at=String(typeof Y=="object"?Y.label:Y);return ne===Q||at===Q});if(H)I[T.name]=typeof H=="object"?H.value:H;else if(T.createApi){let Y={};Y[T.createApi.field]=Q,T.createApi.extra&&Object.assign(Y,T.createApi.extra);let ne=await k(T.createApi.path,{method:"POST",body:JSON.stringify(Y)});if(ne.ok&&ne.data?.id)I[T.name]=ne.data.id;else if(ne.ok&&!ne.data?.id)I[T.name]=Q;else throw new Error(`Gagal membuat master data: ${ne.data?.error||"Unknown error"}`)}}};try{await X(V)}catch(F){G(F.message),D.disabled=!1,D.textContent=y?"Simpan Perubahan":`Tambah ${n}`;return}d&&(I=await d(I,v));let ie=y?"PUT":"POST",ae=y?`${a}/${v.id}`:a,N=await k(ae,{method:ie,body:JSON.stringify(I)});N.ok?(Z(y?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),U(),M()):(G(N.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=y?"Simpan Perubahan":`Tambah ${n}`)}})}function le(v){Me(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let y=await k(`${a}/${v.id}`,{method:"DELETE"});y.ok?(Z(`${n} berhasil dihapus.`),M()):G(y.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return M(),M}O();O();var Se=null,Ye=null;async function we(t=!1){if(Se&&!t)return console.log("Employees Raw (Cache Hit)",Se.slice(0,5)),Se;let e=await k(`/api/employees?limit=${ye}&status=Aktif`);return Se=(e.data?.data||[]).map(r=>({value:r.id,label:r.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",Se.slice(0,5)),Se}async function ee(t=!1){let r=(await we(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",r.slice(0,5)),r}async function K(t=!1){return Ye&&!t||(Ye=((await k("/api/branches?all=1")).data?.data||[]).map(r=>({value:r.id,label:r.full_name}))),Ye}function j(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function ot(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function _e(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function lt(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function de(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}R();function ct(t,e){if(t.period!=="Q3")return!1;let r=String(t.status||"").toLowerCase();if(r!=="selesai"&&r!=="completed"&&r!=="done")return!1;let a=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?a.includes("inspeksi"):e==="gcdc"?a.includes("general cleaning")||a.includes("deep cleaning"):!1}O();R();function St(t,e){let r=String(t.status||"").toLowerCase();return e==="active"?r==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&r==="aktif":!1}O();R();function dt(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let o=new Date(a);o.setDate(a.getDate()+30);let c=new Date(t.end_date);return c.setHours(0,0,0,0),c>=a&&c<=o}return!1}O();R();function wt(t,e){let r=String(t.status||"").toLowerCase();return e==="open"?r==="open":!1}O();function _t(t,e){let r=String(t.status||"").toLowerCase();return e==="pending"?r==="pending":!1}var he={};function Ee(t){if(he[t]){try{he[t].destroy()}catch{}delete he[t]}}function ma(){Object.keys(he).forEach(Ee)}var be=(t,e=0)=>{let r=Number(t);return isNaN(r)||t===null||t===void 0?e:r},Ce=(t,e="\u2014")=>{if(t==null||t==="")return e;let r=String(t).trim();return r===""||r==="[object Object]"?e:r};var ga=t=>{if(!t||typeof t!="string")return"";try{let[e,r]=t.split("-");return new Date(Number(e),Number(r)-1).toLocaleDateString("id-ID",{month:"short",year:"2-digit"})}catch{return t}};function Tt(t,e,r=900){if(!t)return;let a=Math.max(0,Math.round(be(e)));if(a===0){t.textContent="0";return}let o=Date.now(),c=()=>{let s=Math.min((Date.now()-o)/r,1),i=1-Math.pow(1-s,3);t.textContent=Math.round(i*a).toLocaleString("id-ID"),s<1?requestAnimationFrame(c):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(c)}var ba={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},ha=t=>{let e=Ce(t,"\u2014");return`<span class="status-pill ${ba[e]||"pill-neutral"}">${e}</span>`};var pe={family:"Inter",size:11},Te="#94A3B8",Ve="#F1F5F9",pt=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],ya=()=>window.innerWidth<768;function Xe(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:ya()?"bottom":"top",labels:{font:pe,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:pe,titleFont:{...pe,weight:"700"}}},scales:{x:{grid:{color:Ve},ticks:{font:pe,color:Te,maxRotation:0}},y:{grid:{color:Ve},ticks:{font:pe,color:Te},beginAtZero:!0}},...t}}var fa=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),va=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Ct(t=3){return Array(t).fill(0).map((e,r)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${r<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function re(t,e,r=8e3){try{let a=new AbortController,o=setTimeout(()=>a.abort(),r),c=await k(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(o),!c||!c.ok)return e;let s=c.data;return s?s.data!==void 0?s.data??e:s:e}catch{return e}}function ka(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-schbar"].forEach(a=>{let o=document.getElementById(a);o&&(o.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-schbar"].forEach(a=>{let o=document.getElementById(a);if(o&&o.style.display==="none"){o.style.display="block";let c=o.parentElement;if(c&&!c.querySelector(".chart-empty")){let s=document.createElement("div");s.className="chart-empty",s.textContent="Belum ada data",o.style.display="none",c.appendChild(s)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&$t({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Dt({}),["table-contracts","table-issues"].forEach(a=>{let o=document.getElementById(a);o&&o.querySelector(".skeleton")&&(o.innerHTML='<div class="chart-empty">Belum ada data</div>')});let r=document.getElementById("activity-log");r&&r.querySelector(".skeleton")&&(r.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Et(t){ma(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">


      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${fa()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${va()}</div>

      <!-- Charts Row (All 3 in 1 Row) -->
      <div class="charts-row" style="grid-template-columns: 3fr 2fr 2fr;">
        
        <!-- Jadwal Kegiatan per Bulan (Tahun Berjalan) -->
        <div class="chart-card">
          <div class="chart-card-header" style="justify-content: space-between;">
            <div class="chart-card-title">\u{1F4CA} Jadwal Kegiatan per Bulan (<span id="sch-year-label">${new Date().getFullYear()}</span>)</div>
            <select id="filter-sch-year" class="btn-ghost" style="padding:4px 8px; font-size:0.8rem; border:1px solid var(--border); border-radius:6px; cursor:pointer; font-weight:700;">
              <option value="${new Date().getFullYear()}">${new Date().getFullYear()}</option>
              <option value="${new Date().getFullYear()-1}">${new Date().getFullYear()-1}</option>
              <option value="${new Date().getFullYear()-2}">${new Date().getFullYear()-2}</option>
            </select>
          </div>
          <div class="chart-canvas-wrap" style="height:170px; position:relative; margin-top:10px;">
            <div id="skel-schbar" class="skeleton" style="position:absolute; inset:0; border-radius:12px"></div>
            <canvas id="chart-schbar" style="display:none"></canvas>
          </div>
        </div>
        <div class="chart-card" style="display:flex; flex-direction:column;">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan per Kategori</div>
          </div>
          <div style="display:flex; flex-direction:column; gap:12px; align-items:center; justify-content:center; padding: 10px 0; margin-bottom: auto;">
            <div class="chart-canvas-wrap" style="width:110px; height:110px; position:relative">
              <div id="skel-donut" class="skeleton" style="position:absolute;inset:0;border-radius:50%"></div>
              <canvas id="chart-donut" style="display:none"></canvas>
            </div>
            <div id="donut-legend" class="donut-legend" style="width:100%; display:flex; justify-content:center; flex-wrap:wrap; gap:12px;"></div>
          </div>
          <div style="text-align:center; font-size:0.7rem; color:var(--text-3); margin-top:auto">
            Periode: 22 Juni - 22 Juli 2026
          </div>
        </div>
        <div class="chart-card" style="display:flex; flex-direction:column;">
          <div class="chart-card-header" style="flex-direction:column; align-items:flex-start; gap:10px;">
            <div class="chart-card-title">Trend Permasalahan</div>
            <div style="display:flex;align-items:center;gap:12px;font-size:0.75rem;font-weight:600;color:var(--text-2)">
               <div style="display:flex;align-items:center;gap:4px"><div style="width:12px;height:6px;border:2px solid #EF4444;border-radius:2px"></div> Open</div>
               <div style="display:flex;align-items:center;gap:4px"><div style="width:12px;height:6px;border:2px solid #10B981;border-radius:2px"></div> Closed</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="width:100%; height:120px; position:relative; margin-top:auto; margin-bottom:auto;">
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>ut(t)),document.getElementById("filter-insp-month")?.addEventListener("change",async r=>{let a=r.target.value,o=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",c=document.getElementById("skel-insp"),s=document.getElementById("chart-insp");c&&(c.style.display="block",c.style.position="absolute"),s&&(s.style.display="none");let i=await re(o,{},8e3);try{Pt(i)}catch(n){console.warn("InspBar render:",n),ue("skel-insp","chart-insp")}}),document.getElementById("filter-sch-year")?.addEventListener("change",async r=>{let a=r.target.value;document.getElementById("sch-year-label").textContent=a;let o=`/api/dashboard/schedule-monthly?year=${a}`,c=document.getElementById("skel-schbar"),s=document.getElementById("chart-schbar");c&&(c.style.display="block",c.style.position="absolute"),s&&(s.style.display="none");let i=await re(o,{data:[]},8e3);try{It(i)}catch(n){console.warn("SchBar render:",n),ue("skel-schbar","chart-schbar")}}),t._skelTimeout=setTimeout(()=>ka(),5e3),await ut(t),t._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?ut(t):clearInterval(t._dashRefresh)},6e4)}async function ut(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,r,a,o,c,s,i,n,p,m,l,d,u]=await Promise.all([re("/api/dashboard/kpi",{},8e3),re("/api/dashboard/issues-trend",{},8e3),re("/api/dashboard/issues-summary",{},8e3),re("/api/dashboard/inspection-bar",{},8e3),re("/api/dashboard/stats",{},8e3),re("/api/dashboard/calendar",[],8e3),re("/api/schedule?limit=10000",{data:[]},8e3),re("/api/employees?limit=10000",{data:[]},8e3),re("/api/contracts?limit=10000",{data:[]},8e3),re("/api/issues?limit=10000",{data:[]},8e3),re("/api/one-on-one?limit=10000",{data:[]},8e3),re("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3),re(`/api/dashboard/schedule-monthly?year=${document.getElementById("filter-sch-year")?.value||new Date().getFullYear()}`,{data:[]},8e3)]);if(e){let b=Array.isArray(i?.data)?i.data:Array.isArray(i)?i:[],g=Array.isArray(n?.data)?n.data:Array.isArray(n)?n:[],h=Array.isArray(p?.data)?p.data:Array.isArray(p)?p:[],f=Array.isArray(m?.data)?m.data:Array.isArray(m)?m:[],x=Array.isArray(l?.data)?l.data:Array.isArray(l)?l:[];e.employees&&(e.employees.current=g.filter(_=>St(_,"active")).length),e.contracts&&(e.contracts.current=h.filter(_=>dt(_,"active")).length),e.expiring30&&(e.expiring30={current:h.filter(_=>dt(_,"expiring30")).length}),e.issues&&(e.issues.current=f.filter(_=>wt(_,"open")).length),e.one_on_one&&(e.one_on_one.current=x.filter(_=>_t(_,"pending")).length),e.inspection_month&&(e.inspection_month.current=b.filter(_=>ct(_,"inspeksi")).length),e.cleaning_month&&(e.cleaning_month.current=b.filter(_=>ct(_,"gcdc")).length)}try{$t(e)}catch(b){console.warn("KPI render:",b)}try{Dt(e)}catch(b){console.warn("MiniStats render:",b)}try{xa(Array.isArray(a?.by_category)?a.by_category:[])}catch(b){console.warn("Donut render:",b),ue("skel-donut","chart-donut")}try{It(u)}catch(b){console.warn("SchBar render:",b),ue("skel-schbar","chart-schbar")}try{Sa(r)}catch(b){console.warn("Trend render:",b),ue("skel-trend","chart-trend")}try{Pt(o)}catch(b){console.warn("InspBar render:",b),ue("skel-insp","chart-insp")}try{let b=Array.isArray(c)?c:Array.isArray(c?.recent_issues)?c.recent_issues:[];_a(b)}catch(b){console.warn("IssuesTable render:",b)}try{let b=Array.isArray(c?.expiring_contracts)?c.expiring_contracts:[];wa(d)}catch(b){console.warn("ContractsTable render:",b)}try{Ca(Array.isArray(s)?s:[])}catch(b){console.warn("Agenda render:",b)}try{Ta()}catch(b){console.warn("Quick Actions render:",b)}}function $t(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let r=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=r.map(a=>{let o=be(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${o}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${o}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{Tt(a,parseInt(a.dataset.target)||0)})}function Dt(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let r=[{icon:"\u{1F4C5}",label:"Jadwal",val:t.schedule?.current,href:"#/timeline",color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F504}",label:"Reliefer",val:t.reliever_completed?.current,href:"#/relievers?dash_filter=reliever",color:"mini-teal"},{icon:"\u{1F50D}",label:"Inspeksi",val:t.inspection_month?.current,href:"#/timeline?dash_filter=inspeksi",color:"mini-blue"},{icon:"\u{1F9F9}",label:"GCDC",val:t.cleaning_month?.current,href:"#/timeline?dash_filter=gcdc",color:"mini-green"},{icon:"\u{1F4A8}",label:"Fogging",val:t.fogging_month?.current,href:"#/reports/fogging",color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=r.map(a=>`
    <a href="${a.href}" class="mini-stat ${a.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${a.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${be(a.val)}">0</div>
        <div class="mini-stat-text">${a.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(a=>Tt(a,parseInt(a.dataset.target)||0,700))}function xa(t){ue("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),r=document.getElementById("donut-legend");if(!e||!r)return;Ee("donut");let a=(t||[]).filter(n=>be(n.count)>0);if(!a.length){He(e,"Belum ada data permasalahan");return}let o=a.map(n=>`${Ce(n.category,"Lainnya")}`),c=a.map(n=>be(n.count)),s=c.reduce((n,p)=>n+p,0);r.innerHTML=a.map((n,p)=>{let m=pt[p%pt.length],l=s>0?Math.round(n.count/s*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${m}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${l}%)</span></div>
          <div class="donut-legend-label">${o[p]}</div>
        </div>
      </div>
    `}).join("");let i={id:"centerText",beforeDraw:function(n){let p=n.width,m=n.height,l=n.ctx;l.restore();let d=(m/80).toFixed(2);l.font="bold "+d+"em Inter",l.textBaseline="middle",l.fillStyle="#1E293B";let u=s.toString(),b=Math.round((p-l.measureText(u).width)/2),g=m/2;l.fillText(u,b,g-10),l.font="600 "+(d*.35).toFixed(2)+"em Inter",l.fillStyle="#64748B";let h="Total",f=Math.round((p-l.measureText(h).width)/2);l.fillText(h,f,g+15),l.save()}};he.donut=new Chart(e,{type:"doughnut",data:{labels:o,datasets:[{data:c,backgroundColor:pt,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:pe,titleFont:{...pe,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutout:"75%"},plugins:[i]})}function It(t){ue("skel-schbar","chart-schbar");let e=document.getElementById("chart-schbar");if(!e)return;if(Ee("schbar"),!t||!t.data||!t.data.length){He(e,"Belum ada jadwal pada tahun ini.");return}let r=t.data,o=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],c={};r.forEach(d=>{let u=d.activity_type||"Lainnya";u.toLowerCase().includes("inspeksi")&&(u="Inspeksi"),(u.toLowerCase()==="gcdc"||u.toLowerCase().includes("general cleaning"))&&(u="General Cleaning"),u.toLowerCase().includes("deep cleaning")&&(u="Deep Cleaning"),u.toLowerCase().includes("fogging")&&(u="Fogging"),c[u]||(c[u]=Array(12).fill(0));let b=parseInt(d.month,10)-1;b>=0&&b<12&&(c[u][b]+=d.c||0)});let s={Inspeksi:"#60A5FA","General Cleaning":"#4ADE80","Deep Cleaning":"#FB923C",Fogging:"#818CF8"},i=d=>{let u=0;for(let b=0;b<d.length;b++)u=d.charCodeAt(b)+((u<<5)-u);return`hsl(${Math.abs(u)%360}, 65%, 60%)`},n=["Inspeksi","General Cleaning","Deep Cleaning","Fogging"],p=Object.keys(c).filter(d=>!n.includes(d)),l=[...n,...p].filter(d=>c[d]).map(d=>({label:d,data:c[d],backgroundColor:s[d]||i(d),borderWidth:0,borderRadius:4}));he.schbar=new Chart(e,{type:"bar",data:{labels:o,datasets:l},options:Xe({responsive:!0,maintainAspectRatio:!1,interaction:{mode:"index",intersect:!1},plugins:{legend:{display:!0,position:"bottom",align:"center",labels:{boxWidth:12,usePointStyle:!0,padding:20,font:{family:"Inter",size:12}},padding:{top:20}},tooltip:{callbacks:{title:d=>{let u=d[0].label;return{Jan:"Januari",Feb:"Februari",Mar:"Maret",Apr:"April",Mei:"Mei",Jun:"Juni",Jul:"Juli",Agu:"Agustus",Sep:"September",Okt:"Oktober",Nov:"November",Des:"Desember"}[u]||u},label:d=>` ${d.dataset.label}: ${d.raw} Jadwal`}}},scales:{x:{stacked:!0,grid:{display:!1}},y:{stacked:!0,border:{display:!1},grid:{color:"#f3f4f6",drawTicks:!1}}}})})}function Sa(t){ue("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;Ee("trend"),t=t||{};let r=(t.labels||[]).map(ga),a=(t.open||[]).map(c=>be(c)),o=(t.closed||[]).map(c=>be(c));if(!r.length){He(e,"Belum ada data trend");return}he.trend=new Chart(e,{type:"line",data:{labels:r,datasets:[{label:"Open",data:a,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:o,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#10B981",borderWidth:2}]},options:Xe({plugins:{legend:{display:!1}}})})}function Pt(t){ue("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;Ee("inspBar"),t=t||{};let r=t.labels||[],a=(t.fc||[]).map(c=>be(c)),o=(t.spv||[]).map(c=>be(c));if(!r.length){He(e,"Belum ada data inspeksi");return}he.inspBar=new Chart(e,{type:"bar",data:{labels:r,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:o,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:Xe({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:pe,color:Te,maxRotation:45,minRotation:30}},y:{grid:{color:Ve},ticks:{font:pe,color:Te},min:0,max:100}}})})}function wa(t){ue("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;Ee("contractMiniBar"),t=t||{};let r={"06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(s=>{let i=s.split("-")[1];return r[i]||s}),o=(t.data||[]).map(s=>be(s));if(!a.length){He(e,"Belum ada data");return}let c=e.getContext("2d");he.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:o,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:Xe({onClick:(s,i)=>{if(i&&i.length>0){let n=i[0].index,p=(t.labels||[])[n];p&&(window.location.hash="#/contracts?month_expiry="+p)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:pe,color:Te,maxRotation:0,autoSkip:!1}},y:{grid:{color:Ve,borderDash:[4,4],drawBorder:!1},ticks:{font:pe,color:Te,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function _a(t){let e=document.getElementById("table-issues");if(!e)return;let r=(t||[]).slice(0,8);if(!r.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${r.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${ha(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Ce(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Ce(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function Ca(t){let e=document.getElementById("widget-agenda");if(!e)return;let r=new Date,a=`${r.getFullYear()}-${String(r.getMonth()+1).padStart(2,"0")}-${String(r.getDate()).padStart(2,"0")}`,c=(t||[]).filter(s=>(s.event_date||"").startsWith(a)).slice(0,10);if(!c.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${c.map(s=>{let i="#3B82F6",n="#EFF6FF",p="Agenda",m=(s.title||"").toLowerCase();return m.includes("inspeksi")?(i="#10B981",n="#ECFDF5",p="Inspeksi"):m.includes("cleaning")||m.includes("gcdc")?(i="#3B82F6",n="#EFF6FF",p="Cleaning"):m.includes("reliefer")?(i="#F59E0B",n="#FFFBEB",p="Reliefer"):m.includes("fogging")&&(i="#8B5CF6",n="#F5F3FF",p="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(s.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${i};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Ce(s.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Ce(s.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${i}">${p}</div>
        </div>
      `}).join("")}
    </div>
  `}function Ta(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(r=>`
    <a href="${r.href}" class="action-btn">
      <div class="action-icon" style="background:${r.bg}">${r.icon}</div>
      ${r.label}
    </a>
  `).join("")}function ue(t,e){let r=document.getElementById(t),a=document.getElementById(e);if(r&&(r.style.display="none",r.style.position=""),a){a.style.display="block";let o=a.parentElement;if(o){let c=o.querySelector(".chart-empty");c&&c.remove()}}}function He(t,e="Belum ada data"){if(!t)return;t.style.display="none";let r=t.parentElement;if(!r)return;if(!r.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent=e,r.appendChild(o)}}O();async function Bt(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),r=document.getElementById("login-error"),a=document.getElementById("login-btn"),o=document.getElementById("toggle-password"),c=document.getElementById("login-password");o?.addEventListener("click",()=>{let s=c.type==="text";c.type=s?"password":"text",o.style.color=s?"":"var(--primary)"}),e?.addEventListener("submit",async s=>{s.preventDefault(),r.style.display="none";let i=e.username.value.trim(),n=e.password.value;if(!i||!n){r.textContent="Username dan password wajib diisi.",r.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let p=await k("/api/auth/login",{method:"POST",body:JSON.stringify({username:i,password:n})});p.ok&&p.data.success?(rt(p.data.data.token),Ne(p.data.data.user),Z("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(r.textContent=p.data.error||"Username atau password salah.",r.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{r.textContent="Gagal terhubung ke server. Periksa koneksi internet.",r.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}O();R();async function Ea(){return await K()}function $a(t,e){let r=String(t.status||"").toLowerCase();return e==="active"?r==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&r==="aktif":!1}async function Lt(t,e){let r=await Ea(),a=e?e.get("dash_filter"):null;A({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:o=>a?o.filter(c=>$a(c,a)):o,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:o=>_e(o)},{key:"phone",label:"No. HP",render:o=>o?`<a href="tel:${o}">${o}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>j(o)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:o=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:o?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:o?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:r,value:o?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:o?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:o?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:o?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let o=await k(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(o.ok){let c=o.data.data.map(s=>({"Nama Lengkap":s.full_name,Cabang:s.branch_name||"",Divisi:s.division||"","No. HP":s.phone||"","Tgl Masuk":s.join_date||"",Status:s.status||""}));L(c,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async o=>{let c=n=>{if(!n)return null;let p=String(n||"").toLowerCase(),m=r.find(l=>String(l.label||"").toLowerCase()===p);return m?m.value:null},s=o.map(n=>({full_name:String(n["Nama Lengkap"]||"").trim(),branch_id:c(String(n.Cabang||"").trim()),division:String(n.Divisi||"").trim()||"FACILITY CARE",phone:String(n["No. HP"]||"").trim(),join_date:String(n["Tgl Masuk"]||"").trim(),status:String(n.Status||"").trim(),notes:String(n.Catatan||"").trim()})).filter(n=>n.full_name),i=await k("/api/import/employees",{method:"POST",body:JSON.stringify({rows:s,onDuplicate:"update"})});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}O();R();var gt=[],At=[];async function Da(){gt=await K(),At=await we()}var mt=async t=>{let e=[],r=1;for(;;){let o=await(await Promise.resolve().then(()=>(O(),ke))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${r}`);if(!o.ok)break;let c=o.data?.data||o.data||[],s=Array.isArray(c)?c:[];if(e=e.concat(s),s.length<100||o.data?.pagination&&r>=o.data.pagination.pages)break;r++}return e};function Ia(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let o=new Date(a);o.setDate(a.getDate()+30);let c=new Date(t.end_date);return c.setHours(0,0,0,0),c>=a&&c<=o}return!1}async function Nt(t,e){await Da();let r=e?e.get("dash_filter"):null;A({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",onDataLoaded:a=>r?a.filter(o=>Ia(o,r)):a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>_e(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,o)=>o.end_date&&String(o.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':ot(a)},{key:"status",label:"Status",render:a=>j(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:gt},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"month_expiry",label:"Bulan Habis",options:[{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"},{value:"2027-01",label:"Jan 2027"},{value:"2027-02",label:"Feb 2027"},{value:"2027-03",label:"Mar 2027"},{value:"2027-04",label:"Apr 2027"},{value:"2027-05",label:"Mei 2027"},{value:"2027-06",label:"Jun 2027"},{value:"2027-07",label:"Jul 2027"}]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[c,s]=await Promise.all([mt("/api/employees?status=Aktif"),mt("/api/contracts")]);if(c.length>0){let i=s.filter(l=>l.status==="Aktif"),n=new Set(i.map(l=>l.employee_id)),p=c.filter(l=>!n.has(l.id)),m=`<p style="margin-bottom:12px">Data yang terbaca: <b>${c.length}</b> Karyawan Aktif, dan <b>${i.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${p.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;p.forEach(l=>{let d=s.filter(b=>b.employee_id===l.id),u='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(d.length>0){let b=d[0];u=`Pernah ada kontrak (Status: <b style="color:#EF4444">${b.status}</b>, Selesai: ${window.formatDate(b.end_date)})`}m+=`<li style="margin-bottom:8px"><b>${l.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${l.branch_name||"-"} | ${u}</span></li>`}),m+="</ul>",Promise.resolve().then(()=>(ve(),xt)).then(l=>l.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:m,cancelText:"Tutup"}))}}catch(c){console.error(c)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let o=document.querySelector(".page-actions");o&&o.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:At,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:gt,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await k(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let o=a.data.data.map(c=>({"Nama Lengkap":c.employee_name,Cabang:c.branch_name||"","Div / Bagian":c.division||"","Tanggal Mulai":c.start_date||"","Tanggal Selesai":c.end_date&&String(c.end_date).startsWith("2099")?"":c.end_date||"","Sisa Kontrak":c.end_date&&String(c.end_date).startsWith("2099")?"Tetap":c.days_remaining!==null&&c.days_remaining!==void 0?`${c.days_remaining} Hari`:"",Status:c.status||""}));L(o,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[o,c]=await Promise.all([k("/api/branches?limit=10000"),mt("/api/employees")]),s=o.data?.data||[],i=c||[];console.log(`Total employee yang berhasil dimuat dari database : ${i.length}`),i.length>0&&(console.log("Contoh 5 employee pertama:"),i.slice(0,5).forEach((g,h)=>{console.log(`${h+1}. ID: ${g.id}, Name: ${g.full_name}, Status: ${g.status}`)}));let n=g=>{if(!g)return null;let h=String(g||"").replace(/\s+/g," ").toLowerCase().trim(),f=s.find(x=>String(x.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===h||String(x.code||"").replace(/\s+/g," ").toLowerCase().trim()===h||String(x.name||"").replace(/\s+/g," ").toLowerCase().trim()===h);return f?f.id:null},p=(g,h)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${h}`),console.log(`Nama dari Excel : "${g}"`),!g)return console.log("Alasan gagal mapping : Nama kosong"),null;let f=String(g||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${f}"`),console.log(`Jumlah employee di database : ${i.length}`);let x=i.find(_=>String(_.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===f);return x?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${x.id}`),x.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},m=g=>{if(!g)return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let h=String(g).trim();if(/^\d{4,5}(\.\d+)?$/.test(h)){let x=Math.floor(Number(h));if(x>2e4&&x<99999){let _=new Date(Date.UTC(1899,11,30)+x*864e5);return isNaN(_.getTime())?"":_.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);let f=h.split(/[\/\-\.]/);if(f.length===3){let[x,_,C]=f.map(S=>S.trim());if(x.length===4&&_.length<=2&&C.length<=2)return`${x}-${_.padStart(2,"0")}-${C.padStart(2,"0")}`;if(C.length===4&&_.length<=2&&x.length<=2)return`${C}-${_.padStart(2,"0")}-${x.padStart(2,"0")}`}return h},l=a.map((g,h)=>{let f=h+2,x=String(g["Nama Lengkap"]||"").trim(),_=g["Tanggal Mulai"],C=m(_);if(!C){let B=a.__worksheet,E=a.__headers||[],q=E.indexOf("Tanggal Mulai"),M="N/A",ge="N/A",se="N/A";if(q!==-1&&B&&window.XLSX){let v=window.XLSX.utils.encode_cell({c:q,r:f-1});se=v;let y=B[v];y?(M=y.t||"undefined",ge=y.w||"undefined"):M="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let le="Unknown";_==null||_===""?le="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":_ instanceof Date&&isNaN(_.getTime())?le="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":le="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${f}`),console.log(`Employee Name : ${x}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${q})`),console.log(`Raw Cell Value : "${_}"`),console.log(`JavaScript Type : ${typeof _}`),console.log(`SheetJS Cell Type : ${M}`),console.log(`SheetJS Formatted Value : "${ge}"`),console.log(`Value After Trim : "${String(_||"").trim()}"`),console.log(`Value After Date Parser : "${C}"`),console.log(`Is Empty : ${!_}`),console.log(`Is Invalid Date : ${_ instanceof Date?isNaN(_.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${le}`),console.log(`Workbook Sheet : ${B?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${se}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(g,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(E)),console.log(`==========================
`)}let S=p(x,f),w=null;return S||(w="Karyawan tidak ditemukan di Database"),{isValid:!!S,invalidReason:w,rowNum:f,data:{employee_id:S,branch_id:n(String(g.Cabang||"").trim()),division:String(g["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:C,end_date:m(g["Tanggal Selesai"])||"2099-12-31",status:String(g.Status||"").trim(),_rawName:x}}}),d=[],u=[];if(l.forEach(g=>{g.isValid?d.push(g.data):u.push({rowNum:g.rowNum,name:g.data._rawName,reason:g.invalidReason})}),console.log(`Split Validation - Valid: ${d.length}, Invalid: ${u.length}`),d.length===0)return{inserted:0,skipped:a.length,failed:a.length};let b=await k("/api/import/contracts",{method:"POST",body:JSON.stringify({rows:d,onDuplicate:"update"})});if(!b.ok)throw new Error(b.data?.error||"Import gagal");return b.data}}})}O();R();var bt=[],qe=[];function Pa(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let r of e)if(t.some(a=>a.period===r))return r;return"Q3"}function Ba(t,e){if(t.period!=="Q3")return!1;let r=String(t.status||"").toLowerCase();if(r!=="selesai"&&r!=="completed"&&r!=="done")return!1;let a=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?a.includes("inspeksi"):e==="gcdc"?a.includes("general cleaning")||a.includes("deep cleaning"):!1}async function Ft(t,e){bt=await K();let r=await ee(),o=(await k(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`)).data?.data||[],c=new Set;o.forEach(l=>{l.pic&&l.pic.trim()&&c.add(l.pic.trim())}),qe=Array.from(c).sort();let s=l=>l&&!qe.find(d=>(typeof d=="object"?d.value:d)===l)?[...qe,l]:qe,i=l=>{if(!l||l==="-"||String(l).trim()==="")return"";let d=String(l).split("-");return d.length===3&&d[0].length===4?`${d[2]}-${d[1]}-${d[0]}`:l},p=Pa(o),m=e?e.get("dash_filter"):null;A({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",defaultFilters:m?{period:"Q3"}:{period:p},onDataLoaded:l=>(m&&(l=l.filter(d=>Ba(d,m))),l.sort((d,u)=>{let b=d.opening_date?new Date(d.opening_date).getTime():0;return(u.opening_date?new Date(u.opening_date).getTime():0)-b})),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:l=>lt(l)},{key:"period",label:"Periode",render:l=>de(l)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:l=>i(l)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:l=>i(l)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:l=>i(l)},{key:"status",label:"Status",render:l=>j(l)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:bt},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:qe}],formFields:l=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:bt,value:l?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:l?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:l?.period},{name:"pic",label:"PIC",type:"combobox",options:s(l?.pic),value:l?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:l?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:l?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:l?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:l?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let l=await k(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(l.ok){let d=l.data.data.map(u=>({Cabang:u.branch_name||"",Kegiatan:u.activity_type||"",Periode:u.period||"",PIC:u.pic||"","Tgl Opening":u.opening_date||"","Tgl Target":u.target_date||"","Tgl Selesai":u.completion_date||"",Status:u.status||""}));L(d,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async l=>{let u=(await k("/api/branches?all=1")).data?.data||[],b=x=>{if(!x)return null;let _=String(x||"").toLowerCase(),C=u.find(S=>String(S.full_name||"").toLowerCase()===_||String(S.code||"").toLowerCase()===_||String(S.name||"").toLowerCase()===_);return C?C.id:null},g=x=>{if(x==null||x==="")return"";if(x instanceof Date&&!isNaN(x.getTime()))return x.toISOString().slice(0,10);let _=String(x).trim();if(_===""||_==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(_))return _.slice(0,10);if(/^\d{4,5}$/.test(_)){let S=Number(_);if(S>2e4&&S<99999){let w=new Date(Date.UTC(1899,11,30)+S*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}let C=_.split(/[\/\-\.]/);if(C.length===3){let[S,w,B]=C.map(E=>E.trim());if(S.length===4&&w.length<=2&&B.length<=2)return`${S}-${w.padStart(2,"0")}-${B.padStart(2,"0")}`;if(B.length===4&&w.length<=2&&S.length<=2)return`${B}-${w.padStart(2,"0")}-${S.padStart(2,"0")}`}return _},h=l.map(x=>({branch_id:b(String(x.Cabang||"").trim()),activity_type:String(x.Kegiatan||"").trim(),period:String(x.Periode||"").trim(),pic:String(x.PIC||x.Pic||"").trim(),opening_date:g(x["Tgl Opening"]||x["Tanggal Opening"]||x["Tgl Openir"]),target_date:g(x["Tgl Target"]||x["Tanggal Target"]),completion_date:g(x["Tgl Selesai"]||x["Tanggal Selesai"]),status:String(x.Status||"").trim(),notes:String(x.Catatan||x.Keterangan||"").trim()})).filter(x=>x.activity_type&&x.period),f=await k("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:h,onDuplicate:"update"})});if(!f.ok)throw new Error(f.data?.error||"Import gagal");return f.data}}})}O();R();var ht=[],Ze=[];function La(t,e){let r=String(t.status||"").toLowerCase();return e==="open"?r==="open":!1}async function Mt(t,e){let r=e?e.get("dash_filter"):null;ht=await K(),Ze=await ee();let a=s=>s&&!Ze.find(i=>i.value===s)?[...Ze,{value:s,label:s}]:Ze,o=new Date().getFullYear();A({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:s=>r?s.filter(i=>La(i,r)):s,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:s=>`<span class="badge badge-secondary">${s}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:s=>`<span title="${s}">${s?.length>50?s.slice(0,50)+"\u2026":s}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:s=>`<span title="${s||""}">${s?.length>40?s.slice(0,40)+"\u2026":s||"-"}</span>`},{key:"status",label:"Status",render:s=>j(s)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:s=>window.formatDate(s)},{key:"day_count",label:"Hari",render:s=>s??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:ht},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:["2025","2026","2027","2028","2029","2030"]}],formFields:s=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:s?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:ht,value:s?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:s?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:s?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:s?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(s?.employee_name),value:s?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(s?.fc_specialist),value:s?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:s?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:s?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let s=await k(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(s.ok){let i=s.data.data.map(n=>({Tanggal:n.report_date||"",Cabang:n.branch_name||"",Kategori:n.category||"",Sumber:n.source||"",Keluhan:n.complaint||"","Nama FC":n.employee_name||"","FC Spesialis":n.fc_specialist||"",Solusi:n.solution||"","Tgl Selesai":n.completion_date||"",Status:n.status||""}));L(i,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async s=>{let n=(await k("/api/branches?all=1")).data?.data||[],p=d=>{if(!d)return null;let u=String(d||"").toLowerCase(),b=n.find(g=>String(g.full_name||"").toLowerCase()===u||String(g.code||"").toLowerCase()===u||String(g.name||"").toLowerCase()===u);return b?b.id:null},m=s.map(d=>({branch_id:p(String(d.Cabang||"").trim()),report_date:String(d.Tanggal||"").trim(),category:String(d.Kategori||"").trim(),source:String(d.Sumber||"").trim(),complaint:String(d.Keluhan||"").trim(),employee_name:String(d["Nama FC"]||"").trim(),fc_specialist:String(d["FC Spesialis"]||"").trim(),solution:String(d.Solusi||"").trim(),completion_date:String(d["Tgl Selesai"]||"").trim(),status:String(d.Status||"").trim()})).filter(d=>d.report_date&&d.complaint&&d.category),l=await k("/api/import/issues",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}O();var $e=[];function Aa(t,e){let r=String(t.status||"").toLowerCase();return e==="pending"?r==="pending":!1}async function Ot(t,e){let r=e?e.get("dash_filter"):null;$e=await K();let a=await ee(),o=["Ade","Berlin"],c=i=>i&&!a.find(n=>n.value===i)?[...a,{value:i,label:i}]:a,s=i=>i&&!o.find(n=>(typeof n=="object"?n.value:n)===i)?[...o,i]:o;A({container:t,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:i=>r?i.filter(n=>Aa(n,r)):i,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:i=>`<span title="${i||""}">${i?.length>50?i.slice(0,50)+"\u2026":i||"-"}</span>`},{key:"solution",label:"Solusi",render:i=>`<span title="${i||""}">${i?.length>40?i.slice(0,40)+"\u2026":i||"-"}</span>`},{key:"status",label:"Status",render:i=>j(i)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:i=>window.formatDate(i)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:i=>i?`<a href="${i}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"combobox",name:"branch_id",label:"Cabang",options:$e},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async i=>{let n=new URLSearchParams(i||{}).toString(),p=await k(`/api/one-on-one?limit=10000&${n}`);if(p.ok){let m=p.data.data.map(d=>({Tanggal:d.meeting_date||"",Cabang:d.branch_name||"","Nama Karyawan":d.employee_name||"",PIC:d.pic||"",Masalah:d.problem||"",Solusi:d.solution||"",Status:d.status||"","Tgl Selesai":d.completion_date||"",Dokumen:d.document_link||""})),{downloadExcel:l}=await Promise.resolve().then(()=>(R(),oe));l(m,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let i=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(R(),oe));n(i,"Template_Import_OneOnOne")},onImport:async i=>{let n=d=>{if(!d)return null;let u=String(d||"").toLowerCase(),b=$e.find(g=>String(g.label||"").toLowerCase()===u);return b?b.value:null},p=d=>{if(!d)return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let u=String(d).trim();if(/^\d{4,5}$/.test(u)){let g=Number(u);if(g>2e4&&g<99999){let h=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);let b=u.split(/[\/\-\.]/);if(b.length===3){let[g,h,f]=b.map(x=>x.trim());if(g.length===4&&h.length<=2&&f.length<=2)return`${g}-${h.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&h.length<=2&&g.length<=2)return`${f}-${h.padStart(2,"0")}-${g.padStart(2,"0")}`}return u},m=i.map(d=>({meeting_date:p(d.Tanggal),employee_name:String(d["Nama Karyawan"]||"").trim(),branch_id:n(String(d.Cabang||"").trim()),pic:String(d.PIC||"").trim(),problem:String(d.Masalah||"").trim(),solution:String(d.Solusi||"").trim(),status:String(d.Status||"").trim(),completion_date:p(d["Tgl Selesai"]),document_link:String(d.Dokumen||"").trim()})).filter(d=>d.meeting_date&&d.employee_name&&d.branch_id),l=await k("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},formFields:i=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:i?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:i?.branch_id&&!$e.find(n=>n.value==i.branch_id)?[...$e,{value:i.branch_id,label:i.branch_name||i.branch_id}]:$e,createApi:{path:"/api/branches",field:"full_name"},value:i?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"combobox",required:!0,options:c(i?.employee_name),value:i?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:s(i?.pic),createApi:{path:"/api/pic",field:"name"},value:i?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:i?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:i?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:i?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:i?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:i?.document_link}]})}O();async function Rt(t){let e=await K(),r=await ee(),a=["Ade","Berlin"],o=i=>i&&!r.find(n=>n.value===i)?[...r,{value:i,label:i}]:r,c=i=>i&&!a.find(n=>(typeof n=="object"?n.value:n)===i)?[...a,i]:a,s=Array.from({length:5},(i,n)=>String(new Date().getFullYear()-n));A({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:i=>{try{let n=JSON.parse(i);return Array.isArray(n)?n.join(", "):i||"-"}catch{return i||"-"}}},{key:"score",label:"Nilai",render:i=>i!=null?`<strong>${i}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:i=>i?`<a href="${i}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:s}],exportOptions:{moduleName:"training",onExport:async i=>{let n=new URLSearchParams(i||{}).toString(),p=await k(`/api/training?limit=10000&${n}`);if(p.ok){let m=p.data.data.map(d=>{let u=d.participants||"";try{let b=JSON.parse(u);u=Array.isArray(b)?b.join(", "):u}catch{}return{Tanggal:d.training_date||"",Batch:d.batch||"",Materi:d.subject||"",Cabang:d.branch_name||"",Trainer:d.trainer||"",Peserta:u,Nilai:d.score!==null&&d.score!==void 0?d.score:"",Dokumen:d.document_link||""}}),{downloadExcel:l}=await Promise.resolve().then(()=>(R(),oe));l(m,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let i=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(R(),oe));n(i,"Template_Import_Training")},onImport:async i=>{let n=d=>{if(!d)return null;let u=String(d||"").toLowerCase(),b=e.find(g=>String(g.label||"").toLowerCase()===u);return b?b.value:null},p=d=>{if(!d)return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let u=String(d).trim();if(/^\d{4,5}$/.test(u)){let g=Number(u);if(g>2e4&&g<99999){let h=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);let b=u.split(/[\/\-\.]/);if(b.length===3){let[g,h,f]=b.map(x=>x.trim());if(g.length===4&&h.length<=2&&f.length<=2)return`${g}-${h.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&h.length<=2&&g.length<=2)return`${f}-${h.padStart(2,"0")}-${g.padStart(2,"0")}`}return u},m=i.map(d=>({training_date:p(d.Tanggal),batch:String(d.Batch||"").trim(),subject:String(d.Materi||"").trim(),branch_id:n(String(d.Cabang||"").trim()),trainer:String(d.Trainer||"").trim(),participants:String(d.Peserta||"").trim(),score:d.Nilai?Number(d.Nilai):null,document_link:String(d.Dokumen||"").trim()})).filter(d=>d.training_date&&d.subject&&d.branch_id),l=await k("/api/import/training",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},formFields:i=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:i?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:i?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:i?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:i?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:c(i?.trainer),value:i?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let n=JSON.parse(i?.participants);return Array.isArray(n)?n.join(", "):i?.participants||""}catch{return i?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:i?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:i?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:i?.notes}],onBeforeSubmit:async i=>(i.participants&&(i.participants=JSON.stringify(i.participants.split(",").map(n=>n.trim()).filter(Boolean))),i)})}O();ve();R();function Kt({container:t,title:e,icon:r,apiPath:a,columns:o,formFields:c,filterFields:s,defaultFilters:i={},itemLabel:n="Data",canCreate:p=!0,canEdit:m=!0,canDelete:l=!0,onBeforeSubmit:d,onAfterLoad:u,onDataLoaded:b,extraActions:g=[],initialSearch:h="",exportOptions:f=null,bulkDelete:x=!1,paginationMode:_="server"}){let C=1,S={...i};h&&(S.search=h);let w=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${r} ${e}</h1>
      <div class="page-actions">
        ${p?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
      </div>
    </div>

    ${x?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="align-items:center; background:var(--bg-card); border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${f?Ke(f.moduleName):""}

    ${s&&s.length>0?`
    <div class="filter-bar card" style="padding: 1rem; display:flex; flex-wrap:wrap; gap:1rem; align-items:center;">
        ${s.map(v=>{if(v.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${S.search||""}"></div>`;if(v.type==="search-combo"){let y="dl-filter-search",$=(v.options||[]).map(P=>`<option value="${typeof P=="object"?P.label:P}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${y}" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${S.search||""}"><datalist id="${y}">${$}</datalist></div>`}return v.type==="select"?`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${v.name}" id="filter-${v.name}"><option value="">-- ${v.label} --</option>${(v.options||[]).map(y=>`<option value="${typeof y=="object"?y.value:y}" ${S[v.name]===(typeof y=="object"?y.value:y)?"selected":""}>${typeof y=="object"?y.label:y}</option>`).join("")}</select>`:""}).join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function B(){if(!document.getElementById("bulk-toolbar"))return;let y=document.getElementById("bulk-count"),$=document.getElementById("btn-bulk-delete"),P=document.getElementById("btn-bulk-cancel");y.textContent=`${w.size} item dipilih`,w.size>0?($.disabled=!1,P.disabled=!1):($.disabled=!0,P.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{w.clear(),document.querySelectorAll(".row-checkbox").forEach(y=>y.checked=!1);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1),B()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(w.size===0)return;let v=[...w],y=document.createElement("div");y.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",y.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${v.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${v.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(y),y.querySelector("#bulk-cancel-btn").addEventListener("click",()=>y.remove()),y.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let $=y.querySelector("#bulk-confirm-btn");$.disabled=!0,$.textContent="Menghapus...";let P=await k(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:v})});y.remove(),P.ok?(Z(`${v.length} ${n} berhasil dihapus.`),w.clear(),B(),M()):G(P.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),q;if(E?.addEventListener("input",v=>{clearTimeout(q),q=setTimeout(()=>{S.search=v.target.value,C=1,w.clear(),M()},400)}),s?.forEach(v=>{v.type==="select"&&document.getElementById(`filter-${v.name}`)?.addEventListener("change",y=>{S[v.name]=y.target.value,C=1,w.clear(),M()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{S={...i},E&&(E.value=""),s?.forEach(v=>{let y=document.getElementById(`filter-${v.name}`);y&&(y.value="")}),C=1,w.clear(),M()}),document.getElementById("btn-create")?.addEventListener("click",()=>se(null)),f){document.getElementById(`btn-export-${f.moduleName}`)?.addEventListener("click",async y=>{let $=y.target,P=$.innerHTML;$.innerHTML="\u23F3 Loading...",$.disabled=!0;try{await f.onExport()}catch{G("Gagal export data")}finally{$.innerHTML=P,$.disabled=!1}}),document.getElementById(`btn-template-${f.moduleName}`)?.addEventListener("click",()=>{f.onTemplate()});let v=document.getElementById(`input-import-${f.moduleName}`);v?.addEventListener("change",async y=>{let $=y.target.files[0];if(!$)return;v.disabled=!0;let P=document.createElement("div");P.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",P.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(P);let W=P.querySelector("#import-progress-text"),U=P.querySelector("#import-progress-bar"),D=P.querySelector("#import-summary"),I=P.querySelector("#import-close-btn");I.addEventListener("click",()=>{P.remove(),M()});try{let V=await Re($);if(V.length===0)throw new Error("File kosong atau format salah");let X=500,ie=0,ae=0,N=0,F=V.length;W.textContent=`Ditemukan ${F} baris data. Memulai import...`;for(let T=0;T<F;T+=X){let Q=V.slice(T,T+X);W.textContent=`Mengimport baris ${T+1} - ${Math.min(T+X,F)} dari ${F}...`,U.style.width=`${Math.round(T/F*100)}%`;try{let H=await f.onImport(Q);H?(ie+=H.inserted||H.metrics?.inserted||Q.length,ae+=H.skipped||H.metrics?.updated||0):ie+=Q.length}catch(H){console.error("Chunk import failed:",H),N+=Q.length}}U.style.width="100%",W.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',D.style.display="block",D.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${F}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ie}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${N}</strong></div>
        `,N>0&&(D.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),I.style.display="block",v.value=""}catch(V){W.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${V.message}`,U.style.background="var(--danger)",U.style.width="100%",I.style.display="block",v.value=""}finally{v.disabled=!1}})}async function M(){B();let v=document.getElementById("table-container");if(!v)return;v.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let y=_==="client",$=y?1:C,P=y?ye:20,W=new URLSearchParams({page:$,limit:P,...Object.fromEntries(Object.entries(S).filter(([,N])=>N))}),U=await k(`${a}?${W}`);if(!U.ok){v.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${U.data?.error||"Error"}</p></div>`;return}let D=U.data?.data||U.data||[],I=U.data?.pagination,V=D.length,X=D;if(y){D=b(D),X=D;let N=D.length,F=20,T=Math.ceil(N/F);C>T&&T>0&&(C=T);let Q=(C-1)*F,H=C*F;D=D.slice(Q,H),I={page:C,limit:F,total:N,pages:T}}!1,u&&u(D);let ie=Ge({columns:o,data:D,fullData:X,onEdit:m?N=>se(N):null,actions:g.map(N=>({...N,handler:F=>N.handler(F,M)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:x?{selectedIds:w,onToggle:B}:null});v.innerHTML="",v.appendChild(ie);let ae=document.getElementById("pagination-container");if(ae&&(ae.innerHTML="",I&&I.pages>1)){let N=ze({page:I.page,pages:I.pages,total:I.total,limit:I.limit,onPage:F=>{C=F,M()}});N&&ae.appendChild(N)}}function ge(v){let y=typeof c=="function"?c(v):c;return Oe(y)}function se(v){let y=!!v,$=document.createElement("form");if($.noValidate=!0,$.innerHTML=ge(v),y){let W=typeof c=="function"?c(v):c;We($,v)}let{close:P}=ce({title:y?`Edit ${n}`:`Tambah ${n}`,content:$,size:"lg",confirmText:y?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(W,U)=>{if(!$.reportValidity())return;let D=W.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let I=Qe($),V=typeof c=="function"?c(v):c,X=async F=>{for(let T of F)if(T.type==="row")await X(T.fields);else if(T.type==="combobox"&&I[T.name]){let Q=I[T.name],H=(T.options||[]).find(Y=>{let ne=String(typeof Y=="object"?Y.value:Y),at=String(typeof Y=="object"?Y.label:Y);return ne===Q||at===Q});if(H)I[T.name]=typeof H=="object"?H.value:H;else if(T.createApi){let Y={};Y[T.createApi.field]=Q,T.createApi.extra&&Object.assign(Y,T.createApi.extra);let ne=await k(T.createApi.path,{method:"POST",body:JSON.stringify(Y)});if(ne.ok&&ne.data?.id)I[T.name]=ne.data.id;else if(ne.ok&&!ne.data?.id)I[T.name]=Q;else throw new Error(`Gagal membuat master data: ${ne.data?.error||"Unknown error"}`)}}};try{await X(V)}catch(F){G(F.message),D.disabled=!1,D.textContent=y?"Simpan Perubahan":`Tambah ${n}`;return}d&&(I=await d(I,v));let ie=y?"PUT":"POST",ae=y?`${a}/${v.id}`:a,N=await k(ae,{method:ie,body:JSON.stringify(I)});N.ok?(Z(y?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),U(),M()):(G(N.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=y?"Simpan Perubahan":`Tambah ${n}`)}})}function le(v){Me(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let y=await k(`${a}/${v.id}`,{method:"DELETE"});y.ok?(Z(`${n} berhasil dihapus.`),M()):G(y.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return M(),M}O();R();async function Ht(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let r=await K(),a=await ee(),o=e?e.get("dash_filter"):null;console.log("RAW",await we()),console.log("OPTIONS",a);let c=n=>n&&!a.find(p=>p.value===n)?[...a,{value:n,label:n}]:a,s=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],i=n=>n&&!s.includes(n)?[...s,n]:s;Kt({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",onDataLoaded:n=>{if(o==="reliever"){let p=new Date,m=p.getFullYear(),l=String(p.getMonth()+1).padStart(2,"0");return n.filter(d=>{if(String(d.status||"").toLowerCase()!=="done")return!1;let u=d.backup_date||"";if(u.includes("/")){let b=u.split("/");if(b.length===3&&(b[2].length===4?b[2]:`20${b[2]}`)==m&&b[1].padStart(2,"0")==l)return!0}else if(u.includes("-")&&u.startsWith(`${m}-${l}`))return!0;return!1})}return n},columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:n=>de(n)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:n=>window.formatDate(n)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:n=>n?`<span class="badge badge-info">${n}</span>`:"-"},{key:"status",label:"Status",render:n=>j(n)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:s},{type:"select",name:"branch_id",label:"Cabang",options:r},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:r,value:n?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:c(n?.original_fc_name),value:n?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:i(n?.reliever_name),value:n?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:n?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:n?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:n?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:n?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let n=await k(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let p=n.data.data.map(m=>({Cabang:m.branch_name||"","Nama Facility care":m.original_fc_name||"",Periode:m.period||"",Relifer:m.reliever_name||"","Tanggal Back Up":m.backup_date||"","Tanggal Selesai":m.completion_date||"",Keterangan:m.reason||"",Shift:m.shift||"",Status:m.status||""}));p.length===0&&p.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),L(p,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async n=>{let m=(await k("/api/branches?all=1")).data?.data||[],l=b=>{if(!b)return null;let g=String(b||"").toLowerCase(),h=m.find(f=>String(f.full_name||"").toLowerCase()===g||String(f.code||"").toLowerCase()===g||String(f.name||"").toLowerCase()===g);return h?h.id:null},d=n.map(b=>({branch_name:String(b.Cabang||"").trim(),backup_date:String(b["Tanggal Back Up"]||b["Tanggal Backup"]||"").trim(),original_fc_name:String(b["Nama Facility care"]||b["FC Digantikan"]||"").trim(),reliever_name:String(b.Relifer||b.Reliefer||"").trim(),period:String(b.Periode||"").trim(),reason:String(b.Keterangan||"").trim(),shift:String(b.Shift||"").trim(),completion_date:String(b["Tanggal Selesai"]||"").trim(),status:String(b.Status||"").trim()})).filter(b=>b.reliever_name&&b.backup_date),u=await k("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:d,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}O();R();async function qt(t){let e=await K(),r=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));A({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>de(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>j(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),c=await k(`/api/reports/inspection?limit=10000&${o}`);if(c.ok){let s=c.data.data.map(i=>({Cabang:i.branch_name||"",Periode:i.period||"",Tanggal:i.inspection_date||"","Point FC":i.fc_score!==null&&i.fc_score!==void 0?i.fc_score:"","Point SPV":i.spv_score!==null&&i.spv_score!==void 0?i.spv_score:"",Status:i.status||"","Link Dokumen":i.document_link||""}));L(s,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let o=n=>{if(!n)return null;let p=String(n||"").toLowerCase(),m=e.find(l=>String(l.label||"").toLowerCase()===p);return m?m.value:null},c=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let p=String(n).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let l=Number(p);if(l>2e4&&l<99999){let d=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let m=p.split(/[\/\-\.]/);if(m.length===3){let[l,d,u]=m.map(b=>b.trim());if(l.length===4&&d.length<=2&&u.length<=2)return`${l}-${d.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&d.length<=2&&l.length<=2)return`${u}-${d.padStart(2,"0")}-${l.padStart(2,"0")}`}return p},s=a.map(n=>({branch_id:o(String(n.Cabang||"").trim()),period:String(n.Periode||"").trim(),inspection_date:c(n.Tanggal),fc_score:n["Point FC"]!==void 0&&n["Point FC"]!==""?Number(n["Point FC"]):null,spv_score:n["Point SPV"]!==void 0&&n["Point SPV"]!==""?Number(n["Point SPV"]):null,status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.inspection_date),i=await k("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:s,onDuplicate:"update"})});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}O();R();async function jt(t){let e=await K(),r=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));A({container:t,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>de(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>j(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),c=await k(`/api/reports/cleaning?limit=10000&${o}`);if(c.ok){let s=c.data.data.map(i=>({Cabang:i.branch_name||"",Jenis:i.activity_type||"",Periode:i.period||"",Tanggal:i.activity_date||"",Status:i.status||"","Link Dokumen":i.document_link||""}));L(s,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let o=n=>{if(!n)return null;let p=String(n||"").toLowerCase(),m=e.find(l=>String(l.label||"").toLowerCase()===p);return m?m.value:null},c=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let p=String(n).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let l=Number(p);if(l>2e4&&l<99999){let d=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let m=p.split(/[\/\-\.]/);if(m.length===3){let[l,d,u]=m.map(b=>b.trim());if(l.length===4&&d.length<=2&&u.length<=2)return`${l}-${d.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&d.length<=2&&l.length<=2)return`${u}-${d.padStart(2,"0")}-${l.padStart(2,"0")}`}return p},s=a.map(n=>({branch_id:o(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:c(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),i=await k("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:s,onDuplicate:"update"})});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}O();R();async function Jt(t){let e=await K(),r=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));A({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge badge-warning">${a}</span>`},{key:"period",label:"Periode",render:a=>de(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>j(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:a=>a||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:a?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),c=await k(`/api/reports/fogging?limit=10000&${o}`);if(c.ok){let s=c.data.data.map(i=>({Cabang:i.branch_name||"",Jenis:i.activity_type||"Fogging",Periode:i.period||"",Tanggal:i.activity_date||"",Status:i.status||"","Link Dokumen":i.document_link||""}));L(s,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async a=>{let o=n=>{if(!n)return null;let p=String(n||"").toLowerCase(),m=e.find(l=>String(l.label||"").toLowerCase()===p);return m?m.value:null},c=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let p=String(n).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let l=Number(p);if(l>2e4&&l<99999){let d=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let m=p.split(/[\/\-\.]/);if(m.length===3){let[l,d,u]=m.map(b=>b.trim());if(l.length===4&&d.length<=2&&u.length<=2)return`${l}-${d.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&d.length<=2&&l.length<=2)return`${u}-${d.padStart(2,"0")}-${l.padStart(2,"0")}`}return p},s=a.map(n=>({branch_id:o(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"Fogging").trim(),period:String(n.Periode||"").trim(),activity_date:c(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.activity_date),i=await k("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(s)});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}O();R();async function Ut(t){let e=await K(),r=await ee(),a=r,o=s=>s&&!r.find(i=>i.value===s)?[...r,{value:s,label:s}]:r,c=s=>s&&!a.find(i=>i.value===s)?[...a,{value:s,label:s}]:a;A({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:s=>`<span title="${s||""}">${s?.length>60?s.slice(0,60)+"\u2026":s||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:s=>window.formatDate(s)},{key:"status",label:"Status",render:s=>j(s)},{key:"notes",label:"Keterangan",render:s=>s?.length>40?s.slice(0,40)+"\u2026":s||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]}],formFields:s=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:s?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:c(s?.pic),value:s?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:s?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:s?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:s?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:s?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:s?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async s=>{let i=new URLSearchParams(s||{}).toString(),n=await k(`/api/reports/basecamp?limit=10000&${i}`);if(n.ok){let p=n.data.data.map(m=>({"Tgl Info":m.info_date||"",Cabang:m.branch_name||"",Permasalahan:m.problem||"",PIC:m.pic||"","Tgl Done":m.done_date||"",Status:m.status||"",Keterangan:m.notes||""}));L(p,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async s=>{let i=l=>{if(!l)return null;let d=String(l||"").toLowerCase(),u=e.find(b=>String(b.label||"").toLowerCase()===d);return u?u.value:null},n=l=>{if(l==null||l==="")return"";if(l instanceof Date&&!isNaN(l.getTime()))return l.toISOString().slice(0,10);let d=String(l).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let b=Number(d);if(b>2e4&&b<99999){let g=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}let u=d.split(/[\/\-\.]/);if(u.length===3){let[b,g,h]=u.map(f=>f.trim());if(b.length===4&&g.length<=2&&h.length<=2)return`${b}-${g.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&g.length<=2&&b.length<=2)return`${h}-${g.padStart(2,"0")}-${b.padStart(2,"0")}`}return d},p=s.map(l=>({info_date:n(l["Tgl Info"]||l["Tanggal Info"]),branch_id:i(String(l.Cabang||"").trim()),problem:String(l.Permasalahan||"").trim(),pic:String(l.PIC||"").trim(),done_date:n(l["Tgl Done"]||l["Tanggal Done"]),status:String(l.Status||"").trim(),notes:String(l.Keterangan||l.Catatan||"").trim()})).filter(l=>l.info_date&&l.branch_id&&l.problem),m=await k("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(p)});if(!m.ok)throw new Error(m.data?.error||"Import gagal");return m.data}}})}async function Gt(t){A({container:t,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(O(),ke)),o=await a(`/api/sop?limit=10000&${r}`);if(o.ok){let c=o.data.data.map(i=>({"Nama SOP":i.name||"",Kategori:i.category||"",Dokumen:i.document_link||"",Catatan:i.notes||i.description||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(R(),oe));s(c,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:r}=await Promise.resolve().then(()=>(R(),oe));r(e,"Template_Import_SOP")},onImport:async e=>{let r=e.map(c=>({name:String(c["Nama SOP"]||"").trim(),category:String(c.Kategori||"").trim(),document_link:String(c.Dokumen||"").trim(),description:String(c.Catatan||"").trim()})).filter(c=>c.name),{apiFetch:a}=await Promise.resolve().then(()=>(O(),ke)),o=await a("/api/sop/import",{method:"POST",body:JSON.stringify(r)});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function zt(t){A({container:t,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(O(),ke)),o=await a(`/api/checklist?limit=10000&${r}`);if(o.ok){let c=o.data.data.map(i=>({"Nama Checklist":i.name||"",Kategori:i.category||"",Dokumen:i.document_link||"",Deskripsi:i.description||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(R(),oe));s(c,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:r}=await Promise.resolve().then(()=>(R(),oe));r(e,"Template_Import_Checklist")},onImport:async e=>{let r=e.map(c=>({name:String(c["Nama Checklist"]||"").trim(),category:String(c.Kategori||"").trim(),document_link:String(c.Dokumen||"").trim(),description:String(c.Deskripsi||"").trim()})).filter(c=>c.name),{apiFetch:a}=await Promise.resolve().then(()=>(O(),ke)),o=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(r)});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}O();ve();R();async function yt(t,e="forms"){if(e==="supply")return Fa(t);Na(t)}function Na(t){A({container:t,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await k(`/api/forms?limit=10000&${r}`);a.data?.data?L(a.data.data,"Data_Master_Form"):G("Gagal export data master form")},onImport:async e=>{let r=await k("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!r.ok)throw new Error(r.data?.error||"Import failed");return r.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function Fa(t){let r=((await k("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));A({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,o)=>o.branch_name_ref||o.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let o=JSON.parse(a);return Array.isArray(o)?o.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let o=JSON.parse(a);return Array.isArray(o)?o.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>j(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let o=a?.tools_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}let c=a?.chemical_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:a?.branch_id&&!r.find(s=>s.value==a.branch_id)?[...r,{value:a.branch_id,label:a.branch_name||a.branch_id}]:r,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:o},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:c},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),c=await k(`/api/reports/supply?limit=10000&${o}`);if(c.ok){let s=c.data.data.map(i=>{let n=i.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let p=i.chemical_items;try{p=Array.isArray(JSON.parse(p))?JSON.parse(p).join(", "):p}catch{}return{Waktu:i.submitted_at||"",Pengirim:i.submitter_name||"",Cabang:i.branch_name_ref||i.branch_name||"","Alat/Barang":n||"",Chemical:p||"",Catatan:i.additional_notes||"",Status:i.status||"","Diproses Oleh":i.processed_by||""}});L(s,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let c=(await k("/api/branches?all=1")).data?.data||[],s=m=>{if(!m)return null;let l=String(m||"").toLowerCase(),d=c.find(u=>String(u.full_name||"").toLowerCase()===l||String(u.code||"").toLowerCase()===l||String(u.name||"").toLowerCase()===l);return d?d.id:null},i=m=>{if(m==null||m==="")return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let l=String(m).trim();if(l===""||l==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);if(/^\d{4,5}$/.test(l)){let u=Number(l);if(u>2e4&&u<99999){let b=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}let d=l.split(/[\/\-\.]/);if(d.length===3){let[u,b,g]=d.map(h=>h.trim());if(u.length===4&&b.length<=2&&g.length<=2)return`${u}-${b.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&b.length<=2&&u.length<=2)return`${g}-${b.padStart(2,"0")}-${u.padStart(2,"0")}`}return l},n=a.map(m=>({submitted_at:i(m.Waktu||m.Tanggal),submitter_name:String(m.Pengirim||"").trim(),branch_id:s(String(m.Cabang||"").trim()),tools_items:String(m["Alat/Barang"]||m.Alat||"").trim(),chemical_items:String(m.Chemical||"").trim(),additional_notes:String(m.Catatan||m.Keterangan||"").trim(),status:String(m.Status||"").trim(),processed_by:String(m["Diproses Oleh"]||m.PIC||"").trim()})).filter(m=>m.submitted_at&&m.submitter_name&&m.branch_id),p=await k("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,o)=>{let c=ce({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(s,i)=>{let n=s.querySelector("#supply-status").value,p=s.querySelector("#supply-processed-by").value;(await k(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:p})})).ok?(Z("Status diperbarui."),i(),o()):G("Gagal update status.")}})}}]})}O();R();async function Qt(t){let e=fe();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}A({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:r=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[r]||"badge-neutral"}">${r}</span>`},{key:"is_active",label:"Status",render:r=>r?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:r=>r?new Date(r).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:r=>{let a=!!r;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:r?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:r?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:r?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:r?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?r?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let r=await k(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let a=r.data.data.map(o=>({"Nama Lengkap":o.full_name||"",Username:o.username||"",Email:o.email||"",Role:o.role||"",Status:o.is_active?"Aktif":"Nonaktif"}));L(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async r=>{let a=r.map(c=>({full_name:String(c["Nama Lengkap"]||"").trim(),username:String(c.Username||"").trim(),email:String(c.Email||"").trim(),role:String(c.Role||"").trim()||"viewer",password:String(c.Password||"").trim()})).filter(c=>c.username&&c.password&&c.email&&c.full_name),o=await k("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}}})}O();R();async function Wt(t){A({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await k(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)L(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let r=e.map(o=>({code:String(o["Kode Cabang"]||"").trim(),name:String(o["Nama Pendek"]||"").trim(),full_name:String(o["Nama Lengkap"]||"").trim(),city:String(o.Kota||"").trim()})).filter(o=>o.code&&o.name),a=await k("/api/branches/import",{method:"POST",body:JSON.stringify(r)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}O();var me={schedule:{bg:"#3B82F6",text:"#fff",icon:"\u{1F5D3}",label:"Jadwal"},issue:{bg:"#FB923C",text:"#fff",icon:"\u26A0\uFE0F",label:"Permasalahan"},reliever:{bg:"#34D399",text:"#fff",icon:"\u{1F504}",label:"Reliefer"},training:{bg:"#A78BFA",text:"#fff",icon:"\u{1F393}",label:"Training"},contract_expiry:{bg:"#F87171",text:"#fff",icon:"\u{1F4C4}",label:"Kontrak Habis"},one_on_one:{bg:"#F472B6",text:"#fff",icon:"\u{1F4AC}",label:"One on One"},inspection:{bg:"#38BDF8",text:"#fff",icon:"\u{1F50D}",label:"Inspeksi"},cleaning:{bg:"#2DD4BF",text:"#fff",icon:"\u{1F9F9}",label:"Cleaning"},fogging:{bg:"#818CF8",text:"#fff",icon:"\u{1F4A8}",label:"Fogging"},basecamp:{bg:"#A8A29E",text:"#fff",icon:"\u{1F3D5}",label:"Basecamp"},supply:{bg:"#60A5FA",text:"#fff",icon:"\u{1F4E6}",label:"Permintaan"}};function Yt(t){return me[t]||{bg:"#6B7280",icon:"\u{1F4CC}",label:t}}async function Vt(t){let e=new Date,r=[],a=[{value:"schedule",...me.schedule},{value:"issue",...me.issue},{value:"reliever",...me.reliever},{value:"training",...me.training},{value:"contract_expiry",...me.contract_expiry},{value:"one_on_one",...me.one_on_one},{value:"inspection",...me.inspection},{value:"cleaning",...me.cleaning},{value:"fogging",...me.fogging},{value:"basecamp",...me.basecamp}];t.innerHTML=`
    <style>
      .cal-cell { transition: transform 0.15s ease, box-shadow 0.15s ease; }
      .cal-cell:hover { transform: scale(1.01); box-shadow: 0 4px 16px rgba(0,0,0,0.10); z-index: 2; position: relative; }
      .cal-badge {
        display: block;
        padding: 3px 7px;
        border-radius: 8px;
        font-size: 0.64rem;
        font-weight: 700;
        color: #fff;
        margin-bottom: 3px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        max-width: 100%;
        cursor: pointer;
        line-height: 1.35;
        letter-spacing: 0.01em;
      }
      .cal-badge:hover { filter: brightness(1.12); transform: scale(1.03); }
      .cal-more-btn {
        display: inline-block;
        margin-top: 2px;
        font-size: 0.6rem;
        font-weight: 700;
        color: #374151;
        background: #F3F4F6;
        border-radius: 6px;
        padding: 2px 6px;
        cursor: pointer;
      }
      .cal-popup-overlay {
        position: fixed; inset: 0; background: rgba(0,0,0,0.45);
        z-index: 1000; display: flex; align-items: center; justify-content: center;
        animation: calFadeIn 0.18s ease;
      }
      @keyframes calFadeIn { from { opacity: 0; } to { opacity: 1; } }
      .cal-popup {
        background: #fff; border-radius: 16px; width: min(480px, 95vw);
        max-height: 80vh; display: flex; flex-direction: column;
        box-shadow: 0 20px 60px rgba(0,0,0,0.25);
        animation: calSlideUp 0.2s ease;
      }
      @keyframes calSlideUp { from { transform: translateY(24px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
      .cal-popup-head {
        display: flex; justify-content: space-between; align-items: center;
        padding: 16px 20px; border-bottom: 1px solid #E5E7EB;
      }
      .cal-popup-title { font-size: 1rem; font-weight: 800; color: #111827; }
      .cal-popup-date { font-size: 0.78rem; color: #6B7280; margin-top: 2px; }
      .cal-popup-close {
        width: 32px; height: 32px; border-radius: 50%; border: none;
        background: #F3F4F6; cursor: pointer; font-size: 1.1rem;
        display: flex; align-items: center; justify-content: center;
        transition: background 0.15s;
      }
      .cal-popup-close:hover { background: #E5E7EB; }
      .cal-popup-body { overflow-y: auto; padding: 16px; display: flex; flex-direction: column; gap: 8px; }
      .cal-popup-event {
        display: flex; gap: 12px; align-items: flex-start;
        padding: 10px 14px; border-radius: 10px; background: #F9FAFB;
        border-left: 5px solid;
      }
      .cal-popup-icon { font-size: 1.2rem; line-height: 1; flex-shrink: 0; margin-top: 1px; }
      .cal-popup-info { flex: 1; min-width: 0; }
      .cal-popup-type { font-size: 0.62rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; opacity: 0.7; margin-bottom: 2px; }
      .cal-popup-event-title { font-size: 0.85rem; font-weight: 700; color: #111827; word-break: break-word; }
      .cal-popup-branch { font-size: 0.73rem; color: #6B7280; margin-top: 2px; }
      .cal-popup-status { font-size: 0.7rem; font-weight: 700; margin-top: 4px; }
      .cal-popup-sisa { font-size: 0.7rem; color: #6366F1; font-weight: 700; margin-top: 2px; }

      /* Legend pills di header */
      .cal-legend-pill {
        display: inline-flex; align-items: center; gap: 5px;
        padding: 4px 10px; border-radius: 20px; font-size: 0.72rem;
        font-weight: 700; color: #fff; cursor: pointer; user-select: none;
        transition: opacity 0.15s, transform 0.12s;
        white-space: nowrap;
      }
      .cal-legend-pill.inactive { opacity: 0.35; }
      .cal-legend-pill input[type="checkbox"] { display: none; }
    </style>

    <div class="page-header">
      <h1 class="page-title">\u{1F4C5} Kalender FCMS</h1>
    </div>

    <div class="card" style="overflow:visible;">
      <!-- Header Nav + Filters -->
      <div style="padding:14px 20px;border-bottom:1px solid var(--border,#e5e7eb);display:flex;flex-wrap:wrap;gap:12px;align-items:center;">
        <!-- Navigasi bulan -->
        <div style="display:flex;align-items:center;gap:8px;flex-shrink:0;">
          <button class="btn btn-ghost btn-sm" id="cal-prev" style="font-size:1rem;padding:6px 12px;">\u2039</button>
          <span id="cal-month-label" style="min-width:170px;text-align:center;font-size:1.1rem;font-weight:800;color:#111827;"></span>
          <button class="btn btn-ghost btn-sm" id="cal-next" style="font-size:1rem;padding:6px 12px;">\u203A</button>
        </div>

        <!-- Legend / filter pills -->
        <div style="display:flex;flex-wrap:wrap;gap:6px;align-items:center;flex:1;">
          ${a.map(i=>`
            <label class="cal-legend-pill" style="background:${i.bg};" data-filter="${i.value}" title="${i.label}">
              <input type="checkbox" class="cal-filter" value="${i.value}" checked>
              <span>${i.icon} ${i.label}</span>
            </label>
          `).join("")}
        </div>
      </div>

      <!-- Grid -->
      <div class="card-body p-0">
        <div id="calendar-grid" style="min-height:420px;"></div>
      </div>
    </div>
  `,t.querySelectorAll(".cal-legend-pill").forEach(i=>{let n=i.querySelector('input[type="checkbox"]');i.addEventListener("click",p=>{p.preventDefault(),n.checked=!n.checked,i.classList.toggle("inactive",!n.checked),c()})}),document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),c()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),c()});async function o(){try{let i=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0");r=(await k(`/api/dashboard/calendar?month=${i}-${n}`)).data?.data||[]}catch(i){console.warn("[Calendar] Gagal memuat events:",i),r=[]}}async function c(){let i=document.getElementById("calendar-grid");if(i){i.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:1px;background:#E5E7EB;">
      ${Array(42).fill('<div style="background:#F9FAFB;min-height:90px;"></div>').join("")}
    </div>`,await o();try{let n=e.getFullYear(),p=e.getMonth(),m=document.getElementById("cal-month-label");m&&(m.textContent=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}));let l=new Set(Array.from(t.querySelectorAll(".cal-filter:checked")).map(S=>S.value)),d=r.filter(S=>l.has(S.type)),u={};d.forEach(S=>{let w=(S.event_date||"").slice(0,10);w&&(u[w]||(u[w]=[]),u[w].push(S))});let b=new Date(n,p,1).getDay(),g=new Date(n,p+1,0).getDate(),h=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],f=new Date().toISOString().slice(0,10),x='<div style="display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:1px;background:#E5E7EB;border-radius:0 0 12px 12px;overflow:hidden;width:100%;">';h.forEach((S,w)=>{x+=`<div style="background:#F1F5F9;padding:10px 4px;text-align:center;font-size:0.68rem;font-weight:800;color:${w===0||w===6?"#EF4444":"#6B7280"};text-transform:uppercase;letter-spacing:0.06em;">${S}</div>`});for(let S=0;S<b;S++)x+='<div style="background:#FAFAFA;min-height:90px;"></div>';for(let S=1;S<=g;S++){let w=`${n}-${String(p+1).padStart(2,"0")}-${String(S).padStart(2,"0")}`,B=u[w]||[],E=w===f,q=new Date(n,p,S).getDay(),M=q===0||q===6,ge=4;x+=`
          <div class="cal-cell" data-date="${w}"
               style="background:${E?"#EFF6FF":"#fff"};min-height:90px;padding:6px 5px 5px;overflow:hidden;
                      border-top:${E?"3px solid #2563EB":"1px solid transparent"};
                      cursor:${B.length?"pointer":"default"};">
            <!-- Nomor tanggal -->
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
              <span style="font-size:${E?"0.9rem":"0.78rem"};font-weight:${E?"900":"600"};
                    color:${E?"#fff":M?"#EF4444":"#374151"};
                    ${E?"background:#2563EB;border-radius:50%;width:22px;height:22px;display:flex;align-items:center;justify-content:center;":""}">
                ${S}
              </span>
              ${B.length>ge?`<span class="cal-more-btn">+${B.length-ge}</span>`:""}
            </div>

            <!-- Badge events -->
            <div style="display:flex;flex-direction:column;gap:3px;overflow:hidden;">
              ${B.slice(0,ge).map(se=>{let le=Yt(se.type),v=Ma(se.title||se.branch_name||le.label,18);return`<div class="cal-badge" style="background:${le.bg};" title="${De(se.title||le.label)} \u2014 ${De(se.branch_name||"")}">
                  ${le.icon} ${De(v)}
                </div>`}).join("")}
            </div>
          </div>`}let C=(b+g)%7;if(C!==0)for(let S=0;S<7-C;S++)x+='<div style="background:#FAFAFA;min-height:90px;"></div>';x+="</div>",i.innerHTML=x,i.querySelectorAll(".cal-cell[data-date]").forEach(S=>{S.addEventListener("click",()=>{let w=S.dataset.date,B=u[w]||[];B.length&&s(w,B)})})}catch(n){console.error("[Calendar] Render error:",n),i.innerHTML=`<div style="padding:60px;text-align:center;color:#9CA3AF;">
        <div style="font-size:2.5rem;margin-bottom:12px;">\u{1F4C5}</div>
        <div style="font-size:1rem;font-weight:700;">Gagal memuat kalender. Silakan refresh.</div>
      </div>`}}}function s(i,n){document.getElementById("cal-popup-overlay")?.remove();let p=new Date(i+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),m=n.map(d=>{let u=Yt(d.type);return`
        <div class="cal-popup-event" style="border-left-color:${u.bg};">
          <div class="cal-popup-icon">${u.icon}</div>
          <div class="cal-popup-info">
            <div class="cal-popup-type" style="color:${u.bg};">${u.label}</div>
            <div class="cal-popup-event-title">${De(d.title||"-")}</div>
            ${d.branch_name?`<div class="cal-popup-branch">\u{1F4CD} ${De(d.branch_name)}</div>`:""}
            ${d.status?`<div class="cal-popup-status" style="color:${d.status==="Done"||d.status==="Selesai"?"#10B981":"#F59E0B"};">\u25CF ${De(d.status)}</div>`:""}
            ${d.days_remaining!==void 0?`<div class="cal-popup-sisa">\u23F3 Sisa: ${d.days_remaining} hari</div>`:""}
          </div>
          <!-- Dot warna -->
          <div style="width:10px;height:10px;border-radius:50%;background:${u.bg};flex-shrink:0;margin-top:3px;"></div>
        </div>`}).join(""),l=document.createElement("div");l.id="cal-popup-overlay",l.className="cal-popup-overlay",l.innerHTML=`
      <div class="cal-popup">
        <div class="cal-popup-head">
          <div>
            <div class="cal-popup-title">\u{1F4C5} ${n.length} Event</div>
            <div class="cal-popup-date">${p}</div>
          </div>
          <button class="cal-popup-close" id="cal-popup-close-btn">\u2715</button>
        </div>
        <div class="cal-popup-body">${m}</div>
      </div>`,document.body.appendChild(l),document.getElementById("cal-popup-close-btn").addEventListener("click",()=>l.remove()),l.addEventListener("click",d=>{d.target===l&&l.remove()})}c()}function Ma(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function De(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}O();async function Xt(t){let e=fe(),r=(e?.full_name||e?.username||"U")[0].toUpperCase(),o={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${o},${o}99)">
            ${r}
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
  `;let c=localStorage.getItem("fm_token"),s=document.getElementById("session-info");if(c&&s)try{let i=JSON.parse(atob(c.split(".")[1])),n=new Date(i.exp*1e3);s.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{s.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async i=>{i.preventDefault();let n=document.getElementById("pwd-error"),p=document.getElementById("pwd-success"),m=document.getElementById("btn-save-pwd");n.style.display="none",p.style.display="none";let l=i.target,d=l.current_password.value,u=l.new_password.value,b=l.confirm_password.value;if(u!==b){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(u.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}m.disabled=!0,m.textContent="\u23F3 Menyimpan...";let g=await k("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:d,new_password:u})});m.disabled=!1,m.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',g.ok?(p.textContent="\u2705 Password berhasil diubah.",p.style.display="block",l.reset(),Z("Password berhasil diubah.")):(n.textContent=g.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}O();var et={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function te(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let o=Number(e);if(o>2e4&&o<99999){let c=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(c.getTime())?null:c.toISOString().slice(0,10)}}let r=e.split(/[\/\-\.]/);if(r.length===3){let[o,c,s]=r.map(m=>m.trim()),i=Number(o),n=Number(c),p=Number(s);if(o.length===4&&i>1900)return`${o}-${c.padStart(2,"0")}-${s.padStart(2,"0")}`;if(s.length===4&&p>1900)return i>12?`${s}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`:n>12?`${s}-${o.padStart(2,"0")}-${c.padStart(2,"0")}`:`${s}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`;if(s.length===2&&!isNaN(p)){let m=p>=50?`19${s}`:`20${s}`;return i>12?`${m}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`:`${m}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function Zt(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Oa={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",phone:t["No. Hp"],join_date:te(t["Tanggal Masuk"]),status:t.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:te(t["Tanggal Mulai"]),end_date:te(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:te(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:te(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:te(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:te(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:te(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:te(t["Tanggal Target"]||t["Tgl Target"]),completion_date:te(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:te(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:te(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:te(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:te(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:te(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:te(t["Tanggal Back Up"]),completion_date:te(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:te(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:te(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Ra(t,e){let r=et[t];if(!r)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Oa[r.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let o=[],c=[],s=[];return e.filter(n=>!Zt(n)).forEach((n,p)=>{let m=e.indexOf(n)+2,l=[];a.required.forEach(({key:u,label:b})=>{let g=n[u];if(g==null||String(g).trim()===""){let h=Object.keys(n).filter(f=>f.trim()).join(", ");l.push({column:b,originalValue:g||"",reason:`Kolom "${b}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${h.slice(0,120)}`})}});let d=a.map(n);l.length>0?c.push({row:m,data:d,raw:n,errors:l}):(o.push(n),s.push(d))}),{valid:o,errors:c,mapped:s}}function ea(t){let e=[];return t.SheetNames.forEach(r=>{let a=et[r];if(!a)return;let o=t.Sheets[r],c=window.XLSX.utils.sheet_to_json(o,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),s=Ra(r,c),i=c.filter(n=>!Zt(n));e.push({sheetName:r,module:a.module,label:a.label,total:i.length,valid:s.mapped.length,errorCount:s.errors.length,errors:s.errors,mapped:s.mapped,skipped:!1})}),e}function ta(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,o])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(o),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function aa(t){let e=window.XLSX,r=e.utils.book_new(),a=!1;return t.forEach(o=>{if(!o.errors||o.errors.length===0)return;a=!0;let c=o.errors.map(i=>({"No. Baris":i.row,"Kolom Gagal":(i.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(i.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(i.data||{}).map(([n,p])=>[n,p??""]))})),s=e.utils.json_to_sheet(c);e.utils.book_append_sheet(r,s,o.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(r,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Ka=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function na(t){t.innerHTML=`
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
              ${Object.entries(et).map(([g,{label:h}])=>`<span class="import-sheet-tag">\u{1F4C4} ${g} \u2192 ${h}</span>`).join("")}
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
  `;let e=null,r=null,a=0,o={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function c(g){Object.entries(o).forEach(([h,f])=>{f.style.display=h===g?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let g=document.getElementById("btn-backup-db");g.disabled=!0,g.textContent="\u23F3 Memproses Backup...";try{let h=await k("/api/import/backup");if(h.ok){if(!window.XLSX){G("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let f=window.XLSX,x=f.utils.book_new();Object.entries(h.data.database).forEach(([_,C])=>{let S=C.length>0?C:[{}],w=f.utils.json_to_sheet(S);f.utils.book_append_sheet(x,w,_.substring(0,31))}),f.writeFile(x,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),Z("Backup berhasil diunduh!")}else G("Gagal memproses backup: "+(h.data?.error||"Unknown error"))}catch(h){G("Gagal memproses backup: "+h.message)}finally{g.disabled=!1,g.textContent="\u{1F4E6} Backup Database"}});let s=document.getElementById("btn-sync-google");s&&s.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let g=s.innerHTML;s.innerHTML='<span class="spinner"></span> Menyinkronkan...',s.disabled=!0;try{let h=await k("/api/sync/google-sheets",{method:"POST"});h.ok?alert("Sinkronisasi Berhasil: "+(h.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(h.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{s.innerHTML=g,s.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{ta(),Z("Template Excel berhasil didownload!")});let i=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",g=>{g.stopPropagation(),i.click()}),i.addEventListener("change",g=>{g.target.files[0]&&p(g.target.files[0])}),n.addEventListener("dragover",g=>{g.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",g=>{g.preventDefault(),n.classList.remove("drag-over");let h=g.dataTransfer.files[0];h&&h.name.match(/\.xlsx?$/i)?p(h):G("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,i.value="",document.getElementById("file-info").style.display="none",n.style.display="",c("upload")});async function p(g){e=g,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${g.name} (${(g.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await m(g)}async function m(g){c("validating");let h=document.getElementById("validation-status"),f=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");h.textContent="Membaca file Excel...",f.style.width="20%",await je(200);let x=await g.arrayBuffer(),_=window.XLSX.read(x,{type:"array",cellDates:!0});h.textContent=`Memvalidasi ${_.SheetNames.length} sheet...`,f.style.width="50%",await je(100),r=ea(_),f.style.width="100%",h.textContent="Validasi selesai!",await je(300),l()}catch(x){c("upload"),G("Gagal memproses file: "+x.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function l(){c("preview");let g=r.filter(E=>!E.skipped).length,h=r.reduce((E,q)=>E+q.total,0),f=r.reduce((E,q)=>E+q.valid,0),x=r.reduce((E,q)=>E+q.errorCount,0),_=h>0?Math.round(f/h*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${g} sheet</span>
      <span class="badge badge-secondary">${h} baris</span>
      <span class="badge badge-success">${f} valid (${_}%)</span>
      ${x>0?`<span class="badge badge-danger">${x} error</span>`:""}
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
          ${r.map((E,q)=>`
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
                ${E.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${q}">\u{1F50D} ${E.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,C.querySelectorAll(".btn-detail-error").forEach(E=>{E.addEventListener("click",()=>{let q=r[Number(E.dataset.idx)];d(q)})});let S=document.getElementById("error-detail-section"),w=document.getElementById("error-detail-container");w.innerHTML="",S.style.display="none";let B=document.getElementById("btn-start-import");f===0?(B.disabled=!0,B.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(B.disabled=!1,x>0?(B.innerHTML=`\u{1F680} Import ${f} Data Valid (${x} dilewati)`,B.title="Baris error akan dilewati, baris valid tetap diimport"):B.innerHTML=`\u{1F680} Mulai Import ${f} Data`)}function d(g){let h=document.getElementById("error-detail-section"),f=document.getElementById("error-detail-container");h.style.display="";let x=g.errors.slice(0,100).map(_=>(Array.isArray(_.errors)?_.errors:[]).map(S=>{let w=typeof S=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${_.row}</span></td>
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
            <tbody>${x||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${g.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,h.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{c("upload"),document.getElementById("file-info").style.display="none",n.style.display="",e=null,i.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!r)return;aa(r)?Z("Log error berhasil didownload."):Z("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let g=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";u(g)});async function u(g){c("importing"),a=Date.now();let h=[];Ka.forEach(S=>{let w=r?.find(B=>B.module===S&&B.mapped?.length>0);w&&h.push(w)});let f=document.getElementById("import-steps-list");f.innerHTML=h.map(S=>`
      <div class="import-step-item" id="step-item-${S.module}">
        <span class="step-item-icon" id="step-icon-${S.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${S.label} <span class="step-item-count">(${S.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${S.module}"></span>
      </div>
    `).join("");let x=document.getElementById("import-bar"),_=document.getElementById("import-current-status"),C={totalSheets:h.length,totalRows:h.reduce((S,w)=>S+w.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let S=0;S<h.length;S++){let w=h[S],B=document.getElementById(`step-icon-${w.module}`),E=document.getElementById(`step-status-${w.module}`);B.textContent="\u{1F504}",E.textContent="Mengimport...",_.textContent=`Mengimport ${w.label}...`,x.style.width=`${Math.round(S/h.length*100)}%`;try{let q=await k(`/api/import/${w.module}`,{method:"POST",body:JSON.stringify({rows:w.mapped,onDuplicate:g})});if(q.ok){let M=q.data;C.inserted+=M.inserted||0,C.skipped+=M.skipped||0,C.moduleResults.push({label:w.label,inserted:M.inserted||0,skipped:M.skipped||0,status:"ok"}),B.textContent="\u2705",E.innerHTML=`<span class="badge badge-success">${M.inserted||0} berhasil</span>${M.skipped>0?` <span class="badge badge-neutral">${M.skipped} skip</span>`:""}`}else C.failed++,C.moduleResults.push({label:w.label,inserted:0,skipped:0,status:"error",error:q.data?.error}),B.textContent="\u274C",E.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(q){C.failed++,C.moduleResults.push({label:w.label,inserted:0,skipped:0,status:"error",error:q.message}),B.textContent="\u274C",E.innerHTML='<span class="badge badge-danger">Gagal</span>'}await je(150)}x.style.width="100%",_.textContent="Selesai!",await je(400),b(C)}function b(g){c("summary");let h=((Date.now()-a)/1e3).toFixed(1),f=g.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${f?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${f?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
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
        <div class="stat-value">${h}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${g.moduleResults.map(x=>`
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,r=null,i.value="",document.getElementById("file-info").style.display="none",n.style.display="",c("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function je(t){return new Promise(e=>setTimeout(e,t))}O();var tt=[],ia=[];async function ra(t){tt=await K(),ia=await ee(),A({container:t,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:tt}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await k(`/api/sp?limit=10000&${r}`);if(a.ok){let o=a.data.data.map(s=>({"Nama Karyawan":s.employee_name||"",Divisi:s.division||"",Cabang:s.branch_name||"","Tanggal Sp":s.tanggal||"","Akhir Sp":s.akhir_sp||"","Jenis Sp":s.sp_type||"","Link Document / Foto":s.document_link||""})),{downloadExcel:c}=await Promise.resolve().then(()=>(R(),oe));c(o,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(R(),oe));r(e,"Template_Import_SP")},onImport:async e=>{let r=s=>{if(!s)return null;let i=String(s||"").toLowerCase(),n=tt.find(p=>String(p.label||"").toLowerCase()===i);return n?n.value:null},a=s=>{if(!s)return"";if(s instanceof Date&&!isNaN(s.getTime()))return s.toISOString().slice(0,10);let i=String(s).trim();if(/^\d{4,5}$/.test(i)){let p=Number(i);if(p>2e4&&p<99999){let m=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(i))return i.slice(0,10);let n=i.split(/[\/\-\.]/);if(n.length===3){let[p,m,l]=n.map(d=>d.trim());if(p.length===4&&m.length<=2&&l.length<=2)return`${p}-${m.padStart(2,"0")}-${l.padStart(2,"0")}`;if(l.length===4&&m.length<=2&&p.length<=2)return`${l}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`}return i},o=e.map(s=>({employee_name:String(s["Nama Karyawan"]||"").trim(),division:String(s.Divisi||"").trim(),branch_id:r(String(s.Cabang||"").trim()),tanggal:a(s["Tanggal Sp"]),akhir_sp:a(s["Akhir Sp"]),sp_type:String(s["Jenis Sp"]||"").trim(),document_link:String(s["Link Document / Foto"]||"").trim()})).filter(s=>s.employee_name&&s.branch_id),c=await k("/api/import/sp",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},formFields:[{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:ia},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"combobox",name:"branch_id",label:"Cabang",required:!0,options:tt,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}O();var Ie=[],sa=[];async function oa(t){Ie=await K(),sa=await ee(),A({container:t,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",options:Ie},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",options:Ie}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let r=new URLSearchParams(e||{}).toString(),a=await k(`/api/mutasi?limit=10000&${r}`);if(a.ok){let o=a.data.data.map(s=>({Tanggal:s.tanggal||"","Nama Karyawan":s.employee_name||"","Cabang Asal":s.from_branch_name||"","Cabang Tujuan":s.to_branch_name||"",Status:s.status||"",Dokumen:s.document_link||""})),{downloadExcel:c}=await Promise.resolve().then(()=>(R(),oe));c(o,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(R(),oe));r(e,"Template_Import_Mutasi")},onImport:async e=>{let r=s=>{if(!s)return null;let i=String(s||"").toLowerCase(),n=Ie.find(p=>String(p.label||"").toLowerCase()===i);return n?n.value:null},a=s=>{if(!s)return"";if(s instanceof Date&&!isNaN(s.getTime()))return s.toISOString().slice(0,10);let i=String(s).trim();if(/^\d{4,5}$/.test(i)){let p=Number(i);if(p>2e4&&p<99999){let m=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(i))return i.slice(0,10);let n=i.split(/[\/\-\.]/);if(n.length===3){let[p,m,l]=n.map(d=>d.trim());if(p.length===4&&m.length<=2&&l.length<=2)return`${p}-${m.padStart(2,"0")}-${l.padStart(2,"0")}`;if(l.length===4&&m.length<=2&&p.length<=2)return`${l}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`}return i},o=e.map(s=>({tanggal:a(s.Tanggal),employee_name:String(s["Nama Karyawan"]||"").trim(),from_branch_id:r(String(s["Cabang Asal"]||"").trim()),to_branch_id:r(String(s["Cabang Tujuan"]||"").trim()),status:String(s.Status||"").trim(),document_link:String(s.Dokumen||"").trim()})).filter(s=>s.tanggal&&s.employee_name&&s.from_branch_id&&s.to_branch_id),c=await k("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:sa},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Ie,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Ie,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}O();async function la(t){t.innerHTML=`
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
  `;let e=document.getElementById("btnRefreshSync"),r=document.getElementById("queueStatusFilter");e.addEventListener("click",o),r.addEventListener("change",i),document.getElementById("btnRetryAll").addEventListener("click",()=>a("retry",{allFailed:!0})),document.getElementById("btnResetStuck").addEventListener("click",()=>a("reset-stuck",{})),document.getElementById("btnPauseSync").addEventListener("click",()=>a("pause",{})),document.getElementById("btnResumeSync").addEventListener("click",()=>a("resume",{})),document.getElementById("btnReconcile").addEventListener("click",()=>a("reconcile",{module:"ALL",repairMode:!0})),document.getElementById("btnRetrySelected").addEventListener("click",()=>{let p=Array.from(document.querySelectorAll(".chk-queue:checked")).map(m=>m.value);if(p.length===0)return alert("No items selected");a("retry",{ids:p})}),document.getElementById("chkAllQueue").addEventListener("change",p=>{document.querySelectorAll(".chk-queue").forEach(m=>m.checked=p.target.checked)});async function a(p,m){if(confirm(`Are you sure you want to execute action: ${p}?`)){showLoading();try{let l=await k(`/api/sync/actions/${p}`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(m)});l.ok?(alert(l.data?.message||"Success"),o()):G(l.error||"Action failed")}catch(l){G(l.message)}hideLoading()}}await o();async function o(){showLoading(),await Promise.all([s(),i(),c(),n()]),hideLoading()}async function c(){try{let p=await k("/api/sync/performance");if(!p.ok)return;let{webhook:m,google_api:l,d1:d,queue:u,throughput:b}=p.data;document.getElementById("latencyHistograms").innerHTML=`
        <table class="w-full text-left">
          <thead><tr class="text-gray-500 border-b"><th>Metric</th><th>P50</th><th>P95</th><th>P99</th><th>Max</th></tr></thead>
          <tbody class="divide-y">
            <tr><td class="py-1">Webhook</td><td>${m.P50}ms</td><td>${m.P95}ms</td><td>${m.P99}ms</td><td>${m.Max}ms</td></tr>
            <tr><td class="py-1">Google API</td><td>${l.P50}ms</td><td>${l.P95}ms</td><td>${l.P99}ms</td><td>${l.Max}ms</td></tr>
            <tr><td class="py-1">D1 Execute</td><td>${d.P50}ms</td><td>${d.P95}ms</td><td>${d.P99}ms</td><td>${d.Max}ms</td></tr>
            <tr><td class="py-1">Queue Delay</td><td>${u.P50}ms</td><td>${u.P95}ms</td><td>${u.P99}ms</td><td>${u.Max}ms</td></tr>
          </tbody>
        </table>
        <div class="mt-4 pt-3 border-t text-gray-600 flex justify-between">
          <span>Throughput: <b>${b.events_per_sec}</b> ev/sec</span>
          <span><b>${b.events_per_min}</b> ev/min</span>
        </div>
      `,document.getElementById("lblSnapshotStatus").innerHTML='Checked 10 modules. Status: <b>COMPLETED</b><br><span class="text-xs text-gray-500">Run today at 02:00</span>'}catch(p){console.error(p)}}async function s(){try{let p=await k("/api/sync/health");if(!p.ok)return G("Failed to fetch sync health");let{status:m,queue:l,circuit_breaker:d}=p.data,u=`
        <div class="bg-white p-4 rounded-lg shadow border-l-4 ${m==="HEALTHY"?"border-green-500":m==="WARNING"?"border-yellow-500":"border-red-500"}">
          <p class="text-sm text-gray-500">System Health</p>
          <p class="text-2xl font-bold ${m==="HEALTHY"?"text-green-600":m==="WARNING"?"text-yellow-600":"text-red-600"}">${m}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-blue-500">
          <p class="text-sm text-gray-500">Pending Queue</p>
          <p class="text-2xl font-bold text-gray-800">${l.pending||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-orange-500">
          <p class="text-sm text-gray-500">Failed / Retrying</p>
          <p class="text-2xl font-bold text-gray-800">${l.failed||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-red-500">
          <p class="text-sm text-gray-500">Dead Letters</p>
          <p class="text-2xl font-bold text-gray-800">${l.dead_letter||0}</p>
        </div>
      `;document.getElementById("syncOverviewCards").innerHTML=u;let b=document.getElementById("cbStateBadge"),g=document.getElementById("cbStateDesc"),h=document.getElementById("cbStatusCard");h.className="bg-white rounded-lg shadow p-6 border-l-4",d==="CLOSED"?(b.className="px-3 py-1 rounded-full text-sm font-medium bg-green-100 text-green-800",b.textContent="CLOSED",g.textContent="Traffic is flowing normally to Google Sheets.",h.classList.add("border-green-500")):d==="OPEN"?(b.className="px-3 py-1 rounded-full text-sm font-medium bg-red-100 text-red-800",b.textContent="OPEN",g.textContent="Failing fast. Traffic to Google Sheets is paused due to repeated failures.",h.classList.add("border-red-500")):d==="HALF_OPEN"?(b.className="px-3 py-1 rounded-full text-sm font-medium bg-yellow-100 text-yellow-800",b.textContent="HALF-OPEN",g.textContent="Testing recovery. Permitting limited traffic to verify stability.",h.classList.add("border-yellow-500")):b.textContent=d||"UNKNOWN"}catch(p){console.error(p)}}async function i(){try{let p=document.getElementById("queueStatusFilter").value,m=await k("/api/sync/queue?limit=15"+(p?"&status="+p:""));if(!m.ok)return;let l=document.getElementById("queueTableBody"),d=m.data?.data||m.data||[];if(d.length===0){l.innerHTML='<tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">No events found</td></tr>';return}l.innerHTML=d.map(u=>`
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
      `).join("")}catch(p){console.error(p)}}async function n(){try{let p=await k("/api/sync/metrics");if(!p.ok)return;let m=document.getElementById("metricsTableBody"),l=p.data||[];if(l.length===0){m.innerHTML='<tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">No metrics found in last 24h</td></tr>';return}m.innerHTML=l.map(d=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2 font-medium text-gray-800">${d.module}</td>
          <td class="px-4 py-2 text-gray-600">${d.total_events}</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(d.avg_webhook_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(d.avg_d1_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(d.avg_queue_wait_ms||0)} ms</td>
        </tr>
      `).join("")}catch(p){console.error(p)}}}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let e=Math.floor(Number(t)-25569);return new Date(e*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let e=t.split(/[\/\-]/);return`${e[2]}-${e[1]}-${e[0]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let r=e.split("-");if(r.length===3&&r[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],o=parseInt(r[2],10),c=a[parseInt(r[1],10)-1];return`${o} ${c} ${r[0]}`}return e};function z(t){return async e=>{if(!Le()){xe("/login");return}return t(e)}}var Je=null;function Ha(){Je&&clearInterval(Je);let t=()=>{let e=new Date,r=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),o=document.getElementById("header-clock-time"),c=document.getElementById("header-clock-date");o&&(o.textContent=r),c&&(c.textContent=a)};t(),Je=setInterval(t,1e3)}async function qa(){try{let t=await k("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},r=(a,o)=>{let c=document.getElementById(a);c&&(c.textContent=o>0?o:"",c.style.display=o>0?"inline-flex":"none")};r("badge-issues",e.issues?.current||0),r("badge-contracts",e.expiring30?.current||0),r("badge-oo1",e.one_on_one?.current||0),r("badge-schedule",e.schedule?.current||0),r("badge-supply",e.supply?.current||0)}catch{}}var Pe=[];async function ja(){try{let t=await k("/api/dashboard/notifications");if(!t.ok)return;Pe=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Pe.length>0?"block":"none",e.textContent=Pe.length)}catch{}}function Ja(){if(!Pe.length){ce({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,r)=>r()});return}let t=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${Pe.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;ce({title:`Notifikasi (${Pe.length})`,content:t,confirmText:"Tutup",onConfirm:(e,r)=>r()})}function ca(){let t=fe(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
                <span class="topbar-greeting-time">${(()=>{let p=new Date().getHours();return p>=4&&p<11?"Selamat Pagi":p>=11&&p<15?"Selamat Siang":p>=15&&p<18?"Selamat Sore":"Selamat Malam"})()}, </span><span class="topbar-greeting-name">Berlin Ariansyah</span> \u{1F44B}
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
  `;let r=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),o=document.getElementById("topbar-menu-btn"),c=document.getElementById("sidebar-close"),s=()=>{r.classList.add("open"),a.classList.add("show")},i=()=>{r.classList.remove("open"),a.classList.remove("show")};o?.addEventListener("click",s),c?.addEventListener("click",i),a?.addEventListener("click",i),document.querySelectorAll(".nav-item").forEach(p=>p.addEventListener("click",i));function n(){let p=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(d=>{let u=d.dataset.route;d.classList.toggle("active",p===u||u!=="/dashboard"&&p.startsWith(u))});let m=document.getElementById("topbar-title"),l=document.querySelector(".nav-item.active .nav-label");m&&l&&(m.textContent=l.textContent)}window.addEventListener("hashchange",n),n(),Ha(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await k("/api/auth/logout",{method:"POST"}),Ae(),Je&&clearInterval(Je),xe("/login")}),qa(),ja(),document.getElementById("btn-notif")?.addEventListener("click",p=>{p.preventDefault(),Ja()})}async function Ua(){J("/login",({main:e})=>Bt(e)),J("/dashboard",z(({main:e})=>Et(e))),J("/calendar",z(({main:e})=>Vt(e))),J("/employees",z(({main:e,params:r})=>Lt(e,r))),J("/contracts",z(({main:e,params:r})=>Nt(e,r))),J("/sp",z(({main:e})=>ra(e))),J("/mutasi",z(({main:e})=>oa(e))),J("/sync-dashboard",z(({main:e})=>la(e))),J("/timeline",z(({main:e,params:r})=>Ft(e,r))),J("/issues",z(({main:e,params:r})=>Mt(e,r))),J("/one-on-one",z(({main:e,params:r})=>Ot(e,r))),J("/training",z(({main:e})=>Rt(e))),J("/relievers",z(({main:e,params:r})=>Ht(e,r))),J("/reports/inspection",z(({main:e})=>qt(e))),J("/reports/cleaning",z(({main:e})=>jt(e))),J("/reports/fogging",z(({main:e})=>Jt(e))),J("/reports/basecamp",z(({main:e})=>Ut(e))),J("/reports/supply",z(({main:e})=>yt(e,"supply"))),J("/sop",z(({main:e})=>Gt(e))),J("/checklist",z(({main:e})=>zt(e))),J("/forms",z(({main:e})=>yt(e))),J("/users",z(({main:e})=>Qt(e))),J("/branches",z(({main:e})=>Wt(e))),J("/profile",z(({main:e})=>Xt(e))),J("/settings/import",z(({main:e})=>na(e)));let t=Le();if(!t&&window.location.hash!=="#/login"&&xe("/login"),t){let e=await k("/api/auth/me");e.ok?(Ne(e.data.data),ca()):(Ae(),xe("/login"))}window.addEventListener("fm:login",()=>{ca(),xe("/dashboard")}),vt()}Ua();
