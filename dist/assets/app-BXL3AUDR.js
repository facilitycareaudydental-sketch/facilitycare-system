var oa=Object.defineProperty;var tt=(t,e)=>()=>(t&&(e=t(t=0)),e);var at=(t,e)=>{for(var i in e)oa(t,i,{get:e[i],enumerable:!0})};var ve={};at(ve,{API:()=>yt,CLIENT_SIDE_MAX_ROWS:()=>ge,IS_DEVELOPMENT:()=>De,apiFetch:()=>S,clearToken:()=>Pe,getToken:()=>Ie,getUser:()=>be,setToken:()=>nt,setUser:()=>Be});function Ie(){return localStorage.getItem("fm_token")}function nt(t){localStorage.setItem("fm_token",t)}function Pe(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function be(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Be(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function S(t,e={}){let i=Ie(),a={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...e.headers||{}};try{let l=`cb=${Date.now()}`,d=t.includes("?")?"&":"?",r=`${yt}${t}${d}${l}`,s=await fetch(r,{...e,headers:a}),n;try{let c=await s.text();try{n=JSON.parse(c)}catch{n={error:`Server Error (${s.status}): ${c.substring(0,80)}...`}}}catch{n={error:"Gagal membaca respon dari server"}}return s.status===401&&(Pe(),window.location.hash="#/login"),{ok:s.ok,status:s.status,data:n}}catch(l){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${l.message})`}}}}var De,da,yt,ge,M=tt(()=>{De=!1,da="https://fm-operations-api.facilitycare-audydental.workers.dev",yt=da,ge=1e4});var kt={};at(kt,{confirmDialog:()=>Ne,createModal:()=>le});function le({title:t,content:e,onConfirm:i,onCancel:a,confirmText:l="Simpan",cancelText:d="Batal",size:r="md",confirmClass:s="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},c=document.createElement("div");c.className="modal-overlay",c.innerHTML=`
    <div class="modal" style="max-width:${n[r]||n.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${d}</button>
        ${i?`<button class="btn ${s} modal-confirm">${l}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&c.querySelector(".modal-body").appendChild(e);let u=()=>{c.classList.remove("show"),setTimeout(()=>c.remove(),250)};return c.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),u()}),c.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),u()}),i&&c.querySelector(".modal-confirm").addEventListener("click",()=>i(c,u)),c.addEventListener("click",o=>{o.target===c&&(a&&a(),u())}),document.body.appendChild(c),requestAnimationFrame(()=>c.classList.add("show")),{overlay:c,close:u}}function Ne(t,e,i="Konfirmasi"){return le({title:i,content:`<p>${t}</p>`,onConfirm:(a,l)=>{e(),l()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var he=tt(()=>{});var re={};at(re,{downloadExcel:()=>L,parseExcel:()=>Fe,renderExcelButtons:()=>Me});function Fe(t){return new Promise((e,i)=>{let a=new FileReader;a.onload=l=>{try{let d=new Uint8Array(l.target.result),r=XLSX.read(d,{type:"array"}),s=r.SheetNames[0],n=r.Sheets[s];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${r.SheetNames.join(", ")}`),console.log(`Sheet Used: ${s}`);let c=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),u=c.e.r-c.s.r+1,o=c.e.c-c.s.c+1;console.log(`Total Rows (including empty): ${u}`),console.log(`Total Columns: ${o}`);let p=[];for(let b=c.s.c;b<=c.e.c;++b){let g=n[XLSX.utils.encode_cell({c:b,r:c.s.r})];g&&g.v&&p.push(g.v)}console.log(`Headers Found: ${p.join(", ")}`),console.log("---------------------------");let m=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(m,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(m,"__headers",{value:p,enumerable:!1}),e(m)}catch(d){i(d)}},a.onerror=l=>i(l),a.readAsArrayBuffer(t)})}function L(t,e){try{let i=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,i,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function Me(t){return`
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
  `}var R=tt(()=>{});M();var it={},qe=null;function j(t,e){it[t]=e}function ke(t){window.location.hash=t}function ft(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[i,...a]=e.split("?"),l=it[i];if(!l){for(let[r,s]of Object.entries(it))if(r.endsWith("/*")&&i.startsWith(r.slice(0,-2))){l=s;break}}qe&&(qe(),qe=null);let d=document.getElementById("main-content");if(d&&(d.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),l){let r=new URLSearchParams(a.join("?")),s=i.split("/").filter(Boolean),n=await l({path:i,params:r,segments:s,main:d});n&&(qe=n)}else{let r=d||document.getElementById("app");r&&(r.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var Le;function ca(){return Le||(Le=document.createElement("div"),Le.id="toast-container",document.body.appendChild(Le)),Le}function vt(t,e="info",i=3500){let a=ca(),l=document.createElement("div");l.className=`toast toast-${e}`;let d={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};l.innerHTML=`<span class="toast-icon">${d[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(l),requestAnimationFrame(()=>l.classList.add("show")),setTimeout(()=>{l.classList.remove("show"),setTimeout(()=>l.remove(),350)},i)}var Z=t=>vt(t,"success"),G=t=>vt(t,"error");he();M();M();he();function je({columns:t,data:e,onEdit:i,onDelete:a,onView:l,actions:d=[],emptyText:r="Tidak ada data",bulkSelect:s=null}){let n=document.createElement("div");if(n.className="table-wrapper",!e||e.length===0)return n.innerHTML=`<div class="empty-state"><p>${r}</p></div>`,n;let c=document.createElement("table");c.className="data-table";let u=document.createElement("thead"),o=document.createElement("tr");if(s){let m=document.createElement("th");m.style.width="40px",m.style.textAlign="center";let b=document.createElement("input");b.type="checkbox",b.id="select-all-checkbox",b.title="Pilih semua",b.addEventListener("change",()=>{e.forEach(g=>{b.checked?s.selectedIds.add(g.id):s.selectedIds.delete(g.id)}),n.querySelectorAll(".row-checkbox").forEach(g=>g.checked=b.checked),s.onToggle()}),m.appendChild(b),o.appendChild(m)}if(t.forEach(m=>{let b=document.createElement("th");b.textContent=m.label,m.width&&(b.style.width=m.width),o.appendChild(b)}),i||a||l||d.length>0){let m=document.createElement("th");m.textContent="Aksi",m.style.width="120px",o.appendChild(m)}u.appendChild(o),c.appendChild(u);let p=document.createElement("tbody");return e.forEach(m=>{let b=document.createElement("tr");if(s){let g=document.createElement("td");g.style.textAlign="center",g.style.width="40px";let h=document.createElement("input");h.type="checkbox",h.className="row-checkbox",h.checked=s.selectedIds.has(m.id),h.addEventListener("change",()=>{if(h.checked)s.selectedIds.add(m.id);else{s.selectedIds.delete(m.id);let f=document.getElementById("select-all-checkbox");f&&(f.checked=!1)}s.onToggle()}),g.appendChild(h),b.appendChild(g)}if(t.forEach(g=>{let h=document.createElement("td");if(g.render){let f=g.render(m[g.key],m);f instanceof HTMLElement?h.appendChild(f):h.innerHTML=f||""}else h.textContent=m[g.key]!==null&&m[g.key]!==void 0&&m[g.key]!==""?m[g.key]:"";g.nowrap&&(h.style.whiteSpace="nowrap"),b.appendChild(h)}),i||a||l||d.length>0){let g=document.createElement("td");g.className="actions-cell";let h=document.createElement("div");if(h.className="btn-group",l){let f=document.createElement("button");f.className="btn btn-xs btn-ghost",f.innerHTML="\u{1F441}",f.title="Lihat",f.addEventListener("click",()=>l(m)),h.appendChild(f)}if(i){let f=document.createElement("button");f.className="btn btn-xs btn-secondary",f.innerHTML="\u270F\uFE0F",f.title="Edit",f.addEventListener("click",()=>i(m)),h.appendChild(f)}d.forEach(f=>{let k=document.createElement("button");k.className=`btn btn-xs ${f.class||"btn-ghost"}`,k.innerHTML=f.icon||f.label,k.title=f.label,k.addEventListener("click",()=>f.handler(m)),h.appendChild(k)}),g.appendChild(h),b.appendChild(g)}p.appendChild(b)}),c.appendChild(p),n.appendChild(c),n}function Je({page:t,pages:e,total:i,limit:a,onPage:l}){if(e<=1)return null;let d=document.createElement("div");d.className="pagination";let r=document.createElement("span");r.className="pagination-info",r.textContent=`Total: ${i} data`,d.appendChild(r);let s=document.createElement("div");s.className="pagination-btns";let n=(o,p,m=!1,b=!1)=>{let g=document.createElement("button");g.className=`btn btn-sm ${b?"btn-primary":"btn-ghost"} pagination-btn`,g.textContent=o,g.disabled=m,g.addEventListener("click",()=>l(p)),s.appendChild(g)};n("\xAB",1,t===1),n("\u2039",t-1,t===1);let c=Math.max(1,t-2),u=Math.min(e,t+2);for(let o=c;o<=u;o++)n(o,o,!1,o===t);return n("\u203A",t+1,t===e),n("\xBB",e,t===e),d.appendChild(s),d}he();function Ae(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${Ae(e.fields)}</div>`;let i=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",l="";switch(e.type){case"textarea":l=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${i} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let r=(e.options||[]).map(o=>{let p=typeof o=="object"?o.value:o,m=typeof o=="object"?o.label:o,b=e.value==p?"selected":"";return`<option value="${p}" ${b}>${m}</option>`}).join("");l=`<select name="${e.name}" class="form-control" ${i}><option value="">-- Pilih ${e.label||""} --</option>${r}</select>`;break;case"combobox":let s=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,n=(e.options||[]).map(o=>{let p=typeof o=="object"?o.value:o,m=typeof o=="object"?o.label||o.value||"":o||"";return(m==="undefined"||m==="[object Object]"||m==="null")&&(m=""),m?`<option value="${m}"></option>`:""}).join(""),c=e.value||"";if(e.value){let o=(e.options||[]).find(p=>(typeof p=="object"?p.value:p)==e.value);if(o){let p=typeof o=="object"?o.label||o.value||"":o||"";p&&p!=="undefined"&&p!=="[object Object]"&&p!=="null"&&(c=p)}}l=`
          <input type="text" name="${e.name}" list="${s}" class="form-control" value="${c}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${s}">${n}</datalist>
        `;break;case"checkbox":l=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let u=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";l=`<input type="date" name="${e.name}" class="form-control" value="${u}" ${i}>`;break;case"number":l=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${i}>`;break;case"email":l=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i}>`;break;case"url":l=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${i}>`;break;default:l=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i} autocomplete="off">`}let d=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${l}${d}</div>`}).join("")}function Ue(t){let e={},i=new FormData(t);for(let[a,l]of i.entries())e[a]=l===""?null:l;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function Ge(t,e){e&&Object.entries(e).forEach(([i,a])=>{let l=t.querySelector(`[name="${i}"]`);l&&(l.hasAttribute("list")||(l.type==="checkbox"?l.checked=!!a:l.type==="date"&&a&&window.parseFlexibleDate?l.value=window.parseFlexibleDate(a):l.value=a??""))})}R();function N({container:t,title:e,icon:i,apiPath:a,columns:l,formFields:d,filterFields:r,defaultFilters:s={},itemLabel:n="Data",canCreate:c=!0,canEdit:u=!0,canDelete:o=!0,onBeforeSubmit:p,onAfterLoad:m,onDataLoaded:b,extraActions:g=[],initialSearch:h="",exportOptions:f=null,bulkDelete:k=!1,paginationMode:x="server"}){let _=1,w={...s};h&&(w.search=h);let C=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions">
        ${c?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
      </div>
    </div>

    ${k?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="align-items:center; background:var(--bg-card); border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${f?Me(f.moduleName):""}

    ${r&&r.length>0?`
    <div class="filter-bar card" style="padding: 1rem; display:flex; flex-wrap:wrap; gap:1rem; align-items:center;">
        ${r.map(v=>{if(v.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${w.search||""}"></div>`;if(v.type==="search-combo"){let y="dl-filter-search",$=(v.options||[]).map(B=>`<option value="${typeof B=="object"?B.label:B}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${y}" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${w.search||""}"><datalist id="${y}">${$}</datalist></div>`}return v.type==="select"?`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${v.name}" id="filter-${v.name}"><option value="">-- ${v.label} --</option>${(v.options||[]).map(y=>`<option value="${typeof y=="object"?y.value:y}" ${w[v.name]===(typeof y=="object"?y.value:y)?"selected":""}>${typeof y=="object"?y.label:y}</option>`).join("")}</select>`:""}).join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function P(){if(!document.getElementById("bulk-toolbar"))return;let y=document.getElementById("bulk-count"),$=document.getElementById("btn-bulk-delete"),B=document.getElementById("btn-bulk-cancel");y.textContent=`${C.size} item dipilih`,C.size>0?($.disabled=!1,B.disabled=!1):($.disabled=!0,B.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{C.clear(),document.querySelectorAll(".row-checkbox").forEach(y=>y.checked=!1);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1),P()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(C.size===0)return;let v=[...C],y=document.createElement("div");y.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",y.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${v.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${v.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(y),y.querySelector("#bulk-cancel-btn").addEventListener("click",()=>y.remove()),y.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let $=y.querySelector("#bulk-confirm-btn");$.disabled=!0,$.textContent="Menghapus...";let B=await S(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:v})});y.remove(),B.ok?(Z(`${v.length} ${n} berhasil dihapus.`),C.clear(),P(),O()):G(B.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),J;if(E?.addEventListener("input",v=>{clearTimeout(J),J=setTimeout(()=>{w.search=v.target.value,_=1,C.clear(),O()},400)}),r?.forEach(v=>{v.type==="select"&&document.getElementById(`filter-${v.name}`)?.addEventListener("change",y=>{w[v.name]=y.target.value,_=1,C.clear(),O()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{w={...s},E&&(E.value=""),r?.forEach(v=>{let y=document.getElementById(`filter-${v.name}`);y&&(y.value="")}),_=1,C.clear(),O()}),document.getElementById("btn-create")?.addEventListener("click",()=>pe(null)),f){document.getElementById(`btn-export-${f.moduleName}`)?.addEventListener("click",async y=>{let $=y.target,B=$.innerHTML;$.innerHTML="\u23F3 Loading...",$.disabled=!0;try{await f.onExport()}catch{G("Gagal export data")}finally{$.innerHTML=B,$.disabled=!1}}),document.getElementById(`btn-template-${f.moduleName}`)?.addEventListener("click",()=>{f.onTemplate()});let v=document.getElementById(`input-import-${f.moduleName}`);v?.addEventListener("change",async y=>{let $=y.target.files[0];if(!$)return;v.disabled=!0;let B=document.createElement("div");B.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",B.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(B);let V=B.querySelector("#import-progress-text"),U=B.querySelector("#import-progress-bar"),D=B.querySelector("#import-summary"),I=B.querySelector("#import-close-btn");I.addEventListener("click",()=>{B.remove(),O()});try{let Y=await Fe($);if(Y.length===0)throw new Error("File kosong atau format salah");let X=500,ie=0,ae=0,A=0,F=Y.length;V.textContent=`Ditemukan ${F} baris data. Memulai import...`;for(let T=0;T<F;T+=X){let z=Y.slice(T,T+X);V.textContent=`Mengimport baris ${T+1} - ${Math.min(T+X,F)} dari ${F}...`,U.style.width=`${Math.round(T/F*100)}%`;try{let H=await f.onImport(z);H?(ie+=H.inserted||H.metrics?.inserted||z.length,ae+=H.skipped||H.metrics?.updated||0):ie+=z.length}catch(H){console.error("Chunk import failed:",H),A+=z.length}}U.style.width="100%",V.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',D.style.display="block",D.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${F}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ie}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${A}</strong></div>
        `,A>0&&(D.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),I.style.display="block",v.value=""}catch(Y){V.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${Y.message}`,U.style.background="var(--danger)",U.style.width="100%",I.style.display="block",v.value=""}finally{v.disabled=!1}})}async function O(){P();let v=document.getElementById("table-container");if(!v)return;v.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let y=x==="client",$=y?1:_,B=y?ge:20,V=new URLSearchParams({page:$,limit:B,...Object.fromEntries(Object.entries(w).filter(([,A])=>A))}),U=await S(`${a}?${V}`);if(!U.ok){v.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${U.data?.error||"Error"}</p></div>`;return}let D=U.data?.data||U.data||[],I=U.data?.pagination,Y=D.length,X=D;if(y){D=b(D),X=D;let A=D.length,F=20,T=Math.ceil(A/F);_>T&&T>0&&(_=T);let z=(_-1)*F,H=_*F;D=D.slice(z,H),I={page:_,limit:F,total:A,pages:T}}!1,m&&m(D);let ie=je({columns:l,data:D,fullData:X,onEdit:u?A=>pe(A):null,actions:g.map(A=>({...A,handler:F=>A.handler(F,O)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:k?{selectedIds:C,onToggle:P}:null});v.innerHTML="",v.appendChild(ie);let ae=document.getElementById("pagination-container");if(ae&&(ae.innerHTML="",I&&I.pages>1)){let A=Je({page:I.page,pages:I.pages,total:I.total,limit:I.limit,onPage:F=>{_=F,O()}});A&&ae.appendChild(A)}}function ye(v){let y=typeof d=="function"?d(v):d;return Ae(y)}function pe(v){let y=!!v,$=document.createElement("form");if($.noValidate=!0,$.innerHTML=ye(v),y){let V=typeof d=="function"?d(v):d;Ge($,v)}let{close:B}=le({title:y?`Edit ${n}`:`Tambah ${n}`,content:$,size:"lg",confirmText:y?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(V,U)=>{if(!$.reportValidity())return;let D=V.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let I=Ue($),Y=typeof d=="function"?d(v):d,X=async F=>{for(let T of F)if(T.type==="row")await X(T.fields);else if(T.type==="combobox"&&I[T.name]){let z=I[T.name],H=(T.options||[]).find(W=>{let ne=String(typeof W=="object"?W.value:W),et=String(typeof W=="object"?W.label:W);return ne===z||et===z});if(H)I[T.name]=typeof H=="object"?H.value:H;else if(T.createApi){let W={};W[T.createApi.field]=z,T.createApi.extra&&Object.assign(W,T.createApi.extra);let ne=await S(T.createApi.path,{method:"POST",body:JSON.stringify(W)});if(ne.ok&&ne.data?.id)I[T.name]=ne.data.id;else if(ne.ok&&!ne.data?.id)I[T.name]=z;else throw new Error(`Gagal membuat master data: ${ne.data?.error||"Unknown error"}`)}}};try{await X(Y)}catch(F){G(F.message),D.disabled=!1,D.textContent=y?"Simpan Perubahan":`Tambah ${n}`;return}p&&(I=await p(I,v));let ie=y?"PUT":"POST",ae=y?`${a}/${v.id}`:a,A=await S(ae,{method:ie,body:JSON.stringify(I)});A.ok?(Z(y?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),U(),O()):(G(A.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=y?"Simpan Perubahan":`Tambah ${n}`)}})}function fe(v){Ne(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let y=await S(`${a}/${v.id}`,{method:"DELETE"});y.ok?(Z(`${n} berhasil dihapus.`),O()):G(y.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return O(),O}M();M();var Se=null,Qe=null;async function xe(t=!1){if(Se&&!t)return console.log("Employees Raw (Cache Hit)",Se.slice(0,5)),Se;let e=await S(`/api/employees?limit=${ge}&status=Aktif`);return Se=(e.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",Se.slice(0,5)),Se}async function ee(t=!1){let i=(await xe(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",i.slice(0,5)),i}async function K(t=!1){return Qe&&!t||(Qe=((await S("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),Qe}function q(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function rt(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function we(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function st(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function oe(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}R();function lt(t,e){if(t.period!=="Q3")return!1;let i=String(t.status||"").toLowerCase();if(i!=="selesai"&&i!=="completed"&&i!=="done")return!1;let a=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?a.includes("inspeksi"):e==="gcdc"?a.includes("general cleaning")||a.includes("deep cleaning"):!1}M();R();function St(t,e){let i=String(t.status||"").toLowerCase();return e==="active"?i==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&i==="aktif":!1}M();R();function ot(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let l=new Date(a);l.setDate(a.getDate()+30);let d=new Date(t.end_date);return d.setHours(0,0,0,0),d>=a&&d<=l}return!1}M();R();function xt(t,e){let i=String(t.status||"").toLowerCase();return e==="open"?i==="open":!1}M();function wt(t,e){let i=String(t.status||"").toLowerCase();return e==="pending"?i==="pending":!1}var me={};function Oe(t){if(me[t]){try{me[t].destroy()}catch{}delete me[t]}}function pa(){Object.keys(me).forEach(Oe)}var ce=(t,e=0)=>{let i=Number(t);return isNaN(i)||t===null||t===void 0?e:i},_e=(t,e="\u2014")=>{if(t==null||t==="")return e;let i=String(t).trim();return i===""||i==="[object Object]"?e:i};var ma=t=>{if(!t||typeof t!="string")return"";try{let[e,i]=t.split("-");return new Date(Number(e),Number(i)-1).toLocaleDateString("id-ID",{month:"short",year:"2-digit"})}catch{return t}};function Ct(t,e,i=900){if(!t)return;let a=Math.max(0,Math.round(ce(e)));if(a===0){t.textContent="0";return}let l=Date.now(),d=()=>{let r=Math.min((Date.now()-l)/i,1),s=1-Math.pow(1-r,3);t.textContent=Math.round(s*a).toLocaleString("id-ID"),r<1?requestAnimationFrame(d):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(d)}var ua={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},ga=t=>{let e=_e(t,"\u2014");return`<span class="status-pill ${ua[e]||"pill-neutral"}">${e}</span>`};var de={family:"Inter",size:11},Ce="#94A3B8",ze="#F1F5F9",dt=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],ba=()=>window.innerWidth<768;function pt(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:ba()?"bottom":"top",labels:{font:de,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:de,titleFont:{...de,weight:"700"}}},scales:{x:{grid:{color:ze},ticks:{font:de,color:Ce,maxRotation:0}},y:{grid:{color:ze},ticks:{font:de,color:Ce},beginAtZero:!0}},...t}}var ha=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),ya=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function _t(t=3){return Array(t).fill(0).map((e,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function se(t,e,i=8e3){try{let a=new AbortController,l=setTimeout(()=>a.abort(),i),d=await S(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(l),!d||!d.ok)return e;let r=d.data;return r?r.data!==void 0?r.data??e:r:e}catch{return e}}function fa(){["skel-donut","skel-trend","skel-insp","skel-contract"].forEach(a=>{let l=document.getElementById(a);l&&(l.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract"].forEach(a=>{let l=document.getElementById(a);if(l&&l.style.display==="none"){l.style.display="block";let d=l.parentElement;if(d&&!d.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent="Belum ada data",l.style.display="none",d.appendChild(r)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Et({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&$t({}),["table-contracts","table-issues"].forEach(a=>{let l=document.getElementById(a);l&&l.querySelector(".skeleton")&&(l.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Tt(t){pa(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">


      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${ha()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${ya()}</div>

      <!-- Charts Row -->
      <div class="charts-row" style="grid-template-columns: repeat(2, minmax(0, 1fr));">
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan per Kategori</div>
          </div>
          <div style="display:flex; gap:20px; align-items:center; height:100px">
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
          <div class="chart-canvas-wrap" style="height:100px;position:relative">
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>ct(t)),document.getElementById("filter-insp-month")?.addEventListener("change",async i=>{let a=i.target.value,l=a?`/api/dashboard/inspection-bar?month=${a}`:"/api/dashboard/inspection-bar",d=document.getElementById("skel-insp"),r=document.getElementById("chart-insp");d&&(d.style.display="block",d.style.position="absolute"),r&&(r.style.display="none");let s=await se(l,{},8e3);try{Dt(s)}catch(n){console.warn("InspBar render:",n),ue("skel-insp","chart-insp")}}),t._skelTimeout=setTimeout(()=>fa(),5e3),await ct(t),t._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?ct(t):clearInterval(t._dashRefresh)},6e4)}async function ct(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,i,a,l,d,r,s,n,c,u,o,p]=await Promise.all([se("/api/dashboard/kpi",{},8e3),se("/api/dashboard/issues-trend",{},8e3),se("/api/dashboard/issues-summary",{},8e3),se("/api/dashboard/inspection-bar",{},8e3),se("/api/dashboard/stats",{},8e3),se("/api/dashboard/calendar",[],8e3),se("/api/schedule?limit=10000",{data:[]},8e3),se("/api/employees?limit=10000",{data:[]},8e3),se("/api/contracts?limit=10000",{data:[]},8e3),se("/api/issues?limit=10000",{data:[]},8e3),se("/api/one-on-one?limit=10000",{data:[]},8e3),se("/api/dashboard/contracts-chart",{labels:[],data:[]},8e3)]);if(e){let m=Array.isArray(s?.data)?s.data:Array.isArray(s)?s:[],b=Array.isArray(n?.data)?n.data:Array.isArray(n)?n:[],g=Array.isArray(c?.data)?c.data:Array.isArray(c)?c:[],h=Array.isArray(u?.data)?u.data:Array.isArray(u)?u:[],f=Array.isArray(o?.data)?o.data:Array.isArray(o)?o:[];e.employees&&(e.employees.current=b.filter(k=>St(k,"active")).length),e.contracts&&(e.contracts.current=g.filter(k=>ot(k,"active")).length),e.expiring30&&(e.expiring30={current:g.filter(k=>ot(k,"expiring30")).length}),e.issues&&(e.issues.current=h.filter(k=>xt(k,"open")).length),e.one_on_one&&(e.one_on_one.current=f.filter(k=>wt(k,"pending")).length),e.inspection_month&&(e.inspection_month.current=m.filter(k=>lt(k,"inspeksi")).length),e.cleaning_month&&(e.cleaning_month.current=m.filter(k=>lt(k,"gcdc")).length)}try{Et(e)}catch(m){console.warn("KPI render:",m)}try{$t(e)}catch(m){console.warn("MiniStats render:",m)}try{va(Array.isArray(a?.by_category)?a.by_category:[])}catch(m){console.warn("Donut render:",m),ue("skel-donut","chart-donut")}try{ka(i)}catch(m){console.warn("Trend render:",m),ue("skel-trend","chart-trend")}try{Dt(l)}catch(m){console.warn("InspBar render:",m),ue("skel-insp","chart-insp")}try{let m=Array.isArray(d)?d:Array.isArray(d?.recent_issues)?d.recent_issues:[];xa(m)}catch(m){console.warn("IssuesTable render:",m)}try{let m=Array.isArray(d?.expiring_contracts)?d.expiring_contracts:[];Sa(p)}catch(m){console.warn("ContractsTable render:",m)}try{wa(Array.isArray(r)?r:[])}catch(m){console.warn("Agenda render:",m)}try{_a()}catch(m){console.warn("Quick Actions render:",m)}}function Et(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let i=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=i.map(a=>{let l=ce(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${l}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${l}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{Ct(a,parseInt(a.dataset.target)||0)})}function $t(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let i=[{icon:"\u{1F4C5}",label:"Jadwal",val:t.schedule?.current,href:"#/timeline",color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F504}",label:"Reliefer",val:t.reliever_completed?.current,href:"#/relievers?dash_filter=reliever",color:"mini-teal"},{icon:"\u{1F50D}",label:"Inspeksi",val:t.inspection_month?.current,href:"#/timeline?dash_filter=inspeksi",color:"mini-blue"},{icon:"\u{1F9F9}",label:"GCDC",val:t.cleaning_month?.current,href:"#/timeline?dash_filter=gcdc",color:"mini-green"},{icon:"\u{1F4A8}",label:"Fogging",val:t.fogging_month?.current,href:"#/reports/fogging",color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=i.map(a=>`
    <a href="${a.href}" class="mini-stat ${a.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${a.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${ce(a.val)}">0</div>
        <div class="mini-stat-text">${a.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(a=>Ct(a,parseInt(a.dataset.target)||0,700))}function va(t){ue("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!e||!i)return;Oe("donut");let a=(t||[]).filter(n=>ce(n.count)>0);if(!a.length){Ve(e,"Belum ada data permasalahan");return}let l=a.map(n=>`${_e(n.category,"Lainnya")}`),d=a.map(n=>ce(n.count)),r=d.reduce((n,c)=>n+c,0);i.innerHTML=a.map((n,c)=>{let u=dt[c%dt.length],o=r>0?Math.round(n.count/r*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${u}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${o}%)</span></div>
          <div class="donut-legend-label">${l[c]}</div>
        </div>
      </div>
    `}).join("");let s={id:"centerText",beforeDraw:function(n){let c=n.width,u=n.height,o=n.ctx;o.restore();let p=(u/80).toFixed(2);o.font="bold "+p+"em Inter",o.textBaseline="middle",o.fillStyle="#1E293B";let m=r.toString(),b=Math.round((c-o.measureText(m).width)/2),g=u/2;o.fillText(m,b,g-10),o.font="600 "+(p*.35).toFixed(2)+"em Inter",o.fillStyle="#64748B";let h="Total",f=Math.round((c-o.measureText(h).width)/2);o.fillText(h,f,g+15),o.save()}};me.donut=new Chart(e,{type:"doughnut",data:{labels:l,datasets:[{data:d,backgroundColor:dt,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:de,titleFont:{...de,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutout:"75%"},plugins:[s]})}function ka(t){ue("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;Oe("trend"),t=t||{};let i=(t.labels||[]).map(ma),a=(t.open||[]).map(d=>ce(d)),l=(t.closed||[]).map(d=>ce(d));if(!i.length){Ve(e,"Belum ada data trend");return}me.trend=new Chart(e,{type:"line",data:{labels:i,datasets:[{label:"Open",data:a,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:l,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#10B981",borderWidth:2}]},options:pt({plugins:{legend:{display:!1}}})})}function Dt(t){ue("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;Oe("inspBar"),t=t||{};let i=t.labels||[],a=(t.fc||[]).map(d=>ce(d)),l=(t.spv||[]).map(d=>ce(d));if(!i.length){Ve(e,"Belum ada data inspeksi");return}me.inspBar=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:l,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:pt({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:de,color:Ce,maxRotation:45,minRotation:30}},y:{grid:{color:ze},ticks:{font:de,color:Ce},min:0,max:100}}})})}function Sa(t){ue("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;Oe("contractMiniBar"),t=t||{};let i={"06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(r=>{let s=r.split("-")[1];return i[s]||r}),l=(t.data||[]).map(r=>ce(r));if(!a.length){Ve(e,"Belum ada data");return}let d=e.getContext("2d");me.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:l,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:pt({onClick:(r,s)=>{if(s&&s.length>0){let n=s[0].index,c=(t.labels||[])[n];c&&(window.location.hash="#/contracts?month_expiry="+c)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:de,color:Ce,maxRotation:0,autoSkip:!1}},y:{grid:{color:ze,borderDash:[4,4],drawBorder:!1},ticks:{font:de,color:Ce,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function xa(t){let e=document.getElementById("table-issues");if(!e)return;let i=(t||[]).slice(0,8);if(!i.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${i.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${ga(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${_e(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${_e(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function wa(t){let e=document.getElementById("widget-agenda");if(!e)return;let i=new Date,a=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,d=(t||[]).filter(r=>(r.event_date||"").startsWith(a)).slice(0,10);if(!d.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${d.map(r=>{let s="#3B82F6",n="#EFF6FF",c="Agenda",u=(r.title||"").toLowerCase();return u.includes("inspeksi")?(s="#10B981",n="#ECFDF5",c="Inspeksi"):u.includes("cleaning")||u.includes("gcdc")?(s="#3B82F6",n="#EFF6FF",c="Cleaning"):u.includes("reliefer")?(s="#F59E0B",n="#FFFBEB",c="Reliefer"):u.includes("fogging")&&(s="#8B5CF6",n="#F5F3FF",c="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(r.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${s};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${_e(r.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${_e(r.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${s}">${c}</div>
        </div>
      `}).join("")}
    </div>
  `}function _a(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function ue(t,e){let i=document.getElementById(t),a=document.getElementById(e);if(i&&(i.style.display="none",i.style.position=""),a){a.style.display="block";let l=a.parentElement;if(l){let d=l.querySelector(".chart-empty");d&&d.remove()}}}function Ve(t,e="Belum ada data"){if(!t)return;t.style.display="none";let i=t.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let l=document.createElement("div");l.className="chart-empty",l.textContent=e,i.appendChild(l)}}M();async function It(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),i=document.getElementById("login-error"),a=document.getElementById("login-btn"),l=document.getElementById("toggle-password"),d=document.getElementById("login-password");l?.addEventListener("click",()=>{let r=d.type==="text";d.type=r?"password":"text",l.style.color=r?"":"var(--primary)"}),e?.addEventListener("submit",async r=>{r.preventDefault(),i.style.display="none";let s=e.username.value.trim(),n=e.password.value;if(!s||!n){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let c=await S("/api/auth/login",{method:"POST",body:JSON.stringify({username:s,password:n})});c.ok&&c.data.success?(nt(c.data.data.token),Be(c.data.data.user),Z("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=c.data.error||"Username atau password salah.",i.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}M();R();async function Ca(){return await K()}function Ta(t,e){let i=String(t.status||"").toLowerCase();return e==="active"?i==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&i==="aktif":!1}async function Pt(t,e){let i=await Ca(),a=e?e.get("dash_filter"):null;N({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:l=>a?l.filter(d=>Ta(d,a)):l,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:l=>we(l)},{key:"phone",label:"No. HP",render:l=>l?`<a href="tel:${l}">${l}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>q(l)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:l=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:l?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:l?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:i,value:l?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:l?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:l?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:l?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let l=await S(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(l.ok){let d=l.data.data.map(r=>({"Nama Lengkap":r.full_name,Cabang:r.branch_name||"",Divisi:r.division||"","No. HP":r.phone||"","Tgl Masuk":r.join_date||"",Status:r.status||""}));L(d,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async l=>{let d=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),u=i.find(o=>String(o.label||"").toLowerCase()===c);return u?u.value:null},r=l.map(n=>({full_name:String(n["Nama Lengkap"]||"").trim(),branch_id:d(String(n.Cabang||"").trim()),division:String(n.Divisi||"").trim()||"FACILITY CARE",phone:String(n["No. HP"]||"").trim(),join_date:String(n["Tgl Masuk"]||"").trim(),status:String(n.Status||"").trim(),notes:String(n.Catatan||"").trim()})).filter(n=>n.full_name),s=await S("/api/import/employees",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}M();R();var ut=[],Bt=[];async function Ea(){ut=await K(),Bt=await xe()}var mt=async t=>{let e=[],i=1;for(;;){let l=await(await Promise.resolve().then(()=>(M(),ve))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${i}`);if(!l.ok)break;let d=l.data?.data||l.data||[],r=Array.isArray(d)?d:[];if(e=e.concat(r),r.length<100||l.data?.pagination&&i>=l.data.pagination.pages)break;i++}return e};function $a(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let l=new Date(a);l.setDate(a.getDate()+30);let d=new Date(t.end_date);return d.setHours(0,0,0,0),d>=a&&d<=l}return!1}async function Lt(t,e){await Ea();let i=e?e.get("dash_filter"):null;N({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",onDataLoaded:a=>i?a.filter(l=>$a(l,i)):a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>we(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,l)=>l.end_date&&String(l.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':rt(a)},{key:"status",label:"Status",render:a=>q(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:ut},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"month_expiry",label:"Bulan Habis",options:[{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"},{value:"2027-01",label:"Jan 2027"},{value:"2027-02",label:"Feb 2027"},{value:"2027-03",label:"Mar 2027"},{value:"2027-04",label:"Apr 2027"},{value:"2027-05",label:"Mei 2027"},{value:"2027-06",label:"Jun 2027"},{value:"2027-07",label:"Jul 2027"}]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[d,r]=await Promise.all([mt("/api/employees?status=Aktif"),mt("/api/contracts")]);if(d.length>0){let s=r.filter(o=>o.status==="Aktif"),n=new Set(s.map(o=>o.employee_id)),c=d.filter(o=>!n.has(o.id)),u=`<p style="margin-bottom:12px">Data yang terbaca: <b>${d.length}</b> Karyawan Aktif, dan <b>${s.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${c.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;c.forEach(o=>{let p=r.filter(b=>b.employee_id===o.id),m='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(p.length>0){let b=p[0];m=`Pernah ada kontrak (Status: <b style="color:#EF4444">${b.status}</b>, Selesai: ${window.formatDate(b.end_date)})`}u+=`<li style="margin-bottom:8px"><b>${o.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${o.branch_name||"-"} | ${m}</span></li>`}),u+="</ul>",Promise.resolve().then(()=>(he(),kt)).then(o=>o.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:u,cancelText:"Tutup"}))}}catch(d){console.error(d)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let l=document.querySelector(".page-actions");l&&l.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:Bt,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:ut,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await S(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let l=a.data.data.map(d=>({"Nama Lengkap":d.employee_name,Cabang:d.branch_name||"","Div / Bagian":d.division||"","Tanggal Mulai":d.start_date||"","Tanggal Selesai":d.end_date&&String(d.end_date).startsWith("2099")?"":d.end_date||"","Sisa Kontrak":d.end_date&&String(d.end_date).startsWith("2099")?"Tetap":d.days_remaining!==null&&d.days_remaining!==void 0?`${d.days_remaining} Hari`:"",Status:d.status||""}));L(l,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[l,d]=await Promise.all([S("/api/branches?limit=10000"),mt("/api/employees")]),r=l.data?.data||[],s=d||[];console.log(`Total employee yang berhasil dimuat dari database : ${s.length}`),s.length>0&&(console.log("Contoh 5 employee pertama:"),s.slice(0,5).forEach((g,h)=>{console.log(`${h+1}. ID: ${g.id}, Name: ${g.full_name}, Status: ${g.status}`)}));let n=g=>{if(!g)return null;let h=String(g||"").replace(/\s+/g," ").toLowerCase().trim(),f=r.find(k=>String(k.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===h||String(k.code||"").replace(/\s+/g," ").toLowerCase().trim()===h||String(k.name||"").replace(/\s+/g," ").toLowerCase().trim()===h);return f?f.id:null},c=(g,h)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${h}`),console.log(`Nama dari Excel : "${g}"`),!g)return console.log("Alasan gagal mapping : Nama kosong"),null;let f=String(g||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${f}"`),console.log(`Jumlah employee di database : ${s.length}`);let k=s.find(x=>String(x.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===f);return k?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${k.id}`),k.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},u=g=>{if(!g)return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let h=String(g).trim();if(/^\d{4,5}(\.\d+)?$/.test(h)){let k=Math.floor(Number(h));if(k>2e4&&k<99999){let x=new Date(Date.UTC(1899,11,30)+k*864e5);return isNaN(x.getTime())?"":x.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);let f=h.split(/[\/\-\.]/);if(f.length===3){let[k,x,_]=f.map(w=>w.trim());if(k.length===4&&x.length<=2&&_.length<=2)return`${k}-${x.padStart(2,"0")}-${_.padStart(2,"0")}`;if(_.length===4&&x.length<=2&&k.length<=2)return`${_}-${x.padStart(2,"0")}-${k.padStart(2,"0")}`}return h},o=a.map((g,h)=>{let f=h+2,k=String(g["Nama Lengkap"]||"").trim(),x=g["Tanggal Mulai"],_=u(x);if(!_){let P=a.__worksheet,E=a.__headers||[],J=E.indexOf("Tanggal Mulai"),O="N/A",ye="N/A",pe="N/A";if(J!==-1&&P&&window.XLSX){let v=window.XLSX.utils.encode_cell({c:J,r:f-1});pe=v;let y=P[v];y?(O=y.t||"undefined",ye=y.w||"undefined"):O="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let fe="Unknown";x==null||x===""?fe="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":x instanceof Date&&isNaN(x.getTime())?fe="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":fe="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${f}`),console.log(`Employee Name : ${k}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${J})`),console.log(`Raw Cell Value : "${x}"`),console.log(`JavaScript Type : ${typeof x}`),console.log(`SheetJS Cell Type : ${O}`),console.log(`SheetJS Formatted Value : "${ye}"`),console.log(`Value After Trim : "${String(x||"").trim()}"`),console.log(`Value After Date Parser : "${_}"`),console.log(`Is Empty : ${!x}`),console.log(`Is Invalid Date : ${x instanceof Date?isNaN(x.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${fe}`),console.log(`Workbook Sheet : ${P?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${pe}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(g,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(E)),console.log(`==========================
`)}let w=c(k,f),C=null;return w||(C="Karyawan tidak ditemukan di Database"),{isValid:!!w,invalidReason:C,rowNum:f,data:{employee_id:w,branch_id:n(String(g.Cabang||"").trim()),division:String(g["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:_,end_date:u(g["Tanggal Selesai"])||"2099-12-31",status:String(g.Status||"").trim(),_rawName:k}}}),p=[],m=[];if(o.forEach(g=>{g.isValid?p.push(g.data):m.push({rowNum:g.rowNum,name:g.data._rawName,reason:g.invalidReason})}),console.log(`Split Validation - Valid: ${p.length}, Invalid: ${m.length}`),p.length===0)return{inserted:0,skipped:a.length,failed:a.length};let b=await S("/api/import/contracts",{method:"POST",body:JSON.stringify({rows:p,onDuplicate:"update"})});if(!b.ok)throw new Error(b.data?.error||"Import gagal");return b.data}}})}M();R();var gt=[],Re=[];function Da(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let i of e)if(t.some(a=>a.period===i))return i;return"Q3"}function Ia(t,e){if(t.period!=="Q3")return!1;let i=String(t.status||"").toLowerCase();if(i!=="selesai"&&i!=="completed"&&i!=="done")return!1;let a=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?a.includes("inspeksi"):e==="gcdc"?a.includes("general cleaning")||a.includes("deep cleaning"):!1}async function Nt(t,e){gt=await K();let i=await ee(),l=(await S(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`)).data?.data||[],d=new Set;l.forEach(o=>{o.pic&&o.pic.trim()&&d.add(o.pic.trim())}),Re=Array.from(d).sort();let r=o=>o&&!Re.find(p=>(typeof p=="object"?p.value:p)===o)?[...Re,o]:Re,s=o=>{if(!o||o==="-"||String(o).trim()==="")return"";let p=String(o).split("-");return p.length===3&&p[0].length===4?`${p[2]}-${p[1]}-${p[0]}`:o},c=Da(l),u=e?e.get("dash_filter"):null;N({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",defaultFilters:u?{period:"Q3"}:{period:c},onDataLoaded:o=>(u&&(o=o.filter(p=>Ia(p,u))),o.sort((p,m)=>{let b=p.opening_date?new Date(p.opening_date).getTime():0;return(m.opening_date?new Date(m.opening_date).getTime():0)-b})),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:o=>st(o)},{key:"period",label:"Periode",render:o=>oe(o)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:o=>s(o)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:o=>s(o)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:o=>s(o)},{key:"status",label:"Status",render:o=>q(o)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:gt},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Re}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:gt,value:o?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:o?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:o?.period},{name:"pic",label:"PIC",type:"combobox",options:r(o?.pic),value:o?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:o?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:o?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:o?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:o?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let o=await S(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(o.ok){let p=o.data.data.map(m=>({Cabang:m.branch_name||"",Kegiatan:m.activity_type||"",Periode:m.period||"",PIC:m.pic||"","Tgl Opening":m.opening_date||"","Tgl Target":m.target_date||"","Tgl Selesai":m.completion_date||"",Status:m.status||""}));L(p,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async o=>{let m=(await S("/api/branches?all=1")).data?.data||[],b=k=>{if(!k)return null;let x=String(k||"").toLowerCase(),_=m.find(w=>String(w.full_name||"").toLowerCase()===x||String(w.code||"").toLowerCase()===x||String(w.name||"").toLowerCase()===x);return _?_.id:null},g=k=>{if(k==null||k==="")return"";if(k instanceof Date&&!isNaN(k.getTime()))return k.toISOString().slice(0,10);let x=String(k).trim();if(x===""||x==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(x))return x.slice(0,10);if(/^\d{4,5}$/.test(x)){let w=Number(x);if(w>2e4&&w<99999){let C=new Date(Date.UTC(1899,11,30)+w*864e5);return isNaN(C.getTime())?"":C.toISOString().slice(0,10)}}let _=x.split(/[\/\-\.]/);if(_.length===3){let[w,C,P]=_.map(E=>E.trim());if(w.length===4&&C.length<=2&&P.length<=2)return`${w}-${C.padStart(2,"0")}-${P.padStart(2,"0")}`;if(P.length===4&&C.length<=2&&w.length<=2)return`${P}-${C.padStart(2,"0")}-${w.padStart(2,"0")}`}return x},h=o.map(k=>({branch_id:b(String(k.Cabang||"").trim()),activity_type:String(k.Kegiatan||"").trim(),period:String(k.Periode||"").trim(),pic:String(k.PIC||k.Pic||"").trim(),opening_date:g(k["Tgl Opening"]||k["Tanggal Opening"]||k["Tgl Openir"]),target_date:g(k["Tgl Target"]||k["Tanggal Target"]),completion_date:g(k["Tgl Selesai"]||k["Tanggal Selesai"]),status:String(k.Status||"").trim(),notes:String(k.Catatan||k.Keterangan||"").trim()})).filter(k=>k.activity_type&&k.period),f=await S("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:h,onDuplicate:"update"})});if(!f.ok)throw new Error(f.data?.error||"Import gagal");return f.data}}})}M();R();var bt=[],We=[];function Pa(t,e){let i=String(t.status||"").toLowerCase();return e==="open"?i==="open":!1}async function At(t,e){let i=e?e.get("dash_filter"):null;bt=await K(),We=await ee();let a=r=>r&&!We.find(s=>s.value===r)?[...We,{value:r,label:r}]:We,l=new Date().getFullYear();N({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:r=>i?r.filter(s=>Pa(s,i)):r,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:r=>`<span class="badge badge-secondary">${r}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:r=>`<span title="${r}">${r?.length>50?r.slice(0,50)+"\u2026":r}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:r=>`<span title="${r||""}">${r?.length>40?r.slice(0,40)+"\u2026":r||"-"}</span>`},{key:"status",label:"Status",render:r=>q(r)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:r=>window.formatDate(r)},{key:"day_count",label:"Hari",render:r=>r??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:bt},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:["2025","2026","2027","2028","2029","2030"]}],formFields:r=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:r?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:bt,value:r?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:r?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:r?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:r?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(r?.employee_name),value:r?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(r?.fc_specialist),value:r?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:r?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:r?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:r?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let r=await S(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let s=r.data.data.map(n=>({Tanggal:n.report_date||"",Cabang:n.branch_name||"",Kategori:n.category||"",Sumber:n.source||"",Keluhan:n.complaint||"","Nama FC":n.employee_name||"","FC Spesialis":n.fc_specialist||"",Solusi:n.solution||"","Tgl Selesai":n.completion_date||"",Status:n.status||""}));L(s,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async r=>{let n=(await S("/api/branches?all=1")).data?.data||[],c=p=>{if(!p)return null;let m=String(p||"").toLowerCase(),b=n.find(g=>String(g.full_name||"").toLowerCase()===m||String(g.code||"").toLowerCase()===m||String(g.name||"").toLowerCase()===m);return b?b.id:null},u=r.map(p=>({branch_id:c(String(p.Cabang||"").trim()),report_date:String(p.Tanggal||"").trim(),category:String(p.Kategori||"").trim(),source:String(p.Sumber||"").trim(),complaint:String(p.Keluhan||"").trim(),employee_name:String(p["Nama FC"]||"").trim(),fc_specialist:String(p["FC Spesialis"]||"").trim(),solution:String(p.Solusi||"").trim(),completion_date:String(p["Tgl Selesai"]||"").trim(),status:String(p.Status||"").trim()})).filter(p=>p.report_date&&p.complaint&&p.category),o=await S("/api/import/issues",{method:"POST",body:JSON.stringify({rows:u,onDuplicate:"update"})});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}}})}M();var Te=[];function Ba(t,e){let i=String(t.status||"").toLowerCase();return e==="pending"?i==="pending":!1}async function Ft(t,e){let i=e?e.get("dash_filter"):null;Te=await K();let a=await ee(),l=["Ade","Berlin"],d=s=>s&&!a.find(n=>n.value===s)?[...a,{value:s,label:s}]:a,r=s=>s&&!l.find(n=>(typeof n=="object"?n.value:n)===s)?[...l,s]:l;N({container:t,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:s=>i?s.filter(n=>Ba(n,i)):s,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:s=>`<span title="${s||""}">${s?.length>50?s.slice(0,50)+"\u2026":s||"-"}</span>`},{key:"solution",label:"Solusi",render:s=>`<span title="${s||""}">${s?.length>40?s.slice(0,40)+"\u2026":s||"-"}</span>`},{key:"status",label:"Status",render:s=>q(s)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:s=>window.formatDate(s)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:s=>s?`<a href="${s}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"combobox",name:"branch_id",label:"Cabang",options:Te},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async s=>{let n=new URLSearchParams(s||{}).toString(),c=await S(`/api/one-on-one?limit=10000&${n}`);if(c.ok){let u=c.data.data.map(p=>({Tanggal:p.meeting_date||"",Cabang:p.branch_name||"","Nama Karyawan":p.employee_name||"",PIC:p.pic||"",Masalah:p.problem||"",Solusi:p.solution||"",Status:p.status||"","Tgl Selesai":p.completion_date||"",Dokumen:p.document_link||""})),{downloadExcel:o}=await Promise.resolve().then(()=>(R(),re));o(u,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let s=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(R(),re));n(s,"Template_Import_OneOnOne")},onImport:async s=>{let n=p=>{if(!p)return null;let m=String(p||"").toLowerCase(),b=Te.find(g=>String(g.label||"").toLowerCase()===m);return b?b.value:null},c=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let m=String(p).trim();if(/^\d{4,5}$/.test(m)){let g=Number(m);if(g>2e4&&g<99999){let h=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);let b=m.split(/[\/\-\.]/);if(b.length===3){let[g,h,f]=b.map(k=>k.trim());if(g.length===4&&h.length<=2&&f.length<=2)return`${g}-${h.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&h.length<=2&&g.length<=2)return`${f}-${h.padStart(2,"0")}-${g.padStart(2,"0")}`}return m},u=s.map(p=>({meeting_date:c(p.Tanggal),employee_name:String(p["Nama Karyawan"]||"").trim(),branch_id:n(String(p.Cabang||"").trim()),pic:String(p.PIC||"").trim(),problem:String(p.Masalah||"").trim(),solution:String(p.Solusi||"").trim(),status:String(p.Status||"").trim(),completion_date:c(p["Tgl Selesai"]),document_link:String(p.Dokumen||"").trim()})).filter(p=>p.meeting_date&&p.employee_name&&p.branch_id),o=await S("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:u,onDuplicate:"update"})});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}},formFields:s=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:s?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:s?.branch_id&&!Te.find(n=>n.value==s.branch_id)?[...Te,{value:s.branch_id,label:s.branch_name||s.branch_id}]:Te,createApi:{path:"/api/branches",field:"full_name"},value:s?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"combobox",required:!0,options:d(s?.employee_name),value:s?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:r(s?.pic),createApi:{path:"/api/pic",field:"name"},value:s?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:s?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:s?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:s?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:s?.document_link}]})}M();async function Mt(t){let e=await K(),i=await ee(),a=["Ade","Berlin"],l=s=>s&&!i.find(n=>n.value===s)?[...i,{value:s,label:s}]:i,d=s=>s&&!a.find(n=>(typeof n=="object"?n.value:n)===s)?[...a,s]:a,r=Array.from({length:5},(s,n)=>String(new Date().getFullYear()-n));N({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:s=>window.formatDate(s)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:s=>{try{let n=JSON.parse(s);return Array.isArray(n)?n.join(", "):s||"-"}catch{return s||"-"}}},{key:"score",label:"Nilai",render:s=>s!=null?`<strong>${s}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:s=>s?`<a href="${s}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:r}],exportOptions:{moduleName:"training",onExport:async s=>{let n=new URLSearchParams(s||{}).toString(),c=await S(`/api/training?limit=10000&${n}`);if(c.ok){let u=c.data.data.map(p=>{let m=p.participants||"";try{let b=JSON.parse(m);m=Array.isArray(b)?b.join(", "):m}catch{}return{Tanggal:p.training_date||"",Batch:p.batch||"",Materi:p.subject||"",Cabang:p.branch_name||"",Trainer:p.trainer||"",Peserta:m,Nilai:p.score!==null&&p.score!==void 0?p.score:"",Dokumen:p.document_link||""}}),{downloadExcel:o}=await Promise.resolve().then(()=>(R(),re));o(u,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let s=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(R(),re));n(s,"Template_Import_Training")},onImport:async s=>{let n=p=>{if(!p)return null;let m=String(p||"").toLowerCase(),b=e.find(g=>String(g.label||"").toLowerCase()===m);return b?b.value:null},c=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let m=String(p).trim();if(/^\d{4,5}$/.test(m)){let g=Number(m);if(g>2e4&&g<99999){let h=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);let b=m.split(/[\/\-\.]/);if(b.length===3){let[g,h,f]=b.map(k=>k.trim());if(g.length===4&&h.length<=2&&f.length<=2)return`${g}-${h.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&h.length<=2&&g.length<=2)return`${f}-${h.padStart(2,"0")}-${g.padStart(2,"0")}`}return m},u=s.map(p=>({training_date:c(p.Tanggal),batch:String(p.Batch||"").trim(),subject:String(p.Materi||"").trim(),branch_id:n(String(p.Cabang||"").trim()),trainer:String(p.Trainer||"").trim(),participants:String(p.Peserta||"").trim(),score:p.Nilai?Number(p.Nilai):null,document_link:String(p.Dokumen||"").trim()})).filter(p=>p.training_date&&p.subject&&p.branch_id),o=await S("/api/import/training",{method:"POST",body:JSON.stringify({rows:u,onDuplicate:"update"})});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}},formFields:s=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:s?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:s?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:s?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:s?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:d(s?.trainer),value:s?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let n=JSON.parse(s?.participants);return Array.isArray(n)?n.join(", "):s?.participants||""}catch{return s?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:s?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:s?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:s?.notes}],onBeforeSubmit:async s=>(s.participants&&(s.participants=JSON.stringify(s.participants.split(",").map(n=>n.trim()).filter(Boolean))),s)})}M();he();R();function Ot({container:t,title:e,icon:i,apiPath:a,columns:l,formFields:d,filterFields:r,defaultFilters:s={},itemLabel:n="Data",canCreate:c=!0,canEdit:u=!0,canDelete:o=!0,onBeforeSubmit:p,onAfterLoad:m,onDataLoaded:b,extraActions:g=[],initialSearch:h="",exportOptions:f=null,bulkDelete:k=!1,paginationMode:x="server"}){let _=1,w={...s};h&&(w.search=h);let C=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions">
        ${c?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
      </div>
    </div>

    ${k?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="align-items:center; background:var(--bg-card); border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${f?Me(f.moduleName):""}

    ${r&&r.length>0?`
    <div class="filter-bar card" style="padding: 1rem; display:flex; flex-wrap:wrap; gap:1rem; align-items:center;">
        ${r.map(v=>{if(v.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${w.search||""}"></div>`;if(v.type==="search-combo"){let y="dl-filter-search",$=(v.options||[]).map(B=>`<option value="${typeof B=="object"?B.label:B}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${y}" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${w.search||""}"><datalist id="${y}">${$}</datalist></div>`}return v.type==="select"?`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${v.name}" id="filter-${v.name}"><option value="">-- ${v.label} --</option>${(v.options||[]).map(y=>`<option value="${typeof y=="object"?y.value:y}" ${w[v.name]===(typeof y=="object"?y.value:y)?"selected":""}>${typeof y=="object"?y.label:y}</option>`).join("")}</select>`:""}).join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function P(){if(!document.getElementById("bulk-toolbar"))return;let y=document.getElementById("bulk-count"),$=document.getElementById("btn-bulk-delete"),B=document.getElementById("btn-bulk-cancel");y.textContent=`${C.size} item dipilih`,C.size>0?($.disabled=!1,B.disabled=!1):($.disabled=!0,B.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{C.clear(),document.querySelectorAll(".row-checkbox").forEach(y=>y.checked=!1);let v=document.getElementById("select-all-checkbox");v&&(v.checked=!1),P()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(C.size===0)return;let v=[...C],y=document.createElement("div");y.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",y.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${v.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${v.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(y),y.querySelector("#bulk-cancel-btn").addEventListener("click",()=>y.remove()),y.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let $=y.querySelector("#bulk-confirm-btn");$.disabled=!0,$.textContent="Menghapus...";let B=await S(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:v})});y.remove(),B.ok?(Z(`${v.length} ${n} berhasil dihapus.`),C.clear(),P(),O()):G(B.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),J;if(E?.addEventListener("input",v=>{clearTimeout(J),J=setTimeout(()=>{w.search=v.target.value,_=1,C.clear(),O()},400)}),r?.forEach(v=>{v.type==="select"&&document.getElementById(`filter-${v.name}`)?.addEventListener("change",y=>{w[v.name]=y.target.value,_=1,C.clear(),O()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{w={...s},E&&(E.value=""),r?.forEach(v=>{let y=document.getElementById(`filter-${v.name}`);y&&(y.value="")}),_=1,C.clear(),O()}),document.getElementById("btn-create")?.addEventListener("click",()=>pe(null)),f){document.getElementById(`btn-export-${f.moduleName}`)?.addEventListener("click",async y=>{let $=y.target,B=$.innerHTML;$.innerHTML="\u23F3 Loading...",$.disabled=!0;try{await f.onExport()}catch{G("Gagal export data")}finally{$.innerHTML=B,$.disabled=!1}}),document.getElementById(`btn-template-${f.moduleName}`)?.addEventListener("click",()=>{f.onTemplate()});let v=document.getElementById(`input-import-${f.moduleName}`);v?.addEventListener("change",async y=>{let $=y.target.files[0];if(!$)return;v.disabled=!0;let B=document.createElement("div");B.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",B.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(B);let V=B.querySelector("#import-progress-text"),U=B.querySelector("#import-progress-bar"),D=B.querySelector("#import-summary"),I=B.querySelector("#import-close-btn");I.addEventListener("click",()=>{B.remove(),O()});try{let Y=await Fe($);if(Y.length===0)throw new Error("File kosong atau format salah");let X=500,ie=0,ae=0,A=0,F=Y.length;V.textContent=`Ditemukan ${F} baris data. Memulai import...`;for(let T=0;T<F;T+=X){let z=Y.slice(T,T+X);V.textContent=`Mengimport baris ${T+1} - ${Math.min(T+X,F)} dari ${F}...`,U.style.width=`${Math.round(T/F*100)}%`;try{let H=await f.onImport(z);H?(ie+=H.inserted||H.metrics?.inserted||z.length,ae+=H.skipped||H.metrics?.updated||0):ie+=z.length}catch(H){console.error("Chunk import failed:",H),A+=z.length}}U.style.width="100%",V.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',D.style.display="block",D.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${F}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ie}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${A}</strong></div>
        `,A>0&&(D.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),I.style.display="block",v.value=""}catch(Y){V.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${Y.message}`,U.style.background="var(--danger)",U.style.width="100%",I.style.display="block",v.value=""}finally{v.disabled=!1}})}async function O(){P();let v=document.getElementById("table-container");if(!v)return;v.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let y=x==="client",$=y?1:_,B=y?ge:20,V=new URLSearchParams({page:$,limit:B,...Object.fromEntries(Object.entries(w).filter(([,A])=>A))}),U=await S(`${a}?${V}`);if(!U.ok){v.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${U.data?.error||"Error"}</p></div>`;return}let D=U.data?.data||U.data||[],I=U.data?.pagination,Y=D.length,X=D;if(y){D=b(D),X=D;let A=D.length,F=20,T=Math.ceil(A/F);_>T&&T>0&&(_=T);let z=(_-1)*F,H=_*F;D=D.slice(z,H),I={page:_,limit:F,total:A,pages:T}}!1,m&&m(D);let ie=je({columns:l,data:D,fullData:X,onEdit:u?A=>pe(A):null,actions:g.map(A=>({...A,handler:F=>A.handler(F,O)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:k?{selectedIds:C,onToggle:P}:null});v.innerHTML="",v.appendChild(ie);let ae=document.getElementById("pagination-container");if(ae&&(ae.innerHTML="",I&&I.pages>1)){let A=Je({page:I.page,pages:I.pages,total:I.total,limit:I.limit,onPage:F=>{_=F,O()}});A&&ae.appendChild(A)}}function ye(v){let y=typeof d=="function"?d(v):d;return Ae(y)}function pe(v){let y=!!v,$=document.createElement("form");if($.noValidate=!0,$.innerHTML=ye(v),y){let V=typeof d=="function"?d(v):d;Ge($,v)}let{close:B}=le({title:y?`Edit ${n}`:`Tambah ${n}`,content:$,size:"lg",confirmText:y?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(V,U)=>{if(!$.reportValidity())return;let D=V.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let I=Ue($),Y=typeof d=="function"?d(v):d,X=async F=>{for(let T of F)if(T.type==="row")await X(T.fields);else if(T.type==="combobox"&&I[T.name]){let z=I[T.name],H=(T.options||[]).find(W=>{let ne=String(typeof W=="object"?W.value:W),et=String(typeof W=="object"?W.label:W);return ne===z||et===z});if(H)I[T.name]=typeof H=="object"?H.value:H;else if(T.createApi){let W={};W[T.createApi.field]=z,T.createApi.extra&&Object.assign(W,T.createApi.extra);let ne=await S(T.createApi.path,{method:"POST",body:JSON.stringify(W)});if(ne.ok&&ne.data?.id)I[T.name]=ne.data.id;else if(ne.ok&&!ne.data?.id)I[T.name]=z;else throw new Error(`Gagal membuat master data: ${ne.data?.error||"Unknown error"}`)}}};try{await X(Y)}catch(F){G(F.message),D.disabled=!1,D.textContent=y?"Simpan Perubahan":`Tambah ${n}`;return}p&&(I=await p(I,v));let ie=y?"PUT":"POST",ae=y?`${a}/${v.id}`:a,A=await S(ae,{method:ie,body:JSON.stringify(I)});A.ok?(Z(y?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),U(),O()):(G(A.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=y?"Simpan Perubahan":`Tambah ${n}`)}})}function fe(v){Ne(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let y=await S(`${a}/${v.id}`,{method:"DELETE"});y.ok?(Z(`${n} berhasil dihapus.`),O()):G(y.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return O(),O}M();R();async function Rt(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await K(),a=await ee(),l=e?e.get("dash_filter"):null;console.log("RAW",await xe()),console.log("OPTIONS",a);let d=n=>n&&!a.find(c=>c.value===n)?[...a,{value:n,label:n}]:a,r=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],s=n=>n&&!r.includes(n)?[...r,n]:r;Ot({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",onDataLoaded:n=>{if(l==="reliever"){let c=new Date,u=c.getFullYear(),o=String(c.getMonth()+1).padStart(2,"0");return n.filter(p=>{if(String(p.status||"").toLowerCase()!=="done")return!1;let m=p.backup_date||"";if(m.includes("/")){let b=m.split("/");if(b.length===3&&(b[2].length===4?b[2]:`20${b[2]}`)==u&&b[1].padStart(2,"0")==o)return!0}else if(m.includes("-")&&m.startsWith(`${u}-${o}`))return!0;return!1})}return n},columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:n=>oe(n)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:n=>window.formatDate(n)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:n=>n?`<span class="badge badge-info">${n}</span>`:"-"},{key:"status",label:"Status",render:n=>q(n)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:r},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:n?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:d(n?.original_fc_name),value:n?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:s(n?.reliever_name),value:n?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:n?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:n?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:n?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:n?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let n=await S(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let c=n.data.data.map(u=>({Cabang:u.branch_name||"","Nama Facility care":u.original_fc_name||"",Periode:u.period||"",Relifer:u.reliever_name||"","Tanggal Back Up":u.backup_date||"","Tanggal Selesai":u.completion_date||"",Keterangan:u.reason||"",Shift:u.shift||"",Status:u.status||""}));c.length===0&&c.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),L(c,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async n=>{let u=(await S("/api/branches?all=1")).data?.data||[],o=b=>{if(!b)return null;let g=String(b||"").toLowerCase(),h=u.find(f=>String(f.full_name||"").toLowerCase()===g||String(f.code||"").toLowerCase()===g||String(f.name||"").toLowerCase()===g);return h?h.id:null},p=n.map(b=>({branch_name:String(b.Cabang||"").trim(),backup_date:String(b["Tanggal Back Up"]||b["Tanggal Backup"]||"").trim(),original_fc_name:String(b["Nama Facility care"]||b["FC Digantikan"]||"").trim(),reliever_name:String(b.Relifer||b.Reliefer||"").trim(),period:String(b.Periode||"").trim(),reason:String(b.Keterangan||"").trim(),shift:String(b.Shift||"").trim(),completion_date:String(b["Tanggal Selesai"]||"").trim(),status:String(b.Status||"").trim()})).filter(b=>b.reliever_name&&b.backup_date),m=await S("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:p,onDuplicate:"update"})});if(!m.ok)throw new Error(m.data?.error||"Import gagal");return m.data}}})}M();R();async function Kt(t){let e=await K(),i=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));N({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>oe(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>q(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),d=await S(`/api/reports/inspection?limit=10000&${l}`);if(d.ok){let r=d.data.data.map(s=>({Cabang:s.branch_name||"",Periode:s.period||"",Tanggal:s.inspection_date||"","Point FC":s.fc_score!==null&&s.fc_score!==void 0?s.fc_score:"","Point SPV":s.spv_score!==null&&s.spv_score!==void 0?s.spv_score:"",Status:s.status||"","Link Dokumen":s.document_link||""}));L(r,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let l=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),u=e.find(o=>String(o.label||"").toLowerCase()===c);return u?u.value:null},d=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let c=String(n).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let o=Number(c);if(o>2e4&&o<99999){let p=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let u=c.split(/[\/\-\.]/);if(u.length===3){let[o,p,m]=u.map(b=>b.trim());if(o.length===4&&p.length<=2&&m.length<=2)return`${o}-${p.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&p.length<=2&&o.length<=2)return`${m}-${p.padStart(2,"0")}-${o.padStart(2,"0")}`}return c},r=a.map(n=>({branch_id:l(String(n.Cabang||"").trim()),period:String(n.Periode||"").trim(),inspection_date:d(n.Tanggal),fc_score:n["Point FC"]!==void 0&&n["Point FC"]!==""?Number(n["Point FC"]):null,spv_score:n["Point SPV"]!==void 0&&n["Point SPV"]!==""?Number(n["Point SPV"]):null,status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.inspection_date),s=await S("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}M();R();async function Ht(t){let e=await K(),i=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));N({container:t,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>oe(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>q(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),d=await S(`/api/reports/cleaning?limit=10000&${l}`);if(d.ok){let r=d.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||""}));L(r,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let l=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),u=e.find(o=>String(o.label||"").toLowerCase()===c);return u?u.value:null},d=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let c=String(n).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let o=Number(c);if(o>2e4&&o<99999){let p=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let u=c.split(/[\/\-\.]/);if(u.length===3){let[o,p,m]=u.map(b=>b.trim());if(o.length===4&&p.length<=2&&m.length<=2)return`${o}-${p.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&p.length<=2&&o.length<=2)return`${m}-${p.padStart(2,"0")}-${o.padStart(2,"0")}`}return c},r=a.map(n=>({branch_id:l(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:d(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),s=await S("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}M();R();async function qt(t){let e=await K(),i=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));N({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge badge-warning">${a}</span>`},{key:"period",label:"Periode",render:a=>oe(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>q(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:a=>a||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:a?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),d=await S(`/api/reports/fogging?limit=10000&${l}`);if(d.ok){let r=d.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"Fogging",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||""}));L(r,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async a=>{let l=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),u=e.find(o=>String(o.label||"").toLowerCase()===c);return u?u.value:null},d=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let c=String(n).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let o=Number(c);if(o>2e4&&o<99999){let p=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let u=c.split(/[\/\-\.]/);if(u.length===3){let[o,p,m]=u.map(b=>b.trim());if(o.length===4&&p.length<=2&&m.length<=2)return`${o}-${p.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&p.length<=2&&o.length<=2)return`${m}-${p.padStart(2,"0")}-${o.padStart(2,"0")}`}return c},r=a.map(n=>({branch_id:l(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"Fogging").trim(),period:String(n.Periode||"").trim(),activity_date:d(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.activity_date),s=await S("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(r)});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}}})}M();R();async function jt(t){let e=await K(),i=await ee(),a=i,l=r=>r&&!i.find(s=>s.value===r)?[...i,{value:r,label:r}]:i,d=r=>r&&!a.find(s=>s.value===r)?[...a,{value:r,label:r}]:a;N({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:r=>`<span title="${r||""}">${r?.length>60?r.slice(0,60)+"\u2026":r||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:r=>window.formatDate(r)},{key:"status",label:"Status",render:r=>q(r)},{key:"notes",label:"Keterangan",render:r=>r?.length>40?r.slice(0,40)+"\u2026":r||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]}],formFields:r=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:r?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:d(r?.pic),value:r?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:r?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:r?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:r?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:r?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async r=>{let s=new URLSearchParams(r||{}).toString(),n=await S(`/api/reports/basecamp?limit=10000&${s}`);if(n.ok){let c=n.data.data.map(u=>({"Tgl Info":u.info_date||"",Cabang:u.branch_name||"",Permasalahan:u.problem||"",PIC:u.pic||"","Tgl Done":u.done_date||"",Status:u.status||"",Keterangan:u.notes||""}));L(c,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async r=>{let s=o=>{if(!o)return null;let p=String(o||"").toLowerCase(),m=e.find(b=>String(b.label||"").toLowerCase()===p);return m?m.value:null},n=o=>{if(o==null||o==="")return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let p=String(o).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let b=Number(p);if(b>2e4&&b<99999){let g=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}let m=p.split(/[\/\-\.]/);if(m.length===3){let[b,g,h]=m.map(f=>f.trim());if(b.length===4&&g.length<=2&&h.length<=2)return`${b}-${g.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&g.length<=2&&b.length<=2)return`${h}-${g.padStart(2,"0")}-${b.padStart(2,"0")}`}return p},c=r.map(o=>({info_date:n(o["Tgl Info"]||o["Tanggal Info"]),branch_id:s(String(o.Cabang||"").trim()),problem:String(o.Permasalahan||"").trim(),pic:String(o.PIC||"").trim(),done_date:n(o["Tgl Done"]||o["Tanggal Done"]),status:String(o.Status||"").trim(),notes:String(o.Keterangan||o.Catatan||"").trim()})).filter(o=>o.info_date&&o.branch_id&&o.problem),u=await S("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(c)});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}async function Jt(t){N({container:t,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(M(),ve)),l=await a(`/api/sop?limit=10000&${i}`);if(l.ok){let d=l.data.data.map(s=>({"Nama SOP":s.name||"",Kategori:s.category||"",Dokumen:s.document_link||"",Catatan:s.notes||s.description||""})),{downloadExcel:r}=await Promise.resolve().then(()=>(R(),re));r(d,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(R(),re));i(e,"Template_Import_SOP")},onImport:async e=>{let i=e.map(d=>({name:String(d["Nama SOP"]||"").trim(),category:String(d.Kategori||"").trim(),document_link:String(d.Dokumen||"").trim(),description:String(d.Catatan||"").trim()})).filter(d=>d.name),{apiFetch:a}=await Promise.resolve().then(()=>(M(),ve)),l=await a("/api/sop/import",{method:"POST",body:JSON.stringify(i)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function Ut(t){N({container:t,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(M(),ve)),l=await a(`/api/checklist?limit=10000&${i}`);if(l.ok){let d=l.data.data.map(s=>({"Nama Checklist":s.name||"",Kategori:s.category||"",Dokumen:s.document_link||"",Deskripsi:s.description||""})),{downloadExcel:r}=await Promise.resolve().then(()=>(R(),re));r(d,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(R(),re));i(e,"Template_Import_Checklist")},onImport:async e=>{let i=e.map(d=>({name:String(d["Nama Checklist"]||"").trim(),category:String(d.Kategori||"").trim(),document_link:String(d.Dokumen||"").trim(),description:String(d.Deskripsi||"").trim()})).filter(d=>d.name),{apiFetch:a}=await Promise.resolve().then(()=>(M(),ve)),l=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(i)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}M();he();R();async function ht(t,e="forms"){if(e==="supply")return Na(t);La(t)}function La(t){N({container:t,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await S(`/api/forms?limit=10000&${i}`);a.data?.data?L(a.data.data,"Data_Master_Form"):G("Gagal export data master form")},onImport:async e=>{let i=await S("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!i.ok)throw new Error(i.data?.error||"Import failed");return i.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function Na(t){let i=((await S("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));N({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,l)=>l.branch_name_ref||l.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>q(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let l=a?.tools_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}let d=a?.chemical_items;try{d=Array.isArray(JSON.parse(d))?JSON.parse(d).join(", "):d}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:a?.branch_id&&!i.find(r=>r.value==a.branch_id)?[...i,{value:a.branch_id,label:a.branch_name||a.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:l},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:d},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let l=new URLSearchParams(a||{}).toString(),d=await S(`/api/reports/supply?limit=10000&${l}`);if(d.ok){let r=d.data.data.map(s=>{let n=s.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let c=s.chemical_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}return{Waktu:s.submitted_at||"",Pengirim:s.submitter_name||"",Cabang:s.branch_name_ref||s.branch_name||"","Alat/Barang":n||"",Chemical:c||"",Catatan:s.additional_notes||"",Status:s.status||"","Diproses Oleh":s.processed_by||""}});L(r,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let d=(await S("/api/branches?all=1")).data?.data||[],r=u=>{if(!u)return null;let o=String(u||"").toLowerCase(),p=d.find(m=>String(m.full_name||"").toLowerCase()===o||String(m.code||"").toLowerCase()===o||String(m.name||"").toLowerCase()===o);return p?p.id:null},s=u=>{if(u==null||u==="")return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let o=String(u).trim();if(o===""||o==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(o))return o.slice(0,10);if(/^\d{4,5}$/.test(o)){let m=Number(o);if(m>2e4&&m<99999){let b=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}let p=o.split(/[\/\-\.]/);if(p.length===3){let[m,b,g]=p.map(h=>h.trim());if(m.length===4&&b.length<=2&&g.length<=2)return`${m}-${b.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&b.length<=2&&m.length<=2)return`${g}-${b.padStart(2,"0")}-${m.padStart(2,"0")}`}return o},n=a.map(u=>({submitted_at:s(u.Waktu||u.Tanggal),submitter_name:String(u.Pengirim||"").trim(),branch_id:r(String(u.Cabang||"").trim()),tools_items:String(u["Alat/Barang"]||u.Alat||"").trim(),chemical_items:String(u.Chemical||"").trim(),additional_notes:String(u.Catatan||u.Keterangan||"").trim(),status:String(u.Status||"").trim(),processed_by:String(u["Diproses Oleh"]||u.PIC||"").trim()})).filter(u=>u.submitted_at&&u.submitter_name&&u.branch_id),c=await S("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!c.ok)throw new Error(c.data?.error||"Import gagal");return c.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,l)=>{let d=le({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(r,s)=>{let n=r.querySelector("#supply-status").value,c=r.querySelector("#supply-processed-by").value;(await S(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:c})})).ok?(Z("Status diperbarui."),s(),l()):G("Gagal update status.")}})}}]})}M();R();async function Gt(t){let e=be();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}N({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let a=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await S(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let a=i.data.data.map(l=>({"Nama Lengkap":l.full_name||"",Username:l.username||"",Email:l.email||"",Role:l.role||"",Status:l.is_active?"Aktif":"Nonaktif"}));L(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let a=i.map(d=>({full_name:String(d["Nama Lengkap"]||"").trim(),username:String(d.Username||"").trim(),email:String(d.Email||"").trim(),role:String(d.Role||"").trim()||"viewer",password:String(d.Password||"").trim()})).filter(d=>d.username&&d.password&&d.email&&d.full_name),l=await S("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!l.ok)throw new Error(l.data?.error||"Import gagal");return l.data}}})}M();R();async function Qt(t){N({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await S(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)L(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let i=e.map(l=>({code:String(l["Kode Cabang"]||"").trim(),name:String(l["Nama Pendek"]||"").trim(),full_name:String(l["Nama Lengkap"]||"").trim(),city:String(l.Kota||"").trim()})).filter(l=>l.code&&l.name),a=await S("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}M();async function Vt(t){let e=new Date,i=[];t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F4C5} Kalender</h1>
    </div>
    <div class="card" style="overflow:visible;">
      <div class="card-header calendar-nav" style="flex-wrap:wrap;gap:10px;padding:14px 20px;border-bottom:1px solid var(--border);">
        <div style="display:flex;align-items:center;gap:8px;flex-shrink:0;">
          <button class="btn btn-ghost btn-sm" id="cal-prev">\u2039 Prev</button>
          <span class="calendar-month-label" id="cal-month-label" style="min-width:160px;text-align:center;font-size:1.1rem;font-weight:700;"></span>
          <button class="btn btn-ghost btn-sm" id="cal-next">Next \u203A</button>
        </div>
        <div class="calendar-filters" style="display:flex;flex-wrap:wrap;gap:6px 14px;align-items:center;">
          <label class="filter-check"><input type="checkbox" value="schedule"        checked class="cal-filter"> \u{1F5D3} Jadwal</label>
          <label class="filter-check"><input type="checkbox" value="issue"           checked class="cal-filter"> \u26A0\uFE0F Permasalahan</label>
          <label class="filter-check"><input type="checkbox" value="reliever"        checked class="cal-filter"> \u{1F504} Reliefer</label>
          <label class="filter-check"><input type="checkbox" value="training"        checked class="cal-filter"> \u{1F393} Training</label>
          <label class="filter-check"><input type="checkbox" value="contract_expiry" checked class="cal-filter"> \u{1F4CB} Kontrak Habis</label>
          <label class="filter-check"><input type="checkbox" value="one_on_one"      checked class="cal-filter"> \u{1F4AC} One on One</label>
          <label class="filter-check"><input type="checkbox" value="inspection"      checked class="cal-filter"> \u{1F50D} Inspeksi</label>
          <label class="filter-check"><input type="checkbox" value="cleaning"        checked class="cal-filter"> \u{1F9F9} Cleaning</label>
          <label class="filter-check"><input type="checkbox" value="fogging"         checked class="cal-filter"> \u{1F4A8} Fogging</label>
          <label class="filter-check"><input type="checkbox" value="basecamp"        checked class="cal-filter"> \u{1F3D5} Basecamp</label>
        </div>
      </div>
      <div class="card-body p-0">
        <div id="calendar-grid" style="min-height:420px;"></div>
      </div>
    </div>
    <!-- Sidebar detail event -->
    <div id="cal-event-list" class="cal-event-sidebar" style="display:none;">
      <div class="cal-event-header" style="display:flex;justify-content:space-between;align-items:center;padding:12px 16px;border-bottom:1px solid var(--border);font-weight:700;">
        <span id="cal-event-date"></span>
        <button class="btn btn-ghost btn-sm" id="cal-event-close">&times;</button>
      </div>
      <div id="cal-event-items" style="padding:12px;display:flex;flex-direction:column;gap:8px;max-height:60vh;overflow-y:auto;"></div>
    </div>
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),l()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),l()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(d=>d.addEventListener("change",l));async function a(){try{let d=e.getFullYear(),r=String(e.getMonth()+1).padStart(2,"0");i=(await S(`/api/dashboard/calendar?month=${d}-${r}`)).data?.data||[]}catch(d){console.warn("[Calendar] Gagal memuat events:",d),i=[]}}async function l(){let d=document.getElementById("calendar-grid");if(d){d.innerHTML=`
      <div style="display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:1px;background:var(--border-color,#e5e7eb);">
        ${Array(42).fill('<div style="background:#f8fafc;min-height:72px;"></div>').join("")}
      </div>`,await a();try{let r=e.getFullYear(),s=e.getMonth(),n=document.getElementById("cal-month-label");n&&(n.textContent=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}));let c=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(x=>x.value)),u=i.filter(x=>c.has(x.type)),o={};u.forEach(x=>{let _=(x.event_date||"").slice(0,10);_&&(o[_]||(o[_]=[]),o[_].push(x))});let p=new Date(r,s,1).getDay(),m=new Date(r,s+1,0).getDate(),b=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],g=new Date().toISOString().slice(0,10),h='<div class="calendar-grid" style="display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:1px;background:var(--border-color,#e5e7eb);border-radius:0 0 var(--radius,8px) var(--radius,8px);overflow:hidden;width:100%;">';b.forEach(x=>{h+=`<div class="cal-day-header" style="background:#f8fafc;padding:8px 4px;text-align:center;font-size:0.7rem;font-weight:700;color:#6b7280;text-transform:uppercase;letter-spacing:0.04em;">${x}</div>`});for(let x=0;x<p;x++)h+='<div style="background:#f8fafc;min-height:80px;opacity:0.5;"></div>';for(let x=1;x<=m;x++){let _=`${r}-${String(s+1).padStart(2,"0")}-${String(x).padStart(2,"0")}`,w=o[_]||[],C=_===g,P=w.length>0;h+=`
          <div class="cal-cell${C?" cal-today":""}${P?" cal-has-events":""}"
               data-date="${_}" tabindex="0" role="button"
               style="background:${C?"#EFF6FF":"var(--bg-card,#fff)"};min-height:80px;padding:6px;cursor:${P?"pointer":"default"};overflow:hidden;transition:background 0.15s;border-top:2px solid ${C?"#3B82F6":"transparent"};">
            <div style="font-size:0.78rem;font-weight:${C?"800":"600"};color:${C?"#2563EB":"#374151"};margin-bottom:3px;">${x}</div>
            <div style="display:flex;flex-direction:column;gap:2px;overflow:hidden;">
              ${w.slice(0,3).map(E=>`
                <div class="cal-event-dot cal-color-${E.color||"gray"}"
                     style="font-size:0.6rem;font-weight:600;padding:2px 4px;border-radius:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%;"
                     title="${Ye(E.title||E.type)}">
                  ${Aa(E.title||E.branch_name||zt(E.type),16)}
                </div>
              `).join("")}
              ${w.length>3?`<div style="font-size:0.6rem;color:#6b7280;font-weight:600;margin-top:1px;">+${w.length-3} lagi</div>`:""}
            </div>
          </div>`}let k=(p+m)%7;if(k!==0)for(let x=0;x<7-k;x++)h+='<div style="background:#f8fafc;min-height:80px;opacity:0.5;"></div>';h+="</div>",d.innerHTML=h,d.querySelectorAll(".cal-cell[data-date]").forEach(x=>{x.addEventListener("click",()=>{let _=x.dataset.date,w=o[_]||[];if(!w.length)return;let C=document.getElementById("cal-event-list");document.getElementById("cal-event-date").textContent=new Date(_+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),document.getElementById("cal-event-items").innerHTML=w.map(P=>`
            <div class="cal-event-item cal-color-border-${P.color||"gray"}"
                 style="padding:10px 12px;border-left:4px solid;border-radius:6px;background:var(--bg-body,#f9fafb);">
              <div style="font-size:0.7rem;font-weight:700;color:var(--text-3,#9ca3af);text-transform:uppercase;letter-spacing:0.05em;margin-bottom:3px;">${zt(P.type)}</div>
              <div style="font-size:0.85rem;font-weight:600;color:var(--text-1,#111827);margin-bottom:2px;">${Ye(P.title||"-")}</div>
              ${P.branch_name?`<div style="font-size:0.75rem;color:var(--text-2,#6b7280);">\u{1F4CD} ${Ye(P.branch_name)}</div>`:""}
              ${P.status?`<div style="font-size:0.72rem;font-weight:600;margin-top:3px;color:${P.status==="Done"||P.status==="Selesai"?"#10b981":"#f59e0b"};">${Ye(P.status)}</div>`:""}
              ${P.days_remaining!==void 0?`<div style="font-size:0.72rem;color:#6366F1;font-weight:700;margin-top:3px;">Sisa: ${P.days_remaining} hari</div>`:""}
            </div>
          `).join(""),C.style.display="block"}),x.addEventListener("mouseenter",()=>{o[x.dataset.date]?.length&&(x.style.background="#F0F9FF")}),x.addEventListener("mouseleave",()=>{x.style.background=x.classList.contains("cal-today")?"#EFF6FF":"var(--bg-card,#fff)"})})}catch(r){console.error("[Calendar] Render error:",r),d&&(d.innerHTML=`
          <div style="padding:60px;text-align:center;color:var(--text-3,#9ca3af);">
            <div style="font-size:2.5rem;margin-bottom:12px;">\u{1F4C5}</div>
            <div style="font-size:1rem;font-weight:600;">Gagal memuat kalender.</div>
            <div style="font-size:0.85rem;margin-top:6px;">Silakan refresh halaman.</div>
          </div>`)}}}l()}function Aa(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function Ye(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function zt(t){return{schedule:"\u{1F5D3} Jadwal Kegiatan",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis",one_on_one:"\u{1F4AC} One on One",inspection:"\u{1F50D} Inspeksi Hygiene",cleaning:"\u{1F9F9} General/Deep Cleaning",fogging:"\u{1F4A8} Fogging",basecamp:"\u{1F3D5} Basecamp",supply:"\u{1F4E6} Permintaan Barang"}[t]||t}M();async function Wt(t){let e=be(),i=(e?.full_name||e?.username||"U")[0].toUpperCase(),l={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${l},${l}99)">
            ${i}
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
  `;let d=localStorage.getItem("fm_token"),r=document.getElementById("session-info");if(d&&r)try{let s=JSON.parse(atob(d.split(".")[1])),n=new Date(s.exp*1e3);r.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{r.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async s=>{s.preventDefault();let n=document.getElementById("pwd-error"),c=document.getElementById("pwd-success"),u=document.getElementById("btn-save-pwd");n.style.display="none",c.style.display="none";let o=s.target,p=o.current_password.value,m=o.new_password.value,b=o.confirm_password.value;if(m!==b){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(m.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}u.disabled=!0,u.textContent="\u23F3 Menyimpan...";let g=await S("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:p,new_password:m})});u.disabled=!1,u.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',g.ok?(c.textContent="\u2705 Password berhasil diubah.",c.style.display="block",o.reset(),Z("Password berhasil diubah.")):(n.textContent=g.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}M();var Xe={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function te(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let l=Number(e);if(l>2e4&&l<99999){let d=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(d.getTime())?null:d.toISOString().slice(0,10)}}let i=e.split(/[\/\-\.]/);if(i.length===3){let[l,d,r]=i.map(u=>u.trim()),s=Number(l),n=Number(d),c=Number(r);if(l.length===4&&s>1900)return`${l}-${d.padStart(2,"0")}-${r.padStart(2,"0")}`;if(r.length===4&&c>1900)return s>12?`${r}-${d.padStart(2,"0")}-${l.padStart(2,"0")}`:n>12?`${r}-${l.padStart(2,"0")}-${d.padStart(2,"0")}`:`${r}-${d.padStart(2,"0")}-${l.padStart(2,"0")}`;if(r.length===2&&!isNaN(c)){let u=c>=50?`19${r}`:`20${r}`;return s>12?`${u}-${d.padStart(2,"0")}-${l.padStart(2,"0")}`:`${u}-${d.padStart(2,"0")}-${l.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function Yt(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Fa={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",phone:t["No. Hp"],join_date:te(t["Tanggal Masuk"]),status:t.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:te(t["Tanggal Mulai"]),end_date:te(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:te(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:te(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:te(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:te(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:te(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:te(t["Tanggal Target"]||t["Tgl Target"]),completion_date:te(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:te(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:te(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:te(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:te(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:te(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:te(t["Tanggal Back Up"]),completion_date:te(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:te(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:te(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Ma(t,e){let i=Xe[t];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Fa[i.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let l=[],d=[],r=[];return e.filter(n=>!Yt(n)).forEach((n,c)=>{let u=e.indexOf(n)+2,o=[];a.required.forEach(({key:m,label:b})=>{let g=n[m];if(g==null||String(g).trim()===""){let h=Object.keys(n).filter(f=>f.trim()).join(", ");o.push({column:b,originalValue:g||"",reason:`Kolom "${b}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${h.slice(0,120)}`})}});let p=a.map(n);o.length>0?d.push({row:u,data:p,raw:n,errors:o}):(l.push(n),r.push(p))}),{valid:l,errors:d,mapped:r}}function Xt(t){let e=[];return t.SheetNames.forEach(i=>{let a=Xe[i];if(!a)return;let l=t.Sheets[i],d=window.XLSX.utils.sheet_to_json(l,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),r=Ma(i,d),s=d.filter(n=>!Yt(n));e.push({sheetName:i,module:a.module,label:a.label,total:s.length,valid:r.mapped.length,errorCount:r.errors.length,errors:r.errors,mapped:r.mapped,skipped:!1})}),e}function Zt(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,l])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(l),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function ea(t){let e=window.XLSX,i=e.utils.book_new(),a=!1;return t.forEach(l=>{if(!l.errors||l.errors.length===0)return;a=!0;let d=l.errors.map(s=>({"No. Baris":s.row,"Kolom Gagal":(s.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(s.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(s.data||{}).map(([n,c])=>[n,c??""]))})),r=e.utils.json_to_sheet(d);e.utils.book_append_sheet(i,r,l.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Oa=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function ta(t){t.innerHTML=`
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
              ${Object.entries(Xe).map(([g,{label:h}])=>`<span class="import-sheet-tag">\u{1F4C4} ${g} \u2192 ${h}</span>`).join("")}
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
  `;let e=null,i=null,a=0,l={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function d(g){Object.entries(l).forEach(([h,f])=>{f.style.display=h===g?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let g=document.getElementById("btn-backup-db");g.disabled=!0,g.textContent="\u23F3 Memproses Backup...";try{let h=await S("/api/import/backup");if(h.ok){if(!window.XLSX){G("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let f=window.XLSX,k=f.utils.book_new();Object.entries(h.data.database).forEach(([x,_])=>{let w=_.length>0?_:[{}],C=f.utils.json_to_sheet(w);f.utils.book_append_sheet(k,C,x.substring(0,31))}),f.writeFile(k,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),Z("Backup berhasil diunduh!")}else G("Gagal memproses backup: "+(h.data?.error||"Unknown error"))}catch(h){G("Gagal memproses backup: "+h.message)}finally{g.disabled=!1,g.textContent="\u{1F4E6} Backup Database"}});let r=document.getElementById("btn-sync-google");r&&r.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let g=r.innerHTML;r.innerHTML='<span class="spinner"></span> Menyinkronkan...',r.disabled=!0;try{let h=await S("/api/sync/google-sheets",{method:"POST"});h.ok?alert("Sinkronisasi Berhasil: "+(h.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(h.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{r.innerHTML=g,r.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Zt(),Z("Template Excel berhasil didownload!")});let s=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",g=>{g.stopPropagation(),s.click()}),s.addEventListener("change",g=>{g.target.files[0]&&c(g.target.files[0])}),n.addEventListener("dragover",g=>{g.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",g=>{g.preventDefault(),n.classList.remove("drag-over");let h=g.dataTransfer.files[0];h&&h.name.match(/\.xlsx?$/i)?c(h):G("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,s.value="",document.getElementById("file-info").style.display="none",n.style.display="",d("upload")});async function c(g){e=g,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${g.name} (${(g.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await u(g)}async function u(g){d("validating");let h=document.getElementById("validation-status"),f=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");h.textContent="Membaca file Excel...",f.style.width="20%",await Ke(200);let k=await g.arrayBuffer(),x=window.XLSX.read(k,{type:"array",cellDates:!0});h.textContent=`Memvalidasi ${x.SheetNames.length} sheet...`,f.style.width="50%",await Ke(100),i=Xt(x),f.style.width="100%",h.textContent="Validasi selesai!",await Ke(300),o()}catch(k){d("upload"),G("Gagal memproses file: "+k.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function o(){d("preview");let g=i.filter(E=>!E.skipped).length,h=i.reduce((E,J)=>E+J.total,0),f=i.reduce((E,J)=>E+J.valid,0),k=i.reduce((E,J)=>E+J.errorCount,0),x=h>0?Math.round(f/h*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${g} sheet</span>
      <span class="badge badge-secondary">${h} baris</span>
      <span class="badge badge-success">${f} valid (${x}%)</span>
      ${k>0?`<span class="badge badge-danger">${k} error</span>`:""}
    `;let _=document.getElementById("preview-table-container");_.innerHTML=`
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
          ${i.map((E,J)=>`
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
                ${E.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${J}">\u{1F50D} ${E.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,_.querySelectorAll(".btn-detail-error").forEach(E=>{E.addEventListener("click",()=>{let J=i[Number(E.dataset.idx)];p(J)})});let w=document.getElementById("error-detail-section"),C=document.getElementById("error-detail-container");C.innerHTML="",w.style.display="none";let P=document.getElementById("btn-start-import");f===0?(P.disabled=!0,P.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(P.disabled=!1,k>0?(P.innerHTML=`\u{1F680} Import ${f} Data Valid (${k} dilewati)`,P.title="Baris error akan dilewati, baris valid tetap diimport"):P.innerHTML=`\u{1F680} Mulai Import ${f} Data`)}function p(g){let h=document.getElementById("error-detail-section"),f=document.getElementById("error-detail-container");h.style.display="";let k=g.errors.slice(0,100).map(x=>(Array.isArray(x.errors)?x.errors:[]).map(w=>{let C=typeof w=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${x.row}</span></td>
            <td><strong>${C?w.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${C&&w.originalValue!==void 0?w.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${C?w.reason:w}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${C&&w.aliases?`Gunakan salah satu nama kolom:<br><em>${w.aliases}</em>`:C&&w.hint?w.hint:""}
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
            <tbody>${k||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${g.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,h.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{d("upload"),document.getElementById("file-info").style.display="none",n.style.display="",e=null,s.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;ea(i)?Z("Log error berhasil didownload."):Z("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let g=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";m(g)});async function m(g){d("importing"),a=Date.now();let h=[];Oa.forEach(w=>{let C=i?.find(P=>P.module===w&&P.mapped?.length>0);C&&h.push(C)});let f=document.getElementById("import-steps-list");f.innerHTML=h.map(w=>`
      <div class="import-step-item" id="step-item-${w.module}">
        <span class="step-item-icon" id="step-icon-${w.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${w.label} <span class="step-item-count">(${w.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${w.module}"></span>
      </div>
    `).join("");let k=document.getElementById("import-bar"),x=document.getElementById("import-current-status"),_={totalSheets:h.length,totalRows:h.reduce((w,C)=>w+C.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let w=0;w<h.length;w++){let C=h[w],P=document.getElementById(`step-icon-${C.module}`),E=document.getElementById(`step-status-${C.module}`);P.textContent="\u{1F504}",E.textContent="Mengimport...",x.textContent=`Mengimport ${C.label}...`,k.style.width=`${Math.round(w/h.length*100)}%`;try{let J=await S(`/api/import/${C.module}`,{method:"POST",body:JSON.stringify({rows:C.mapped,onDuplicate:g})});if(J.ok){let O=J.data;_.inserted+=O.inserted||0,_.skipped+=O.skipped||0,_.moduleResults.push({label:C.label,inserted:O.inserted||0,skipped:O.skipped||0,status:"ok"}),P.textContent="\u2705",E.innerHTML=`<span class="badge badge-success">${O.inserted||0} berhasil</span>${O.skipped>0?` <span class="badge badge-neutral">${O.skipped} skip</span>`:""}`}else _.failed++,_.moduleResults.push({label:C.label,inserted:0,skipped:0,status:"error",error:J.data?.error}),P.textContent="\u274C",E.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(J){_.failed++,_.moduleResults.push({label:C.label,inserted:0,skipped:0,status:"error",error:J.message}),P.textContent="\u274C",E.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Ke(150)}k.style.width="100%",x.textContent="Selesai!",await Ke(400),b(_)}function b(g){d("summary");let h=((Date.now()-a)/1e3).toFixed(1),f=g.failed===0;document.getElementById("summary-status-icon").innerHTML=`
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
          ${g.moduleResults.map(k=>`
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,i=null,s.value="",document.getElementById("file-info").style.display="none",n.style.display="",d("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Ke(t){return new Promise(e=>setTimeout(e,t))}M();var Ze=[],aa=[];async function na(t){Ze=await K(),aa=await ee(),N({container:t,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:Ze}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await S(`/api/sp?limit=10000&${i}`);if(a.ok){let l=a.data.data.map(r=>({"Nama Karyawan":r.employee_name||"",Divisi:r.division||"",Cabang:r.branch_name||"","Tanggal Sp":r.tanggal||"","Akhir Sp":r.akhir_sp||"","Jenis Sp":r.sp_type||"","Link Document / Foto":r.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(R(),re));d(l,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(R(),re));i(e,"Template_Import_SP")},onImport:async e=>{let i=r=>{if(!r)return null;let s=String(r||"").toLowerCase(),n=Ze.find(c=>String(c.label||"").toLowerCase()===s);return n?n.value:null},a=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let s=String(r).trim();if(/^\d{4,5}$/.test(s)){let c=Number(s);if(c>2e4&&c<99999){let u=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);let n=s.split(/[\/\-\.]/);if(n.length===3){let[c,u,o]=n.map(p=>p.trim());if(c.length===4&&u.length<=2&&o.length<=2)return`${c}-${u.padStart(2,"0")}-${o.padStart(2,"0")}`;if(o.length===4&&u.length<=2&&c.length<=2)return`${o}-${u.padStart(2,"0")}-${c.padStart(2,"0")}`}return s},l=e.map(r=>({employee_name:String(r["Nama Karyawan"]||"").trim(),division:String(r.Divisi||"").trim(),branch_id:i(String(r.Cabang||"").trim()),tanggal:a(r["Tanggal Sp"]),akhir_sp:a(r["Akhir Sp"]),sp_type:String(r["Jenis Sp"]||"").trim(),document_link:String(r["Link Document / Foto"]||"").trim()})).filter(r=>r.employee_name&&r.branch_id),d=await S("/api/import/sp",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:[{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:aa},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"combobox",name:"branch_id",label:"Cabang",required:!0,options:Ze,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}M();var Ee=[],ia=[];async function ra(t){Ee=await K(),ia=await ee(),N({container:t,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",options:Ee},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",options:Ee}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await S(`/api/mutasi?limit=10000&${i}`);if(a.ok){let l=a.data.data.map(r=>({Tanggal:r.tanggal||"","Nama Karyawan":r.employee_name||"","Cabang Asal":r.from_branch_name||"","Cabang Tujuan":r.to_branch_name||"",Status:r.status||"",Dokumen:r.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(R(),re));d(l,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(R(),re));i(e,"Template_Import_Mutasi")},onImport:async e=>{let i=r=>{if(!r)return null;let s=String(r||"").toLowerCase(),n=Ee.find(c=>String(c.label||"").toLowerCase()===s);return n?n.value:null},a=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let s=String(r).trim();if(/^\d{4,5}$/.test(s)){let c=Number(s);if(c>2e4&&c<99999){let u=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);let n=s.split(/[\/\-\.]/);if(n.length===3){let[c,u,o]=n.map(p=>p.trim());if(c.length===4&&u.length<=2&&o.length<=2)return`${c}-${u.padStart(2,"0")}-${o.padStart(2,"0")}`;if(o.length===4&&u.length<=2&&c.length<=2)return`${o}-${u.padStart(2,"0")}-${c.padStart(2,"0")}`}return s},l=e.map(r=>({tanggal:a(r.Tanggal),employee_name:String(r["Nama Karyawan"]||"").trim(),from_branch_id:i(String(r["Cabang Asal"]||"").trim()),to_branch_id:i(String(r["Cabang Tujuan"]||"").trim()),status:String(r.Status||"").trim(),document_link:String(r.Dokumen||"").trim()})).filter(r=>r.tanggal&&r.employee_name&&r.from_branch_id&&r.to_branch_id),d=await S("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:l,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:ia},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Ee,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Ee,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}M();async function sa(t){t.innerHTML=`
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
  `;let e=document.getElementById("btnRefreshSync"),i=document.getElementById("queueStatusFilter");e.addEventListener("click",l),i.addEventListener("change",s),document.getElementById("btnRetryAll").addEventListener("click",()=>a("retry",{allFailed:!0})),document.getElementById("btnResetStuck").addEventListener("click",()=>a("reset-stuck",{})),document.getElementById("btnPauseSync").addEventListener("click",()=>a("pause",{})),document.getElementById("btnResumeSync").addEventListener("click",()=>a("resume",{})),document.getElementById("btnReconcile").addEventListener("click",()=>a("reconcile",{module:"ALL",repairMode:!0})),document.getElementById("btnRetrySelected").addEventListener("click",()=>{let c=Array.from(document.querySelectorAll(".chk-queue:checked")).map(u=>u.value);if(c.length===0)return alert("No items selected");a("retry",{ids:c})}),document.getElementById("chkAllQueue").addEventListener("change",c=>{document.querySelectorAll(".chk-queue").forEach(u=>u.checked=c.target.checked)});async function a(c,u){if(confirm(`Are you sure you want to execute action: ${c}?`)){showLoading();try{let o=await S(`/api/sync/actions/${c}`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(u)});o.ok?(alert(o.data?.message||"Success"),l()):G(o.error||"Action failed")}catch(o){G(o.message)}hideLoading()}}await l();async function l(){showLoading(),await Promise.all([r(),s(),d(),n()]),hideLoading()}async function d(){try{let c=await S("/api/sync/performance");if(!c.ok)return;let{webhook:u,google_api:o,d1:p,queue:m,throughput:b}=c.data;document.getElementById("latencyHistograms").innerHTML=`
        <table class="w-full text-left">
          <thead><tr class="text-gray-500 border-b"><th>Metric</th><th>P50</th><th>P95</th><th>P99</th><th>Max</th></tr></thead>
          <tbody class="divide-y">
            <tr><td class="py-1">Webhook</td><td>${u.P50}ms</td><td>${u.P95}ms</td><td>${u.P99}ms</td><td>${u.Max}ms</td></tr>
            <tr><td class="py-1">Google API</td><td>${o.P50}ms</td><td>${o.P95}ms</td><td>${o.P99}ms</td><td>${o.Max}ms</td></tr>
            <tr><td class="py-1">D1 Execute</td><td>${p.P50}ms</td><td>${p.P95}ms</td><td>${p.P99}ms</td><td>${p.Max}ms</td></tr>
            <tr><td class="py-1">Queue Delay</td><td>${m.P50}ms</td><td>${m.P95}ms</td><td>${m.P99}ms</td><td>${m.Max}ms</td></tr>
          </tbody>
        </table>
        <div class="mt-4 pt-3 border-t text-gray-600 flex justify-between">
          <span>Throughput: <b>${b.events_per_sec}</b> ev/sec</span>
          <span><b>${b.events_per_min}</b> ev/min</span>
        </div>
      `,document.getElementById("lblSnapshotStatus").innerHTML='Checked 10 modules. Status: <b>COMPLETED</b><br><span class="text-xs text-gray-500">Run today at 02:00</span>'}catch(c){console.error(c)}}async function r(){try{let c=await S("/api/sync/health");if(!c.ok)return G("Failed to fetch sync health");let{status:u,queue:o,circuit_breaker:p}=c.data,m=`
        <div class="bg-white p-4 rounded-lg shadow border-l-4 ${u==="HEALTHY"?"border-green-500":u==="WARNING"?"border-yellow-500":"border-red-500"}">
          <p class="text-sm text-gray-500">System Health</p>
          <p class="text-2xl font-bold ${u==="HEALTHY"?"text-green-600":u==="WARNING"?"text-yellow-600":"text-red-600"}">${u}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-blue-500">
          <p class="text-sm text-gray-500">Pending Queue</p>
          <p class="text-2xl font-bold text-gray-800">${o.pending||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-orange-500">
          <p class="text-sm text-gray-500">Failed / Retrying</p>
          <p class="text-2xl font-bold text-gray-800">${o.failed||0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-red-500">
          <p class="text-sm text-gray-500">Dead Letters</p>
          <p class="text-2xl font-bold text-gray-800">${o.dead_letter||0}</p>
        </div>
      `;document.getElementById("syncOverviewCards").innerHTML=m;let b=document.getElementById("cbStateBadge"),g=document.getElementById("cbStateDesc"),h=document.getElementById("cbStatusCard");h.className="bg-white rounded-lg shadow p-6 border-l-4",p==="CLOSED"?(b.className="px-3 py-1 rounded-full text-sm font-medium bg-green-100 text-green-800",b.textContent="CLOSED",g.textContent="Traffic is flowing normally to Google Sheets.",h.classList.add("border-green-500")):p==="OPEN"?(b.className="px-3 py-1 rounded-full text-sm font-medium bg-red-100 text-red-800",b.textContent="OPEN",g.textContent="Failing fast. Traffic to Google Sheets is paused due to repeated failures.",h.classList.add("border-red-500")):p==="HALF_OPEN"?(b.className="px-3 py-1 rounded-full text-sm font-medium bg-yellow-100 text-yellow-800",b.textContent="HALF-OPEN",g.textContent="Testing recovery. Permitting limited traffic to verify stability.",h.classList.add("border-yellow-500")):b.textContent=p||"UNKNOWN"}catch(c){console.error(c)}}async function s(){try{let c=document.getElementById("queueStatusFilter").value,u=await S("/api/sync/queue?limit=15"+(c?"&status="+c:""));if(!u.ok)return;let o=document.getElementById("queueTableBody"),p=u.data?.data||u.data||[];if(p.length===0){o.innerHTML='<tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">No events found</td></tr>';return}o.innerHTML=p.map(m=>`
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
      `).join("")}catch(c){console.error(c)}}async function n(){try{let c=await S("/api/sync/metrics");if(!c.ok)return;let u=document.getElementById("metricsTableBody"),o=c.data||[];if(o.length===0){u.innerHTML='<tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">No metrics found in last 24h</td></tr>';return}u.innerHTML=o.map(p=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2 font-medium text-gray-800">${p.module}</td>
          <td class="px-4 py-2 text-gray-600">${p.total_events}</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(p.avg_webhook_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(p.avg_d1_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(p.avg_queue_wait_ms||0)} ms</td>
        </tr>
      `).join("")}catch(c){console.error(c)}}}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let e=Math.floor(Number(t)-25569);return new Date(e*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let e=t.split(/[\/\-]/);return`${e[2]}-${e[1]}-${e[0]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let i=e.split("-");if(i.length===3&&i[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],l=parseInt(i[2],10),d=a[parseInt(i[1],10)-1];return`${l} ${d} ${i[0]}`}return e};function Q(t){return async e=>{if(!Ie()){ke("/login");return}return t(e)}}var He=null;function Ra(){He&&clearInterval(He);let t=()=>{let e=new Date,i=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),l=document.getElementById("header-clock-time"),d=document.getElementById("header-clock-date");l&&(l.textContent=i),d&&(d.textContent=a)};t(),He=setInterval(t,1e3)}async function Ka(){try{let t=await S("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},i=(a,l)=>{let d=document.getElementById(a);d&&(d.textContent=l>0?l:"",d.style.display=l>0?"inline-flex":"none")};i("badge-issues",e.issues?.current||0),i("badge-contracts",e.expiring30?.current||0),i("badge-oo1",e.one_on_one?.current||0),i("badge-schedule",e.schedule?.current||0),i("badge-supply",e.supply?.current||0)}catch{}}var $e=[];async function Ha(){try{let t=await S("/api/dashboard/notifications");if(!t.ok)return;$e=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=$e.length>0?"block":"none",e.textContent=$e.length)}catch{}}function qa(){if(!$e.length){le({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,i)=>i()});return}let t=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${$e.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;le({title:`Notifikasi (${$e.length})`,content:t,confirmText:"Tutup",onConfirm:(e,i)=>i()})}function la(){let t=be(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
  `;let i=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),l=document.getElementById("topbar-menu-btn"),d=document.getElementById("sidebar-close"),r=()=>{i.classList.add("open"),a.classList.add("show")},s=()=>{i.classList.remove("open"),a.classList.remove("show")};l?.addEventListener("click",r),d?.addEventListener("click",s),a?.addEventListener("click",s),document.querySelectorAll(".nav-item").forEach(c=>c.addEventListener("click",s));function n(){let c=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(p=>{let m=p.dataset.route;p.classList.toggle("active",c===m||m!=="/dashboard"&&c.startsWith(m))});let u=document.getElementById("topbar-title"),o=document.querySelector(".nav-item.active .nav-label");u&&o&&(u.textContent=o.textContent)}window.addEventListener("hashchange",n),n(),Ra(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await S("/api/auth/logout",{method:"POST"}),Pe(),He&&clearInterval(He),ke("/login")}),Ka(),Ha(),document.getElementById("btn-notif")?.addEventListener("click",c=>{c.preventDefault(),qa()})}async function ja(){j("/login",({main:e})=>It(e)),j("/dashboard",Q(({main:e})=>Tt(e))),j("/calendar",Q(({main:e})=>Vt(e))),j("/employees",Q(({main:e,params:i})=>Pt(e,i))),j("/contracts",Q(({main:e,params:i})=>Lt(e,i))),j("/sp",Q(({main:e})=>na(e))),j("/mutasi",Q(({main:e})=>ra(e))),j("/sync-dashboard",Q(({main:e})=>sa(e))),j("/timeline",Q(({main:e,params:i})=>Nt(e,i))),j("/issues",Q(({main:e,params:i})=>At(e,i))),j("/one-on-one",Q(({main:e,params:i})=>Ft(e,i))),j("/training",Q(({main:e})=>Mt(e))),j("/relievers",Q(({main:e,params:i})=>Rt(e,i))),j("/reports/inspection",Q(({main:e})=>Kt(e))),j("/reports/cleaning",Q(({main:e})=>Ht(e))),j("/reports/fogging",Q(({main:e})=>qt(e))),j("/reports/basecamp",Q(({main:e})=>jt(e))),j("/reports/supply",Q(({main:e})=>ht(e,"supply"))),j("/sop",Q(({main:e})=>Jt(e))),j("/checklist",Q(({main:e})=>Ut(e))),j("/forms",Q(({main:e})=>ht(e))),j("/users",Q(({main:e})=>Gt(e))),j("/branches",Q(({main:e})=>Qt(e))),j("/profile",Q(({main:e})=>Wt(e))),j("/settings/import",Q(({main:e})=>ta(e)));let t=Ie();if(!t&&window.location.hash!=="#/login"&&ke("/login"),t){let e=await S("/api/auth/me");e.ok?(Be(e.data.data),la()):(Pe(),ke("/login"))}window.addEventListener("fm:login",()=>{la(),ke("/dashboard")}),ft()}ja();
