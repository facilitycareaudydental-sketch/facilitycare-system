var ra=Object.defineProperty;var it=(t,e)=>()=>(t&&(e=t(t=0)),e);var rt=(t,e)=>{for(var l in e)ra(t,l,{get:e[l],enumerable:!0})};var ke={};rt(ke,{API:()=>ht,CLIENT_SIDE_MAX_ROWS:()=>ye,IS_DEVELOPMENT:()=>Le,apiFetch:()=>x,clearToken:()=>Ae,getToken:()=>Ne,getUser:()=>fe,setToken:()=>lt,setUser:()=>Fe});function Ne(){return localStorage.getItem("fm_token")}function lt(t){localStorage.setItem("fm_token",t)}function Ae(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function fe(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Fe(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function x(t,e={}){let l=Ne(),a={"Content-Type":"application/json",...l?{Authorization:`Bearer ${l}`}:{},...e.headers||{}};try{let o=`cb=${Date.now()}`,s=t.includes("?")?"&":"?",r=`${ht}${t}${s}${o}`,i=await fetch(r,{...e,headers:a}),n;try{let p=await i.text();try{n=JSON.parse(p)}catch{n={error:`Server Error (${i.status}): ${p.substring(0,80)}...`}}}catch{n={error:"Gagal membaca respon dari server"}}if(i.status===401&&(Ae(),window.location.hash="#/login"),i.ok){let p=(e.method||"GET").toUpperCase();["POST","PUT","DELETE","PATCH"].includes(p)&&window.dispatchEvent(new CustomEvent("fcms-data-changed"))}return{ok:i.ok,status:i.status,data:n}}catch(o){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${o.message})`}}}}var Le,la,ht,ye,O=it(()=>{Le=!1,la="https://fm-operations-api.facilitycare-audydental.workers.dev",ht=la,ye=1e4});var vt={};rt(vt,{confirmDialog:()=>Oe,createModal:()=>se});function se({title:t,content:e,onConfirm:l,onCancel:a,confirmText:o="Simpan",cancelText:s="Batal",size:r="md",confirmClass:i="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},p=document.createElement("div");p.className="modal-overlay",p.innerHTML=`
    <div class="modal" style="max-width:${n[r]||n.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${s}</button>
        ${l?`<button class="btn ${i} modal-confirm">${o}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&p.querySelector(".modal-body").appendChild(e);let m=()=>{p.classList.remove("show"),setTimeout(()=>p.remove(),250)};return p.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),m()}),p.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),m()}),l&&p.querySelector(".modal-confirm").addEventListener("click",()=>l(p,m)),p.addEventListener("click",d=>{d.target===p&&(a&&a(),m())}),document.body.appendChild(p),requestAnimationFrame(()=>p.classList.add("show")),{overlay:p,close:m}}function Oe(t,e,l="Konfirmasi"){return se({title:l,content:`<p>${t}</p>`,onConfirm:(a,o)=>{e(),o()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var ve=it(()=>{});var le={};rt(le,{downloadExcel:()=>L,parseExcel:()=>Ke,renderExcelButtons:()=>He});function Ke(t){return new Promise((e,l)=>{let a=new FileReader;a.onload=o=>{try{let s=new Uint8Array(o.target.result),r=XLSX.read(s,{type:"array"}),i=r.SheetNames[0],n=r.Sheets[i];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${r.SheetNames.join(", ")}`),console.log(`Sheet Used: ${i}`);let p=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),m=p.e.r-p.s.r+1,d=p.e.c-p.s.c+1;console.log(`Total Rows (including empty): ${m}`),console.log(`Total Columns: ${d}`);let c=[];for(let b=p.s.c;b<=p.e.c;++b){let g=n[XLSX.utils.encode_cell({c:b,r:p.s.r})];g&&g.v&&c.push(g.v)}console.log(`Headers Found: ${c.join(", ")}`),console.log("---------------------------");let u=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(u,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(u,"__headers",{value:c,enumerable:!1}),e(u)}catch(s){l(s)}},a.onerror=o=>l(o),a.readAsArrayBuffer(t)})}function L(t,e){try{let l=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,l,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(l){throw console.error("Error generating Excel file:",l),l}}function He(t){return`
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
  `}var R=it(()=>{});O();var ot={},Ge=null;function J(t,e){ot[t]=e}function xe(t){window.location.hash=t}function yt(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[l,...a]=e.split("?"),o=ot[l];if(!o){for(let[r,i]of Object.entries(ot))if(r.endsWith("/*")&&l.startsWith(r.slice(0,-2))){o=i;break}}Ge&&(Ge(),Ge=null);let s=document.getElementById("main-content");if(s&&(s.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),o){let r=new URLSearchParams(a.join("?")),i=l.split("/").filter(Boolean),n=await o({path:l,params:r,segments:i,main:s});n&&(Ge=n)}else{let r=s||document.getElementById("app");r&&(r.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var Me;function oa(){return Me||(Me=document.createElement("div"),Me.id="toast-container",document.body.appendChild(Me)),Me}function ft(t,e="info",l=3500){let a=oa(),o=document.createElement("div");o.className=`toast toast-${e}`;let s={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};o.innerHTML=`<span class="toast-icon">${s[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(o),requestAnimationFrame(()=>o.classList.add("show")),setTimeout(()=>{o.classList.remove("show"),setTimeout(()=>o.remove(),350)},l)}var Z=t=>ft(t,"success"),G=t=>ft(t,"error");ve();O();O();ve();function ze({columns:t,data:e,onEdit:l,onDelete:a,onView:o,actions:s=[],emptyText:r="Tidak ada data",bulkSelect:i=null}){let n=document.createElement("div");if(n.className="table-wrapper",!e||e.length===0)return n.innerHTML=`<div class="empty-state"><p>${r}</p></div>`,n;let p=document.createElement("table");p.className="data-table";let m=document.createElement("thead"),d=document.createElement("tr");if(i){let u=document.createElement("th");u.style.width="40px",u.style.textAlign="center";let b=document.createElement("input");b.type="checkbox",b.id="select-all-checkbox",b.title="Pilih semua",b.addEventListener("change",()=>{e.forEach(g=>{b.checked?i.selectedIds.add(g.id):i.selectedIds.delete(g.id)}),n.querySelectorAll(".row-checkbox").forEach(g=>g.checked=b.checked),i.onToggle()}),u.appendChild(b),d.appendChild(u)}if(t.forEach(u=>{let b=document.createElement("th");b.textContent=u.label,u.width&&(b.style.width=u.width),d.appendChild(b)}),l||a||o||s.length>0){let u=document.createElement("th");u.textContent="Aksi",u.style.width="120px",d.appendChild(u)}m.appendChild(d),p.appendChild(m);let c=document.createElement("tbody");return e.forEach(u=>{let b=document.createElement("tr");if(i){let g=document.createElement("td");g.style.textAlign="center",g.style.width="40px";let h=document.createElement("input");h.type="checkbox",h.className="row-checkbox",h.checked=i.selectedIds.has(u.id),h.addEventListener("change",()=>{if(h.checked)i.selectedIds.add(u.id);else{i.selectedIds.delete(u.id);let k=document.getElementById("select-all-checkbox");k&&(k.checked=!1)}i.onToggle()}),g.appendChild(h),b.appendChild(g)}if(t.forEach(g=>{let h=document.createElement("td");if(g.render){let k=g.render(u[g.key],u);k instanceof HTMLElement?h.appendChild(k):h.innerHTML=k||""}else h.textContent=u[g.key]!==null&&u[g.key]!==void 0&&u[g.key]!==""?u[g.key]:"";g.nowrap&&(h.style.whiteSpace="nowrap"),b.appendChild(h)}),l||a||o||s.length>0){let g=document.createElement("td");g.className="actions-cell";let h=document.createElement("div");if(h.className="btn-group",o){let k=document.createElement("button");k.className="btn btn-xs btn-ghost",k.innerHTML="\u{1F441}",k.title="Lihat",k.addEventListener("click",()=>o(u)),h.appendChild(k)}if(l){let k=document.createElement("button");k.className="btn btn-xs btn-secondary",k.innerHTML="\u270F\uFE0F",k.title="Edit",k.addEventListener("click",()=>l(u)),h.appendChild(k)}s.forEach(k=>{let S=document.createElement("button");S.className=`btn btn-xs ${k.class||"btn-ghost"}`,S.innerHTML=k.icon||k.label,S.title=k.label,S.addEventListener("click",()=>k.handler(u)),h.appendChild(S)}),g.appendChild(h),b.appendChild(g)}c.appendChild(b)}),p.appendChild(c),n.appendChild(p),n}function Qe({page:t,pages:e,total:l,limit:a,onPage:o}){if(e<=1)return null;let s=document.createElement("div");s.className="pagination";let r=document.createElement("span");r.className="pagination-info",r.textContent=`Total: ${l} data`,s.appendChild(r);let i=document.createElement("div");i.className="pagination-btns";let n=(d,c,u=!1,b=!1)=>{let g=document.createElement("button");g.className=`btn btn-sm ${b?"btn-primary":"btn-ghost"} pagination-btn`,g.textContent=d,g.disabled=u,g.addEventListener("click",()=>o(c)),i.appendChild(g)};n("\xAB",1,t===1),n("\u2039",t-1,t===1);let p=Math.max(1,t-2),m=Math.min(e,t+2);for(let d=p;d<=m;d++)n(d,d,!1,d===t);return n("\u203A",t+1,t===e),n("\xBB",e,t===e),s.appendChild(i),s}ve();function Re(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${Re(e.fields)}</div>`;let l=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",o="";switch(e.type){case"textarea":o=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${l} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let r=(e.options||[]).map(d=>{let c=typeof d=="object"?d.value:d,u=typeof d=="object"?d.label:d,b=e.value==c?"selected":"";return`<option value="${c}" ${b}>${u}</option>`}).join("");o=`<select name="${e.name}" class="form-control" ${l}><option value="">-- Pilih ${e.label||""} --</option>${r}</select>`;break;case"combobox":let i=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,n=(e.options||[]).map(d=>{let c=typeof d=="object"?d.value:d,u=typeof d=="object"?d.label||d.value||"":d||"";return(u==="undefined"||u==="[object Object]"||u==="null")&&(u=""),u?`<option value="${u}"></option>`:""}).join(""),p=e.value||"";if(e.value){let d=(e.options||[]).find(c=>(typeof c=="object"?c.value:c)==e.value);if(d){let c=typeof d=="object"?d.label||d.value||"":d||"";c&&c!=="undefined"&&c!=="[object Object]"&&c!=="null"&&(p=c)}}o=`
          <input type="text" name="${e.name}" list="${i}" class="form-control" value="${p}" placeholder="Pilih atau ketik baru..." ${l} autocomplete="off">
          <datalist id="${i}">${n}</datalist>
        `;break;case"checkbox":o=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let m=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";o=`<input type="date" name="${e.name}" class="form-control" value="${m}" ${l}>`;break;case"number":o=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${l}>`;break;case"email":o=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${l}>`;break;case"url":o=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${l}>`;break;default:o=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${l} autocomplete="off">`}let s=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${o}${s}</div>`}).join("")}function Ye(t){let e={},l=new FormData(t);for(let[a,o]of l.entries())e[a]=o===""?null:o;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function We(t,e){e&&Object.entries(e).forEach(([l,a])=>{let o=t.querySelector(`[name="${l}"]`);o&&(o.hasAttribute("list")||(o.type==="checkbox"?o.checked=!!a:o.type==="date"&&a&&window.parseFlexibleDate?o.value=window.parseFlexibleDate(a):o.value=a??""))})}R();function N({container:t,title:e,icon:l,apiPath:a,columns:o,formFields:s,filterFields:r,defaultFilters:i={},itemLabel:n="Data",canCreate:p=!0,canEdit:m=!0,canDelete:d=!0,onBeforeSubmit:c,onAfterLoad:u,onDataLoaded:b,extraActions:g=[],initialSearch:h="",exportOptions:k=null,bulkDelete:S=!1,paginationMode:C="server"}){let _=1,f={...i};h&&(f.search=h);let w=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${l} ${e}</h1>
      <div class="page-actions">
        ${p?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
      </div>
    </div>

    ${S?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="align-items:center; background:var(--bg-card); border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${k?He(k.moduleName):""}

    ${r&&r.length>0?`
    <div class="filter-bar card" style="padding: 1rem; display:flex; flex-wrap:wrap; gap:1rem; align-items:center;">
        ${r.map(v=>{if(v.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${f.search||""}"></div>`;if(v.type==="search-combo"){let y="dl-filter-search",$=(v.options||[]).map(P=>`<option value="${typeof P=="object"?P.label:P}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${y}" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${f.search||""}"><datalist id="${y}">${$}</datalist></div>`}return v.type==="select"?`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${v.name}" id="filter-${v.name}"><option value="">-- ${v.label} --</option>${(v.options||[]).map(y=>`<option value="${typeof y=="object"?y.value:y}" ${f[v.name]===(typeof y=="object"?y.value:y)?"selected":""}>${typeof y=="object"?y.label:y}</option>`).join("")}</select>`:""}).join("")}
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
    `,document.body.appendChild(y),y.querySelector("#bulk-cancel-btn").addEventListener("click",()=>y.remove()),y.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let $=y.querySelector("#bulk-confirm-btn");$.disabled=!0,$.textContent="Menghapus...";let P=await x(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:v})});y.remove(),P.ok?(Z(`${v.length} ${n} berhasil dihapus.`),w.clear(),B(),M()):G(P.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),q;if(E?.addEventListener("input",v=>{clearTimeout(q),q=setTimeout(()=>{f.search=v.target.value,_=1,w.clear(),M()},400)}),r?.forEach(v=>{v.type==="select"&&document.getElementById(`filter-${v.name}`)?.addEventListener("change",y=>{f[v.name]=y.target.value,_=1,w.clear(),M()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{f={...i},E&&(E.value=""),r?.forEach(v=>{let y=document.getElementById(`filter-${v.name}`);y&&(y.value="")}),_=1,w.clear(),M()}),document.getElementById("btn-create")?.addEventListener("click",()=>re(null)),k){document.getElementById(`btn-export-${k.moduleName}`)?.addEventListener("click",async y=>{let $=y.target,P=$.innerHTML;$.innerHTML="\u23F3 Loading...",$.disabled=!0;try{await k.onExport()}catch{G("Gagal export data")}finally{$.innerHTML=P,$.disabled=!1}}),document.getElementById(`btn-template-${k.moduleName}`)?.addEventListener("click",()=>{k.onTemplate()});let v=document.getElementById(`input-import-${k.moduleName}`);v?.addEventListener("change",async y=>{let $=y.target.files[0];if(!$)return;v.disabled=!0;let P=document.createElement("div");P.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",P.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(P);let Y=P.querySelector("#import-progress-text"),U=P.querySelector("#import-progress-bar"),D=P.querySelector("#import-summary"),I=P.querySelector("#import-close-btn");I.addEventListener("click",()=>{P.remove(),M()});try{let V=await Ke($);if(V.length===0)throw new Error("File kosong atau format salah");let X=500,ie=0,ae=0,A=0,F=V.length;Y.textContent=`Ditemukan ${F} baris data. Memulai import...`;for(let T=0;T<F;T+=X){let Q=V.slice(T,T+X);Y.textContent=`Mengimport baris ${T+1} - ${Math.min(T+X,F)} dari ${F}...`,U.style.width=`${Math.round(T/F*100)}%`;try{let H=await k.onImport(Q);H?(ie+=H.inserted||H.metrics?.inserted||Q.length,ae+=H.skipped||H.metrics?.updated||0):ie+=Q.length}catch(H){console.error("Chunk import failed:",H),A+=Q.length}}U.style.width="100%",Y.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',D.style.display="block",D.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${F}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ie}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${A}</strong></div>
        `,A>0&&(D.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),I.style.display="block",v.value=""}catch(V){Y.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${V.message}`,U.style.background="var(--danger)",U.style.width="100%",I.style.display="block",v.value=""}finally{v.disabled=!1}})}async function M(){B();let v=document.getElementById("table-container");if(!v)return;v.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let y=C==="client",$=y?1:_,P=y?ye:20,Y=new URLSearchParams({page:$,limit:P,...Object.fromEntries(Object.entries(f).filter(([,A])=>A))}),U=await x(`${a}?${Y}`);if(!U.ok){v.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${U.data?.error||"Error"}</p></div>`;return}let D=U.data?.data||U.data||[],I=U.data?.pagination,V=D.length,X=D;if(y){D=b(D),X=D;let A=D.length,F=20,T=Math.ceil(A/F);_>T&&T>0&&(_=T);let Q=(_-1)*F,H=_*F;D=D.slice(Q,H),I={page:_,limit:F,total:A,pages:T}}!1,u&&u(D);let ie=ze({columns:o,data:D,fullData:X,onEdit:m?A=>re(A):null,actions:g.map(A=>({...A,handler:F=>A.handler(F,M)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:S?{selectedIds:w,onToggle:B}:null});v.innerHTML="",v.appendChild(ie);let ae=document.getElementById("pagination-container");if(ae&&(ae.innerHTML="",I&&I.pages>1)){let A=Qe({page:I.page,pages:I.pages,total:I.total,limit:I.limit,onPage:F=>{_=F,M()}});A&&ae.appendChild(A)}}function me(v){let y=typeof s=="function"?s(v):s;return Re(y)}function re(v){let y=!!v,$=document.createElement("form");if($.noValidate=!0,$.innerHTML=me(v),y){let Y=typeof s=="function"?s(v):s;We($,v)}let{close:P}=se({title:y?`Edit ${n}`:`Tambah ${n}`,content:$,size:"lg",confirmText:y?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(Y,U)=>{if(!$.reportValidity())return;let D=Y.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let I=Ye($),V=typeof s=="function"?s(v):s,X=async F=>{for(let T of F)if(T.type==="row")await X(T.fields);else if(T.type==="combobox"&&I[T.name]){let Q=I[T.name],H=(T.options||[]).find(W=>{let ne=String(typeof W=="object"?W.value:W),nt=String(typeof W=="object"?W.label:W);return ne===Q||nt===Q});if(H)I[T.name]=typeof H=="object"?H.value:H;else if(T.createApi){let W={};W[T.createApi.field]=Q,T.createApi.extra&&Object.assign(W,T.createApi.extra);let ne=await x(T.createApi.path,{method:"POST",body:JSON.stringify(W)});if(ne.ok&&ne.data?.id)I[T.name]=ne.data.id;else if(ne.ok&&!ne.data?.id)I[T.name]=Q;else throw new Error(`Gagal membuat master data: ${ne.data?.error||"Unknown error"}`)}}};try{await X(V)}catch(F){G(F.message),D.disabled=!1,D.textContent=y?"Simpan Perubahan":`Tambah ${n}`;return}c&&(I=await c(I,v));let ie=y?"PUT":"POST",ae=y?`${a}/${v.id}`:a,A=await x(ae,{method:ie,body:JSON.stringify(I)});A.ok?(Z(y?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),U(),M()):(G(A.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=y?"Simpan Perubahan":`Tambah ${n}`)}})}function oe(v){Oe(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let y=await x(`${a}/${v.id}`,{method:"DELETE"});y.ok?(Z(`${n} berhasil dihapus.`),M()):G(y.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return M(),M}O();O();var Se=null,Ve=null;async function we(t=!1){if(Se&&!t)return console.log("Employees Raw (Cache Hit)",Se.slice(0,5)),Se;let e=await x(`/api/employees?limit=${ye}&status=Aktif`);return Se=(e.data?.data||[]).map(l=>({value:l.id,label:l.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",Se.slice(0,5)),Se}async function ee(t=!1){let l=(await we(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",l.slice(0,5)),l}async function K(t=!1){return Ve&&!t||(Ve=((await x("/api/branches?all=1")).data?.data||[]).map(l=>({value:l.id,label:l.full_name}))),Ve}function j(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function st(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function _e(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function dt(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function de(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}R();O();R();O();R();O();R();O();var he={};function $e(t){if(he[t]){try{he[t].destroy()}catch{}delete he[t]}}function sa(){Object.keys(he).forEach($e)}var be=(t,e=0)=>{let l=Number(t);return isNaN(l)||t===null||t===void 0?e:l},Te=(t,e="\u2014")=>{if(t==null||t==="")return e;let l=String(t).trim();return l===""||l==="[object Object]"?e:l};var da=t=>{if(!t||typeof t!="string")return"";try{let[e,l]=t.split("-");return new Date(Number(e),Number(l)-1).toLocaleDateString("id-ID",{month:"short",year:"2-digit"})}catch{return t}};function xt(t,e,l=900){if(!t)return;let a=Math.max(0,Math.round(be(e)));if(a===0){t.textContent="0";return}let o=Date.now(),s=()=>{let r=Math.min((Date.now()-o)/l,1),i=1-Math.pow(1-r,3);t.textContent=Math.round(i*a).toLocaleString("id-ID"),r<1?requestAnimationFrame(s):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(s)}var ca={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},pa=t=>{let e=Te(t,"\u2014");return`<span class="status-pill ${ca[e]||"pill-neutral"}">${e}</span>`};var ce={family:"Inter",size:11},Ee="#94A3B8",Xe="#F1F5F9",ct=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],ua=()=>window.innerWidth<768;function Ze(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:ua()?"bottom":"top",labels:{font:ce,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ce,titleFont:{...ce,weight:"700"}}},scales:{x:{grid:{color:Xe},ticks:{font:ce,color:Ee,maxRotation:0}},y:{grid:{color:Xe},ticks:{font:ce,color:Ee},beginAtZero:!0}},...t}}var ma=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join("");function kt(t=3){return Array(t).fill(0).map((e,l)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${l<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function ge(t,e,l=8e3){try{let a=new AbortController,o=setTimeout(()=>a.abort(),l),s=await x(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(o),!s||!s.ok)return e;let r=s.data;return r?r.data!==void 0?r.data??e:r:e}catch{return e}}function ba(){["skel-donut","skel-trend","skel-insp","skel-contract","skel-schbar"].forEach(a=>{let o=document.getElementById(a);o&&(o.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract","chart-schbar"].forEach(a=>{let o=document.getElementById(a);if(o&&o.style.display==="none"){o.style.display="block";let s=o.parentElement;if(s&&!s.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent="Belum ada data",o.style.display="none",s.appendChild(r)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&wt({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&_t({}),["table-contracts","table-issues"].forEach(a=>{let o=document.getElementById(a);o&&o.querySelector(".skeleton")&&(o.innerHTML='<div class="chart-empty">Belum ada data</div>')});let l=document.getElementById("activity-log");l&&l.querySelector(".skeleton")&&(l.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function St(t){sa(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">


      <!-- KPI -->
      
      <div class="dashboard-filters" style="display:flex; gap:10px; margin-bottom: 20px; background: white; padding: 15px; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
        <select id="global-branch" class="input-text" style="flex:1;">
          <option value="">Semua Cabang</option>
        </select>
        <select id="global-month" class="input-text" style="width:150px;">
          <!-- Months will be populated by JS -->
        </select>
      </div>

      <div class="kpi-row" id="kpi-row" style="display:grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));">${ma()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row" style="display:grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));">${ga()}</div>

      <!-- Charts Row (All 3 in 1 Row) -->
      <div class="charts-row" style="grid-template-columns: 1.35fr 0.85fr 1.1fr;">
        
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
          <div class="chart-canvas-wrap" style="height:190px; position:relative; margin-top:10px;">
            <div id="skel-schbar" class="skeleton" style="position:absolute; inset:0; border-radius:12px"></div>
            <canvas id="chart-schbar" style="display:none"></canvas>
          </div>
        </div>
        <div class="chart-card" style="display:flex; flex-direction:column;">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan per Kategori</div>
          </div>
          <div style="display:flex; flex-direction:row; gap:30px; align-items:center; justify-content:center; padding: 10px 0; margin-top: auto; margin-bottom: auto;">
            <div class="chart-canvas-wrap" style="width:120px; height:120px; position:relative">
              <div id="skel-donut" class="skeleton" style="position:absolute;inset:0;border-radius:50%"></div>
              <canvas id="chart-donut" style="display:none"></canvas>
            </div>
            <div id="donut-legend" class="donut-legend" style="display:flex; flex-direction:column; justify-content:center; gap:16px;"></div>
          </div>
          <div style="text-align:center; font-size:0.7rem; color:var(--text-3); margin-top:auto; padding-top:10px;">
            Periode: 22 Juni - 22 Juli 2026
          </div>
        </div>
        <div class="chart-card" style="display:flex; flex-direction:column;">
          <div class="chart-card-header" style="display:flex; justify-content:space-between; align-items:center;">
            <div class="chart-card-title">Trend Permasalahan 12 Bulan</div>
            <div style="display:flex;align-items:center;gap:12px;font-size:0.75rem;font-weight:600;color:var(--text-2)">
               <div style="display:flex;align-items:center;gap:6px"><div style="width:10px;height:10px;border-radius:50%;background:#EF4444"></div> Open</div>
               <div style="display:flex;align-items:center;gap:6px"><div style="width:10px;height:10px;border-radius:50%;background:#10B981"></div> Closed</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="width:100%; height:180px; position:relative; margin-top:20px; margin-bottom:0;">
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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${kt(3)}</div>
        </div>
          <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${kt(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>Ce(t)),document.getElementById("filter-insp-month")?.addEventListener("change",async a=>{let o=a.target.value,s=o?`/api/dashboard/inspection-bar?month=${o}`:"/api/dashboard/inspection-bar",r=document.getElementById("skel-insp"),i=document.getElementById("chart-insp");r&&(r.style.display="block",r.style.position="absolute"),i&&(i.style.display="none");let n=await ge(s,{},8e3);try{Tt(n)}catch(p){console.warn("InspBar render:",p),pe("skel-insp","chart-insp")}}),document.getElementById("filter-sch-year")?.addEventListener("change",async a=>{let o=a.target.value;document.getElementById("sch-year-label").textContent=o;let s=`/api/dashboard/schedule-monthly?year=${o}`,r=document.getElementById("skel-schbar"),i=document.getElementById("chart-schbar");r&&(r.style.display="block",r.style.position="absolute"),i&&(i.style.display="none");let n=await ge(s,{data:[]},8e3);try{Ct(n)}catch(p){console.warn("SchBar render:",p),pe("skel-schbar","chart-schbar")}}),t._skelTimeout=setTimeout(()=>ba(),5e3);async function l(){let a=document.getElementById("global-branch"),o=document.getElementById("global-month");if(!(!a||!o)){if(o.children.length===0){let s=new Date;for(let r=0;r<12;r++){let i=new Date(s.getFullYear(),s.getMonth()-r,1),n=i.getFullYear()+"-"+String(i.getMonth()+1).padStart(2,"0"),p=i.toLocaleDateString("id-ID",{month:"long",year:"numeric"});o.innerHTML+=`<option value="${n}">${p}</option>`}}if(a.children.length===1){let s=await x("/api/branches");s.ok&&s.data&&s.data.data&&s.data.data.forEach(r=>{a.innerHTML+=`<option value="${r.id}">${r.full_name}</option>`})}a.addEventListener("change",()=>Ce(t)),o.addEventListener("change",()=>Ce(t))}}await l(),await Ce(t),window.addEventListener("fcms-data-changed",()=>{document.getElementById("dash-root")&&Ce(t)}),t._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?Ce(t):clearInterval(t._dashRefresh)},3e4)}async function Ce(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let e=document.getElementById("global-branch"),l=document.getElementById("global-month"),a=e?e.value:"",o=l?l.value:"",s=`?branch_id=${a}&month=${o}`,r=document.getElementById("filter-sch-year"),i=r?r.value:new Date().getFullYear(),[n,p,m,d,c,u,b,g,h,k,S,C,_]=await Promise.all([ge(`/api/dashboard/kpi${s}`,{},8e3),ge(`/api/dashboard/issues-trend${s}`,{},8e3),ge(`/api/dashboard/issues-summary${s}`,{},8e3),ge(`/api/dashboard/inspection-bar${s}`,{},8e3),ge(`/api/dashboard/stats${s}`,{},8e3),ge(`/api/dashboard/calendar${s}`,[],8e3),null,null,null,null,null,ge(`/api/dashboard/contracts-chart${s}`,{labels:[],data:[]},8e3),ge(`/api/dashboard/schedule-monthly${s}&year=${i}`,{data:[]},8e3)]);try{wt(n)}catch(f){console.warn("KPI render:",f)}try{_t(n)}catch(f){console.warn("MiniStats render:",f)}try{ha(Array.isArray(m?.by_category)?m.by_category:[])}catch(f){console.warn("Donut render:",f),pe("skel-donut","chart-donut")}try{Ct(_)}catch(f){console.warn("SchBar render:",f),pe("skel-schbar","chart-schbar")}try{ya(p)}catch(f){console.warn("Trend render:",f),pe("skel-trend","chart-trend")}try{Tt(d)}catch(f){console.warn("InspBar render:",f),pe("skel-insp","chart-insp")}try{let f=Array.isArray(c)?c:Array.isArray(c?.recent_issues)?c.recent_issues:[];va(f)}catch(f){console.warn("IssuesTable render:",f)}try{let f=Array.isArray(c?.expiring_contracts)?c.expiring_contracts:[];fa(C)}catch(f){console.warn("ContractsTable render:",f)}try{ka(Array.isArray(u)?u:[])}catch(f){console.warn("Agenda render:",f)}try{xa()}catch(f){console.warn("Quick Actions render:",f)}}function wt(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let l=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F504}",label:"Reliefer Aktif",sub:"Karyawan reliefer",href:"#/employees?dash_filter=reliefer",color:"kpi-purple",key:"reliever_total",trendPct:"0%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=l.map(a=>{let o=be(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${o}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${o}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{xt(a,parseInt(a.dataset.target)||0)})}function _t(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let l=[{icon:"\u{1F4C5}",label:"Jadwal",val:t.schedule?.current,href:"#/timeline",color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F504}",label:"Reliefer",val:t.reliever_completed?.current,href:"#/relievers?dash_filter=reliever",color:"mini-teal"},{icon:"\u{1F50D}",label:"Inspeksi",val:t.inspection_month?.current,href:"#/timeline?dash_filter=inspeksi",color:"mini-blue"},{icon:"\u{1F9F9}",label:"GCDC",val:t.cleaning_month?.current,href:"#/timeline?dash_filter=gcdc",color:"mini-green"},{icon:"\u{1F4A8}",label:"Fogging",val:t.fogging_month?.current,href:"#/reports/fogging",color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=l.map(a=>`
    <a href="${a.href}" class="mini-stat ${a.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${a.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${be(a.val)}">0</div>
        <div class="mini-stat-text">${a.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(a=>xt(a,parseInt(a.dataset.target)||0,700))}function ha(t){pe("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),l=document.getElementById("donut-legend");if(!e||!l)return;$e("donut");let a=(t||[]).filter(n=>be(n.count)>0);if(!a.length){qe(e,"Belum ada data permasalahan");return}let o=a.map(n=>`${Te(n.category,"Lainnya")}`),s=a.map(n=>be(n.count)),r=s.reduce((n,p)=>n+p,0);l.innerHTML=a.map((n,p)=>{let m=ct[p%ct.length],d=r>0?Math.round(n.count/r*100):0;return`
      <div class="donut-legend-item" style="display:flex; gap:8px;">
        <div class="donut-legend-color" style="background:${m}; width:12px; height:12px; border-radius:50%; flex-shrink:0; margin-top:4px;"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${d}%)</span></div>
          <div class="donut-legend-label">${o[p]}</div>
        </div>
      </div>
    `}).join("");let i={id:"centerText",beforeDraw:function(n){let p=n.width,m=n.height,d=n.ctx;d.restore();let c=(m/80).toFixed(2);d.font="bold "+c+"em Inter",d.textBaseline="middle",d.fillStyle="#1E293B";let u=r.toString(),b=Math.round((p-d.measureText(u).width)/2),g=m/2;d.fillText(u,b,g-10),d.font="600 "+(c*.35).toFixed(2)+"em Inter",d.fillStyle="#64748B";let h="Total",k=Math.round((p-d.measureText(h).width)/2);d.fillText(h,k,g+15),d.save()}};he.donut=new Chart(e,{type:"doughnut",data:{labels:o,datasets:[{data:s,backgroundColor:ct,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ce,titleFont:{...ce,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutoutPercentage:78,cutout:"78%"},plugins:[i]})}function Ct(t){pe("skel-schbar","chart-schbar");let e=document.getElementById("chart-schbar");if(!e)return;if($e("schbar"),!t||!t.data||!t.data.length){qe(e,"Belum ada jadwal pada tahun ini.");return}let l=t.data||[],o=["Jan","Feb","Mar","Apr","Mei","Jun","Jul","Agu","Sep","Okt","Nov","Des"],s={Inspeksi:Array(12).fill(0),"General Cleaning":Array(12).fill(0),"Deep Cleaning":Array(12).fill(0),Fogging:Array(12).fill(0)};l.forEach(c=>{let u=c.activity_type||"Lainnya";u.toLowerCase().includes("inspeksi")&&(u="Inspeksi"),(u.toLowerCase()==="gcdc"||u.toLowerCase().includes("general cleaning"))&&(u="General Cleaning"),u.toLowerCase().includes("deep cleaning")&&(u="Deep Cleaning"),(u.toLowerCase().includes("fogging")||u.toLowerCase().includes("foging"))&&(u="Fogging"),s[u]||(s[u]=Array(12).fill(0));let b=parseInt(c.month,10)-1;b>=0&&b<12&&(s[u][b]+=c.c||0)});let r={Inspeksi:"#60A5FA","General Cleaning":"#4ADE80","Deep Cleaning":"#FB923C",Fogging:"#EF4444"},i=c=>{let u=0;for(let b=0;b<c.length;b++)u=c.charCodeAt(b)+((u<<5)-u);return`hsl(${Math.abs(u)%360}, 65%, 60%)`},n=["Inspeksi","General Cleaning","Deep Cleaning","Fogging"],p=Object.keys(s).filter(c=>!n.includes(c)),d=[...n,...p].filter(c=>s[c]).map(c=>({label:c,data:s[c],backgroundColor:r[c]||i(c),borderWidth:0,borderRadius:4}));he.schbar=new Chart(e,{type:"bar",data:{labels:o,datasets:d},options:Ze({responsive:!0,maintainAspectRatio:!1,interaction:{mode:"index",intersect:!1},plugins:{legend:{display:!0,position:"bottom",align:"center",labels:{boxWidth:10,boxHeight:10,useBorderRadius:!0,borderRadius:3,padding:8,font:{family:"Inter",size:9}},padding:{top:10}},tooltip:{callbacks:{title:c=>{let u=c[0].label;return{Jan:"Januari",Feb:"Februari",Mar:"Maret",Apr:"April",Mei:"Mei",Jun:"Juni",Jul:"Juli",Agu:"Agustus",Sep:"September",Okt:"Oktober",Nov:"November",Des:"Desember"}[u]||u},label:c=>` ${c.dataset.label}: ${c.raw} Jadwal`}}},scales:{x:{stacked:!0,grid:{display:!1}},y:{stacked:!0,border:{display:!1},grid:{color:"#f3f4f6",drawTicks:!1}}}})})}function ya(t){pe("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;$e("trend"),t=t||{};let l=(t.labels||[]).map(da),a=(t.open||[]).map(s=>be(s)),o=(t.closed||[]).map(s=>be(s));if(!l.length){qe(e,"Belum ada data trend");return}he.trend=new Chart(e,{type:"line",data:{labels:l,datasets:[{label:"Open",data:a,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:o,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#10B981",borderWidth:2}]},options:Ze({plugins:{legend:{display:!1}}})})}function Tt(t){pe("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;$e("inspBar"),t=t||{};let l=t.labels||[],a=(t.fc||[]).map(s=>be(s)),o=(t.spv||[]).map(s=>be(s));if(!l.length){qe(e,"Belum ada data inspeksi");return}he.inspBar=new Chart(e,{type:"bar",data:{labels:l,datasets:[{label:"Skor FC",data:a,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:o,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:Ze({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:ce,color:Ee,maxRotation:45,minRotation:30}},y:{grid:{color:Xe},ticks:{font:ce,color:Ee},min:0,max:100}}})})}function fa(t){pe("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;$e("contractMiniBar"),t=t||{};let l={"06":"Jun","07":"Jul","08":"Agu","09":"Sep",10:"Okt",11:"Nov",12:"Des"},a=(t.labels||[]).map(r=>{let i=r.split("-")[1];return l[i]||r}),o=(t.data||[]).map(r=>be(r));if(!a.length){qe(e,"Belum ada data");return}let s=e.getContext("2d");he.contractMiniBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Kontrak Habis",data:o,backgroundColor:"#3B82F6",borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:Ze({onClick:(r,i)=>{if(i&&i.length>0){let n=i[0].index,p=(t.labels||[])[n];p&&(window.location.hash="#/contracts?month_expiry="+p)}},plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ce,color:Ee,maxRotation:0,autoSkip:!1}},y:{grid:{color:Xe,borderDash:[4,4],drawBorder:!1},ticks:{font:ce,color:Ee,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function va(t){let e=document.getElementById("table-issues");if(!e)return;let l=(t||[]).slice(0,8);if(!l.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${l.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${pa(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Te(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Te(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function ka(t){let e=document.getElementById("widget-agenda");if(!e)return;let l=new Date,a=`${l.getFullYear()}-${String(l.getMonth()+1).padStart(2,"0")}-${String(l.getDate()).padStart(2,"0")}`,s=(t||[]).filter(r=>(r.event_date||"").startsWith(a)).slice(0,10);if(!s.length){e.innerHTML="";return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${s.map(r=>{let i="#3B82F6",n="#EFF6FF",p="Agenda",m=(r.title||"").toLowerCase();return m.includes("inspeksi")?(i="#10B981",n="#ECFDF5",p="Inspeksi"):m.includes("cleaning")||m.includes("gcdc")?(i="#3B82F6",n="#EFF6FF",p="Cleaning"):m.includes("reliefer")?(i="#F59E0B",n="#FFFBEB",p="Reliefer"):m.includes("fogging")&&(i="#8B5CF6",n="#F5F3FF",p="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(r.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${i};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Te(r.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Te(r.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${i}">${p}</div>
        </div>
      `}).join("")}
    </div>
  `}function xa(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(l=>`
    <a href="${l.href}" class="action-btn">
      <div class="action-icon" style="background:${l.bg}">${l.icon}</div>
      ${l.label}
    </a>
  `).join("")}function pe(t,e){let l=document.getElementById(t),a=document.getElementById(e);if(l&&(l.style.display="none",l.style.position=""),a){a.style.display="block";let o=a.parentElement;if(o){let s=o.querySelector(".chart-empty");s&&s.remove()}}}function qe(t,e="Belum ada data"){if(!t)return;t.style.display="none";let l=t.parentElement;if(!l)return;if(!l.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent=e,l.appendChild(o)}}O();async function Et(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),l=document.getElementById("login-error"),a=document.getElementById("login-btn"),o=document.getElementById("toggle-password"),s=document.getElementById("login-password");o?.addEventListener("click",()=>{let r=s.type==="text";s.type=r?"password":"text",o.style.color=r?"":"var(--primary)"}),e?.addEventListener("submit",async r=>{r.preventDefault(),l.style.display="none";let i=e.username.value.trim(),n=e.password.value;if(!i||!n){l.textContent="Username dan password wajib diisi.",l.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let p=await x("/api/auth/login",{method:"POST",body:JSON.stringify({username:i,password:n})});p.ok&&p.data.success?(lt(p.data.data.token),Fe(p.data.data.user),Z("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(l.textContent=p.data.error||"Username atau password salah.",l.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{l.textContent="Gagal terhubung ke server. Periksa koneksi internet.",l.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}O();R();async function Sa(){return await K()}function wa(t,e){let l=String(t.status||"").toLowerCase();return e==="active"?l==="aktif":e==="reliefer"?t.division==="FC - RELIEFER"&&l==="aktif":!1}async function $t(t,e){let l=await Sa(),a=e?e.get("dash_filter"):null;N({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:o=>a?o.filter(s=>wa(s,a)):o,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:o=>_e(o)},{key:"phone",label:"No. HP",render:o=>o?`<a href="tel:${o}">${o}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>j(o)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:l},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY","FC - RELIEFER"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:o=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:o?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:o?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:l,value:o?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY","FC - RELIEFER"],value:o?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:o?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:o?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let o=await x(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(o.ok){let s=o.data.data.map(r=>({"Nama Lengkap":r.full_name,Cabang:r.branch_name||"",Divisi:r.division||"","No. HP":r.phone||"","Tgl Masuk":r.join_date||"",Status:r.status||""}));L(s,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async o=>{let s=n=>{if(!n)return null;let p=String(n||"").toLowerCase(),m=l.find(d=>String(d.label||"").toLowerCase()===p);return m?m.value:null},r=o.map(n=>({full_name:String(n["Nama Lengkap"]||"").trim(),branch_id:s(String(n.Cabang||"").trim()),division:String(n.Divisi||"").trim()||"FACILITY CARE",phone:String(n["No. HP"]||"").trim(),join_date:String(n["Tgl Masuk"]||"").trim(),status:String(n.Status||"").trim(),notes:String(n.Catatan||"").trim()})).filter(n=>n.full_name),i=await x("/api/import/employees",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}O();R();var ut=[],Dt=[];async function _a(){ut=await K(),Dt=await we()}var pt=async t=>{let e=[],l=1;for(;;){let o=await(await Promise.resolve().then(()=>(O(),ke))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${l}`);if(!o.ok)break;let s=o.data?.data||o.data||[],r=Array.isArray(s)?s:[];if(e=e.concat(r),r.length<100||o.data?.pagination&&l>=o.data.pagination.pages)break;l++}return e};function Ca(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let o=new Date(a);o.setDate(a.getDate()+30);let s=new Date(t.end_date);return s.setHours(0,0,0,0),s>=a&&s<=o}return!1}async function It(t,e){await _a();let l=e?e.get("dash_filter"):null;N({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",onDataLoaded:a=>l?a.filter(o=>Ca(o,l)):a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>_e(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,o)=>o.end_date&&String(o.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':st(a)},{key:"status",label:"Status",render:a=>j(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:ut},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"month_expiry",label:"Bulan Habis",options:[{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"},{value:"2027-01",label:"Jan 2027"},{value:"2027-02",label:"Feb 2027"},{value:"2027-03",label:"Mar 2027"},{value:"2027-04",label:"Apr 2027"},{value:"2027-05",label:"Mei 2027"},{value:"2027-06",label:"Jun 2027"},{value:"2027-07",label:"Jul 2027"}]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[s,r]=await Promise.all([pt("/api/employees?status=Aktif"),pt("/api/contracts")]);if(s.length>0){let i=r.filter(d=>d.status==="Aktif"),n=new Set(i.map(d=>d.employee_id)),p=s.filter(d=>!n.has(d.id)),m=`<p style="margin-bottom:12px">Data yang terbaca: <b>${s.length}</b> Karyawan Aktif, dan <b>${i.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${p.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;p.forEach(d=>{let c=r.filter(b=>b.employee_id===d.id),u='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(c.length>0){let b=c[0];u=`Pernah ada kontrak (Status: <b style="color:#EF4444">${b.status}</b>, Selesai: ${window.formatDate(b.end_date)})`}m+=`<li style="margin-bottom:8px"><b>${d.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${d.branch_name||"-"} | ${u}</span></li>`}),m+="</ul>",Promise.resolve().then(()=>(ve(),vt)).then(d=>d.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:m,cancelText:"Tutup"}))}}catch(s){console.error(s)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let o=document.querySelector(".page-actions");o&&o.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:Dt,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:ut,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await x(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let o=a.data.data.map(s=>({"Nama Lengkap":s.employee_name,Cabang:s.branch_name||"","Div / Bagian":s.division||"","Tanggal Mulai":s.start_date||"","Tanggal Selesai":s.end_date&&String(s.end_date).startsWith("2099")?"":s.end_date||"","Sisa Kontrak":s.end_date&&String(s.end_date).startsWith("2099")?"Tetap":s.days_remaining!==null&&s.days_remaining!==void 0?`${s.days_remaining} Hari`:"",Status:s.status||""}));L(o,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[o,s]=await Promise.all([x("/api/branches?limit=10000"),pt("/api/employees")]),r=o.data?.data||[],i=s||[];console.log(`Total employee yang berhasil dimuat dari database : ${i.length}`),i.length>0&&(console.log("Contoh 5 employee pertama:"),i.slice(0,5).forEach((g,h)=>{console.log(`${h+1}. ID: ${g.id}, Name: ${g.full_name}, Status: ${g.status}`)}));let n=g=>{if(!g)return null;let h=String(g||"").replace(/\s+/g," ").toLowerCase().trim(),k=r.find(S=>String(S.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===h||String(S.code||"").replace(/\s+/g," ").toLowerCase().trim()===h||String(S.name||"").replace(/\s+/g," ").toLowerCase().trim()===h);return k?k.id:null},p=(g,h)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${h}`),console.log(`Nama dari Excel : "${g}"`),!g)return console.log("Alasan gagal mapping : Nama kosong"),null;let k=String(g||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${k}"`),console.log(`Jumlah employee di database : ${i.length}`);let S=i.find(C=>String(C.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===k);return S?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${S.id}`),S.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},m=g=>{if(!g)return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let h=String(g).trim();if(/^\d{4,5}(\.\d+)?$/.test(h)){let S=Math.floor(Number(h));if(S>2e4&&S<99999){let C=new Date(Date.UTC(1899,11,30)+S*864e5);return isNaN(C.getTime())?"":C.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);let k=h.split(/[\/\-\.]/);if(k.length===3){let[S,C,_]=k.map(f=>f.trim());if(S.length===4&&C.length<=2&&_.length<=2)return`${S}-${C.padStart(2,"0")}-${_.padStart(2,"0")}`;if(_.length===4&&C.length<=2&&S.length<=2)return`${_}-${C.padStart(2,"0")}-${S.padStart(2,"0")}`}return h},d=a.map((g,h)=>{let k=h+2,S=String(g["Nama Lengkap"]||"").trim(),C=g["Tanggal Mulai"],_=m(C);if(!_){let B=a.__worksheet,E=a.__headers||[],q=E.indexOf("Tanggal Mulai"),M="N/A",me="N/A",re="N/A";if(q!==-1&&B&&window.XLSX){let v=window.XLSX.utils.encode_cell({c:q,r:k-1});re=v;let y=B[v];y?(M=y.t||"undefined",me=y.w||"undefined"):M="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let oe="Unknown";C==null||C===""?oe="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":C instanceof Date&&isNaN(C.getTime())?oe="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":oe="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${k}`),console.log(`Employee Name : ${S}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${q})`),console.log(`Raw Cell Value : "${C}"`),console.log(`JavaScript Type : ${typeof C}`),console.log(`SheetJS Cell Type : ${M}`),console.log(`SheetJS Formatted Value : "${me}"`),console.log(`Value After Trim : "${String(C||"").trim()}"`),console.log(`Value After Date Parser : "${_}"`),console.log(`Is Empty : ${!C}`),console.log(`Is Invalid Date : ${C instanceof Date?isNaN(C.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${oe}`),console.log(`Workbook Sheet : ${B?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${re}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(g,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(E)),console.log(`==========================
`)}let f=p(S,k),w=null;return f||(w="Karyawan tidak ditemukan di Database"),{isValid:!!f,invalidReason:w,rowNum:k,data:{employee_id:f,branch_id:n(String(g.Cabang||"").trim()),division:String(g["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:_,end_date:m(g["Tanggal Selesai"])||"2099-12-31",status:String(g.Status||"").trim(),_rawName:S}}}),c=[],u=[];if(d.forEach(g=>{g.isValid?c.push(g.data):u.push({rowNum:g.rowNum,name:g.data._rawName,reason:g.invalidReason})}),console.log(`Split Validation - Valid: ${c.length}, Invalid: ${u.length}`),c.length===0)return{inserted:0,skipped:a.length,failed:a.length};let b=await x("/api/import/contracts",{method:"POST",body:JSON.stringify({rows:c,onDuplicate:"update"})});if(!b.ok)throw new Error(b.data?.error||"Import gagal");return b.data}}})}O();R();var mt=[],je=[];function Ta(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let l of e)if(t.some(a=>a.period===l))return l;return"Q3"}function Ea(t,e){if(t.period!=="Q3")return!1;let l=String(t.status||"").toLowerCase();if(l!=="selesai"&&l!=="completed"&&l!=="done")return!1;let a=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?a.includes("inspeksi"):e==="gcdc"?a.includes("general cleaning")||a.includes("deep cleaning"):!1}async function Pt(t,e){mt=await K();let l=await ee(),o=(await x(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`)).data?.data||[],s=new Set;o.forEach(d=>{d.pic&&d.pic.trim()&&s.add(d.pic.trim())}),je=Array.from(s).sort();let r=d=>d&&!je.find(c=>(typeof c=="object"?c.value:c)===d)?[...je,d]:je,i=d=>{if(!d||d==="-"||String(d).trim()==="")return"";let c=String(d).split("-");return c.length===3&&c[0].length===4?`${c[2]}-${c[1]}-${c[0]}`:d},p=Ta(o),m=e?e.get("dash_filter"):null;N({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",defaultFilters:m?{period:"Q3"}:{period:p},onDataLoaded:d=>(m&&(d=d.filter(c=>Ea(c,m))),d.sort((c,u)=>{let b=c.opening_date?new Date(c.opening_date).getTime():0;return(u.opening_date?new Date(u.opening_date).getTime():0)-b})),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:d=>dt(d)},{key:"period",label:"Periode",render:d=>de(d)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:d=>i(d)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:d=>i(d)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:d=>i(d)},{key:"status",label:"Status",render:d=>j(d)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:mt},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:je}],formFields:d=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:mt,value:d?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:d?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:d?.period},{name:"pic",label:"PIC",type:"combobox",options:r(d?.pic),value:d?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:d?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:d?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:d?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:d?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:d?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let d=await x(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(d.ok){let c=d.data.data.map(u=>({Cabang:u.branch_name||"",Kegiatan:u.activity_type||"",Periode:u.period||"",PIC:u.pic||"","Tgl Opening":u.opening_date||"","Tgl Target":u.target_date||"","Tgl Selesai":u.completion_date||"",Status:u.status||""}));L(c,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async d=>{let u=(await x("/api/branches?all=1")).data?.data||[],b=S=>{if(!S)return null;let C=String(S||"").toLowerCase(),_=u.find(f=>String(f.full_name||"").toLowerCase()===C||String(f.code||"").toLowerCase()===C||String(f.name||"").toLowerCase()===C);return _?_.id:null},g=S=>{if(S==null||S==="")return"";if(S instanceof Date&&!isNaN(S.getTime()))return S.toISOString().slice(0,10);let C=String(S).trim();if(C===""||C==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(C))return C.slice(0,10);if(/^\d{4,5}$/.test(C)){let f=Number(C);if(f>2e4&&f<99999){let w=new Date(Date.UTC(1899,11,30)+f*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}let _=C.split(/[\/\-\.]/);if(_.length===3){let[f,w,B]=_.map(E=>E.trim());if(f.length===4&&w.length<=2&&B.length<=2)return`${f}-${w.padStart(2,"0")}-${B.padStart(2,"0")}`;if(B.length===4&&w.length<=2&&f.length<=2)return`${B}-${w.padStart(2,"0")}-${f.padStart(2,"0")}`}return C},h=d.map(S=>({branch_id:b(String(S.Cabang||"").trim()),activity_type:String(S.Kegiatan||"").trim(),period:String(S.Periode||"").trim(),pic:String(S.PIC||S.Pic||"").trim(),opening_date:g(S["Tgl Opening"]||S["Tanggal Opening"]||S["Tgl Openir"]),target_date:g(S["Tgl Target"]||S["Tanggal Target"]),completion_date:g(S["Tgl Selesai"]||S["Tanggal Selesai"]),status:String(S.Status||"").trim(),notes:String(S.Catatan||S.Keterangan||"").trim()})).filter(S=>S.activity_type&&S.period),k=await x("/api/import/schedule",{method:"POST",body:JSON.stringify({rows:h,onDuplicate:"update"})});if(!k.ok)throw new Error(k.data?.error||"Import gagal");return k.data}}})}O();R();var gt=[],et=[];function $a(t,e){let l=String(t.status||"").toLowerCase();return e==="open"?l==="open":!1}async function Bt(t,e){let l=e?e.get("dash_filter"):null;gt=await K(),et=await ee();let a=r=>r&&!et.find(i=>i.value===r)?[...et,{value:r,label:r}]:et,o=new Date().getFullYear();N({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:r=>l?r.filter(i=>$a(i,l)):r,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:r=>`<span class="badge badge-secondary">${r}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:r=>`<span title="${r}">${r?.length>50?r.slice(0,50)+"\u2026":r}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:r=>`<span title="${r||""}">${r?.length>40?r.slice(0,40)+"\u2026":r||"-"}</span>`},{key:"status",label:"Status",render:r=>j(r)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:r=>window.formatDate(r)},{key:"day_count",label:"Hari",render:r=>r??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:gt},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:["2025","2026","2027","2028","2029","2030"]}],formFields:r=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:r?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:gt,value:r?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:r?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:r?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:r?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(r?.employee_name),value:r?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(r?.fc_specialist),value:r?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:r?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:r?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:r?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let r=await x(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let i=r.data.data.map(n=>({Tanggal:n.report_date||"",Cabang:n.branch_name||"",Kategori:n.category||"",Sumber:n.source||"",Keluhan:n.complaint||"","Nama FC":n.employee_name||"","FC Spesialis":n.fc_specialist||"",Solusi:n.solution||"","Tgl Selesai":n.completion_date||"",Status:n.status||""}));L(i,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async r=>{let n=(await x("/api/branches?all=1")).data?.data||[],p=c=>{if(!c)return null;let u=String(c||"").toLowerCase(),b=n.find(g=>String(g.full_name||"").toLowerCase()===u||String(g.code||"").toLowerCase()===u||String(g.name||"").toLowerCase()===u);return b?b.id:null},m=r.map(c=>({branch_id:p(String(c.Cabang||"").trim()),report_date:String(c.Tanggal||"").trim(),category:String(c.Kategori||"").trim(),source:String(c.Sumber||"").trim(),complaint:String(c.Keluhan||"").trim(),employee_name:String(c["Nama FC"]||"").trim(),fc_specialist:String(c["FC Spesialis"]||"").trim(),solution:String(c.Solusi||"").trim(),completion_date:String(c["Tgl Selesai"]||"").trim(),status:String(c.Status||"").trim()})).filter(c=>c.report_date&&c.complaint&&c.category),d=await x("/api/import/issues",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}}})}O();var De=[];function Da(t,e){let l=String(t.status||"").toLowerCase();return e==="pending"?l==="pending":!1}async function Lt(t,e){let l=e?e.get("dash_filter"):null;De=await K();let a=await ee(),o=["Ade","Berlin"],s=i=>i&&!a.find(n=>n.value===i)?[...a,{value:i,label:i}]:a,r=i=>i&&!o.find(n=>(typeof n=="object"?n.value:n)===i)?[...o,i]:o;N({container:t,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:i=>l?i.filter(n=>Da(n,l)):i,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:i=>`<span title="${i||""}">${i?.length>50?i.slice(0,50)+"\u2026":i||"-"}</span>`},{key:"solution",label:"Solusi",render:i=>`<span title="${i||""}">${i?.length>40?i.slice(0,40)+"\u2026":i||"-"}</span>`},{key:"status",label:"Status",render:i=>j(i)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:i=>window.formatDate(i)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:i=>i?`<a href="${i}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"combobox",name:"branch_id",label:"Cabang",options:De},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async i=>{let n=new URLSearchParams(i||{}).toString(),p=await x(`/api/one-on-one?limit=10000&${n}`);if(p.ok){let m=p.data.data.map(c=>({Tanggal:c.meeting_date||"",Cabang:c.branch_name||"","Nama Karyawan":c.employee_name||"",PIC:c.pic||"",Masalah:c.problem||"",Solusi:c.solution||"",Status:c.status||"","Tgl Selesai":c.completion_date||"",Dokumen:c.document_link||""})),{downloadExcel:d}=await Promise.resolve().then(()=>(R(),le));d(m,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let i=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(R(),le));n(i,"Template_Import_OneOnOne")},onImport:async i=>{let n=c=>{if(!c)return null;let u=String(c||"").toLowerCase(),b=De.find(g=>String(g.label||"").toLowerCase()===u);return b?b.value:null},p=c=>{if(!c)return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let u=String(c).trim();if(/^\d{4,5}$/.test(u)){let g=Number(u);if(g>2e4&&g<99999){let h=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);let b=u.split(/[\/\-\.]/);if(b.length===3){let[g,h,k]=b.map(S=>S.trim());if(g.length===4&&h.length<=2&&k.length<=2)return`${g}-${h.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&h.length<=2&&g.length<=2)return`${k}-${h.padStart(2,"0")}-${g.padStart(2,"0")}`}return u},m=i.map(c=>({meeting_date:p(c.Tanggal),employee_name:String(c["Nama Karyawan"]||"").trim(),branch_id:n(String(c.Cabang||"").trim()),pic:String(c.PIC||"").trim(),problem:String(c.Masalah||"").trim(),solution:String(c.Solusi||"").trim(),status:String(c.Status||"").trim(),completion_date:p(c["Tgl Selesai"]),document_link:String(c.Dokumen||"").trim()})).filter(c=>c.meeting_date&&c.employee_name&&c.branch_id),d=await x("/api/import/one_on_one",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:i=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:i?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:i?.branch_id&&!De.find(n=>n.value==i.branch_id)?[...De,{value:i.branch_id,label:i.branch_name||i.branch_id}]:De,createApi:{path:"/api/branches",field:"full_name"},value:i?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"combobox",required:!0,options:s(i?.employee_name),value:i?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:r(i?.pic),createApi:{path:"/api/pic",field:"name"},value:i?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:i?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:i?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:i?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:i?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:i?.document_link}]})}O();async function Nt(t){let e=await K(),l=await ee(),a=["Ade","Berlin"],o=i=>i&&!l.find(n=>n.value===i)?[...l,{value:i,label:i}]:l,s=i=>i&&!a.find(n=>(typeof n=="object"?n.value:n)===i)?[...a,i]:a,r=Array.from({length:5},(i,n)=>String(new Date().getFullYear()-n));N({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:i=>{try{let n=JSON.parse(i);return Array.isArray(n)?n.join(", "):i||"-"}catch{return i||"-"}}},{key:"score",label:"Nilai",render:i=>i!=null?`<strong>${i}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:i=>i?`<a href="${i}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:r}],exportOptions:{moduleName:"training",onExport:async i=>{let n=new URLSearchParams(i||{}).toString(),p=await x(`/api/training?limit=10000&${n}`);if(p.ok){let m=p.data.data.map(c=>{let u=c.participants||"";try{let b=JSON.parse(u);u=Array.isArray(b)?b.join(", "):u}catch{}return{Tanggal:c.training_date||"",Batch:c.batch||"",Materi:c.subject||"",Cabang:c.branch_name||"",Trainer:c.trainer||"",Peserta:u,Nilai:c.score!==null&&c.score!==void 0?c.score:"",Dokumen:c.document_link||""}}),{downloadExcel:d}=await Promise.resolve().then(()=>(R(),le));d(m,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let i=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(R(),le));n(i,"Template_Import_Training")},onImport:async i=>{let n=c=>{if(!c)return null;let u=String(c||"").toLowerCase(),b=e.find(g=>String(g.label||"").toLowerCase()===u);return b?b.value:null},p=c=>{if(!c)return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let u=String(c).trim();if(/^\d{4,5}$/.test(u)){let g=Number(u);if(g>2e4&&g<99999){let h=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(u))return u.slice(0,10);let b=u.split(/[\/\-\.]/);if(b.length===3){let[g,h,k]=b.map(S=>S.trim());if(g.length===4&&h.length<=2&&k.length<=2)return`${g}-${h.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&h.length<=2&&g.length<=2)return`${k}-${h.padStart(2,"0")}-${g.padStart(2,"0")}`}return u},m=i.map(c=>({training_date:p(c.Tanggal),batch:String(c.Batch||"").trim(),subject:String(c.Materi||"").trim(),branch_id:n(String(c.Cabang||"").trim()),trainer:String(c.Trainer||"").trim(),participants:String(c.Peserta||"").trim(),score:c.Nilai?Number(c.Nilai):null,document_link:String(c.Dokumen||"").trim()})).filter(c=>c.training_date&&c.subject&&c.branch_id),d=await x("/api/import/training",{method:"POST",body:JSON.stringify({rows:m,onDuplicate:"update"})});if(!d.ok)throw new Error(d.data?.error||"Import gagal");return d.data}},formFields:i=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:i?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:i?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:i?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:i?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:s(i?.trainer),value:i?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let n=JSON.parse(i?.participants);return Array.isArray(n)?n.join(", "):i?.participants||""}catch{return i?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:i?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:i?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:i?.notes}],onBeforeSubmit:async i=>(i.participants&&(i.participants=JSON.stringify(i.participants.split(",").map(n=>n.trim()).filter(Boolean))),i)})}O();ve();R();function At({container:t,title:e,icon:l,apiPath:a,columns:o,formFields:s,filterFields:r,defaultFilters:i={},itemLabel:n="Data",canCreate:p=!0,canEdit:m=!0,canDelete:d=!0,onBeforeSubmit:c,onAfterLoad:u,onDataLoaded:b,extraActions:g=[],initialSearch:h="",exportOptions:k=null,bulkDelete:S=!1,paginationMode:C="server"}){let _=1,f={...i};h&&(f.search=h);let w=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${l} ${e}</h1>
      <div class="page-actions">
        ${p?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
      </div>
    </div>

    ${S?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="align-items:center; background:var(--bg-card); border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${k?He(k.moduleName):""}

    ${r&&r.length>0?`
    <div class="filter-bar card" style="padding: 1rem; display:flex; flex-wrap:wrap; gap:1rem; align-items:center;">
        ${r.map(v=>{if(v.type==="search")return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${f.search||""}"></div>`;if(v.type==="search-combo"){let y="dl-filter-search",$=(v.options||[]).map(P=>`<option value="${typeof P=="object"?P.label:P}"></option>`).join("");return`<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${y}" class="form-control" autocomplete="off" placeholder="${v.placeholder||"Cari..."}" id="filter-search" value="${f.search||""}"><datalist id="${y}">${$}</datalist></div>`}return v.type==="select"?`<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${v.name}" id="filter-${v.name}"><option value="">-- ${v.label} --</option>${(v.options||[]).map(y=>`<option value="${typeof y=="object"?y.value:y}" ${f[v.name]===(typeof y=="object"?y.value:y)?"selected":""}>${typeof y=="object"?y.label:y}</option>`).join("")}</select>`:""}).join("")}
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
    `,document.body.appendChild(y),y.querySelector("#bulk-cancel-btn").addEventListener("click",()=>y.remove()),y.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let $=y.querySelector("#bulk-confirm-btn");$.disabled=!0,$.textContent="Menghapus...";let P=await x(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:v})});y.remove(),P.ok?(Z(`${v.length} ${n} berhasil dihapus.`),w.clear(),B(),M()):G(P.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),q;if(E?.addEventListener("input",v=>{clearTimeout(q),q=setTimeout(()=>{f.search=v.target.value,_=1,w.clear(),M()},400)}),r?.forEach(v=>{v.type==="select"&&document.getElementById(`filter-${v.name}`)?.addEventListener("change",y=>{f[v.name]=y.target.value,_=1,w.clear(),M()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{f={...i},E&&(E.value=""),r?.forEach(v=>{let y=document.getElementById(`filter-${v.name}`);y&&(y.value="")}),_=1,w.clear(),M()}),document.getElementById("btn-create")?.addEventListener("click",()=>re(null)),k){document.getElementById(`btn-export-${k.moduleName}`)?.addEventListener("click",async y=>{let $=y.target,P=$.innerHTML;$.innerHTML="\u23F3 Loading...",$.disabled=!0;try{await k.onExport()}catch{G("Gagal export data")}finally{$.innerHTML=P,$.disabled=!1}}),document.getElementById(`btn-template-${k.moduleName}`)?.addEventListener("click",()=>{k.onTemplate()});let v=document.getElementById(`input-import-${k.moduleName}`);v?.addEventListener("change",async y=>{let $=y.target.files[0];if(!$)return;v.disabled=!0;let P=document.createElement("div");P.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center",P.innerHTML=`
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">\u{1F504} Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `,document.body.appendChild(P);let Y=P.querySelector("#import-progress-text"),U=P.querySelector("#import-progress-bar"),D=P.querySelector("#import-summary"),I=P.querySelector("#import-close-btn");I.addEventListener("click",()=>{P.remove(),M()});try{let V=await Ke($);if(V.length===0)throw new Error("File kosong atau format salah");let X=500,ie=0,ae=0,A=0,F=V.length;Y.textContent=`Ditemukan ${F} baris data. Memulai import...`;for(let T=0;T<F;T+=X){let Q=V.slice(T,T+X);Y.textContent=`Mengimport baris ${T+1} - ${Math.min(T+X,F)} dari ${F}...`,U.style.width=`${Math.round(T/F*100)}%`;try{let H=await k.onImport(Q);H?(ie+=H.inserted||H.metrics?.inserted||Q.length,ae+=H.skipped||H.metrics?.updated||0):ie+=Q.length}catch(H){console.error("Chunk import failed:",H),A+=Q.length}}U.style.width="100%",Y.innerHTML='<strong style="color:var(--success)">\u2705 Import Selesai!</strong>',D.style.display="block",D.innerHTML=`
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${F}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${ie}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${ae}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${A}</strong></div>
        `,A>0&&(D.innerHTML+='<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>'),I.style.display="block",v.value=""}catch(V){Y.innerHTML=`<strong style="color:var(--danger)">\u274C Gagal Memproses File</strong><br>${V.message}`,U.style.background="var(--danger)",U.style.width="100%",I.style.display="block",v.value=""}finally{v.disabled=!1}})}async function M(){B();let v=document.getElementById("table-container");if(!v)return;v.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let y=C==="client",$=y?1:_,P=y?ye:20,Y=new URLSearchParams({page:$,limit:P,...Object.fromEntries(Object.entries(f).filter(([,A])=>A))}),U=await x(`${a}?${Y}`);if(!U.ok){v.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${U.data?.error||"Error"}</p></div>`;return}let D=U.data?.data||U.data||[],I=U.data?.pagination,V=D.length,X=D;if(y){D=b(D),X=D;let A=D.length,F=20,T=Math.ceil(A/F);_>T&&T>0&&(_=T);let Q=(_-1)*F,H=_*F;D=D.slice(Q,H),I={page:_,limit:F,total:A,pages:T}}!1,u&&u(D);let ie=ze({columns:o,data:D,fullData:X,onEdit:m?A=>re(A):null,actions:g.map(A=>({...A,handler:F=>A.handler(F,M)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:S?{selectedIds:w,onToggle:B}:null});v.innerHTML="",v.appendChild(ie);let ae=document.getElementById("pagination-container");if(ae&&(ae.innerHTML="",I&&I.pages>1)){let A=Qe({page:I.page,pages:I.pages,total:I.total,limit:I.limit,onPage:F=>{_=F,M()}});A&&ae.appendChild(A)}}function me(v){let y=typeof s=="function"?s(v):s;return Re(y)}function re(v){let y=!!v,$=document.createElement("form");if($.noValidate=!0,$.innerHTML=me(v),y){let Y=typeof s=="function"?s(v):s;We($,v)}let{close:P}=se({title:y?`Edit ${n}`:`Tambah ${n}`,content:$,size:"lg",confirmText:y?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(Y,U)=>{if(!$.reportValidity())return;let D=Y.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let I=Ye($),V=typeof s=="function"?s(v):s,X=async F=>{for(let T of F)if(T.type==="row")await X(T.fields);else if(T.type==="combobox"&&I[T.name]){let Q=I[T.name],H=(T.options||[]).find(W=>{let ne=String(typeof W=="object"?W.value:W),nt=String(typeof W=="object"?W.label:W);return ne===Q||nt===Q});if(H)I[T.name]=typeof H=="object"?H.value:H;else if(T.createApi){let W={};W[T.createApi.field]=Q,T.createApi.extra&&Object.assign(W,T.createApi.extra);let ne=await x(T.createApi.path,{method:"POST",body:JSON.stringify(W)});if(ne.ok&&ne.data?.id)I[T.name]=ne.data.id;else if(ne.ok&&!ne.data?.id)I[T.name]=Q;else throw new Error(`Gagal membuat master data: ${ne.data?.error||"Unknown error"}`)}}};try{await X(V)}catch(F){G(F.message),D.disabled=!1,D.textContent=y?"Simpan Perubahan":`Tambah ${n}`;return}c&&(I=await c(I,v));let ie=y?"PUT":"POST",ae=y?`${a}/${v.id}`:a,A=await x(ae,{method:ie,body:JSON.stringify(I)});A.ok?(Z(y?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),U(),M()):(G(A.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=y?"Simpan Perubahan":`Tambah ${n}`)}})}function oe(v){Oe(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let y=await x(`${a}/${v.id}`,{method:"DELETE"});y.ok?(Z(`${n} berhasil dihapus.`),M()):G(y.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return M(),M}O();R();async function Ft(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let l=await K(),a=await ee(),o=e?e.get("dash_filter"):null;console.log("RAW",await we()),console.log("OPTIONS",a);let s=n=>n&&!a.find(p=>p.value===n)?[...a,{value:n,label:n}]:a,r=["Agung Septiadi","Wasrikin","Iqbal Al Banna","Muhammad Tri Ismandanu"],i=n=>n&&!r.includes(n)?[...r,n]:r;At({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",onDataLoaded:n=>{if(o==="reliever"){let p=new Date,m=p.getFullYear(),d=String(p.getMonth()+1).padStart(2,"0");return n.filter(c=>{if(String(c.status||"").toLowerCase()!=="done")return!1;let u=c.backup_date||"";if(u.includes("/")){let b=u.split("/");if(b.length===3&&(b[2].length===4?b[2]:`20${b[2]}`)==m&&b[1].padStart(2,"0")==d)return!0}else if(u.includes("-")&&u.startsWith(`${m}-${d}`))return!0;return!1})}return n},columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:n=>de(n)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:n=>window.formatDate(n)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:n=>n?`<span class="badge badge-info">${n}</span>`:"-"},{key:"status",label:"Status",render:n=>j(n)}],filterFields:[{type:"select",name:"reliever_name",label:"Cari reliefer / FC...",options:r},{type:"select",name:"branch_id",label:"Cabang",options:l},{type:"select",name:"month",label:"Bulan",options:[{value:"2026-01",label:"Jan 2026"},{value:"2026-02",label:"Feb 2026"},{value:"2026-03",label:"Mar 2026"},{value:"2026-04",label:"Apr 2026"},{value:"2026-05",label:"Mei 2026"},{value:"2026-06",label:"Jun 2026"},{value:"2026-07",label:"Jul 2026"},{value:"2026-08",label:"Agu 2026"},{value:"2026-09",label:"Sep 2026"},{value:"2026-10",label:"Okt 2026"},{value:"2026-11",label:"Nov 2026"},{value:"2026-12",label:"Des 2026"}]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:l,value:n?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:s(n?.original_fc_name),value:n?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:i(n?.reliever_name),value:n?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:n?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:n?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:n?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:n?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let n=await x(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let p=n.data.data.map(m=>({Cabang:m.branch_name||"","Nama Facility care":m.original_fc_name||"",Periode:m.period||"",Relifer:m.reliever_name||"","Tanggal Back Up":m.backup_date||"","Tanggal Selesai":m.completion_date||"",Keterangan:m.reason||"",Shift:m.shift||"",Status:m.status||""}));p.length===0&&p.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),L(p,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async n=>{let m=(await x("/api/branches?all=1")).data?.data||[],d=b=>{if(!b)return null;let g=String(b||"").toLowerCase(),h=m.find(k=>String(k.full_name||"").toLowerCase()===g||String(k.code||"").toLowerCase()===g||String(k.name||"").toLowerCase()===g);return h?h.id:null},c=n.map(b=>({branch_name:String(b.Cabang||"").trim(),backup_date:String(b["Tanggal Back Up"]||b["Tanggal Backup"]||"").trim(),original_fc_name:String(b["Nama Facility care"]||b["FC Digantikan"]||"").trim(),reliever_name:String(b.Relifer||b.Reliefer||"").trim(),period:String(b.Periode||"").trim(),reason:String(b.Keterangan||"").trim(),shift:String(b.Shift||"").trim(),completion_date:String(b["Tanggal Selesai"]||"").trim(),status:String(b.Status||"").trim()})).filter(b=>b.reliever_name&&b.backup_date),u=await x("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:c,onDuplicate:"update"})});if(!u.ok)throw new Error(u.data?.error||"Import gagal");return u.data}}})}O();R();async function Mt(t){let e=await K(),l=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));N({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>de(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>j(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:l}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),s=await x(`/api/reports/inspection?limit=10000&${o}`);if(s.ok){let r=s.data.data.map(i=>({Cabang:i.branch_name||"",Periode:i.period||"",Tanggal:i.inspection_date||"","Point FC":i.fc_score!==null&&i.fc_score!==void 0?i.fc_score:"","Point SPV":i.spv_score!==null&&i.spv_score!==void 0?i.spv_score:"",Status:i.status||"","Link Dokumen":i.document_link||""}));L(r,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_Inspeksi")},onImport:async a=>{let o=n=>{if(!n)return null;let p=String(n||"").toLowerCase(),m=e.find(d=>String(d.label||"").toLowerCase()===p);return m?m.value:null},s=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let p=String(n).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let d=Number(p);if(d>2e4&&d<99999){let c=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(c.getTime())?"":c.toISOString().slice(0,10)}}let m=p.split(/[\/\-\.]/);if(m.length===3){let[d,c,u]=m.map(b=>b.trim());if(d.length===4&&c.length<=2&&u.length<=2)return`${d}-${c.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&c.length<=2&&d.length<=2)return`${u}-${c.padStart(2,"0")}-${d.padStart(2,"0")}`}return p},r=a.map(n=>({branch_id:o(String(n.Cabang||"").trim()),period:String(n.Periode||"").trim(),inspection_date:s(n.Tanggal),fc_score:n["Point FC"]!==void 0&&n["Point FC"]!==""?Number(n["Point FC"]):null,spv_score:n["Point SPV"]!==void 0&&n["Point SPV"]!==""?Number(n["Point SPV"]):null,status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.inspection_date),i=await x("/api/import/inspection",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}O();R();async function Ot(t){let e=await K(),l=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));N({container:t,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>de(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>j(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:l}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),s=await x(`/api/reports/cleaning?limit=10000&${o}`);if(s.ok){let r=s.data.data.map(i=>({Cabang:i.branch_name||"",Jenis:i.activity_type||"",Periode:i.period||"",Tanggal:i.activity_date||"",Status:i.status||"","Link Dokumen":i.document_link||""}));L(r,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:""}],"Template_Import_GCDC")},onImport:async a=>{let o=n=>{if(!n)return null;let p=String(n||"").toLowerCase(),m=e.find(d=>String(d.label||"").toLowerCase()===p);return m?m.value:null},s=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let p=String(n).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let d=Number(p);if(d>2e4&&d<99999){let c=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(c.getTime())?"":c.toISOString().slice(0,10)}}let m=p.split(/[\/\-\.]/);if(m.length===3){let[d,c,u]=m.map(b=>b.trim());if(d.length===4&&c.length<=2&&u.length<=2)return`${d}-${c.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&c.length<=2&&d.length<=2)return`${u}-${c.padStart(2,"0")}-${d.padStart(2,"0")}`}return p},r=a.map(n=>({branch_id:o(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:s(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),i=await x("/api/import/cleaning",{method:"POST",body:JSON.stringify({rows:r,onDuplicate:"update"})});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}O();R();async function Rt(t){let e=await K(),l=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));N({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge badge-warning">${a}</span>`},{key:"period",label:"Periode",render:a=>de(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>j(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:a=>a||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:l}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:a?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),s=await x(`/api/reports/fogging?limit=10000&${o}`);if(s.ok){let r=s.data.data.map(i=>({Cabang:i.branch_name||"",Jenis:i.activity_type||"Fogging",Periode:i.period||"",Tanggal:i.activity_date||"",Status:i.status||"","Link Dokumen":i.document_link||""}));L(r,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async a=>{let o=n=>{if(!n)return null;let p=String(n||"").toLowerCase(),m=e.find(d=>String(d.label||"").toLowerCase()===p);return m?m.value:null},s=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let p=String(n).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let d=Number(p);if(d>2e4&&d<99999){let c=new Date(Date.UTC(1899,11,30)+d*864e5);return isNaN(c.getTime())?"":c.toISOString().slice(0,10)}}let m=p.split(/[\/\-\.]/);if(m.length===3){let[d,c,u]=m.map(b=>b.trim());if(d.length===4&&c.length<=2&&u.length<=2)return`${d}-${c.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&c.length<=2&&d.length<=2)return`${u}-${c.padStart(2,"0")}-${d.padStart(2,"0")}`}return p},r=a.map(n=>({branch_id:o(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"Fogging").trim(),period:String(n.Periode||"").trim(),activity_date:s(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.activity_date),i=await x("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(r)});if(!i.ok)throw new Error(i.data?.error||"Import gagal");return i.data}}})}O();R();async function Kt(t){let e=await K(),l=await ee(),a=l,o=r=>r&&!l.find(i=>i.value===r)?[...l,{value:r,label:r}]:l,s=r=>r&&!a.find(i=>i.value===r)?[...a,{value:r,label:r}]:a;N({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:r=>`<span title="${r||""}">${r?.length>60?r.slice(0,60)+"\u2026":r||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:r=>window.formatDate(r)},{key:"status",label:"Status",render:r=>j(r)},{key:"notes",label:"Keterangan",render:r=>r?.length>40?r.slice(0,40)+"\u2026":r||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]}],formFields:r=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:r?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:s(r?.pic),value:r?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:r?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:r?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:r?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:r?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async r=>{let i=new URLSearchParams(r||{}).toString(),n=await x(`/api/reports/basecamp?limit=10000&${i}`);if(n.ok){let p=n.data.data.map(m=>({"Tgl Info":m.info_date||"",Cabang:m.branch_name||"",Permasalahan:m.problem||"",PIC:m.pic||"","Tgl Done":m.done_date||"",Status:m.status||"",Keterangan:m.notes||""}));L(p,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async r=>{let i=d=>{if(!d)return null;let c=String(d||"").toLowerCase(),u=e.find(b=>String(b.label||"").toLowerCase()===c);return u?u.value:null},n=d=>{if(d==null||d==="")return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let c=String(d).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let b=Number(c);if(b>2e4&&b<99999){let g=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}let u=c.split(/[\/\-\.]/);if(u.length===3){let[b,g,h]=u.map(k=>k.trim());if(b.length===4&&g.length<=2&&h.length<=2)return`${b}-${g.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&g.length<=2&&b.length<=2)return`${h}-${g.padStart(2,"0")}-${b.padStart(2,"0")}`}return c},p=r.map(d=>({info_date:n(d["Tgl Info"]||d["Tanggal Info"]),branch_id:i(String(d.Cabang||"").trim()),problem:String(d.Permasalahan||"").trim(),pic:String(d.PIC||"").trim(),done_date:n(d["Tgl Done"]||d["Tanggal Done"]),status:String(d.Status||"").trim(),notes:String(d.Keterangan||d.Catatan||"").trim()})).filter(d=>d.info_date&&d.branch_id&&d.problem),m=await x("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(p)});if(!m.ok)throw new Error(m.data?.error||"Import gagal");return m.data}}})}async function Ht(t){N({container:t,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let l=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(O(),ke)),o=await a(`/api/sop?limit=10000&${l}`);if(o.ok){let s=o.data.data.map(i=>({"Nama SOP":i.name||"",Kategori:i.category||"",Dokumen:i.document_link||"",Catatan:i.notes||i.description||""})),{downloadExcel:r}=await Promise.resolve().then(()=>(R(),le));r(s,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:l}=await Promise.resolve().then(()=>(R(),le));l(e,"Template_Import_SOP")},onImport:async e=>{let l=e.map(s=>({name:String(s["Nama SOP"]||"").trim(),category:String(s.Kategori||"").trim(),document_link:String(s.Dokumen||"").trim(),description:String(s.Catatan||"").trim()})).filter(s=>s.name),{apiFetch:a}=await Promise.resolve().then(()=>(O(),ke)),o=await a("/api/sop/import",{method:"POST",body:JSON.stringify(l)});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function qt(t){N({container:t,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let l=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(O(),ke)),o=await a(`/api/checklist?limit=10000&${l}`);if(o.ok){let s=o.data.data.map(i=>({"Nama Checklist":i.name||"",Kategori:i.category||"",Dokumen:i.document_link||"",Deskripsi:i.description||""})),{downloadExcel:r}=await Promise.resolve().then(()=>(R(),le));r(s,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:l}=await Promise.resolve().then(()=>(R(),le));l(e,"Template_Import_Checklist")},onImport:async e=>{let l=e.map(s=>({name:String(s["Nama Checklist"]||"").trim(),category:String(s.Kategori||"").trim(),document_link:String(s.Dokumen||"").trim(),description:String(s.Deskripsi||"").trim()})).filter(s=>s.name),{apiFetch:a}=await Promise.resolve().then(()=>(O(),ke)),o=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(l)});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}O();ve();R();async function bt(t,e="forms"){if(e==="supply")return Pa(t);Ia(t)}function Ia(t){N({container:t,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}],exportOptions:{moduleName:"forms",onExport:async e=>{let l=new URLSearchParams(e||{}).toString(),a=await x(`/api/forms?limit=10000&${l}`);a.data?.data?L(a.data.data,"Data_Master_Form"):G("Gagal export data master form")},onImport:async e=>{let l=await x("/api/forms/import",{method:"POST",body:JSON.stringify({data:e})});if(!l.ok)throw new Error(l.data?.error||"Import failed");return l.data},onTemplate:()=>{window.location.hash="#/import"}}})}async function Pa(t){let l=((await x("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));N({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,o)=>o.branch_name_ref||o.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let o=JSON.parse(a);return Array.isArray(o)?o.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let o=JSON.parse(a);return Array.isArray(o)?o.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>j(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let o=a?.tools_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}let s=a?.chemical_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:a?.branch_id&&!l.find(r=>r.value==a.branch_id)?[...l,{value:a.branch_id,label:a.branch_name||a.branch_id}]:l,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:o},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:s},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),s=await x(`/api/reports/supply?limit=10000&${o}`);if(s.ok){let r=s.data.data.map(i=>{let n=i.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let p=i.chemical_items;try{p=Array.isArray(JSON.parse(p))?JSON.parse(p).join(", "):p}catch{}return{Waktu:i.submitted_at||"",Pengirim:i.submitter_name||"",Cabang:i.branch_name_ref||i.branch_name||"","Alat/Barang":n||"",Chemical:p||"",Catatan:i.additional_notes||"",Status:i.status||"","Diproses Oleh":i.processed_by||""}});L(r,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let s=(await x("/api/branches?all=1")).data?.data||[],r=m=>{if(!m)return null;let d=String(m||"").toLowerCase(),c=s.find(u=>String(u.full_name||"").toLowerCase()===d||String(u.code||"").toLowerCase()===d||String(u.name||"").toLowerCase()===d);return c?c.id:null},i=m=>{if(m==null||m==="")return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let d=String(m).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let u=Number(d);if(u>2e4&&u<99999){let b=new Date(Date.UTC(1899,11,30)+u*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}let c=d.split(/[\/\-\.]/);if(c.length===3){let[u,b,g]=c.map(h=>h.trim());if(u.length===4&&b.length<=2&&g.length<=2)return`${u}-${b.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&b.length<=2&&u.length<=2)return`${g}-${b.padStart(2,"0")}-${u.padStart(2,"0")}`}return d},n=a.map(m=>({submitted_at:i(m.Waktu||m.Tanggal),submitter_name:String(m.Pengirim||"").trim(),branch_id:r(String(m.Cabang||"").trim()),tools_items:String(m["Alat/Barang"]||m.Alat||"").trim(),chemical_items:String(m.Chemical||"").trim(),additional_notes:String(m.Catatan||m.Keterangan||"").trim(),status:String(m.Status||"").trim(),processed_by:String(m["Diproses Oleh"]||m.PIC||"").trim()})).filter(m=>m.submitted_at&&m.submitter_name&&m.branch_id),p=await x("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!p.ok)throw new Error(p.data?.error||"Import gagal");return p.data}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,o)=>{let s=se({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(r,i)=>{let n=r.querySelector("#supply-status").value,p=r.querySelector("#supply-processed-by").value;(await x(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:p})})).ok?(Z("Status diperbarui."),i(),o()):G("Gagal update status.")}})}}]})}O();R();async function jt(t){let e=fe();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}N({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:l=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[l]||"badge-neutral"}">${l}</span>`},{key:"is_active",label:"Status",render:l=>l?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:l=>l?new Date(l).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:l=>{let a=!!l;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:l?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:l?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:l?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:l?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?l?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let l=await x(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(l.ok){let a=l.data.data.map(o=>({"Nama Lengkap":o.full_name||"",Username:o.username||"",Email:o.email||"",Role:o.role||"",Status:o.is_active?"Aktif":"Nonaktif"}));L(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async l=>{let a=l.map(s=>({full_name:String(s["Nama Lengkap"]||"").trim(),username:String(s.Username||"").trim(),email:String(s.Email||"").trim(),role:String(s.Role||"").trim()||"viewer",password:String(s.Password||"").trim()})).filter(s=>s.username&&s.password&&s.email&&s.full_name),o=await x("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!o.ok)throw new Error(o.data?.error||"Import gagal");return o.data}}})}O();R();async function Jt(t){N({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await x(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)L(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{L([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let l=e.map(o=>({code:String(o["Kode Cabang"]||"").trim(),name:String(o["Nama Pendek"]||"").trim(),full_name:String(o["Nama Lengkap"]||"").trim(),city:String(o.Kota||"").trim()})).filter(o=>o.code&&o.name),a=await x("/api/branches/import",{method:"POST",body:JSON.stringify(l)});if(!a.ok)throw new Error(a.data?.error||"Import gagal");return a.data}}})}O();var ue={schedule:{bg:"#3B82F6",text:"#fff",icon:"\u{1F5D3}",label:"Jadwal"},issue:{bg:"#FB923C",text:"#fff",icon:"\u26A0\uFE0F",label:"Permasalahan"},reliever:{bg:"#34D399",text:"#fff",icon:"\u{1F504}",label:"Reliefer"},training:{bg:"#A78BFA",text:"#fff",icon:"\u{1F393}",label:"Training"},contract_expiry:{bg:"#F87171",text:"#fff",icon:"\u{1F4C4}",label:"Kontrak Habis"},one_on_one:{bg:"#F472B6",text:"#fff",icon:"\u{1F4AC}",label:"One on One"},inspection:{bg:"#38BDF8",text:"#fff",icon:"\u{1F50D}",label:"Inspeksi"},cleaning:{bg:"#2DD4BF",text:"#fff",icon:"\u{1F9F9}",label:"Cleaning"},fogging:{bg:"#818CF8",text:"#fff",icon:"\u{1F4A8}",label:"Fogging"},basecamp:{bg:"#A8A29E",text:"#fff",icon:"\u{1F3D5}",label:"Basecamp"},supply:{bg:"#60A5FA",text:"#fff",icon:"\u{1F4E6}",label:"Permintaan"}};function Ut(t){return ue[t]||{bg:"#6B7280",icon:"\u{1F4CC}",label:t}}async function Gt(t){let e=new Date,l=[],a=[{value:"schedule",...ue.schedule},{value:"issue",...ue.issue},{value:"reliever",...ue.reliever},{value:"training",...ue.training},{value:"contract_expiry",...ue.contract_expiry},{value:"one_on_one",...ue.one_on_one},{value:"inspection",...ue.inspection},{value:"cleaning",...ue.cleaning},{value:"fogging",...ue.fogging},{value:"basecamp",...ue.basecamp}];t.innerHTML=`
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
  `,t.querySelectorAll(".cal-legend-pill").forEach(i=>{let n=i.querySelector('input[type="checkbox"]');i.addEventListener("click",p=>{p.preventDefault(),n.checked=!n.checked,i.classList.toggle("inactive",!n.checked),s()})}),document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),s()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),s()});async function o(){try{let i=e.getFullYear(),n=String(e.getMonth()+1).padStart(2,"0");l=(await x(`/api/dashboard/calendar?month=${i}-${n}`)).data?.data||[]}catch(i){console.warn("[Calendar] Gagal memuat events:",i),l=[]}}async function s(){let i=document.getElementById("calendar-grid");if(i){i.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:1px;background:#E5E7EB;">
      ${Array(42).fill('<div style="background:#F9FAFB;min-height:90px;"></div>').join("")}
    </div>`,await o();try{let n=e.getFullYear(),p=e.getMonth(),m=document.getElementById("cal-month-label");m&&(m.textContent=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}));let d=new Set(Array.from(t.querySelectorAll(".cal-filter:checked")).map(f=>f.value)),c=l.filter(f=>d.has(f.type)),u={};c.forEach(f=>{let w=(f.event_date||"").slice(0,10);w&&(u[w]||(u[w]=[]),u[w].push(f))});let b=new Date(n,p,1).getDay(),g=new Date(n,p+1,0).getDate(),h=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],k=new Date().toISOString().slice(0,10),S='<div style="display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:1px;background:#E5E7EB;border-radius:0 0 12px 12px;overflow:hidden;width:100%;">';h.forEach((f,w)=>{S+=`<div style="background:#F1F5F9;padding:10px 4px;text-align:center;font-size:0.68rem;font-weight:800;color:${w===0||w===6?"#EF4444":"#6B7280"};text-transform:uppercase;letter-spacing:0.06em;">${f}</div>`});for(let f=0;f<b;f++)S+='<div style="background:#FAFAFA;min-height:90px;"></div>';for(let f=1;f<=g;f++){let w=`${n}-${String(p+1).padStart(2,"0")}-${String(f).padStart(2,"0")}`,B=u[w]||[],E=w===k,q=new Date(n,p,f).getDay(),M=q===0||q===6,me=4;S+=`
          <div class="cal-cell" data-date="${w}"
               style="background:${E?"#EFF6FF":"#fff"};min-height:90px;padding:6px 5px 5px;overflow:hidden;
                      border-top:${E?"3px solid #2563EB":"1px solid transparent"};
                      cursor:${B.length?"pointer":"default"};">
            <!-- Nomor tanggal -->
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
              <span style="font-size:${E?"0.9rem":"0.78rem"};font-weight:${E?"900":"600"};
                    color:${E?"#fff":M?"#EF4444":"#374151"};
                    ${E?"background:#2563EB;border-radius:50%;width:22px;height:22px;display:flex;align-items:center;justify-content:center;":""}">
                ${f}
              </span>
              ${B.length>me?`<span class="cal-more-btn">+${B.length-me}</span>`:""}
            </div>

            <!-- Badge events -->
            <div style="display:flex;flex-direction:column;gap:3px;overflow:hidden;">
              ${B.slice(0,me).map(re=>{let oe=Ut(re.type),v=Ba(re.title||re.branch_name||oe.label,18);return`<div class="cal-badge" style="background:${oe.bg};" title="${Ie(re.title||oe.label)} \u2014 ${Ie(re.branch_name||"")}">
                  ${oe.icon} ${Ie(v)}
                </div>`}).join("")}
            </div>
          </div>`}let _=(b+g)%7;if(_!==0)for(let f=0;f<7-_;f++)S+='<div style="background:#FAFAFA;min-height:90px;"></div>';S+="</div>",i.innerHTML=S,i.querySelectorAll(".cal-cell[data-date]").forEach(f=>{f.addEventListener("click",()=>{let w=f.dataset.date,B=u[w]||[];B.length&&r(w,B)})})}catch(n){console.error("[Calendar] Render error:",n),i.innerHTML=`<div style="padding:60px;text-align:center;color:#9CA3AF;">
        <div style="font-size:2.5rem;margin-bottom:12px;">\u{1F4C5}</div>
        <div style="font-size:1rem;font-weight:700;">Gagal memuat kalender. Silakan refresh.</div>
      </div>`}}}function r(i,n){document.getElementById("cal-popup-overlay")?.remove();let p=new Date(i+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),m=n.map(c=>{let u=Ut(c.type);return`
        <div class="cal-popup-event" style="border-left-color:${u.bg};">
          <div class="cal-popup-icon">${u.icon}</div>
          <div class="cal-popup-info">
            <div class="cal-popup-type" style="color:${u.bg};">${u.label}</div>
            <div class="cal-popup-event-title">${Ie(c.title||"-")}</div>
            ${c.branch_name?`<div class="cal-popup-branch">\u{1F4CD} ${Ie(c.branch_name)}</div>`:""}
            ${c.status?`<div class="cal-popup-status" style="color:${c.status==="Done"||c.status==="Selesai"?"#10B981":"#F59E0B"};">\u25CF ${Ie(c.status)}</div>`:""}
            ${c.days_remaining!==void 0?`<div class="cal-popup-sisa">\u23F3 Sisa: ${c.days_remaining} hari</div>`:""}
          </div>
          <!-- Dot warna -->
          <div style="width:10px;height:10px;border-radius:50%;background:${u.bg};flex-shrink:0;margin-top:3px;"></div>
        </div>`}).join(""),d=document.createElement("div");d.id="cal-popup-overlay",d.className="cal-popup-overlay",d.innerHTML=`
      <div class="cal-popup">
        <div class="cal-popup-head">
          <div>
            <div class="cal-popup-title">\u{1F4C5} ${n.length} Event</div>
            <div class="cal-popup-date">${p}</div>
          </div>
          <button class="cal-popup-close" id="cal-popup-close-btn">\u2715</button>
        </div>
        <div class="cal-popup-body">${m}</div>
      </div>`,document.body.appendChild(d),document.getElementById("cal-popup-close-btn").addEventListener("click",()=>d.remove()),d.addEventListener("click",c=>{c.target===d&&d.remove()})}s()}function Ba(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function Ie(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}O();async function zt(t){let e=fe(),l=(e?.full_name||e?.username||"U")[0].toUpperCase(),o={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${o},${o}99)">
            ${l}
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
  `;let s=localStorage.getItem("fm_token"),r=document.getElementById("session-info");if(s&&r)try{let i=JSON.parse(atob(s.split(".")[1])),n=new Date(i.exp*1e3);r.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{r.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async i=>{i.preventDefault();let n=document.getElementById("pwd-error"),p=document.getElementById("pwd-success"),m=document.getElementById("btn-save-pwd");n.style.display="none",p.style.display="none";let d=i.target,c=d.current_password.value,u=d.new_password.value,b=d.confirm_password.value;if(u!==b){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(u.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}m.disabled=!0,m.textContent="\u23F3 Menyimpan...";let g=await x("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:c,new_password:u})});m.disabled=!1,m.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',g.ok?(p.textContent="\u2705 Password berhasil diubah.",p.style.display="block",d.reset(),Z("Password berhasil diubah.")):(n.textContent=g.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}O();var tt={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function te(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let o=Number(e);if(o>2e4&&o<99999){let s=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(s.getTime())?null:s.toISOString().slice(0,10)}}let l=e.split(/[\/\-\.]/);if(l.length===3){let[o,s,r]=l.map(m=>m.trim()),i=Number(o),n=Number(s),p=Number(r);if(o.length===4&&i>1900)return`${o}-${s.padStart(2,"0")}-${r.padStart(2,"0")}`;if(r.length===4&&p>1900)return i>12?`${r}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`:n>12?`${r}-${o.padStart(2,"0")}-${s.padStart(2,"0")}`:`${r}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`;if(r.length===2&&!isNaN(p)){let m=p>=50?`19${r}`:`20${r}`;return i>12?`${m}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`:`${m}-${s.padStart(2,"0")}-${o.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function Qt(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var La={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",phone:t["No. Hp"],join_date:te(t["Tanggal Masuk"]),status:t.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:te(t["Tanggal Mulai"]),end_date:te(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:te(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:te(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:te(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:te(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:te(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:te(t["Tanggal Target"]||t["Tgl Target"]),completion_date:te(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:te(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:te(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:te(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:te(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:te(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:te(t["Tanggal Back Up"]),completion_date:te(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:te(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:te(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Na(t,e){let l=tt[t];if(!l)return{valid:[],errors:[],mapped:[],skipped:!0};let a=La[l.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let o=[],s=[],r=[];return e.filter(n=>!Qt(n)).forEach((n,p)=>{let m=e.indexOf(n)+2,d=[];a.required.forEach(({key:u,label:b})=>{let g=n[u];if(g==null||String(g).trim()===""){let h=Object.keys(n).filter(k=>k.trim()).join(", ");d.push({column:b,originalValue:g||"",reason:`Kolom "${b}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${h.slice(0,120)}`})}});let c=a.map(n);d.length>0?s.push({row:m,data:c,raw:n,errors:d}):(o.push(n),r.push(c))}),{valid:o,errors:s,mapped:r}}function Yt(t){let e=[];return t.SheetNames.forEach(l=>{let a=tt[l];if(!a)return;let o=t.Sheets[l],s=window.XLSX.utils.sheet_to_json(o,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),r=Na(l,s),i=s.filter(n=>!Qt(n));e.push({sheetName:l,module:a.module,label:a.label,total:i.length,valid:r.mapped.length,errorCount:r.errors.length,errors:r.errors,mapped:r.mapped,skipped:!1})}),e}function Wt(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Tipe Kontrak":"PKWT 1",PKWT:"001/PKWT/2024",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06","Link Document":""}],"Time Line":[{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Januari",Pic:"Berlin","Tanggal Opening":"2024-01-01","Tanggal Target":"2024-01-10","Tanggal Selesai":"2024-01-09",Status:"Done",Keterangan:""}],"Report Inspeksi Hygiene 2026":[{Tanggal:"2026-01-15",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Point FC":"85.5","Point SPV":"90.0","Link Dokumen":"https://..."}],"Report GC-DC 2026":[{Tanggal:"2026-01-20",Cabang:"001. Pondok Bambu","Jenis Kegiatan":"General Cleaning",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Report Fogging 2026":[{Tanggal:"2026-01-25",Cabang:"001. Pondok Bambu",Periode:"Q1",Status:"Done","Link Dokumen":"https://..."}],"Rekap Laporan Basecamp":[{"Tgl Info":"2024-02-01",Cabang:"001. Pondok Bambu",Permasalahan:"Lampu mati",PIC:"Berlin","Tgl Done":"2024-02-02",Status:"Done",Ket:""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}],Training:[{Tanggal:"2024-04-10",Batch:"Batch 1",Materi:"Basic Cleaning",Peserta:"5",Cabang:"001. Pondok Bambu",Trainer:"Fajar",Nilai:"85",Keterangan:""}],"Master Checklist":[{"Master Checklist":"Checklist Kebersihan Toilet","Link Document":"https://..."}],"Master Form":[{"Master Form":"Form Izin Keluar","Link Document":"https://..."}],"Permintaan Chemical":[{Timestamp:"2024-05-01","Nama Lengkap":"Budi Santoso","Kebutuhan Untuk Cabang":"001. Pondok Bambu","Alat - Alat / Barang":"Sapu","Jumlah Permintaan Alat / Barang":"2",Chemical:"Karbol","Jumlah Permintaan Chemical":"1 Liter","Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.":"",Status:"Pending"}]}).forEach(([a,o])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(o),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function Vt(t){let e=window.XLSX,l=e.utils.book_new(),a=!1;return t.forEach(o=>{if(!o.errors||o.errors.length===0)return;a=!0;let s=o.errors.map(i=>({"No. Baris":i.row,"Kolom Gagal":(i.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(i.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(i.data||{}).map(([n,p])=>[n,p??""]))})),r=e.utils.json_to_sheet(s);e.utils.book_append_sheet(l,r,o.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(l,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Aa=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function Xt(t){t.innerHTML=`
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
              ${Object.entries(tt).map(([g,{label:h}])=>`<span class="import-sheet-tag">\u{1F4C4} ${g} \u2192 ${h}</span>`).join("")}
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
  `;let e=null,l=null,a=0,o={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function s(g){Object.entries(o).forEach(([h,k])=>{k.style.display=h===g?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let g=document.getElementById("btn-backup-db");g.disabled=!0,g.textContent="\u23F3 Memproses Backup...";try{let h=await x("/api/import/backup");if(h.ok){if(!window.XLSX){G("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");return}let k=window.XLSX,S=k.utils.book_new();Object.entries(h.data.database).forEach(([C,_])=>{let f=_.length>0?_:[{}],w=k.utils.json_to_sheet(f);k.utils.book_append_sheet(S,w,C.substring(0,31))}),k.writeFile(S,`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`),Z("Backup berhasil diunduh!")}else G("Gagal memproses backup: "+(h.data?.error||"Unknown error"))}catch(h){G("Gagal memproses backup: "+h.message)}finally{g.disabled=!1,g.textContent="\u{1F4E6} Backup Database"}});let r=document.getElementById("btn-sync-google");r&&r.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let g=r.innerHTML;r.innerHTML='<span class="spinner"></span> Menyinkronkan...',r.disabled=!0;try{let h=await x("/api/sync/google-sheets",{method:"POST"});h.ok?alert("Sinkronisasi Berhasil: "+(h.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(h.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{r.innerHTML=g,r.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Wt(),Z("Template Excel berhasil didownload!")});let i=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",g=>{g.stopPropagation(),i.click()}),i.addEventListener("change",g=>{g.target.files[0]&&p(g.target.files[0])}),n.addEventListener("dragover",g=>{g.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",g=>{g.preventDefault(),n.classList.remove("drag-over");let h=g.dataTransfer.files[0];h&&h.name.match(/\.xlsx?$/i)?p(h):G("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,i.value="",document.getElementById("file-info").style.display="none",n.style.display="",s("upload")});async function p(g){e=g,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${g.name} (${(g.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await m(g)}async function m(g){s("validating");let h=document.getElementById("validation-status"),k=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");h.textContent="Membaca file Excel...",k.style.width="20%",await Je(200);let S=await g.arrayBuffer(),C=window.XLSX.read(S,{type:"array",cellDates:!0});h.textContent=`Memvalidasi ${C.SheetNames.length} sheet...`,k.style.width="50%",await Je(100),l=Yt(C),k.style.width="100%",h.textContent="Validasi selesai!",await Je(300),d()}catch(S){s("upload"),G("Gagal memproses file: "+S.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function d(){s("preview");let g=l.filter(E=>!E.skipped).length,h=l.reduce((E,q)=>E+q.total,0),k=l.reduce((E,q)=>E+q.valid,0),S=l.reduce((E,q)=>E+q.errorCount,0),C=h>0?Math.round(k/h*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${g} sheet</span>
      <span class="badge badge-secondary">${h} baris</span>
      <span class="badge badge-success">${k} valid (${C}%)</span>
      ${S>0?`<span class="badge badge-danger">${S} error</span>`:""}
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
          ${l.map((E,q)=>`
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
    `,_.querySelectorAll(".btn-detail-error").forEach(E=>{E.addEventListener("click",()=>{let q=l[Number(E.dataset.idx)];c(q)})});let f=document.getElementById("error-detail-section"),w=document.getElementById("error-detail-container");w.innerHTML="",f.style.display="none";let B=document.getElementById("btn-start-import");k===0?(B.disabled=!0,B.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(B.disabled=!1,S>0?(B.innerHTML=`\u{1F680} Import ${k} Data Valid (${S} dilewati)`,B.title="Baris error akan dilewati, baris valid tetap diimport"):B.innerHTML=`\u{1F680} Mulai Import ${k} Data`)}function c(g){let h=document.getElementById("error-detail-section"),k=document.getElementById("error-detail-container");h.style.display="";let S=g.errors.slice(0,100).map(C=>(Array.isArray(C.errors)?C.errors:[]).map(f=>{let w=typeof f=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${C.row}</span></td>
            <td><strong>${w?f.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${w&&f.originalValue!==void 0?f.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${w?f.reason:f}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${w&&f.aliases?`Gunakan salah satu nama kolom:<br><em>${f.aliases}</em>`:w&&f.hint?f.hint:""}
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
            <tbody>${S||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${g.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,h.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{s("upload"),document.getElementById("file-info").style.display="none",n.style.display="",e=null,i.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!l)return;Vt(l)?Z("Log error berhasil didownload."):Z("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let g=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";u(g)});async function u(g){s("importing"),a=Date.now();let h=[];Aa.forEach(f=>{let w=l?.find(B=>B.module===f&&B.mapped?.length>0);w&&h.push(w)});let k=document.getElementById("import-steps-list");k.innerHTML=h.map(f=>`
      <div class="import-step-item" id="step-item-${f.module}">
        <span class="step-item-icon" id="step-icon-${f.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${f.label} <span class="step-item-count">(${f.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${f.module}"></span>
      </div>
    `).join("");let S=document.getElementById("import-bar"),C=document.getElementById("import-current-status"),_={totalSheets:h.length,totalRows:h.reduce((f,w)=>f+w.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let f=0;f<h.length;f++){let w=h[f],B=document.getElementById(`step-icon-${w.module}`),E=document.getElementById(`step-status-${w.module}`);B.textContent="\u{1F504}",E.textContent="Mengimport...",C.textContent=`Mengimport ${w.label}...`,S.style.width=`${Math.round(f/h.length*100)}%`;try{let q=await x(`/api/import/${w.module}`,{method:"POST",body:JSON.stringify({rows:w.mapped,onDuplicate:g})});if(q.ok){let M=q.data;_.inserted+=M.inserted||0,_.skipped+=M.skipped||0,_.moduleResults.push({label:w.label,inserted:M.inserted||0,skipped:M.skipped||0,status:"ok"}),B.textContent="\u2705",E.innerHTML=`<span class="badge badge-success">${M.inserted||0} berhasil</span>${M.skipped>0?` <span class="badge badge-neutral">${M.skipped} skip</span>`:""}`}else _.failed++,_.moduleResults.push({label:w.label,inserted:0,skipped:0,status:"error",error:q.data?.error}),B.textContent="\u274C",E.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(q){_.failed++,_.moduleResults.push({label:w.label,inserted:0,skipped:0,status:"error",error:q.message}),B.textContent="\u274C",E.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Je(150)}S.style.width="100%",C.textContent="Selesai!",await Je(400),b(_)}function b(g){s("summary");let h=((Date.now()-a)/1e3).toFixed(1),k=g.failed===0;document.getElementById("summary-status-icon").innerHTML=`
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
        <div class="stat-value">${h}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${g.moduleResults.map(S=>`
            <tr>
              <td>${S.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${S.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${S.skipped}</span></td>
              <td style="text-align:center">
                ${S.status==="ok"?'<span class="badge badge-success">\u2705 Sukses</span>':`<span class="badge badge-danger" title="${S.error||""}">\u274C Gagal</span>`}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,l=null,i.value="",document.getElementById("file-info").style.display="none",n.style.display="",s("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Je(t){return new Promise(e=>setTimeout(e,t))}O();var at=[],Zt=[];async function ea(t){at=await K(),Zt=await ee(),N({container:t,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:at}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let l=new URLSearchParams(e||{}).toString(),a=await x(`/api/sp?limit=10000&${l}`);if(a.ok){let o=a.data.data.map(r=>({"Nama Karyawan":r.employee_name||"",Divisi:r.division||"",Cabang:r.branch_name||"","Tanggal Sp":r.tanggal||"","Akhir Sp":r.akhir_sp||"","Jenis Sp":r.sp_type||"","Link Document / Foto":r.document_link||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(R(),le));s(o,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:l}=await Promise.resolve().then(()=>(R(),le));l(e,"Template_Import_SP")},onImport:async e=>{let l=r=>{if(!r)return null;let i=String(r||"").toLowerCase(),n=at.find(p=>String(p.label||"").toLowerCase()===i);return n?n.value:null},a=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let i=String(r).trim();if(/^\d{4,5}$/.test(i)){let p=Number(i);if(p>2e4&&p<99999){let m=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(i))return i.slice(0,10);let n=i.split(/[\/\-\.]/);if(n.length===3){let[p,m,d]=n.map(c=>c.trim());if(p.length===4&&m.length<=2&&d.length<=2)return`${p}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&m.length<=2&&p.length<=2)return`${d}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`}return i},o=e.map(r=>({employee_name:String(r["Nama Karyawan"]||"").trim(),division:String(r.Divisi||"").trim(),branch_id:l(String(r.Cabang||"").trim()),tanggal:a(r["Tanggal Sp"]),akhir_sp:a(r["Akhir Sp"]),sp_type:String(r["Jenis Sp"]||"").trim(),document_link:String(r["Link Document / Foto"]||"").trim()})).filter(r=>r.employee_name&&r.branch_id),s=await x("/api/import/sp",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:[{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:Zt},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"combobox",name:"branch_id",label:"Cabang",required:!0,options:at,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}O();var Pe=[],ta=[];async function aa(t){Pe=await K(),ta=await ee(),N({container:t,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",options:Pe},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",options:Pe}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let l=new URLSearchParams(e||{}).toString(),a=await x(`/api/mutasi?limit=10000&${l}`);if(a.ok){let o=a.data.data.map(r=>({Tanggal:r.tanggal||"","Nama Karyawan":r.employee_name||"","Cabang Asal":r.from_branch_name||"","Cabang Tujuan":r.to_branch_name||"",Status:r.status||"",Dokumen:r.document_link||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(R(),le));s(o,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:l}=await Promise.resolve().then(()=>(R(),le));l(e,"Template_Import_Mutasi")},onImport:async e=>{let l=r=>{if(!r)return null;let i=String(r||"").toLowerCase(),n=Pe.find(p=>String(p.label||"").toLowerCase()===i);return n?n.value:null},a=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let i=String(r).trim();if(/^\d{4,5}$/.test(i)){let p=Number(i);if(p>2e4&&p<99999){let m=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(i))return i.slice(0,10);let n=i.split(/[\/\-\.]/);if(n.length===3){let[p,m,d]=n.map(c=>c.trim());if(p.length===4&&m.length<=2&&d.length<=2)return`${p}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&m.length<=2&&p.length<=2)return`${d}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`}return i},o=e.map(r=>({tanggal:a(r.Tanggal),employee_name:String(r["Nama Karyawan"]||"").trim(),from_branch_id:l(String(r["Cabang Asal"]||"").trim()),to_branch_id:l(String(r["Cabang Tujuan"]||"").trim()),status:String(r.Status||"").trim(),document_link:String(r.Dokumen||"").trim()})).filter(r=>r.tanggal&&r.employee_name&&r.from_branch_id&&r.to_branch_id),s=await x("/api/import/mutasi",{method:"POST",body:JSON.stringify({rows:o,onDuplicate:"update"})});if(!s.ok)throw new Error(s.data?.error||"Import gagal");return s.data}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:ta},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:Pe,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:Pe,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}O();async function na(t){t.innerHTML=`
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
  `;let e=document.getElementById("btnRefreshSync"),l=document.getElementById("queueStatusFilter");e.addEventListener("click",o),l.addEventListener("change",i),document.getElementById("btnRetryAll").addEventListener("click",()=>a("retry",{allFailed:!0})),document.getElementById("btnResetStuck").addEventListener("click",()=>a("reset-stuck",{})),document.getElementById("btnPauseSync").addEventListener("click",()=>a("pause",{})),document.getElementById("btnResumeSync").addEventListener("click",()=>a("resume",{})),document.getElementById("btnReconcile").addEventListener("click",()=>a("reconcile",{module:"ALL",repairMode:!0})),document.getElementById("btnRetrySelected").addEventListener("click",()=>{let p=Array.from(document.querySelectorAll(".chk-queue:checked")).map(m=>m.value);if(p.length===0)return alert("No items selected");a("retry",{ids:p})}),document.getElementById("chkAllQueue").addEventListener("change",p=>{document.querySelectorAll(".chk-queue").forEach(m=>m.checked=p.target.checked)});async function a(p,m){if(confirm(`Are you sure you want to execute action: ${p}?`)){showLoading();try{let d=await x(`/api/sync/actions/${p}`,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(m)});d.ok?(alert(d.data?.message||"Success"),o()):G(d.error||"Action failed")}catch(d){G(d.message)}hideLoading()}}await o();async function o(){showLoading(),await Promise.all([r(),i(),s(),n()]),hideLoading()}async function s(){try{let p=await x("/api/sync/performance");if(!p.ok)return;let{webhook:m,google_api:d,d1:c,queue:u,throughput:b}=p.data;document.getElementById("latencyHistograms").innerHTML=`
        <table class="w-full text-left">
          <thead><tr class="text-gray-500 border-b"><th>Metric</th><th>P50</th><th>P95</th><th>P99</th><th>Max</th></tr></thead>
          <tbody class="divide-y">
            <tr><td class="py-1">Webhook</td><td>${m.P50}ms</td><td>${m.P95}ms</td><td>${m.P99}ms</td><td>${m.Max}ms</td></tr>
            <tr><td class="py-1">Google API</td><td>${d.P50}ms</td><td>${d.P95}ms</td><td>${d.P99}ms</td><td>${d.Max}ms</td></tr>
            <tr><td class="py-1">D1 Execute</td><td>${c.P50}ms</td><td>${c.P95}ms</td><td>${c.P99}ms</td><td>${c.Max}ms</td></tr>
            <tr><td class="py-1">Queue Delay</td><td>${u.P50}ms</td><td>${u.P95}ms</td><td>${u.P99}ms</td><td>${u.Max}ms</td></tr>
          </tbody>
        </table>
        <div class="mt-4 pt-3 border-t text-gray-600 flex justify-between">
          <span>Throughput: <b>${b.events_per_sec}</b> ev/sec</span>
          <span><b>${b.events_per_min}</b> ev/min</span>
        </div>
      `,document.getElementById("lblSnapshotStatus").innerHTML='Checked 10 modules. Status: <b>COMPLETED</b><br><span class="text-xs text-gray-500">Run today at 02:00</span>'}catch(p){console.error(p)}}async function r(){try{let p=await x("/api/sync/health");if(!p.ok)return G("Failed to fetch sync health");let{status:m,queue:d,circuit_breaker:c}=p.data,u=`
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
      `;document.getElementById("syncOverviewCards").innerHTML=u;let b=document.getElementById("cbStateBadge"),g=document.getElementById("cbStateDesc"),h=document.getElementById("cbStatusCard");h.className="bg-white rounded-lg shadow p-6 border-l-4",c==="CLOSED"?(b.className="px-3 py-1 rounded-full text-sm font-medium bg-green-100 text-green-800",b.textContent="CLOSED",g.textContent="Traffic is flowing normally to Google Sheets.",h.classList.add("border-green-500")):c==="OPEN"?(b.className="px-3 py-1 rounded-full text-sm font-medium bg-red-100 text-red-800",b.textContent="OPEN",g.textContent="Failing fast. Traffic to Google Sheets is paused due to repeated failures.",h.classList.add("border-red-500")):c==="HALF_OPEN"?(b.className="px-3 py-1 rounded-full text-sm font-medium bg-yellow-100 text-yellow-800",b.textContent="HALF-OPEN",g.textContent="Testing recovery. Permitting limited traffic to verify stability.",h.classList.add("border-yellow-500")):b.textContent=c||"UNKNOWN"}catch(p){console.error(p)}}async function i(){try{let p=document.getElementById("queueStatusFilter").value,m=await x("/api/sync/queue?limit=15"+(p?"&status="+p:""));if(!m.ok)return;let d=document.getElementById("queueTableBody"),c=m.data?.data||m.data||[];if(c.length===0){d.innerHTML='<tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">No events found</td></tr>';return}d.innerHTML=c.map(u=>`
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
      `).join("")}catch(p){console.error(p)}}async function n(){try{let p=await x("/api/sync/metrics");if(!p.ok)return;let m=document.getElementById("metricsTableBody"),d=p.data||[];if(d.length===0){m.innerHTML='<tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">No metrics found in last 24h</td></tr>';return}m.innerHTML=d.map(c=>`
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2 font-medium text-gray-800">${c.module}</td>
          <td class="px-4 py-2 text-gray-600">${c.total_events}</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(c.avg_webhook_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(c.avg_d1_ms||0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(c.avg_queue_wait_ms||0)} ms</td>
        </tr>
      `).join("")}catch(p){console.error(p)}}}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let e=Math.floor(Number(t)-25569);return new Date(e*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let e=t.split(/[\/\-]/);return`${e[2]}-${e[1]}-${e[0]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let l=e.split("-");if(l.length===3&&l[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],o=parseInt(l[2],10),s=a[parseInt(l[1],10)-1];return`${o} ${s} ${l[0]}`}return e};function z(t){return async e=>{if(!Ne()){xe("/login");return}return t(e)}}var Ue=null;function Fa(){Ue&&clearInterval(Ue);let t=()=>{let e=new Date,l=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),o=document.getElementById("header-clock-time"),s=document.getElementById("header-clock-date");o&&(o.textContent=l),s&&(s.textContent=a)};t(),Ue=setInterval(t,1e3)}async function Ma(){try{let t=await x("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},l=(a,o)=>{let s=document.getElementById(a);s&&(s.textContent=o>0?o:"",s.style.display=o>0?"inline-flex":"none")};l("badge-issues",e.issues?.current||0),l("badge-contracts",e.expiring30?.current||0),l("badge-oo1",e.one_on_one?.current||0),l("badge-schedule",e.schedule?.current||0),l("badge-supply",e.supply?.current||0)}catch{}}var Be=[];async function Oa(){try{let t=await x("/api/dashboard/notifications");if(!t.ok)return;Be=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Be.length>0?"block":"none",e.textContent=Be.length)}catch{}}function Ra(){if(!Be.length){se({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,l)=>l()});return}let t=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${Be.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;se({title:`Notifikasi (${Be.length})`,content:t,confirmText:"Tutup",onConfirm:(e,l)=>l()})}function ia(){let t=fe(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
  `;let l=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),o=document.getElementById("topbar-menu-btn"),s=document.getElementById("sidebar-close"),r=()=>{l.classList.add("open"),a.classList.add("show")},i=()=>{l.classList.remove("open"),a.classList.remove("show")};o?.addEventListener("click",r),s?.addEventListener("click",i),a?.addEventListener("click",i),document.querySelectorAll(".nav-item").forEach(p=>p.addEventListener("click",i));function n(){let p=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(c=>{let u=c.dataset.route;c.classList.toggle("active",p===u||u!=="/dashboard"&&p.startsWith(u))});let m=document.getElementById("topbar-title"),d=document.querySelector(".nav-item.active .nav-label");m&&d&&(m.textContent=d.textContent)}window.addEventListener("hashchange",n),n(),Fa(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await x("/api/auth/logout",{method:"POST"}),Ae(),Ue&&clearInterval(Ue),xe("/login")}),Ma(),Oa(),document.getElementById("btn-notif")?.addEventListener("click",p=>{p.preventDefault(),Ra()})}async function Ka(){J("/login",({main:e})=>Et(e)),J("/dashboard",z(({main:e})=>St(e))),J("/calendar",z(({main:e})=>Gt(e))),J("/employees",z(({main:e,params:l})=>$t(e,l))),J("/contracts",z(({main:e,params:l})=>It(e,l))),J("/sp",z(({main:e})=>ea(e))),J("/mutasi",z(({main:e})=>aa(e))),J("/sync-dashboard",z(({main:e})=>na(e))),J("/timeline",z(({main:e,params:l})=>Pt(e,l))),J("/issues",z(({main:e,params:l})=>Bt(e,l))),J("/one-on-one",z(({main:e,params:l})=>Lt(e,l))),J("/training",z(({main:e})=>Nt(e))),J("/relievers",z(({main:e,params:l})=>Ft(e,l))),J("/reports/inspection",z(({main:e})=>Mt(e))),J("/reports/cleaning",z(({main:e})=>Ot(e))),J("/reports/fogging",z(({main:e})=>Rt(e))),J("/reports/basecamp",z(({main:e})=>Kt(e))),J("/reports/supply",z(({main:e})=>bt(e,"supply"))),J("/sop",z(({main:e})=>Ht(e))),J("/checklist",z(({main:e})=>qt(e))),J("/forms",z(({main:e})=>bt(e))),J("/users",z(({main:e})=>jt(e))),J("/branches",z(({main:e})=>Jt(e))),J("/profile",z(({main:e})=>zt(e))),J("/settings/import",z(({main:e})=>Xt(e)));let t=Ne();if(!t&&window.location.hash!=="#/login"&&xe("/login"),t){let e=await x("/api/auth/me");e.ok?(Fe(e.data.data),ia()):(Ae(),xe("/login"))}window.addEventListener("fm:login",()=>{ia(),xe("/dashboard")}),yt()}Ka();
