var it=Object.defineProperty;var Je=(a,e)=>()=>(a&&(e=a(a=0)),e);var Ge=(a,e)=>{for(var i in e)it(a,i,{get:e[i],enumerable:!0})};var de={};Ge(de,{API:()=>ga,CLIENT_SIDE_MAX_ROWS:()=>_e,IS_DEVELOPMENT:()=>Qe,apiFetch:()=>f,clearToken:()=>Ce,getToken:()=>xe,getUser:()=>oe,setToken:()=>Ve,setUser:()=>Te});function xe(){return localStorage.getItem("fm_token")}function Ve(a){localStorage.setItem("fm_token",a)}function Ce(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function oe(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Te(a){localStorage.setItem("fm_user",JSON.stringify(a))}async function f(a,e={}){let i=xe(),t={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...e.headers||{}};try{let o=`cb=${Date.now()}`,l=a.includes("?")?"&":"?",r=`${ga}${a}${l}${o}`,s=await fetch(r,{...e,headers:t}),n;try{let c=await s.text();try{n=JSON.parse(c)}catch{n={error:`Server Error (${s.status}): ${c.substring(0,80)}...`}}}catch{n={error:"Gagal membaca respon dari server"}}return s.status===401&&(Ce(),window.location.hash="#/login"),{ok:s.ok,status:s.status,data:n}}catch(o){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${o.message})`}}}}var Qe,rt,ga,_e,D=Je(()=>{Qe=!1,rt="https://fm-operations-api.facilitycare-audydental.workers.dev",ga=rt,_e=1e4});var ya={};Ge(ya,{confirmDialog:()=>We,createModal:()=>te});function te({title:a,content:e,onConfirm:i,onCancel:t,confirmText:o="Simpan",cancelText:l="Batal",size:r="md",confirmClass:s="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},c=document.createElement("div");c.className="modal-overlay",c.innerHTML=`
    <div class="modal" style="max-width:${n[r]||n.md}">
      <div class="modal-header">
        <h3 class="modal-title">${a}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${l}</button>
        ${i?`<button class="btn ${s} modal-confirm">${o}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&c.querySelector(".modal-body").appendChild(e);let m=()=>{c.classList.remove("show"),setTimeout(()=>c.remove(),250)};return c.querySelector(".modal-close").addEventListener("click",()=>{t&&t(),m()}),c.querySelector(".modal-cancel").addEventListener("click",()=>{t&&t(),m()}),i&&c.querySelector(".modal-confirm").addEventListener("click",()=>i(c,m)),c.addEventListener("click",p=>{p.target===c&&(t&&t(),m())}),document.body.appendChild(c),requestAnimationFrame(()=>c.classList.add("show")),{overlay:c,close:m}}function We(a,e,i="Konfirmasi"){return te({title:i,content:`<p>${a}</p>`,onConfirm:(t,o)=>{e(),o()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var Ee=Je(()=>{});var V={};Ge(V,{downloadExcel:()=>E,parseExcel:()=>Xe,renderExcelButtons:()=>Ze});function Xe(a){return new Promise((e,i)=>{let t=new FileReader;t.onload=o=>{try{let l=new Uint8Array(o.target.result),r=XLSX.read(l,{type:"array"}),s=r.SheetNames[0],n=r.Sheets[s];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${a.name}`),console.log(`File Size: ${(a.size/1024).toFixed(2)} KB`),console.log(`File Type: ${a.type||"unknown"}`),console.log(`Sheets Found: ${r.SheetNames.join(", ")}`),console.log(`Sheet Used: ${s}`);let c=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),m=c.e.r-c.s.r+1,p=c.e.c-c.s.c+1;console.log(`Total Rows (including empty): ${m}`),console.log(`Total Columns: ${p}`);let d=[];for(let y=c.s.c;y<=c.e.c;++y){let g=n[XLSX.utils.encode_cell({c:y,r:c.s.r})];g&&g.v&&d.push(g.v)}console.log(`Headers Found: ${d.join(", ")}`),console.log("---------------------------");let b=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(b,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(b,"__headers",{value:d,enumerable:!1}),e(b)}catch(l){i(l)}},t.onerror=o=>i(o),t.readAsArrayBuffer(a)})}function E(a,e){try{let i=XLSX.utils.json_to_sheet(a),t=XLSX.utils.book_new();XLSX.utils.book_append_sheet(t,i,"Data"),XLSX.writeFile(t,`${e}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function Ze(a){return`
    <div class="excel-actions" style="display:flex;gap:0.5rem;margin-bottom:1rem;">
      <button class="btn btn-secondary btn-sm" id="btn-export-${a}">
        \u{1F4E5} Export Excel
      </button>
      <button class="btn btn-secondary btn-sm" id="btn-template-${a}">
        \u{1F4C4} Download Template
      </button>
      <label class="btn btn-primary btn-sm" style="cursor:pointer;margin:0;" id="label-import-${a}">
        <span class="import-text">\u{1F4E4} Import Excel</span>
        <input type="file" id="input-import-${a}" accept=".xlsx, .xls, .csv" style="display:none;">
      </label>
    </div>
  `}var B=Je(()=>{});D();var Ye={},Ne=null;function O(a,e){Ye[a]=e}function pe(a){window.location.hash=a}function ba(){async function a(){let e=window.location.hash.replace("#","")||"/dashboard",[i,...t]=e.split("?"),o=Ye[i];if(!o){for(let[r,s]of Object.entries(Ye))if(r.endsWith("/*")&&i.startsWith(r.slice(0,-2))){o=s;break}}Ne&&(Ne(),Ne=null);let l=document.getElementById("main-content");if(l&&(l.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),o){let r=new URLSearchParams(t.join("?")),s=i.split("/").filter(Boolean),n=await o({path:i,params:r,segments:s,main:l});n&&(Ne=n)}else{let r=l||document.getElementById("app");r&&(r.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",a),a()}var $e;function st(){return $e||($e=document.createElement("div"),$e.id="toast-container",document.body.appendChild($e)),$e}function ha(a,e="info",i=3500){let t=st(),o=document.createElement("div");o.className=`toast toast-${e}`;let l={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};o.innerHTML=`<span class="toast-icon">${l[e]||"\u2139"}</span><span class="toast-msg">${a}</span>`,t.appendChild(o),requestAnimationFrame(()=>o.classList.add("show")),setTimeout(()=>{o.classList.remove("show"),setTimeout(()=>o.remove(),350)},i)}var J=a=>ha(a,"success"),Q=a=>ha(a,"error");Ee();D();D();function fa({columns:a,data:e,onEdit:i,onDelete:t,onView:o,actions:l=[],emptyText:r="Tidak ada data",bulkSelect:s=null}){let n=document.createElement("div");if(n.className="table-wrapper",!e||e.length===0)return n.innerHTML=`<div class="empty-state"><p>${r}</p></div>`,n;let c=document.createElement("table");c.className="data-table";let m=document.createElement("thead"),p=document.createElement("tr");if(s){let b=document.createElement("th");b.style.width="40px",b.style.textAlign="center";let y=document.createElement("input");y.type="checkbox",y.id="select-all-checkbox",y.title="Pilih semua",y.addEventListener("change",()=>{e.forEach(g=>{y.checked?s.selectedIds.add(g.id):s.selectedIds.delete(g.id)}),n.querySelectorAll(".row-checkbox").forEach(g=>g.checked=y.checked),s.onToggle()}),b.appendChild(y),p.appendChild(b)}if(a.forEach(b=>{let y=document.createElement("th");y.textContent=b.label,b.width&&(y.style.width=b.width),p.appendChild(y)}),i||t||o||l.length>0){let b=document.createElement("th");b.textContent="Aksi",b.style.width="120px",p.appendChild(b)}m.appendChild(p),c.appendChild(m);let d=document.createElement("tbody");return e.forEach(b=>{let y=document.createElement("tr");if(s){let g=document.createElement("td");g.style.textAlign="center",g.style.width="40px";let u=document.createElement("input");u.type="checkbox",u.className="row-checkbox",u.checked=s.selectedIds.has(b.id),u.addEventListener("change",()=>{if(u.checked)s.selectedIds.add(b.id);else{s.selectedIds.delete(b.id);let h=document.getElementById("select-all-checkbox");h&&(h.checked=!1)}s.onToggle()}),g.appendChild(u),y.appendChild(g)}if(a.forEach(g=>{let u=document.createElement("td");if(g.render){let h=g.render(b[g.key],b);h instanceof HTMLElement?u.appendChild(h):u.innerHTML=h||""}else u.textContent=b[g.key]!==null&&b[g.key]!==void 0&&b[g.key]!==""?b[g.key]:"";g.nowrap&&(u.style.whiteSpace="nowrap"),y.appendChild(u)}),i||t||o||l.length>0){let g=document.createElement("td");g.className="actions-cell";let u=document.createElement("div");if(u.className="btn-group",o){let h=document.createElement("button");h.className="btn btn-xs btn-ghost",h.innerHTML="\u{1F441}",h.title="Lihat",h.addEventListener("click",()=>o(b)),u.appendChild(h)}if(i){let h=document.createElement("button");h.className="btn btn-xs btn-secondary",h.innerHTML="\u270F\uFE0F",h.title="Edit",h.addEventListener("click",()=>i(b)),u.appendChild(h)}l.forEach(h=>{let S=document.createElement("button");S.className=`btn btn-xs ${h.class||"btn-ghost"}`,S.innerHTML=h.icon||h.label,S.title=h.label,S.addEventListener("click",()=>h.handler(b)),u.appendChild(S)}),g.appendChild(u),y.appendChild(g)}d.appendChild(y)}),c.appendChild(d),n.appendChild(c),n}function va({page:a,pages:e,total:i,limit:t,onPage:o}){if(e<=1)return null;let l=document.createElement("div");l.className="pagination";let r=document.createElement("span");r.className="pagination-info",r.textContent=`Total: ${i} data`,l.appendChild(r);let s=document.createElement("div");s.className="pagination-btns";let n=(p,d,b=!1,y=!1)=>{let g=document.createElement("button");g.className=`btn btn-sm ${y?"btn-primary":"btn-ghost"} pagination-btn`,g.textContent=p,g.disabled=b,g.addEventListener("click",()=>o(d)),s.appendChild(g)};n("\xAB",1,a===1),n("\u2039",a-1,a===1);let c=Math.max(1,a-2),m=Math.min(e,a+2);for(let p=c;p<=m;p++)n(p,p,!1,p===a);return n("\u203A",a+1,a===e),n("\xBB",e,a===e),l.appendChild(s),l}Ee();function ze(a){return a.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${ze(e.fields)}</div>`;let i=e.required?"required":"",t=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",o="";switch(e.type){case"textarea":o=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${i} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let r=(e.options||[]).map(p=>{let d=typeof p=="object"?p.value:p,b=typeof p=="object"?p.label:p,y=e.value==d?"selected":"";return`<option value="${d}" ${y}>${b}</option>`}).join("");o=`<select name="${e.name}" class="form-control" ${i}><option value="">-- Pilih ${e.label||""} --</option>${r}</select>`;break;case"combobox":let s=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,n=(e.options||[]).map(p=>{let d=typeof p=="object"?p.value:p,b=typeof p=="object"?p.label||p.value||"":p||"";return(b==="undefined"||b==="[object Object]"||b==="null")&&(b=""),b?`<option value="${b}"></option>`:""}).join(""),c=e.value||"";if(e.value){let p=(e.options||[]).find(d=>(typeof d=="object"?d.value:d)==e.value);if(p){let d=typeof p=="object"?p.label||p.value||"":p||"";d&&d!=="undefined"&&d!=="[object Object]"&&d!=="null"&&(c=d)}}o=`
          <input type="text" name="${e.name}" list="${s}" class="form-control" value="${c}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${s}">${n}</datalist>
        `;break;case"checkbox":o=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let m=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";o=`<input type="date" name="${e.name}" class="form-control" value="${m}" ${i}>`;break;case"number":o=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${i}>`;break;case"email":o=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i}>`;break;case"url":o=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${i}>`;break;default:o=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i} autocomplete="off">`}let l=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${t}${o}${l}</div>`}).join("")}function ka(a){let e={},i=new FormData(a);for(let[t,o]of i.entries())e[t]=o===""?null:o;return a.querySelectorAll("input[type=checkbox]").forEach(t=>{t.checked||(e[t.name]=null)}),e}function Sa(a,e){e&&Object.entries(e).forEach(([i,t])=>{let o=a.querySelector(`[name="${i}"]`);o&&(o.hasAttribute("list")||(o.type==="checkbox"?o.checked=!!t:o.type==="date"&&t&&window.parseFlexibleDate?o.value=window.parseFlexibleDate(t):o.value=t??""))})}B();function $({container:a,title:e,icon:i,apiPath:t,columns:o,formFields:l,filterFields:r,defaultFilters:s={},itemLabel:n="Data",canCreate:c=!0,canEdit:m=!0,canDelete:p=!0,onBeforeSubmit:d,onAfterLoad:b,onDataLoaded:y,extraActions:g=[],initialSearch:u="",exportOptions:h=null,bulkDelete:S=!1,paginationMode:w="server"}){let v=1,k={...s};u&&(k.search=u);let T=new Set;a.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions">
        ${c?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
      </div>
    </div>

    ${S?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${h?Ze(h.moduleName):""}

    ${r&&r.length>0?`
    <div class="filter-bar card">
      <div class="filter-bar-inner">
        ${r.map(x=>x.type==="search"?`<div class="filter-search"><input type="search" class="form-control" placeholder="${x.placeholder||"Cari..."}" id="filter-search" value="${k.search||""}"></div>`:x.type==="select"?`<select class="form-control filter-select" name="${x.name}" id="filter-${x.name}"><option value="">-- ${x.label} --</option>${(x.options||[]).map(_=>`<option value="${typeof _=="object"?_.value:_}" ${k[x.name]===(typeof _=="object"?_.value:_)?"selected":""}>${typeof _=="object"?_.label:_}</option>`).join("")}</select>`:"").join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
      </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function L(){if(!document.getElementById("bulk-toolbar"))return;let _=document.getElementById("bulk-count"),A=document.getElementById("btn-bulk-delete"),j=document.getElementById("btn-bulk-cancel");_.textContent=`${T.size} item dipilih`,T.size>0?(A.disabled=!1,j.disabled=!1):(A.disabled=!0,j.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{T.clear(),document.querySelectorAll(".row-checkbox").forEach(_=>_.checked=!1);let x=document.getElementById("select-all-checkbox");x&&(x.checked=!1),L()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(T.size===0)return;let x=[...T],_=document.createElement("div");_.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",_.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${x.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${x.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(_),_.querySelector("#bulk-cancel-btn").addEventListener("click",()=>_.remove()),_.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let A=_.querySelector("#bulk-confirm-btn");A.disabled=!0,A.textContent="Menghapus...";let j=await f(`${t}/bulk`,{method:"DELETE",body:JSON.stringify({ids:x})});_.remove(),j.ok?(J(`${x.length} ${n} berhasil dihapus.`),T.clear(),L(),F()):Q(j.data?.error||"Gagal menghapus data.")})});let C=document.getElementById("filter-search"),P;if(C?.addEventListener("input",x=>{clearTimeout(P),P=setTimeout(()=>{k.search=x.target.value,v=1,F()},400)}),r?.forEach(x=>{x.type==="select"&&document.getElementById(`filter-${x.name}`)?.addEventListener("change",_=>{k[x.name]=_.target.value,v=1,F()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{k={...s},C&&(C.value=""),r?.forEach(x=>{let _=document.getElementById(`filter-${x.name}`);_&&(_.value="")}),v=1,F()}),document.getElementById("btn-create")?.addEventListener("click",()=>Se(null)),h){document.getElementById(`btn-export-${h.moduleName}`)?.addEventListener("click",async _=>{let A=_.target,j=A.innerHTML;A.innerHTML="\u23F3 Loading...",A.disabled=!0;try{await h.onExport()}catch{Q("Gagal export data")}finally{A.innerHTML=j,A.disabled=!1}}),document.getElementById(`btn-template-${h.moduleName}`)?.addEventListener("click",()=>{h.onTemplate()});let x=document.getElementById(`input-import-${h.moduleName}`);x?.addEventListener("change",async _=>{let A=_.target.files[0];if(!A)return;let j=document.getElementById(`label-import-${h.moduleName}`),Y=j?j.querySelector(".import-text"):null,ae=Y?Y.innerText:"";Y&&(Y.innerText="\u231B Memproses..."),j&&(j.style.pointerEvents="none"),x.disabled=!0;try{let M=await Xe(A);if(M.length===0)throw new Error("File kosong atau format salah");await h.onImport(M),J("Import berhasil!"),F()}catch(M){Q(M.message||"Gagal import data")}finally{Y&&(Y.innerText=ae),j&&(j.style.pointerEvents="auto"),x.disabled=!1,x.value=""}})}async function F(){T.clear(),L();let x=document.getElementById("table-container");if(!x)return;x.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let _=w==="client",A=_?1:v,j=_?_e:20,Y=new URLSearchParams({page:A,limit:j,...Object.fromEntries(Object.entries(k).filter(([,G])=>G))}),ae=await f(`${t}?${Y}`);if(!ae.ok){x.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${ae.data?.error||"Error"}</p></div>`;return}let M=ae.data?.data||ae.data||[],R=ae.data?.pagination,Ue=M.length;if(_){M=y(M);let G=M.length,W=20,ie=Math.ceil(G/W);v>ie&&ie>0&&(v=ie);let U=(v-1)*W,ce=v*W;M=M.slice(U,ce),R={page:v,limit:W,total:G,pages:ie}}!1,b&&b(M);let Le=fa({columns:o,data:M,onEdit:m?G=>Se(G):null,actions:g.map(G=>({...G,handler:W=>G.handler(W,F)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:S?{selectedIds:T,onToggle:L}:null});x.innerHTML="",x.appendChild(Le);let we=document.getElementById("pagination-container");if(we&&(we.innerHTML="",R&&R.pages>1)){let G=va({page:R.page,pages:R.pages,total:R.total,limit:R.limit,onPage:W=>{v=W,F()}});G&&we.appendChild(G)}}function ke(x){let _=typeof l=="function"?l(x):l;return ze(_)}function Se(x){let _=!!x,A=document.createElement("form");if(A.noValidate=!0,A.innerHTML=ke(x),_){let Y=typeof l=="function"?l(x):l;Sa(A,x)}let{close:j}=te({title:_?`Edit ${n}`:`Tambah ${n}`,content:A,size:"lg",confirmText:_?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(Y,ae)=>{if(!A.reportValidity())return;let M=Y.querySelector(".modal-confirm");M.disabled=!0,M.textContent="Menyimpan...";let R=ka(A),Ue=typeof l=="function"?l(x):l,Le=async ie=>{for(let U of ie)if(U.type==="row")await Le(U.fields);else if(U.type==="combobox"&&R[U.name]){let ce=R[U.name],Ae=(U.options||[]).find(ee=>{let re=String(typeof ee=="object"?ee.value:ee),nt=String(typeof ee=="object"?ee.label:ee);return re===ce||nt===ce});if(Ae)R[U.name]=typeof Ae=="object"?Ae.value:Ae;else if(U.createApi){let ee={};ee[U.createApi.field]=ce,U.createApi.extra&&Object.assign(ee,U.createApi.extra);let re=await f(U.createApi.path,{method:"POST",body:JSON.stringify(ee)});if(re.ok&&re.data?.id)R[U.name]=re.data.id;else if(re.ok&&!re.data?.id)R[U.name]=ce;else throw new Error(`Gagal membuat master data: ${re.data?.error||"Unknown error"}`)}}};try{await Le(Ue)}catch(ie){Q(ie.message),M.disabled=!1,M.textContent=_?"Simpan Perubahan":`Tambah ${n}`;return}d&&(R=await d(R,x));let we=_?"PUT":"POST",G=_?`${t}/${x.id}`:t,W=await f(G,{method:we,body:JSON.stringify(R)});W.ok?(J(_?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),ae(),F()):(Q(W.data?.error||"Gagal menyimpan data."),M.disabled=!1,M.textContent=_?"Simpan Perubahan":`Tambah ${n}`)}})}function je(x){We(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let _=await f(`${t}/${x.id}`,{method:"DELETE"});_.ok?(J(`${n} berhasil dihapus.`),F()):Q(_.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return F(),F}D();D();var me=null,Oe=null;async function ue(a=!1){if(me&&!a)return console.log("Employees Raw (Cache Hit)",me.slice(0,5)),me;let e=await f(`/api/employees?limit=${_e}&status=Aktif`);return me=(e.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",me.slice(0,5)),me}async function q(a=!1){let i=(await ue(a)).map(t=>({value:t.label,label:t.label}));return console.log("Employee Options",i.slice(0,5)),i}async function I(a=!1){return Oe&&!a||(Oe=((await f("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),Oe}function N(a){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!a||a==="-"||String(a).trim()===""?"":`<span class="badge ${e[a]||"badge-neutral"}">${a}</span>`}function ea(a){return a==null?'<span class="badge badge-neutral">-</span>':a<0?`<span class="badge badge-danger">Expired (${Math.abs(a)}h)</span>`:a<=14?`<span class="badge badge-danger">${a} hari</span>`:a<=30?`<span class="badge badge-warning">${a} hari</span>`:`<span class="badge badge-success">${a} hari</span>`}function ge(a){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[a]||"badge-neutral"}">${a||"-"}</span>`}function aa(a){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[a]||"badge-neutral"}">${a||"-"}</span>`}function z(a){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[a]||"badge-neutral"}">${a||"-"}</span>`}B();function ta(a,e){if(a.period!=="Q3")return!1;let i=String(a.status||"").toLowerCase();if(i!=="selesai"&&i!=="completed"&&i!=="done")return!1;let t=String(a.activity_type||"").toLowerCase();return e==="inspeksi"?t.includes("inspeksi"):e==="gcdc"?t.includes("general cleaning")||t.includes("deep cleaning"):!1}D();B();function wa(a,e){let i=String(a.status||"").toLowerCase();return e==="active"?i==="aktif":!1}D();B();function na(a,e){if(String(a.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!a.end_date)return!1;let t=new Date;t.setHours(0,0,0,0);let o=new Date(t);o.setDate(t.getDate()+30);let l=new Date(a.end_date);return l.setHours(0,0,0,0),l>=t&&l<=o}return!1}D();B();function _a(a,e){let i=String(a.status||"").toLowerCase();return e==="open"?i==="open":!1}D();function xa(a,e){let i=String(a.status||"").toLowerCase();return e==="pending"?i==="pending":!1}var se={};function De(a){if(se[a]){try{se[a].destroy()}catch{}delete se[a]}}function ot(){Object.keys(se).forEach(De)}var ne=(a,e=0)=>{let i=Number(a);return isNaN(i)||a===null||a===void 0?e:i},be=(a,e="\u2014")=>{if(a==null||a==="")return e;let i=String(a).trim();return i===""||i==="[object Object]"?e:i};var lt=a=>{if(!a||typeof a!="string")return"";try{let[e,i]=a.split("-");return new Date(Number(e),Number(i)-1).toLocaleDateString("id-ID",{month:"short",year:"2-digit"})}catch{return a}};function Ca(a,e,i=900){if(!a)return;let t=Math.max(0,Math.round(ne(e)));if(t===0){a.textContent="0";return}let o=Date.now(),l=()=>{let r=Math.min((Date.now()-o)/i,1),s=1-Math.pow(1-r,3);a.textContent=Math.round(s*t).toLocaleString("id-ID"),r<1?requestAnimationFrame(l):a.textContent=t.toLocaleString("id-ID")};requestAnimationFrame(l)}var ct={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},dt=a=>{let e=be(a,"\u2014");return`<span class="status-pill ${ct[e]||"pill-neutral"}">${e}</span>`};var Z={family:"Inter",size:11},he="#94A3B8",Fe="#F1F5F9",ia=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],pt=()=>window.innerWidth<768;function oa(a={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:pt()?"bottom":"top",labels:{font:Z,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:Z,titleFont:{...Z,weight:"700"}}},scales:{x:{grid:{color:Fe},ticks:{font:Z,color:he,maxRotation:0}},y:{grid:{color:Fe},ticks:{font:Z,color:he},beginAtZero:!0}},...a}}var mt=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join(""),ut=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function ra(a=3){return Array(a).fill(0).map((e,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<a-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function X(a,e,i=8e3){try{let t=new AbortController,o=setTimeout(()=>t.abort(),i),l=await f(a,{signal:t.signal}).catch(()=>null);if(clearTimeout(o),!l||!l.ok)return e;let r=l.data;return r?r.data!==void 0?r.data??e:r:e}catch{return e}}function gt(){["skel-donut","skel-trend","skel-insp","skel-contract"].forEach(t=>{let o=document.getElementById(t);o&&(o.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract"].forEach(t=>{let o=document.getElementById(t);if(o&&o.style.display==="none"){o.style.display="block";let l=o.parentElement;if(l&&!l.querySelector(".chart-empty")){let r=document.createElement("div");r.className="chart-empty",r.textContent="Belum ada data",o.style.display="none",l.appendChild(r)}}});let a=document.getElementById("kpi-row");a&&a.querySelector(".skeleton")&&$a({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Ea({}),["table-contracts","table-issues"].forEach(t=>{let o=document.getElementById(t);o&&o.querySelector(".skeleton")&&(o.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Ta(a){ot(),a._dashRefresh&&clearInterval(a._dashRefresh),a._skelTimeout&&clearTimeout(a._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});a.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">


      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${mt()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${ut()}</div>

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
          <div class="chart-card-header" style="flex-wrap: wrap; gap: 8px;">
            <div class="chart-card-title">Jadwal Hari Ini <span style="font-size:0.75rem; font-weight:normal; color:var(--text-3); margin-left:6px">${e}</span></div>
            <a href="#/calendar" class="chart-link">Lihat Kalender</a>
          </div>
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${ra(3)}</div>
        </div>
        <!-- KPI Kebersihan -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">KPI Kebersihan</div>
          </div>
          <div id="widget-kpi-kebersihan" style="margin-top:0px">${ra(4)}</div>
        </div>
        <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${ra(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>sa(a)),a._skelTimeout=setTimeout(()=>gt(),5e3),await sa(a),a._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?sa(a):clearInterval(a._dashRefresh)},6e4)}async function sa(a){a._skelTimeout&&(clearTimeout(a._skelTimeout),a._skelTimeout=null);let[e,i,t,o,l,r,s,n,c,m,p]=await Promise.all([X("/api/dashboard/kpi",{},8e3),X("/api/dashboard/issues-trend",{},8e3),X("/api/dashboard/issues-summary",{},8e3),X("/api/dashboard/inspection-bar",{},8e3),X("/api/dashboard/stats",{},8e3),X("/api/dashboard/calendar",[],8e3),X("/api/schedule?limit=10000",{data:[]},8e3),X("/api/employees?limit=10000",{data:[]},8e3),X("/api/contracts?limit=10000",{data:[]},8e3),X("/api/issues?limit=10000",{data:[]},8e3),X("/api/one_on_one?limit=10000",{data:[]},8e3)]);if(e){let d=Array.isArray(s?.data)?s.data:Array.isArray(s)?s:[],b=Array.isArray(n?.data)?n.data:Array.isArray(n)?n:[],y=Array.isArray(c?.data)?c.data:Array.isArray(c)?c:[],g=Array.isArray(m?.data)?m.data:Array.isArray(m)?m:[],u=Array.isArray(p?.data)?p.data:Array.isArray(p)?p:[];e.employees&&(e.employees.current=b.filter(h=>wa(h,"active")).length),e.contracts&&(e.contracts.current=y.filter(h=>na(h,"active")).length),e.expiring30&&(e.expiring30={current:y.filter(h=>na(h,"expiring30")).length}),e.issues&&(e.issues.current=g.filter(h=>_a(h,"open")).length),e.one_on_one&&(e.one_on_one.current=u.filter(h=>xa(h,"pending")).length),e.inspection_month&&(e.inspection_month.current=d.filter(h=>ta(h,"inspeksi")).length),e.cleaning_month&&(e.cleaning_month.current=d.filter(h=>ta(h,"gcdc")).length)}try{$a(e)}catch(d){console.warn("KPI render:",d)}try{Ea(e)}catch(d){console.warn("MiniStats render:",d)}try{bt(Array.isArray(t?.by_category)?t.by_category:[])}catch(d){console.warn("Donut render:",d),le("skel-donut","chart-donut")}try{ht(i)}catch(d){console.warn("Trend render:",d),le("skel-trend","chart-trend")}try{yt(o)}catch(d){console.warn("InspBar render:",d),le("skel-insp","chart-insp")}try{let d=Array.isArray(l)?l:Array.isArray(l?.recent_issues)?l.recent_issues:[];vt(d)}catch(d){console.warn("IssuesTable render:",d)}try{let d=Array.isArray(l?.expiring_contracts)?l.expiring_contracts:[];ft()}catch(d){console.warn("ContractsTable render:",d)}try{kt(Array.isArray(r)?r:[])}catch(d){console.warn("Agenda render:",d)}try{St(e)}catch(d){console.warn("KPI Kebersihan render:",d)}try{wt()}catch(d){console.warn("Quick Actions render:",d)}}function $a(a){let e=document.getElementById("kpi-row");if(!e)return;a=a||{};let i=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=i.map(t=>{let o=ne(a[t.key]?.current,0);return`
      <a href="${t.href}" class="kpi-card ${t.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${t.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${o}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${o}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${t.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${t.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(t=>{Ca(t,parseInt(t.dataset.target)||0)})}function Ea(a){let e=document.getElementById("mini-stats-row");if(!e)return;a=a||{};let i=[{icon:"\u{1F4C5}",label:"Jadwal",val:a.schedule?.current,href:"#/timeline",color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:a.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F504}",label:"Reliefer",val:a.reliever_total?.current,href:"#/relievers?dash_filter=reliever",color:"mini-teal"},{icon:"\u{1F50D}",label:"Inspeksi",val:a.inspection_month?.current,href:"#/timeline?dash_filter=inspeksi",color:"mini-blue"},{icon:"\u{1F9F9}",label:"GCDC",val:a.cleaning_month?.current,href:"#/timeline?dash_filter=gcdc",color:"mini-green"},{icon:"\u{1F4A8}",label:"Fogging",val:a.fogging_month?.current,href:"#/reports/fogging",color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:a.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=i.map(t=>`
    <a href="${t.href}" class="mini-stat ${t.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${t.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${ne(t.val)}">0</div>
        <div class="mini-stat-text">${t.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(t=>Ca(t,parseInt(t.dataset.target)||0,700))}function bt(a){le("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!e||!i)return;De("donut");let t=(a||[]).filter(n=>ne(n.count)>0);if(!t.length){la(e,"Belum ada data permasalahan");return}let o=t.map(n=>`${be(n.category,"Lainnya")}`),l=t.map(n=>ne(n.count)),r=l.reduce((n,c)=>n+c,0);i.innerHTML=t.map((n,c)=>{let m=ia[c%ia.length],p=r>0?Math.round(n.count/r*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${m}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${p}%)</span></div>
          <div class="donut-legend-label">${o[c]}</div>
        </div>
      </div>
    `}).join("");let s={id:"centerText",beforeDraw:function(n){let c=n.width,m=n.height,p=n.ctx;p.restore();let d=(m/80).toFixed(2);p.font="bold "+d+"em Inter",p.textBaseline="middle",p.fillStyle="#1E293B";let b=r.toString(),y=Math.round((c-p.measureText(b).width)/2),g=m/2;p.fillText(b,y,g-10),p.font="600 "+(d*.35).toFixed(2)+"em Inter",p.fillStyle="#64748B";let u="Total",h=Math.round((c-p.measureText(u).width)/2);p.fillText(u,h,g+15),p.save()}};se.donut=new Chart(e,{type:"doughnut",data:{labels:o,datasets:[{data:l,backgroundColor:ia,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:Z,titleFont:{...Z,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutout:"75%"},plugins:[s]})}function ht(a){le("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;De("trend"),a=a||{};let i=(a.labels||[]).map(lt),t=(a.open||[]).map(l=>ne(l)),o=(a.closed||[]).map(l=>ne(l));if(!i.length){la(e,"Belum ada data trend");return}se.trend=new Chart(e,{type:"line",data:{labels:i,datasets:[{label:"Open",data:t,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:o,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#10B981",borderWidth:2}]},options:oa({plugins:{legend:{display:!1}}})})}function yt(a){le("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;De("inspBar"),a=a||{};let i=a.labels||[],t=(a.fc||[]).map(l=>ne(l)),o=(a.spv||[]).map(l=>ne(l));if(!i.length){la(e,"Belum ada data inspeksi");return}se.inspBar=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Skor FC",data:t,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:o,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:oa({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:Z,color:he,maxRotation:45,minRotation:30}},y:{grid:{color:Fe},ticks:{font:Z,color:he},min:0,max:100}}})})}function ft(){le("skel-contract-mini","chart-contract-mini");let a=document.getElementById("chart-contract-mini");if(!a)return;De("contractMiniBar");let e=["Jun","Jul","Agu","Sep","Okt","Nov","Des"],i=[12,18,9,24,15,30,42],o=a.getContext("2d").createLinearGradient(0,0,0,200);o.addColorStop(0,"#60A5FA"),o.addColorStop(1,"#2563EB"),se.contractMiniBar=new Chart(a,{type:"bar",data:{labels:e,datasets:[{label:"Kontrak Habis",data:i,backgroundColor:o,borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:oa({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:Z,color:he,maxRotation:0}},y:{grid:{color:Fe,borderDash:[4,4],drawBorder:!1},ticks:{font:Z,color:he,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function vt(a){let e=document.getElementById("table-issues");if(!e)return;let i=(a||[]).slice(0,8);if(!i.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${i.map(t=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${dt(t.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${be(t.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${be(t.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function kt(a){let e=document.getElementById("widget-agenda");if(!e)return;let i=new Date,t=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,l=(a||[]).filter(r=>(r.event_date||"").startsWith(t)).slice(0,10);if(!l.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada agenda hari ini</div>';return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${l.map(r=>{let s="#3B82F6",n="#EFF6FF",c="Agenda",m=(r.title||"").toLowerCase();return m.includes("inspeksi")?(s="#10B981",n="#ECFDF5",c="Inspeksi"):m.includes("cleaning")||m.includes("gcdc")?(s="#3B82F6",n="#EFF6FF",c="Cleaning"):m.includes("reliefer")?(s="#F59E0B",n="#FFFBEB",c="Reliefer"):m.includes("fogging")&&(s="#8B5CF6",n="#F5F3FF",c="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(r.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${s};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${be(r.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${be(r.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${s}">${c}</div>
        </div>
      `}).join("")}
    </div>
  `}function St(a){let e=document.getElementById("widget-kpi-kebersihan");if(!e)return;let i=[{label:"Kebersihan Area",val:"97%",target:"Target 95%",icon:"\u{1F9F9}",bg:"#ECFDF5",color:"#10B981"},{label:"Penyelesaian Complaint",val:"100%",target:"Target 100%",icon:"\u23F1\uFE0F",bg:"#ECFDF5",color:"#10B981"},{label:"Kepatuhan Jadwal Cleaning",val:"99%",target:"Target 100%",icon:"\u23F1\uFE0F",bg:"#EFF6FF",color:"#3B82F6"},{label:"Kepatuhan GCDC",val:"100%",target:"Target 100%",icon:"\u{1F9F9}",bg:"#EFF6FF",color:"#3B82F6"},{label:"Complaint Cleaning (\u226410)",val:"4",target:"Target \u226410",icon:"\u{1F4DD}",bg:"#F5F3FF",color:"#8B5CF6"},{label:"Pelaksanaan Fogging",val:"100%",target:"Target 100%",icon:"\u{1F4A8}",bg:"#F5F3FF",color:"#8B5CF6"}];e.innerHTML=`
    <div style="display:grid;grid-template-columns:1fr;gap:12px;padding-bottom:12px;height:160px;overflow-y:auto;padding-right:8px;">
      ${i.map(t=>{let o=t.val.includes("%")?parseInt(t.val):Math.min(100,parseInt(t.val)*10);return`
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
            <div class="prog-bar-fill" style="width:${o}%;background:${t.color}"></div>
          </div>
        </div>
      `}).join("")}
    </div>
  `}function wt(){let a=document.getElementById("quick-actions");if(!a)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];a.innerHTML=e.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function le(a,e){let i=document.getElementById(a),t=document.getElementById(e);i&&(i.style.display="none",i.style.position=""),t&&(t.style.display="block")}function la(a,e="Belum ada data"){if(!a)return;a.style.display="none";let i=a.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent=e,i.appendChild(o)}}D();async function Da(a){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),i=document.getElementById("login-error"),t=document.getElementById("login-btn"),o=document.getElementById("toggle-password"),l=document.getElementById("login-password");o?.addEventListener("click",()=>{let r=l.type==="text";l.type=r?"password":"text",o.style.color=r?"":"var(--primary)"}),e?.addEventListener("submit",async r=>{r.preventDefault(),i.style.display="none";let s=e.username.value.trim(),n=e.password.value;if(!s||!n){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}t.querySelector(".btn-text").style.display="none",t.querySelector(".btn-spinner").style.display="",t.disabled=!0;try{let c=await f("/api/auth/login",{method:"POST",body:JSON.stringify({username:s,password:n})});c.ok&&c.data.success?(Ve(c.data.data.token),Te(c.data.data.user),J("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=c.data.error||"Username atau password salah.",i.style.display="block",t.classList.add("shake"),setTimeout(()=>t.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{t.querySelector(".btn-text").style.display="",t.querySelector(".btn-spinner").style.display="none",t.disabled=!1}})}D();B();async function _t(){return await I()}function xt(a,e){let i=String(a.status||"").toLowerCase();return e==="active"?i==="aktif":!1}async function Ia(a,e){let i=await _t(),t=e?e.get("dash_filter"):null;$({container:a,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",onDataLoaded:o=>t?o.filter(l=>xt(l,t)):o,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:o=>ge(o)},{key:"phone",label:"No. HP",render:o=>o?`<a href="tel:${o}">${o}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:o=>window.formatDate(o)},{key:"status",label:"Status",render:o=>N(o)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:o=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:o?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:o?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:i,value:o?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:o?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:o?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:o?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let o=await f(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(o.ok){let l=o.data.data.map(r=>({"Nama Lengkap":r.full_name,Cabang:r.branch_name||"",Divisi:r.division||"","No. HP":r.phone||"","Tgl Masuk":r.join_date||"",Status:r.status||""}));E(l,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async o=>{let l=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),m=i.find(p=>String(p.label||"").toLowerCase()===c);return m?m.value:null},r=o.map(n=>({full_name:String(n["Nama Lengkap"]||"").trim(),branch_id:l(String(n.Cabang||"").trim()),division:String(n.Divisi||"").trim()||"FACILITY CARE",phone:String(n["No. HP"]||"").trim(),join_date:String(n["Tgl Masuk"]||"").trim(),status:String(n.Status||"").trim(),notes:String(n.Catatan||"").trim()})).filter(n=>n.full_name),s=await f("/api/employees/import",{method:"POST",body:JSON.stringify(r)});if(!s.ok)throw new Error(s.data?.error||"Import gagal")}}})}D();B();var da=[],Pa=[];async function Ct(){da=await I(),Pa=await ue()}var ca=async a=>{let e=[],i=1;for(;;){let o=await(await Promise.resolve().then(()=>(D(),de))).apiFetch(`${a}${a.includes("?")?"&":"?"}limit=100&page=${i}`);if(!o.ok)break;let l=o.data?.data||o.data||[],r=Array.isArray(l)?l:[];if(e=e.concat(r),r.length<100||o.data?.pagination&&i>=o.data.pagination.pages)break;i++}return e};function Tt(a,e){if(String(a.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!a.end_date)return!1;let t=new Date;t.setHours(0,0,0,0);let o=new Date(t);o.setDate(t.getDate()+30);let l=new Date(a.end_date);return l.setHours(0,0,0,0),l>=t&&l<=o}return!1}async function Me(a,e){await Ct();let i=e?e.get("dash_filter"):null;$({container:a,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",onDataLoaded:t=>i?t.filter(o=>Tt(o,i)):t,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:t=>ge(t)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:t=>window.formatDate(t)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:t=>!t||String(t).startsWith("2099")?"Tetap / PKWTT":window.formatDate(t)},{key:"days_remaining",label:"Sisa Kontrak",render:(t,o)=>o.end_date&&String(o.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':ea(t)},{key:"status",label:"Status",render:t=>N(t)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:da},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:t=>(t.end_date||(t.end_date="2099-12-31"),t),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let t=document.createElement("button");t.id="btn-find-missing",t.className="btn btn-ghost",t.innerHTML="\u{1F50D} Cek Selisih Karyawan",t.style.marginLeft="8px",t.style.color="#EF4444",t.style.border="1px solid currentColor",t.onclick=async()=>{t.innerHTML="\u231B Mencari...",t.disabled=!0;try{let[l,r]=await Promise.all([ca("/api/employees?status=Aktif"),ca("/api/contracts")]);if(l.length>0){let s=r.filter(p=>p.status==="Aktif"),n=new Set(s.map(p=>p.employee_id)),c=l.filter(p=>!n.has(p.id)),m=`<p style="margin-bottom:12px">Data yang terbaca: <b>${l.length}</b> Karyawan Aktif, dan <b>${s.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${c.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;c.forEach(p=>{let d=r.filter(y=>y.employee_id===p.id),b='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(d.length>0){let y=d[0];b=`Pernah ada kontrak (Status: <b style="color:#EF4444">${y.status}</b>, Selesai: ${window.formatDate(y.end_date)})`}m+=`<li style="margin-bottom:8px"><b>${p.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${p.branch_name||"-"} | ${b}</span></li>`}),m+="</ul>",Promise.resolve().then(()=>(Ee(),ya)).then(p=>p.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:m,cancelText:"Tutup"}))}}catch(l){console.error(l)}t.innerHTML="\u{1F50D} Cek Selisih Karyawan",t.disabled=!1};let o=document.querySelector(".page-actions");o&&o.appendChild(t)}},formFields:t=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:Pa,value:t?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:da,value:t?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:t?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:t?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:t?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:t?.end_date&&!String(t.end_date).startsWith("2099")?t.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:t?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:t?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let t=await f(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(t.ok){let o=t.data.data.map(l=>({"Nama Lengkap":l.employee_name,Cabang:l.branch_name||"","Div / Bagian":l.division||"","Tanggal Mulai":l.start_date||"","Tanggal Selesai":l.end_date&&String(l.end_date).startsWith("2099")?"":l.end_date||"","Sisa Kontrak":l.end_date&&String(l.end_date).startsWith("2099")?"Tetap":l.days_remaining!==null&&l.days_remaining!==void 0?`${l.days_remaining} Hari`:"",Status:l.status||""}));E(o,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async t=>{let[o,l]=await Promise.all([f("/api/branches?limit=10000"),ca("/api/employees")]),r=o.data?.data||[],s=l||[];console.log(`Total employee yang berhasil dimuat dari database : ${s.length}`),s.length>0&&(console.log("Contoh 5 employee pertama:"),s.slice(0,5).forEach((u,h)=>{console.log(`${h+1}. ID: ${u.id}, Name: ${u.full_name}, Status: ${u.status}`)}));let n=u=>{if(!u)return null;let h=String(u||"").replace(/\s+/g," ").toLowerCase().trim(),S=r.find(w=>String(w.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===h||String(w.code||"").replace(/\s+/g," ").toLowerCase().trim()===h||String(w.name||"").replace(/\s+/g," ").toLowerCase().trim()===h);return S?S.id:null},c=(u,h)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${h}`),console.log(`Nama dari Excel : "${u}"`),!u)return console.log("Alasan gagal mapping : Nama kosong"),null;let S=String(u||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${S}"`),console.log(`Jumlah employee di database : ${s.length}`);let w=s.find(v=>String(v.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===S);return w?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${w.id}`),w.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},m=u=>{if(!u)return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let h=String(u).trim();if(/^\d{4,5}(\.\d+)?$/.test(h)){let w=Math.floor(Number(h));if(w>2e4&&w<99999){let v=new Date(Date.UTC(1899,11,30)+w*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);let S=h.split(/[\/\-\.]/);if(S.length===3){let[w,v,k]=S.map(T=>T.trim());if(w.length===4&&v.length<=2&&k.length<=2)return`${w}-${v.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&v.length<=2&&w.length<=2)return`${k}-${v.padStart(2,"0")}-${w.padStart(2,"0")}`}return h},p=t.map((u,h)=>{let S=h+2,w=String(u["Nama Lengkap"]||"").trim(),v=u["Tanggal Mulai"],k=m(v);if(!k){let C=t.__worksheet,P=t.__headers||[],F=P.indexOf("Tanggal Mulai"),ke="N/A",Se="N/A",je="N/A";if(F!==-1&&C&&window.XLSX){let _=window.XLSX.utils.encode_cell({c:F,r:S-1});je=_;let A=C[_];A?(ke=A.t||"undefined",Se=A.w||"undefined"):ke="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let x="Unknown";v==null||v===""?x="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":v instanceof Date&&isNaN(v.getTime())?x="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":x="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${S}`),console.log(`Employee Name : ${w}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${F})`),console.log(`Raw Cell Value : "${v}"`),console.log(`JavaScript Type : ${typeof v}`),console.log(`SheetJS Cell Type : ${ke}`),console.log(`SheetJS Formatted Value : "${Se}"`),console.log(`Value After Trim : "${String(v||"").trim()}"`),console.log(`Value After Date Parser : "${k}"`),console.log(`Is Empty : ${!v}`),console.log(`Is Invalid Date : ${v instanceof Date?isNaN(v.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${x}`),console.log(`Workbook Sheet : ${C?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${je}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(u,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(P)),console.log(`==========================
`)}let T=c(w,S),L=null;return T?k||(L="Tanggal Mulai kosong atau tidak berformat tanggal"):L="Karyawan tidak ditemukan di Database",{isValid:!!(T&&k),invalidReason:L,rowNum:S,data:{employee_id:T,branch_id:n(String(u.Cabang||"").trim()),division:String(u["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:k,end_date:m(u["Tanggal Selesai"])||"2099-12-31",status:String(u.Status||"").trim(),_rawName:w}}}),d=[],b=[];if(p.forEach(u=>{u.isValid?d.push(u.data):b.push({rowNum:u.rowNum,name:u.data._rawName,reason:u.invalidReason})}),console.log(`Split Validation - Valid: ${d.length}, Invalid: ${b.length}`),d.length===0){let u=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${t.length}
Valid: 0
Invalid: ${b.length}

Daftar Kegagalan (Contoh):
`;b.slice(0,10).forEach(h=>{u+=`- Row ${h.rowNum} | Nama: ${h.name} | Alasan: ${h.reason}
`}),b.length>10&&(u+=`- ... dan ${b.length-10} lainnya.
`),alert(u);return}let y=await f("/api/contracts/import",{method:"POST",body:JSON.stringify(d)}),g=`IMPORT SUMMARY
======================
`;g+=`Total Baris Excel : ${t.length}
`,g+=`Baris Valid       : ${d.length}
`,g+=`Baris Invalid     : ${b.length}

`,y&&y.data&&y.data.metrics?(g+=`Berhasil INSERT   : ${y.data.metrics.inserted}
`,g+=`Berhasil UPDATE   : ${y.data.metrics.updated}
`):g+=`Berhasil diproses : ${d.length}
`,b.length>0&&(g+=`
DAFTAR DATA DILEWATI:
`,b.forEach(u=>{g+=`- Row ${u.rowNum} | ${u.name} | ${u.reason}
`})),alert(g),typeof Me=="function"&&Me()}}})}D();B();var pa=[],Ie=[];function $t(a){if(!Array.isArray(a))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let i of e)if(a.some(t=>t.period===i))return i;return"Q3"}function Et(a,e){if(a.period!=="Q3")return!1;let i=String(a.status||"").toLowerCase();if(i!=="selesai"&&i!=="completed"&&i!=="done")return!1;let t=String(a.activity_type||"").toLowerCase();return e==="inspeksi"?t.includes("inspeksi"):e==="gcdc"?t.includes("general cleaning")||t.includes("deep cleaning"):!1}async function Ba(a,e){pa=await I();let i=await q();Ie=["Ade","Berlin"];let t=c=>c&&!Ie.find(m=>(typeof m=="object"?m.value:m)===c)?[...Ie,c]:Ie,o=await f(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),l=c=>{if(!c||c==="-"||String(c).trim()==="")return"";let m=String(c).split("-");return m.length===3&&m[0].length===4?`${m[2]}-${m[1]}-${m[0]}`:c},r=o.data?.data||[],s=$t(r),n=e?e.get("dash_filter"):null;$({container:a,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",defaultFilters:n?{period:"Q3"}:{period:s},onDataLoaded:c=>(n&&(c=c.filter(m=>Et(m,n))),c.sort((m,p)=>{let d=m.opening_date?new Date(m.opening_date).getTime():0;return(p.opening_date?new Date(p.opening_date).getTime():0)-d})),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:c=>aa(c)},{key:"period",label:"Periode",render:c=>z(c)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:c=>l(c)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:c=>l(c)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:c=>l(c)},{key:"status",label:"Status",render:c=>N(c)}],filterFields:[{type:"combobox",name:"branch_id",label:"Cabang",options:pa},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"combobox",name:"pic",label:"PIC",options:Ie}],formFields:c=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:pa,value:c?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"],value:c?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:c?.period},{name:"pic",label:"PIC",type:"combobox",options:t(c?.pic),value:c?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:c?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:c?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:c?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:c?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:c?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let c=await f(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(c.ok){let m=c.data.data.map(p=>({Cabang:p.branch_name||"",Kegiatan:p.activity_type||"",Periode:p.period||"",PIC:p.pic||"","Tgl Opening":p.opening_date||"","Tgl Target":p.target_date||"","Tgl Selesai":p.completion_date||"",Status:p.status||""}));E(m,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async c=>{let p=(await f("/api/branches?all=1")).data?.data||[],d=u=>{if(!u)return null;let h=String(u||"").toLowerCase(),S=p.find(w=>String(w.full_name||"").toLowerCase()===h||String(w.code||"").toLowerCase()===h||String(w.name||"").toLowerCase()===h);return S?S.id:null},b=u=>{if(u==null||u==="")return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let h=String(u).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let w=Number(h);if(w>2e4&&w<99999){let v=new Date(Date.UTC(1899,11,30)+w*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}let S=h.split(/[\/\-\.]/);if(S.length===3){let[w,v,k]=S.map(T=>T.trim());if(w.length===4&&v.length<=2&&k.length<=2)return`${w}-${v.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&v.length<=2&&w.length<=2)return`${k}-${v.padStart(2,"0")}-${w.padStart(2,"0")}`}return h},y=c.map(u=>({branch_id:d(String(u.Cabang||"").trim()),activity_type:String(u.Kegiatan||"").trim(),period:String(u.Periode||"").trim(),pic:String(u.PIC||u.Pic||"").trim(),opening_date:b(u["Tgl Opening"]||u["Tanggal Opening"]||u["Tgl Openir"]),target_date:b(u["Tgl Target"]||u["Tanggal Target"]),completion_date:b(u["Tgl Selesai"]||u["Tanggal Selesai"]),status:String(u.Status||"").trim(),notes:String(u.Catatan||u.Keterangan||"").trim()})).filter(u=>u.activity_type&&u.period),g=await f("/api/schedule/import",{method:"POST",body:JSON.stringify(y)});if(!g.ok)throw new Error(g.data?.error||"Import gagal")}}})}D();B();var ma=[],Ke=[];function Dt(a,e){let i=String(a.status||"").toLowerCase();return e==="open"?i==="open":!1}async function La(a,e){let i=e?e.get("dash_filter"):null;ma=await I(),Ke=await q();let t=r=>r&&!Ke.find(s=>s.value===r)?[...Ke,{value:r,label:r}]:Ke,o=new Date().getFullYear(),l=Array.from({length:5},(r,s)=>String(o-s));$({container:a,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",paginationMode:"client",onDataLoaded:r=>i?r.filter(s=>Dt(s,i)):r,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:r=>`<span class="badge badge-secondary">${r}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:r=>`<span title="${r}">${r?.length>50?r.slice(0,50)+"\u2026":r}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:r=>`<span title="${r||""}">${r?.length>40?r.slice(0,40)+"\u2026":r||"-"}</span>`},{key:"status",label:"Status",render:r=>N(r)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:r=>window.formatDate(r)},{key:"day_count",label:"Hari",render:r=>r??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:ma},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:l}],formFields:r=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:r?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:ma,value:r?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:r?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:r?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:r?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:t(r?.employee_name),value:r?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:t(r?.fc_specialist),value:r?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:r?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:r?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:r?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let r=await f(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(r.ok){let s=r.data.data.map(n=>({Tanggal:n.report_date||"",Cabang:n.branch_name||"",Kategori:n.category||"",Sumber:n.source||"",Keluhan:n.complaint||"","Nama FC":n.employee_name||"","FC Spesialis":n.fc_specialist||"",Solusi:n.solution||"","Tgl Selesai":n.completion_date||"",Status:n.status||""}));E(s,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async r=>{let n=(await f("/api/branches?all=1")).data?.data||[],c=d=>{if(!d)return null;let b=String(d||"").toLowerCase(),y=n.find(g=>String(g.full_name||"").toLowerCase()===b||String(g.code||"").toLowerCase()===b||String(g.name||"").toLowerCase()===b);return y?y.id:null},m=r.map(d=>({branch_id:c(String(d.Cabang||"").trim()),report_date:String(d.Tanggal||"").trim(),category:String(d.Kategori||"").trim(),source:String(d.Sumber||"").trim(),complaint:String(d.Keluhan||"").trim(),employee_name:String(d["Nama FC"]||"").trim(),fc_specialist:String(d["FC Spesialis"]||"").trim(),solution:String(d.Solusi||"").trim(),completion_date:String(d["Tgl Selesai"]||"").trim(),status:String(d.Status||"").trim()})).filter(d=>d.report_date&&d.complaint&&d.category),p=await f("/api/issues/import",{method:"POST",body:JSON.stringify(m)});if(!p.ok)throw new Error(p.data?.error||"Import gagal")}}})}D();var ye=[];function It(a,e){let i=String(a.status||"").toLowerCase();return e==="pending"?i==="pending":!1}async function Aa(a,e){let i=e?e.get("dash_filter"):null;ye=await I();let t=await q(),o=["Ade","Berlin"],l=s=>s&&!t.find(n=>n.value===s)?[...t,{value:s,label:s}]:t,r=s=>s&&!o.find(n=>(typeof n=="object"?n.value:n)===s)?[...o,s]:o;$({container:a,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",paginationMode:"client",onDataLoaded:s=>i?s.filter(n=>It(n,i)):s,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:s=>`<span title="${s||""}">${s?.length>50?s.slice(0,50)+"\u2026":s||"-"}</span>`},{key:"solution",label:"Solusi",render:s=>`<span title="${s||""}">${s?.length>40?s.slice(0,40)+"\u2026":s||"-"}</span>`},{key:"status",label:"Status",render:s=>N(s)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:s=>window.formatDate(s)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:s=>s?`<a href="${s}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"combobox",name:"branch_id",label:"Cabang",options:ye},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async s=>{let n=new URLSearchParams(s||{}).toString(),c=await f(`/api/one-on-one?limit=10000&${n}`);if(c.ok){let m=c.data.data.map(d=>({Tanggal:d.meeting_date||"",Cabang:d.branch_name||"","Nama Karyawan":d.employee_name||"",PIC:d.pic||"",Masalah:d.problem||"",Solusi:d.solution||"",Status:d.status||"","Tgl Selesai":d.completion_date||"",Dokumen:d.document_link||""})),{downloadExcel:p}=await Promise.resolve().then(()=>(B(),V));p(m,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let s=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(B(),V));n(s,"Template_Import_OneOnOne")},onImport:async s=>{let n=d=>{if(!d)return null;let b=String(d||"").toLowerCase(),y=ye.find(g=>String(g.label||"").toLowerCase()===b);return y?y.value:null},c=d=>{if(!d)return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let b=String(d).trim();if(/^\d{4,5}$/.test(b)){let g=Number(b);if(g>2e4&&g<99999){let u=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(b))return b.slice(0,10);let y=b.split(/[\/\-\.]/);if(y.length===3){let[g,u,h]=y.map(S=>S.trim());if(g.length===4&&u.length<=2&&h.length<=2)return`${g}-${u.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&u.length<=2&&g.length<=2)return`${h}-${u.padStart(2,"0")}-${g.padStart(2,"0")}`}return b},m=s.map(d=>({meeting_date:c(d.Tanggal),employee_name:String(d["Nama Karyawan"]||"").trim(),branch_id:n(String(d.Cabang||"").trim()),pic:String(d.PIC||"").trim(),problem:String(d.Masalah||"").trim(),solution:String(d.Solusi||"").trim(),status:String(d.Status||"").trim(),completion_date:c(d["Tgl Selesai"]),document_link:String(d.Dokumen||"").trim()})).filter(d=>d.meeting_date&&d.employee_name&&d.branch_id),p=await f("/api/one-on-one/import",{method:"POST",body:JSON.stringify(m)});if(!p.ok)throw new Error(p.data?.error||"Import gagal")}},formFields:s=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:s?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:s?.branch_id&&!ye.find(n=>n.value==s.branch_id)?[...ye,{value:s.branch_id,label:s.branch_name||s.branch_id}]:ye,createApi:{path:"/api/branches",field:"full_name"},value:s?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"combobox",required:!0,options:l(s?.employee_name),value:s?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:r(s?.pic),createApi:{path:"/api/pic",field:"name"},value:s?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:s?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:s?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:s?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:s?.document_link}]})}D();async function Na(a){let e=await I(),i=await q(),t=["Ade","Berlin"],o=s=>s&&!i.find(n=>n.value===s)?[...i,{value:s,label:s}]:i,l=s=>s&&!t.find(n=>(typeof n=="object"?n.value:n)===s)?[...t,s]:t,r=Array.from({length:5},(s,n)=>String(new Date().getFullYear()-n));$({container:a,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:s=>window.formatDate(s)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:s=>{try{let n=JSON.parse(s);return Array.isArray(n)?n.join(", "):s||"-"}catch{return s||"-"}}},{key:"score",label:"Nilai",render:s=>s!=null?`<strong>${s}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:s=>s?`<a href="${s}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:r}],exportOptions:{moduleName:"training",onExport:async s=>{let n=new URLSearchParams(s||{}).toString(),c=await f(`/api/training?limit=10000&${n}`);if(c.ok){let m=c.data.data.map(d=>{let b=d.participants||"";try{let y=JSON.parse(b);b=Array.isArray(y)?y.join(", "):b}catch{}return{Tanggal:d.training_date||"",Batch:d.batch||"",Materi:d.subject||"",Cabang:d.branch_name||"",Trainer:d.trainer||"",Peserta:b,Nilai:d.score!==null&&d.score!==void 0?d.score:"",Dokumen:d.document_link||""}}),{downloadExcel:p}=await Promise.resolve().then(()=>(B(),V));p(m,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let s=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(B(),V));n(s,"Template_Import_Training")},onImport:async s=>{let n=d=>{if(!d)return null;let b=String(d||"").toLowerCase(),y=e.find(g=>String(g.label||"").toLowerCase()===b);return y?y.value:null},c=d=>{if(!d)return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let b=String(d).trim();if(/^\d{4,5}$/.test(b)){let g=Number(b);if(g>2e4&&g<99999){let u=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(b))return b.slice(0,10);let y=b.split(/[\/\-\.]/);if(y.length===3){let[g,u,h]=y.map(S=>S.trim());if(g.length===4&&u.length<=2&&h.length<=2)return`${g}-${u.padStart(2,"0")}-${h.padStart(2,"0")}`;if(h.length===4&&u.length<=2&&g.length<=2)return`${h}-${u.padStart(2,"0")}-${g.padStart(2,"0")}`}return b},m=s.map(d=>({training_date:c(d.Tanggal),batch:String(d.Batch||"").trim(),subject:String(d.Materi||"").trim(),branch_id:n(String(d.Cabang||"").trim()),trainer:String(d.Trainer||"").trim(),participants:String(d.Peserta||"").trim(),score:d.Nilai?Number(d.Nilai):null,document_link:String(d.Dokumen||"").trim()})).filter(d=>d.training_date&&d.subject&&d.branch_id),p=await f("/api/training/import",{method:"POST",body:JSON.stringify(m)});if(!p.ok)throw new Error(p.data?.error||"Import gagal")}},formFields:s=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:s?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:s?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:s?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:s?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:l(s?.trainer),value:s?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let n=JSON.parse(s?.participants);return Array.isArray(n)?n.join(", "):s?.participants||""}catch{return s?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:s?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:s?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:s?.notes}],onBeforeSubmit:async s=>(s.participants&&(s.participants=JSON.stringify(s.participants.split(",").map(n=>n.trim()).filter(Boolean))),s)})}D();B();async function Oa(a,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await I(),t=await q(),o=e?e.get("dash_filter"):null;console.log("RAW",await ue()),console.log("OPTIONS",t);let l=n=>n&&!t.find(c=>c.value===n)?[...t,{value:n,label:n}]:t,r=["Agung Septiadi","Wasrikin","IQBAL AL BANNA"],s=n=>n&&!r.includes(n)?[...r,n]:r;$({container:a,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",onDataLoaded:n=>{if(o==="reliever"){let c=new Date,m=c.getFullYear(),p=String(c.getMonth()+1).padStart(2,"0");return n.filter(d=>{if(String(d.status||"").toLowerCase()!=="done")return!1;let b=d.backup_date||"";if(b.includes("/")){let y=b.split("/");if(y.length===3&&(y[2].length===4?y[2]:`20${y[2]}`)==m&&y[1].padStart(2,"0")==p)return!0}else if(b.includes("-")&&b.startsWith(`${m}-${p}`))return!0;return!1})}return n},columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:n=>z(n)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:n=>window.formatDate(n)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:n=>n?`<span class="badge badge-info">${n}</span>`:"-"},{key:"status",label:"Status",render:n=>N(n)}],filterFields:[{type:"search",placeholder:"Cari reliefer / FC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:n?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:l(n?.original_fc_name),value:n?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:s(n?.reliever_name),value:n?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:n?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:n?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:n?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:n?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let n=await f(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let c=n.data.data.map(m=>({Cabang:m.branch_name||"","Nama Facility care":m.original_fc_name||"",Periode:m.period||"",Relifer:m.reliever_name||"","Tanggal Back Up":m.backup_date||"","Tanggal Selesai":m.completion_date||"",Keterangan:m.reason||"",Shift:m.shift||"",Status:m.status||""}));c.length===0&&c.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),E(c,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async n=>{let m=(await f("/api/branches?all=1")).data?.data||[],p=y=>{if(!y)return null;let g=String(y||"").toLowerCase(),u=m.find(h=>String(h.full_name||"").toLowerCase()===g||String(h.code||"").toLowerCase()===g||String(h.name||"").toLowerCase()===g);return u?u.id:null},d=n.map(y=>({branch_name:String(y.Cabang||"").trim(),backup_date:String(y["Tanggal Back Up"]||y["Tanggal Backup"]||"").trim(),original_fc_name:String(y["Nama Facility care"]||y["FC Digantikan"]||"").trim(),reliever_name:String(y.Relifer||y.Reliefer||"").trim(),period:String(y.Periode||"").trim(),reason:String(y.Keterangan||"").trim(),shift:String(y.Shift||"").trim(),completion_date:String(y["Tanggal Selesai"]||"").trim(),status:String(y.Status||"").trim()})).filter(y=>y.reliever_name&&y.backup_date),b=await f("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:d})});if(!b.ok)throw new Error(b.data?.error||"Import gagal")}}})}D();B();async function Fa(a){let e=await I(),i=Array.from({length:4},(t,o)=>String(new Date().getFullYear()-o));$({container:a,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:t=>z(t)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"fc_score",label:"Point FC",render:t=>t!=null?`<strong class="${t>=80?"text-success":t>=60?"text-warning":"text-danger"}">${t}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:t=>t!=null?`<strong>${t}</strong>`:"-"},{key:"status",label:"Status",render:t=>N(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:t?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:t?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:t?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:t?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async t=>{let o=new URLSearchParams(t||{}).toString(),l=await f(`/api/reports/inspection?limit=10000&${o}`);if(l.ok){let r=l.data.data.map(s=>({Cabang:s.branch_name||"",Periode:s.period||"",Tanggal:s.inspection_date||"","Point FC":s.fc_score!==null&&s.fc_score!==void 0?s.fc_score:"","Point SPV":s.spv_score!==null&&s.spv_score!==void 0?s.spv_score:"",Status:s.status||"","Link Dokumen":s.document_link||""}));E(r,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Inspeksi")},onImport:async t=>{let o=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),m=e.find(p=>String(p.label||"").toLowerCase()===c);return m?m.value:null},l=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let c=String(n).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let p=Number(c);if(p>2e4&&p<99999){let d=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let m=c.split(/[\/\-\.]/);if(m.length===3){let[p,d,b]=m.map(y=>y.trim());if(p.length===4&&d.length<=2&&b.length<=2)return`${p}-${d.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&d.length<=2&&p.length<=2)return`${b}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`}return c},r=t.map(n=>({branch_id:o(String(n.Cabang||"").trim()),period:String(n.Periode||"").trim(),inspection_date:l(n.Tanggal),fc_score:n["Point FC"]!==void 0&&n["Point FC"]!==""?Number(n["Point FC"]):null,spv_score:n["Point SPV"]!==void 0&&n["Point SPV"]!==""?Number(n["Point SPV"]):null,status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.inspection_date),s=await f("/api/reports/inspection/import",{method:"POST",body:JSON.stringify(r)});if(!s.ok)throw new Error(s.data?.error||"Import gagal")}}})}D();B();async function Ma(a){let e=await I(),i=Array.from({length:4},(t,o)=>String(new Date().getFullYear()-o));$({container:a,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:t=>`<span class="badge ${t==="Deep Cleaning"?"badge-purple":"badge-success"}">${t}</span>`},{key:"period",label:"Periode",render:t=>z(t)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"status",label:"Status",render:t=>N(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:t?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:t?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:t?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async t=>{let o=new URLSearchParams(t||{}).toString(),l=await f(`/api/reports/cleaning?limit=10000&${o}`);if(l.ok){let r=l.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||""}));E(r,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_GCDC")},onImport:async t=>{let o=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),m=e.find(p=>String(p.label||"").toLowerCase()===c);return m?m.value:null},l=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let c=String(n).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let p=Number(c);if(p>2e4&&p<99999){let d=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let m=c.split(/[\/\-\.]/);if(m.length===3){let[p,d,b]=m.map(y=>y.trim());if(p.length===4&&d.length<=2&&b.length<=2)return`${p}-${d.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&d.length<=2&&p.length<=2)return`${b}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`}return c},r=t.map(n=>({branch_id:o(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:l(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),s=await f("/api/reports/cleaning/import",{method:"POST",body:JSON.stringify(r)});if(!s.ok)throw new Error(s.data?.error||"Import gagal")}}})}D();B();async function Ka(a){let e=await I(),i=Array.from({length:4},(t,o)=>String(new Date().getFullYear()-o));$({container:a,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:t=>`<span class="badge badge-warning">${t}</span>`},{key:"period",label:"Periode",render:t=>z(t)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"status",label:"Status",render:t=>N(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:t=>t||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:t?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:t?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async t=>{let o=new URLSearchParams(t||{}).toString(),l=await f(`/api/reports/fogging?limit=10000&${o}`);if(l.ok){let r=l.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"Fogging",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||""}));E(r,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async t=>{let o=n=>{if(!n)return null;let c=String(n||"").toLowerCase(),m=e.find(p=>String(p.label||"").toLowerCase()===c);return m?m.value:null},l=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let c=String(n).trim();if(c===""||c==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);if(/^\d{4,5}$/.test(c)){let p=Number(c);if(p>2e4&&p<99999){let d=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let m=c.split(/[\/\-\.]/);if(m.length===3){let[p,d,b]=m.map(y=>y.trim());if(p.length===4&&d.length<=2&&b.length<=2)return`${p}-${d.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&d.length<=2&&p.length<=2)return`${b}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`}return c},r=t.map(n=>({branch_id:o(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"Fogging").trim(),period:String(n.Periode||"").trim(),activity_date:l(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.activity_date),s=await f("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(r)});if(!s.ok)throw new Error(s.data?.error||"Import gagal")}}})}D();B();async function Ra(a){let e=await I(),i=await q(),t=i,o=r=>r&&!i.find(s=>s.value===r)?[...i,{value:r,label:r}]:i,l=r=>r&&!t.find(s=>s.value===r)?[...t,{value:r,label:r}]:t;$({container:a,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:r=>`<span title="${r||""}">${r?.length>60?r.slice(0,60)+"\u2026":r||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:r=>window.formatDate(r)},{key:"status",label:"Status",render:r=>N(r)},{key:"notes",label:"Keterangan",render:r=>r?.length>40?r.slice(0,40)+"\u2026":r||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]}],formFields:r=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:r?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:l(r?.pic),value:r?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:r?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:r?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:r?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:r?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async r=>{let s=new URLSearchParams(r||{}).toString(),n=await f(`/api/reports/basecamp?limit=10000&${s}`);if(n.ok){let c=n.data.data.map(m=>({"Tgl Info":m.info_date||"",Cabang:m.branch_name||"",Permasalahan:m.problem||"",PIC:m.pic||"","Tgl Done":m.done_date||"",Status:m.status||"",Keterangan:m.notes||""}));E(c,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async r=>{let s=p=>{if(!p)return null;let d=String(p||"").toLowerCase(),b=e.find(y=>String(y.label||"").toLowerCase()===d);return b?b.value:null},n=p=>{if(p==null||p==="")return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let d=String(p).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let y=Number(d);if(y>2e4&&y<99999){let g=new Date(Date.UTC(1899,11,30)+y*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}let b=d.split(/[\/\-\.]/);if(b.length===3){let[y,g,u]=b.map(h=>h.trim());if(y.length===4&&g.length<=2&&u.length<=2)return`${y}-${g.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&g.length<=2&&y.length<=2)return`${u}-${g.padStart(2,"0")}-${y.padStart(2,"0")}`}return d},c=r.map(p=>({info_date:n(p["Tgl Info"]||p["Tanggal Info"]),branch_id:s(String(p.Cabang||"").trim()),problem:String(p.Permasalahan||"").trim(),pic:String(p.PIC||"").trim(),done_date:n(p["Tgl Done"]||p["Tanggal Done"]),status:String(p.Status||"").trim(),notes:String(p.Keterangan||p.Catatan||"").trim()})).filter(p=>p.info_date&&p.branch_id&&p.problem),m=await f("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(c)});if(!m.ok)throw new Error(m.data?.error||"Import gagal")}}})}async function qa(a){$({container:a,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:t}=await Promise.resolve().then(()=>(D(),de)),o=await t(`/api/sop?limit=10000&${i}`);if(o.ok){let l=o.data.data.map(s=>({"Nama SOP":s.name||"",Kategori:s.category||"",Dokumen:s.document_link||"",Catatan:s.notes||s.description||""})),{downloadExcel:r}=await Promise.resolve().then(()=>(B(),V));r(l,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(B(),V));i(e,"Template_Import_SOP")},onImport:async e=>{let i=e.map(l=>({name:String(l["Nama SOP"]||"").trim(),category:String(l.Kategori||"").trim(),document_link:String(l.Dokumen||"").trim(),description:String(l.Catatan||"").trim()})).filter(l=>l.name),{apiFetch:t}=await Promise.resolve().then(()=>(D(),de)),o=await t("/api/sop/import",{method:"POST",body:JSON.stringify(i)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function Ha(a){$({container:a,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:t}=await Promise.resolve().then(()=>(D(),de)),o=await t(`/api/checklist?limit=10000&${i}`);if(o.ok){let l=o.data.data.map(s=>({"Nama Checklist":s.name||"",Kategori:s.category||"",Dokumen:s.document_link||"",Deskripsi:s.description||""})),{downloadExcel:r}=await Promise.resolve().then(()=>(B(),V));r(l,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(B(),V));i(e,"Template_Import_Checklist")},onImport:async e=>{let i=e.map(l=>({name:String(l["Nama Checklist"]||"").trim(),category:String(l.Kategori||"").trim(),document_link:String(l.Dokumen||"").trim(),description:String(l.Deskripsi||"").trim()})).filter(l=>l.name),{apiFetch:t}=await Promise.resolve().then(()=>(D(),de)),o=await t("/api/checklist/import",{method:"POST",body:JSON.stringify(i)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}D();Ee();B();async function ua(a,e="forms"){if(e==="supply")return Bt(a);Pt(a)}function Pt(a){$({container:a,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}]})}async function Bt(a){let i=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name}));$({container:a,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:t=>t?new Date(t).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(t,o)=>o.branch_name_ref||o.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:t=>{try{let o=JSON.parse(t);return Array.isArray(o)?o.join(", "):t}catch{return t||"-"}}},{key:"chemical_items",label:"Chemical",render:t=>{try{let o=JSON.parse(t);return Array.isArray(o)?o.join(", "):t}catch{return t||"-"}}},{key:"additional_notes",label:"Catatan",render:t=>t?.length>40?t.slice(0,40)+"\u2026":t||"-"},{key:"status",label:"Status",render:t=>N(t)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:t=>{let o=t?.tools_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}let l=t?.chemical_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:t?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:t?.branch_id&&!i.find(r=>r.value==t.branch_id)?[...i,{value:t.branch_id,label:t.branch_name||t.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:o},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:t?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:l},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:t?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:t?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:t?.status||""},{name:"processed_by",label:"Diproses Oleh",value:t?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async t=>{let o=new URLSearchParams(t||{}).toString(),l=await f(`/api/reports/supply?limit=10000&${o}`);if(l.ok){let r=l.data.data.map(s=>{let n=s.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let c=s.chemical_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}return{Waktu:s.submitted_at||"",Pengirim:s.submitter_name||"",Cabang:s.branch_name_ref||s.branch_name||"","Alat/Barang":n||"",Chemical:c||"",Catatan:s.additional_notes||"",Status:s.status||"","Diproses Oleh":s.processed_by||""}});E(r,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async t=>{let l=(await f("/api/branches?all=1")).data?.data||[],r=m=>{if(!m)return null;let p=String(m||"").toLowerCase(),d=l.find(b=>String(b.full_name||"").toLowerCase()===p||String(b.code||"").toLowerCase()===p||String(b.name||"").toLowerCase()===p);return d?d.id:null},s=m=>{if(m==null||m==="")return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let p=String(m).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let b=Number(p);if(b>2e4&&b<99999){let y=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(y.getTime())?"":y.toISOString().slice(0,10)}}let d=p.split(/[\/\-\.]/);if(d.length===3){let[b,y,g]=d.map(u=>u.trim());if(b.length===4&&y.length<=2&&g.length<=2)return`${b}-${y.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&y.length<=2&&b.length<=2)return`${g}-${y.padStart(2,"0")}-${b.padStart(2,"0")}`}return p},n=t.map(m=>({submitted_at:s(m.Waktu||m.Tanggal),submitter_name:String(m.Pengirim||"").trim(),branch_id:r(String(m.Cabang||"").trim()),tools_items:String(m["Alat/Barang"]||m.Alat||"").trim(),chemical_items:String(m.Chemical||"").trim(),additional_notes:String(m.Catatan||m.Keterangan||"").trim(),status:String(m.Status||"").trim(),processed_by:String(m["Diproses Oleh"]||m.PIC||"").trim()})).filter(m=>m.submitted_at&&m.submitter_name&&m.branch_id),c=await f("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(t,o)=>{let l=te({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(r,s)=>{let n=r.querySelector("#supply-status").value,c=r.querySelector("#supply-processed-by").value;(await f(`/api/reports/supply/${t.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:c})})).ok?(J("Status diperbarui."),s(),o()):Q("Gagal update status.")}})}}]})}D();B();async function ja(a){let e=oe();if(!e||!["superadmin","admin"].includes(e.role)){a.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}$({container:a,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let t=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!t,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!t,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:t?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!t,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:t?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await f(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let t=i.data.data.map(o=>({"Nama Lengkap":o.full_name||"",Username:o.username||"",Email:o.email||"",Role:o.role||"",Status:o.is_active?"Aktif":"Nonaktif"}));E(t,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let t=i.map(l=>({full_name:String(l["Nama Lengkap"]||"").trim(),username:String(l.Username||"").trim(),email:String(l.Email||"").trim(),role:String(l.Role||"").trim()||"viewer",password:String(l.Password||"").trim()})).filter(l=>l.username&&l.password&&l.email&&l.full_name),o=await f("/api/users/import",{method:"POST",body:JSON.stringify(t)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}}})}D();B();async function Ua(a){$({container:a,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await f(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)E(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let i=e.map(o=>({code:String(o["Kode Cabang"]||"").trim(),name:String(o["Nama Pendek"]||"").trim(),full_name:String(o["Nama Lengkap"]||"").trim(),city:String(o.Kota||"").trim()})).filter(o=>o.code&&o.name),t=await f("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}}})}D();async function Ja(a){let e=new Date,i=[];a.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),o()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),o()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(l=>l.addEventListener("change",o));async function t(){try{let l=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;i=(await f(`/api/dashboard/calendar?month=${l}`)).data?.data||[]}catch(l){console.warn("[Calendar] Failed to load events, rendering empty grid:",l),i=[]}}async function o(){let l=document.getElementById("calendar-grid");if(l){l.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await t();try{let r=e.getFullYear(),s=e.getMonth(),n=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),c=document.getElementById("cal-month-label");c&&(c.textContent=n);let m=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(v=>v.value)),p=i.filter(v=>m.has(v.type)),d={};p.forEach(v=>{let k=(v.event_date||"").slice(0,10);d[k]||(d[k]=[]),d[k].push(v)});let b=new Date(r,s,1).getDay(),y=new Date(r,s+1,0).getDate(),g=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],u=new Date().toISOString().slice(0,10),h='<div class="calendar-grid">';g.forEach(v=>{h+=`<div class="cal-day-header">${v}</div>`});for(let v=0;v<b;v++)h+='<div class="cal-cell cal-cell-empty"></div>';for(let v=1;v<=y;v++){let k=`${r}-${String(s+1).padStart(2,"0")}-${String(v).padStart(2,"0")}`,T=d[k]||[],L=k===u;h+=`
          <div class="cal-cell ${L?"cal-today":""} ${T.length?"cal-has-events":""}"
               data-date="${k}" tabindex="0" role="button" aria-label="${k}">
            <div class="cal-day-num ${L?"today-num":""}">${v}</div>
            <div class="cal-events-preview">
              ${T.slice(0,3).map(C=>`
                <div class="cal-event-dot cal-color-${C.color||"gray"}" title="${Re(C.title||C.type)}">
                  <span class="cal-event-dot-label">${Lt(C.title||C.branch_name||C.type,18)}</span>
                </div>
              `).join("")}
              ${T.length>3?`<div class="cal-more">+${T.length-3} lagi</div>`:""}
            </div>
          </div>`}let w=(b+y)%7;if(w!==0)for(let v=0;v<7-w;v++)h+='<div class="cal-cell cal-cell-empty"></div>';h+="</div>",l.innerHTML=h,l.querySelectorAll(".cal-cell[data-date]").forEach(v=>{v.addEventListener("click",()=>{let k=v.dataset.date,T=d[k]||[];if(!T.length)return;let L=document.getElementById("cal-event-list"),C=new Date(k+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=C,document.getElementById("cal-event-items").innerHTML=T.map(P=>`
            <div class="cal-event-item cal-color-border-${P.color||"gray"}">
              <div class="cal-event-type">${At(P.type)}</div>
              <div class="cal-event-title">${Re(P.title||"-")}</div>
              <div class="cal-event-branch">${Re(P.branch_name||"")}</div>
              ${P.status?`<div class="cal-event-status">${Re(P.status)}</div>`:""}
              ${P.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${P.days_remaining} hari</div>`:""}
            </div>
          `).join(""),L.style.display="block"})})}catch(r){console.error("[Calendar] Render error:",r),l&&(l.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}o()}function Lt(a,e){return a?a.length>e?a.slice(0,e)+"\u2026":a:""}function Re(a){return a?String(a).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function At(a){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[a]||a}D();async function Ga(a){let e=oe(),i=(e?.full_name||e?.username||"U")[0].toUpperCase(),o={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";a.innerHTML=`
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
  `;let l=localStorage.getItem("fm_token"),r=document.getElementById("session-info");if(l&&r)try{let s=JSON.parse(atob(l.split(".")[1])),n=new Date(s.exp*1e3);r.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{r.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async s=>{s.preventDefault();let n=document.getElementById("pwd-error"),c=document.getElementById("pwd-success"),m=document.getElementById("btn-save-pwd");n.style.display="none",c.style.display="none";let p=s.target,d=p.current_password.value,b=p.new_password.value,y=p.confirm_password.value;if(b!==y){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(b.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}m.disabled=!0,m.textContent="\u23F3 Menyimpan...";let g=await f("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:d,new_password:b})});m.disabled=!1,m.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',g.ok?(c.textContent="\u2705 Password berhasil diubah.",c.style.display="block",p.reset(),J("Password berhasil diubah.")):(n.textContent=g.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}D();var qe={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function H(a){if(a==null||a==="")return null;if(a instanceof Date)return isNaN(a.getTime())?null:a.toISOString().slice(0,10);let e=String(a).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let o=Number(e);if(o>2e4&&o<99999){let l=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(l.getTime())?null:l.toISOString().slice(0,10)}}let i=e.split(/[\/\-\.]/);if(i.length===3){let[o,l,r]=i.map(m=>m.trim()),s=Number(o),n=Number(l),c=Number(r);if(o.length===4&&s>1900)return`${o}-${l.padStart(2,"0")}-${r.padStart(2,"0")}`;if(r.length===4&&c>1900)return s>12?`${r}-${l.padStart(2,"0")}-${o.padStart(2,"0")}`:n>12?`${r}-${o.padStart(2,"0")}-${l.padStart(2,"0")}`:`${r}-${l.padStart(2,"0")}-${o.padStart(2,"0")}`;if(r.length===2&&!isNaN(c)){let m=c>=50?`19${r}`:`20${r}`;return s>12?`${m}-${l.padStart(2,"0")}-${o.padStart(2,"0")}`:`${m}-${l.padStart(2,"0")}-${o.padStart(2,"0")}`}}let t=new Date(e);return isNaN(t.getTime())?null:t.toISOString().slice(0,10)}function Qa(a){return Object.values(a).every(e=>e==null||String(e).trim()==="")}var Nt={validation:{required:[],map:a=>({cabang:a.CABANG,pic:a.PIC,kegiatan:a.KEGIATAN,quartal:a.QUARTAL,masa_pkwt:a["MASA PKWT"],pic_pelapor:a["PIC PELAPOR"],kontrak:a.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:a=>({name:a["Nama SOP"],category:a.Kategori||"Umum",document_link:a["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:a=>({full_name:a["Nama Lengkap"],branch_name:a.Cabang,division:a["Div / Bagian"]||"FACILITY CARE",phone:a["No. Hp"],join_date:H(a["Tanggal Masuk"]),status:a.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:a=>({employee_name:a["Nama Lengkap"],branch_name:a.Cabang,division:a["Div / Bagian"]||"FACILITY CARE",start_date:H(a["Tanggal Mulai"]),end_date:H(a["Tanggal Selesai"]),contract_type:a["Tipe Kontrak"]||"",pkwt_number:a.PKWT||"",status:a.Status||"",notes:a.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:a=>({report_date:H(a["Tanggal Info"]),branch_name:a.Cabang,category:a.Kategori,source:a["Sumber Laporan"],complaint:a.Keluhan,employee_name:a["Nama FC"],fc_specialist:a["FC Spesialis"],solution:a.Solusi,status:a.Status||"",completion_date:H(a["Tanggal Selesai"])})},one_on_one:{required:[],map:a=>({meeting_date:H(a.Tanggal),branch_name:a.Cabang,employee_name:a["Nama Karyawan"],pic:a.Pic,problem:a.Masalah,solution:a.Solusi,status:a.Status||"",completion_date:H(a["Tanggal Selesai"]),document_link:a["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:a=>({branch_name:a.Cabang,activity_type:a.Kegiatan,period:a.Periode,pic:a.Pic||a.PIC,opening_date:H(a["Tanggal Opening"]||a["Tgl Opening"]),target_date:H(a["Tanggal Target"]||a["Tgl Target"]),completion_date:H(a["Tanggal Selesai"]||a["Tgl Selesai"]),status:a.Status||"",notes:a.Keterangan||a.Catatan})},inspection:{required:[],map:a=>({inspection_date:H(a.Tanggal),branch_name:a.Cabang,period:a.Periode,status:a.Status||"",fc_score:a["Point FC SP"]!==void 0&&a["Point FC SP"]!==null?parseFloat(String(a["Point FC SP"]).replace(",",".")):null,spv_score:a["Point SPV"]!==void 0&&a["Point SPV"]!==null?parseFloat(String(a["Point SPV"]).replace(",",".")):null,document_link:a.Link,notes:""})},cleaning:{required:[],map:a=>({activity_date:H(a.Tanggal),branch_name:a.Cabang,activity_type:a["Jenis Kegiatan"]||"General Cleaning",period:a.Periode,status:a.Status||"",document_link:a.Link,notes:""})},fogging:{required:[],map:a=>({activity_date:H(a.Tanggal),branch_name:a.Cabang,period:a.Periode,status:a.Status||"",document_link:a.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:a=>({info_date:H(a["Tgl Info"]),branch_name:a.Cabang,problem:a.Permasalahan,pic:a.PIC,done_date:H(a["Tgl Done"]),status:a.Status||"",notes:a.Ket})},relievers:{required:[],map:a=>({branch_name:a.Cabang,original_fc_name:a["Nama Facility care"],period:a.Periode,reliever_name:a.Relifer,backup_date:H(a["Tanggal Back Up"]),completion_date:H(a["Tanggal Selesai"]),reason:a.Keterangan,shift:a.Shift,status:a.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:a=>({training_date:H(a.Tanggal),batch:a.Batch,subject:a.Materi,participants:a.Peserta,branch_name:a.Cabang,trainer:a.Trainer,score:a.Nilai!==void 0&&a.Nilai!==null?parseFloat(String(a.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:a=>({name:a["Master Checklist"],category:"Umum",document_link:a["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:a=>({name:a["Master Form"],category:"Umum",document_link:a["Link Document"],description:""})},supply:{required:[],map:a=>({submitted_at:H(a.Timestamp),submitter_name:a["Nama Lengkap"],branch_name:a["Kebutuhan Untuk Cabang"],tools_items:a["Alat - Alat / Barang"],tools_quantity:a["Jumlah Permintaan Alat / Barang"],chemical_items:a.Chemical,chemical_quantity:a["Jumlah Permintaan Chemical"],additional_notes:a["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:a.Status||""})}};function Ot(a,e){let i=qe[a];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let t=Nt[i.module];if(!t)return{valid:[],errors:[],mapped:[],skipped:!0};let o=[],l=[],r=[];return e.filter(n=>!Qa(n)).forEach((n,c)=>{let m=e.indexOf(n)+2,p=[];t.required.forEach(({key:b,label:y})=>{let g=n[b];if(g==null||String(g).trim()===""){let u=Object.keys(n).filter(h=>h.trim()).join(", ");p.push({column:y,originalValue:g||"",reason:`Kolom "${y}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${u.slice(0,120)}`})}});let d=t.map(n);p.length>0?l.push({row:m,data:d,raw:n,errors:p}):(o.push(n),r.push(d))}),{valid:o,errors:l,mapped:r}}function Va(a){let e=[];return a.SheetNames.forEach(i=>{let t=qe[i];if(!t)return;let o=a.Sheets[i],l=window.XLSX.utils.sheet_to_json(o,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),r=Ot(i,l),s=l.filter(n=>!Qa(n));e.push({sheetName:i,module:t.module,label:t.label,total:s.length,valid:r.mapped.length,errorCount:r.errors.length,errors:r.errors,mapped:r.mapped,skipped:!1})}),e}function Ya(){let a=window.XLSX,e=a.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu","NAMA KARYAWAN":"Budi Santoso",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02",Day:"1"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06",Day:"1","Link Document":""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}]}).forEach(([t,o])=>{a.utils.book_append_sheet(e,a.utils.json_to_sheet(o),t)}),a.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function Wa(a){let e=window.XLSX,i=e.utils.book_new(),t=!1;return a.forEach(o=>{if(!o.errors||o.errors.length===0)return;t=!0;let l=o.errors.map(s=>({"No. Baris":s.row,"Kolom Gagal":(s.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(s.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(s.data||{}).map(([n,c])=>[n,c??""]))})),r=e.utils.json_to_sheet(l);e.utils.book_append_sheet(i,r,o.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),t?(e.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Ft=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function za(a){a.innerHTML=`
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
              ${Object.entries(qe).map(([g,{label:u}])=>`<span class="import-sheet-tag">\u{1F4C4} ${g} \u2192 ${u}</span>`).join("")}
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
  `;let e=null,i=null,t=0,o={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function l(g){Object.entries(o).forEach(([u,h])=>{h.style.display=u===g?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let g=document.getElementById("btn-backup-db");g.disabled=!0,g.textContent="\u23F3 Memproses Backup...";try{let u=await f("/api/import/backup");if(u.ok){let h=new Blob([JSON.stringify(u.data,null,2)],{type:"application/json"}),S=URL.createObjectURL(h),w=document.createElement("a");w.href=S,w.download=`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.json`,document.body.appendChild(w),w.click(),document.body.removeChild(w),URL.revokeObjectURL(S),J("Backup berhasil diunduh!")}else Q("Gagal memproses backup: "+(u.data?.error||"Unknown error"))}catch(u){Q("Gagal memproses backup: "+u.message)}finally{g.disabled=!1,g.textContent="\u{1F4E6} Backup Database"}});let r=document.getElementById("btn-sync-google");r&&r.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let g=r.innerHTML;r.innerHTML='<span class="spinner"></span> Menyinkronkan...',r.disabled=!0;try{let u=await f("/api/sync/google-sheets",{method:"POST"});u.ok?alert("Sinkronisasi Berhasil: "+(u.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(u.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{r.innerHTML=g,r.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{Ya(),J("Template Excel berhasil didownload!")});let s=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",g=>{g.stopPropagation(),s.click()}),s.addEventListener("change",g=>{g.target.files[0]&&c(g.target.files[0])}),n.addEventListener("dragover",g=>{g.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",g=>{g.preventDefault(),n.classList.remove("drag-over");let u=g.dataTransfer.files[0];u&&u.name.match(/\.xlsx?$/i)?c(u):Q("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,s.value="",document.getElementById("file-info").style.display="none",n.style.display="",l("upload")});async function c(g){e=g,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${g.name} (${(g.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await m(g)}async function m(g){l("validating");let u=document.getElementById("validation-status"),h=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");u.textContent="Membaca file Excel...",h.style.width="20%",await Pe(200);let S=await g.arrayBuffer(),w=window.XLSX.read(S,{type:"array",cellDates:!0});u.textContent=`Memvalidasi ${w.SheetNames.length} sheet...`,h.style.width="50%",await Pe(100),i=Va(w),h.style.width="100%",u.textContent="Validasi selesai!",await Pe(300),p()}catch(S){l("upload"),Q("Gagal memproses file: "+S.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function p(){l("preview");let g=i.filter(C=>!C.skipped).length,u=i.reduce((C,P)=>C+P.total,0),h=i.reduce((C,P)=>C+P.valid,0),S=i.reduce((C,P)=>C+P.errorCount,0),w=u>0?Math.round(h/u*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${g} sheet</span>
      <span class="badge badge-secondary">${u} baris</span>
      <span class="badge badge-success">${h} valid (${w}%)</span>
      ${S>0?`<span class="badge badge-danger">${S} error</span>`:""}
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
          ${i.map((C,P)=>`
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
    `,v.querySelectorAll(".btn-detail-error").forEach(C=>{C.addEventListener("click",()=>{let P=i[Number(C.dataset.idx)];d(P)})});let k=document.getElementById("error-detail-section"),T=document.getElementById("error-detail-container");T.innerHTML="",k.style.display="none";let L=document.getElementById("btn-start-import");h===0?(L.disabled=!0,L.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(L.disabled=!1,S>0?(L.innerHTML=`\u{1F680} Import ${h} Data Valid (${S} dilewati)`,L.title="Baris error akan dilewati, baris valid tetap diimport"):L.innerHTML=`\u{1F680} Mulai Import ${h} Data`)}function d(g){let u=document.getElementById("error-detail-section"),h=document.getElementById("error-detail-container");u.style.display="";let S=g.errors.slice(0,100).map(w=>(Array.isArray(w.errors)?w.errors:[]).map(k=>{let T=typeof k=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${w.row}</span></td>
            <td><strong>${T?k.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${T&&k.originalValue!==void 0?k.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${T?k.reason:k}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${T&&k.aliases?`Gunakan salah satu nama kolom:<br><em>${k.aliases}</em>`:T&&k.hint?k.hint:""}
            </td>
          </tr>
        `}).join("")).join("");h.innerHTML=`
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
    `,u.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{l("upload"),document.getElementById("file-info").style.display="none",n.style.display="",e=null,s.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;Wa(i)?J("Log error berhasil didownload."):J("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let g=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";b(g)});async function b(g){l("importing"),t=Date.now();let u=[];Ft.forEach(k=>{let T=i?.find(L=>L.module===k&&L.mapped?.length>0);T&&u.push(T)});let h=document.getElementById("import-steps-list");h.innerHTML=u.map(k=>`
      <div class="import-step-item" id="step-item-${k.module}">
        <span class="step-item-icon" id="step-icon-${k.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${k.label} <span class="step-item-count">(${k.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${k.module}"></span>
      </div>
    `).join("");let S=document.getElementById("import-bar"),w=document.getElementById("import-current-status"),v={totalSheets:u.length,totalRows:u.reduce((k,T)=>k+T.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let k=0;k<u.length;k++){let T=u[k],L=document.getElementById(`step-icon-${T.module}`),C=document.getElementById(`step-status-${T.module}`);L.textContent="\u{1F504}",C.textContent="Mengimport...",w.textContent=`Mengimport ${T.label}...`,S.style.width=`${Math.round(k/u.length*100)}%`;try{let P=await f(`/api/import/${T.module}`,{method:"POST",body:JSON.stringify({rows:T.mapped,onDuplicate:g})});if(P.ok){let F=P.data;v.inserted+=F.inserted||0,v.skipped+=F.skipped||0,v.moduleResults.push({label:T.label,inserted:F.inserted||0,skipped:F.skipped||0,status:"ok"}),L.textContent="\u2705",C.innerHTML=`<span class="badge badge-success">${F.inserted||0} berhasil</span>${F.skipped>0?` <span class="badge badge-neutral">${F.skipped} skip</span>`:""}`}else v.failed++,v.moduleResults.push({label:T.label,inserted:0,skipped:0,status:"error",error:P.data?.error}),L.textContent="\u274C",C.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(P){v.failed++,v.moduleResults.push({label:T.label,inserted:0,skipped:0,status:"error",error:P.message}),L.textContent="\u274C",C.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Pe(150)}S.style.width="100%",w.textContent="Selesai!",await Pe(400),y(v)}function y(g){l("summary");let u=((Date.now()-t)/1e3).toFixed(1),h=g.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${h?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${h?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
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
        <div class="stat-value">${u}s</div>
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,i=null,s.value="",document.getElementById("file-info").style.display="none",n.style.display="",l("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Pe(a){return new Promise(e=>setTimeout(e,a))}D();var He=[],Xa=[];async function Za(a){He=await I(),Xa=await q(),$({container:a,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:He}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),t=await f(`/api/sp?limit=10000&${i}`);if(t.ok){let o=t.data.data.map(r=>({"Nama Karyawan":r.employee_name||"",Divisi:r.division||"",Cabang:r.branch_name||"","Tanggal Sp":r.tanggal||"","Akhir Sp":r.akhir_sp||"","Jenis Sp":r.sp_type||"","Link Document / Foto":r.document_link||""})),{downloadExcel:l}=await Promise.resolve().then(()=>(B(),V));l(o,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(B(),V));i(e,"Template_Import_SP")},onImport:async e=>{let i=r=>{if(!r)return null;let s=String(r||"").toLowerCase(),n=He.find(c=>String(c.label||"").toLowerCase()===s);return n?n.value:null},t=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let s=String(r).trim();if(/^\d{4,5}$/.test(s)){let c=Number(s);if(c>2e4&&c<99999){let m=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);let n=s.split(/[\/\-\.]/);if(n.length===3){let[c,m,p]=n.map(d=>d.trim());if(c.length===4&&m.length<=2&&p.length<=2)return`${c}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&m.length<=2&&c.length<=2)return`${p}-${m.padStart(2,"0")}-${c.padStart(2,"0")}`}return s},o=e.map(r=>({employee_name:String(r["Nama Karyawan"]||"").trim(),division:String(r.Divisi||"").trim(),branch_id:i(String(r.Cabang||"").trim()),tanggal:t(r["Tanggal Sp"]),akhir_sp:t(r["Akhir Sp"]),sp_type:String(r["Jenis Sp"]||"").trim(),document_link:String(r["Link Document / Foto"]||"").trim()})).filter(r=>r.employee_name&&r.branch_id),l=await f("/api/sp/import",{method:"POST",body:JSON.stringify(o)});if(!l.ok)throw new Error(l.data?.error||"Import gagal")}},formFields:[{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:Xa},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"combobox",name:"branch_id",label:"Cabang",required:!0,options:He,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}D();var fe=[],et=[];async function at(a){fe=await I(),et=await q(),$({container:a,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",options:fe},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",options:fe}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),t=await f(`/api/mutasi?limit=10000&${i}`);if(t.ok){let o=t.data.data.map(r=>({Tanggal:r.tanggal||"","Nama Karyawan":r.employee_name||"","Cabang Asal":r.from_branch_name||"","Cabang Tujuan":r.to_branch_name||"",Status:r.status||"",Dokumen:r.document_link||""})),{downloadExcel:l}=await Promise.resolve().then(()=>(B(),V));l(o,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(B(),V));i(e,"Template_Import_Mutasi")},onImport:async e=>{let i=r=>{if(!r)return null;let s=String(r||"").toLowerCase(),n=fe.find(c=>String(c.label||"").toLowerCase()===s);return n?n.value:null},t=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let s=String(r).trim();if(/^\d{4,5}$/.test(s)){let c=Number(s);if(c>2e4&&c<99999){let m=new Date(Date.UTC(1899,11,30)+c*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);let n=s.split(/[\/\-\.]/);if(n.length===3){let[c,m,p]=n.map(d=>d.trim());if(c.length===4&&m.length<=2&&p.length<=2)return`${c}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&m.length<=2&&c.length<=2)return`${p}-${m.padStart(2,"0")}-${c.padStart(2,"0")}`}return s},o=e.map(r=>({tanggal:t(r.Tanggal),employee_name:String(r["Nama Karyawan"]||"").trim(),from_branch_id:i(String(r["Cabang Asal"]||"").trim()),to_branch_id:i(String(r["Cabang Tujuan"]||"").trim()),status:String(r.Status||"").trim(),document_link:String(r.Dokumen||"").trim()})).filter(r=>r.tanggal&&r.employee_name&&r.from_branch_id&&r.to_branch_id),l=await f("/api/mutasi/import",{method:"POST",body:JSON.stringify(o)});if(!l.ok)throw new Error(l.data?.error||"Import gagal")}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:et},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:fe,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:fe,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}window.parseFlexibleDate=a=>{if(!a||a==="-")return"";if(a=String(a).trim(),/^\d{5}$/.test(a)){let e=Math.floor(Number(a)-25569);return new Date(e*86400*1e3).toISOString().split("T")[0]}if(a.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let e=a.split(/[\/\-]/);return`${e[2]}-${e[1]}-${e[0]}`}return a.split("T")[0]};window.formatDate=a=>{let e=window.parseFlexibleDate(a);if(!e)return"";let i=e.split("-");if(i.length===3&&i[0].length===4){let t=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],o=parseInt(i[2],10),l=t[parseInt(i[1],10)-1];return`${o} ${l} ${i[0]}`}return e};function K(a){return async e=>{if(!xe()){pe("/login");return}return a(e)}}var Be=null;function Mt(){Be&&clearInterval(Be);let a=()=>{let e=new Date,i=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),t=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),o=document.getElementById("header-clock-time"),l=document.getElementById("header-clock-date");o&&(o.textContent=i),l&&(l.textContent=t)};a(),Be=setInterval(a,1e3)}async function Kt(){try{let a=await f("/api/dashboard/kpi");if(!a.ok)return;let e=a.data?.data||a.data||{},i=(t,o)=>{let l=document.getElementById(t);l&&(l.textContent=o>0?o:"",l.style.display=o>0?"inline-flex":"none")};i("badge-issues",e.issues?.current||0),i("badge-contracts",e.expiring30?.current||0),i("badge-oo1",e.one_on_one?.current||0),i("badge-schedule",e.schedule?.current||0),i("badge-supply",e.supply?.current||0)}catch{}}var ve=[];async function Rt(){try{let a=await f("/api/dashboard/notifications");if(!a.ok)return;ve=a.data?.data||a.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=ve.length>0?"block":"none",e.textContent=ve.length)}catch{}}function qt(){if(!ve.length){te({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,i)=>i()});return}let a=`
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
  `;te({title:`Notifikasi (${ve.length})`,content:a,confirmText:"Tutup",onConfirm:(e,i)=>i()})}function tt(){let a=oe(),e=(a?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
          ${a&&(a.role==="superadmin"||a.role==="admin")?`
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
  `;let i=document.getElementById("sidebar"),t=document.getElementById("sidebar-overlay"),o=document.getElementById("topbar-menu-btn"),l=document.getElementById("sidebar-close"),r=()=>{i.classList.add("open"),t.classList.add("show")},s=()=>{i.classList.remove("open"),t.classList.remove("show")};o?.addEventListener("click",r),l?.addEventListener("click",s),t?.addEventListener("click",s),document.querySelectorAll(".nav-item").forEach(c=>c.addEventListener("click",s));function n(){let c=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(d=>{let b=d.dataset.route;d.classList.toggle("active",c===b||b!=="/dashboard"&&c.startsWith(b))});let m=document.getElementById("topbar-title"),p=document.querySelector(".nav-item.active .nav-label");m&&p&&(m.textContent=p.textContent)}window.addEventListener("hashchange",n),n(),Mt(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await f("/api/auth/logout",{method:"POST"}),Ce(),Be&&clearInterval(Be),pe("/login")}),Kt(),Rt(),document.getElementById("btn-notif")?.addEventListener("click",c=>{c.preventDefault(),qt()})}async function Ht(){O("/login",({main:e})=>Da(e)),O("/dashboard",K(({main:e})=>Ta(e))),O("/calendar",K(({main:e})=>Ja(e))),O("/employees",K(({main:e,params:i})=>Ia(e,i))),O("/contracts",K(({main:e,params:i})=>Me(e,i))),O("/sp",K(({main:e})=>Za(e))),O("/mutasi",K(({main:e})=>at(e))),O("/timeline",K(({main:e,params:i})=>Ba(e,i))),O("/issues",K(({main:e,params:i})=>La(e,i))),O("/one-on-one",K(({main:e,params:i})=>Aa(e,i))),O("/training",K(({main:e})=>Na(e))),O("/relievers",K(({main:e,params:i})=>Oa(e,i))),O("/reports/inspection",K(({main:e})=>Fa(e))),O("/reports/cleaning",K(({main:e})=>Ma(e))),O("/reports/fogging",K(({main:e})=>Ka(e))),O("/reports/basecamp",K(({main:e})=>Ra(e))),O("/reports/supply",K(({main:e})=>ua(e,"supply"))),O("/sop",K(({main:e})=>qa(e))),O("/checklist",K(({main:e})=>Ha(e))),O("/forms",K(({main:e})=>ua(e))),O("/users",K(({main:e})=>ja(e))),O("/branches",K(({main:e})=>Ua(e))),O("/profile",K(({main:e})=>Ga(e))),O("/settings/import",K(({main:e})=>za(e)));let a=xe();if(!a&&window.location.hash!=="#/login"&&pe("/login"),a){let e=await f("/api/auth/me");e.ok?(Te(e.data.data),tt()):(Ce(),pe("/login"))}window.addEventListener("fm:login",()=>{tt(),pe("/dashboard")}),ba()}Ht();
