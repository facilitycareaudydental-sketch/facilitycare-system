var ca=Object.defineProperty;var ze=(t,e)=>()=>(t&&(e=t(t=0)),e);var We=(t,e)=>{for(var i in e)ca(t,i,{get:e[i],enumerable:!0})};var he={};We(he,{API:()=>yt,CLIENT_SIDE_MAX_ROWS:()=>$e,IS_DEVELOPMENT:()=>Ye,apiFetch:()=>v,clearToken:()=>De,getToken:()=>Ee,getUser:()=>ue,setToken:()=>Xe,setUser:()=>Ie});function Ee(){return localStorage.getItem("fm_token")}function Xe(t){localStorage.setItem("fm_token",t)}function De(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ue(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function Ie(t){localStorage.setItem("fm_user",JSON.stringify(t))}async function v(t,e={}){let i=Ee(),a={"Content-Type":"application/json",...i?{Authorization:`Bearer ${i}`}:{},...e.headers||{}};try{let o=t;o.startsWith("/api/")&&!o.startsWith("/api/v1/")&&(o="/api/v1/"+o.substring(5));let c=`cb=${Date.now()}`,s=o.includes("?")?"&":"?",r=`${yt}${o}${s}${c}`,n=await fetch(r,{...e,headers:a}),l;try{let m=await n.text();try{l=JSON.parse(m)}catch{l={error:`Server Error (${n.status}): ${m.substring(0,80)}...`}}}catch{l={error:"Gagal membaca respon dari server"}}return n.status===401&&(De(),window.location.hash="#/login"),{ok:n.ok,status:n.status,data:l}}catch(o){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${o.message})`}}}}var Ye,da,yt,$e,I=ze(()=>{Ye=!1,da="https://fm-operations-api.facilitycare-audydental.workers.dev",yt=da,$e=1e4});var kt={};We(kt,{confirmDialog:()=>et,createModal:()=>se});function se({title:t,content:e,onConfirm:i,onCancel:a,confirmText:o="Simpan",cancelText:c="Batal",size:s="md",confirmClass:r="btn-primary"}){let n={sm:"400px",md:"560px",lg:"720px",xl:"900px"},l=document.createElement("div");l.className="modal-overlay",l.innerHTML=`
    <div class="modal" style="max-width:${n[s]||n.md}">
      <div class="modal-header">
        <h3 class="modal-title">${t}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${c}</button>
        ${i?`<button class="btn ${r} modal-confirm">${o}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&l.querySelector(".modal-body").appendChild(e);let m=()=>{l.classList.remove("show"),setTimeout(()=>l.remove(),250)};return l.querySelector(".modal-close").addEventListener("click",()=>{a&&a(),m()}),l.querySelector(".modal-cancel").addEventListener("click",()=>{a&&a(),m()}),i&&l.querySelector(".modal-confirm").addEventListener("click",()=>i(l,m)),l.addEventListener("click",p=>{p.target===l&&(a&&a(),m())}),document.body.appendChild(l),requestAnimationFrame(()=>l.classList.add("show")),{overlay:l,close:m}}function et(t,e,i="Konfirmasi"){return se({title:i,content:`<p>${t}</p>`,onConfirm:(a,o)=>{e(),o()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}var Be=ze(()=>{});var X={};We(X,{downloadExcel:()=>E,parseExcel:()=>at,renderExcelButtons:()=>nt});function at(t){return new Promise((e,i)=>{let a=new FileReader;a.onload=o=>{try{let c=new Uint8Array(o.target.result),s=XLSX.read(c,{type:"array"}),r=s.SheetNames[0],n=s.Sheets[r];console.log("--- START EXCEL PARSING ---"),console.log(`File Name: ${t.name}`),console.log(`File Size: ${(t.size/1024).toFixed(2)} KB`),console.log(`File Type: ${t.type||"unknown"}`),console.log(`Sheets Found: ${s.SheetNames.join(", ")}`),console.log(`Sheet Used: ${r}`);let l=XLSX.utils.decode_range(n["!ref"]||"A1:A1"),m=l.e.r-l.s.r+1,p=l.e.c-l.s.c+1;console.log(`Total Rows (including empty): ${m}`),console.log(`Total Columns: ${p}`);let d=[];for(let h=l.s.c;h<=l.e.c;++h){let b=n[XLSX.utils.encode_cell({c:h,r:l.s.r})];b&&b.v&&d.push(b.v)}console.log(`Headers Found: ${d.join(", ")}`),console.log("---------------------------");let g=XLSX.utils.sheet_to_json(n,{defval:""});Object.defineProperty(g,"__worksheet",{value:n,enumerable:!1}),Object.defineProperty(g,"__headers",{value:d,enumerable:!1}),e(g)}catch(c){i(c)}},a.onerror=o=>i(o),a.readAsArrayBuffer(t)})}function E(t,e){try{let i=XLSX.utils.json_to_sheet(t),a=XLSX.utils.book_new();XLSX.utils.book_append_sheet(a,i,"Data"),XLSX.writeFile(a,`${e}.xlsx`)}catch(i){throw console.error("Error generating Excel file:",i),i}}function nt(t){return`
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
  `}var L=ze(()=>{});I();var Ze={},Re=null;function M(t,e){Ze[t]=e}function ye(t){window.location.hash=t}function ft(){async function t(){let e=window.location.hash.replace("#","")||"/dashboard",[i,...a]=e.split("?"),o=Ze[i];if(!o){for(let[s,r]of Object.entries(Ze))if(s.endsWith("/*")&&i.startsWith(s.slice(0,-2))){o=r;break}}Re&&(Re(),Re=null);let c=document.getElementById("main-content");if(c&&(c.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),o){let s=new URLSearchParams(a.join("?")),r=i.split("/").filter(Boolean),n=await o({path:i,params:s,segments:r,main:c});n&&(Re=n)}else{let s=c||document.getElementById("app");s&&(s.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",t),t()}var Pe;function pa(){return Pe||(Pe=document.createElement("div"),Pe.id="toast-container",document.body.appendChild(Pe)),Pe}function vt(t,e="info",i=3500){let a=pa(),o=document.createElement("div");o.className=`toast toast-${e}`;let c={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};o.innerHTML=`<span class="toast-icon">${c[e]||"\u2139"}</span><span class="toast-msg">${t}</span>`,a.appendChild(o),requestAnimationFrame(()=>o.classList.add("show")),setTimeout(()=>{o.classList.remove("show"),setTimeout(()=>o.remove(),350)},i)}var V=t=>vt(t,"success"),W=t=>vt(t,"error");Be();I();I();function St({columns:t,data:e,onEdit:i,onDelete:a,onView:o,actions:c=[],emptyText:s="Tidak ada data",bulkSelect:r=null}){let n=document.createElement("div");if(n.className="table-wrapper",!e||e.length===0)return n.innerHTML=`<div class="empty-state"><p>${s}</p></div>`,n;let l=document.createElement("table");l.className="data-table";let m=document.createElement("thead"),p=document.createElement("tr");if(r){let g=document.createElement("th");g.style.width="40px",g.style.textAlign="center";let h=document.createElement("input");h.type="checkbox",h.id="select-all-checkbox",h.title="Pilih semua",h.addEventListener("change",()=>{e.forEach(b=>{h.checked?r.selectedIds.add(b.id):r.selectedIds.delete(b.id)}),n.querySelectorAll(".row-checkbox").forEach(b=>b.checked=h.checked),r.onToggle()}),g.appendChild(h),p.appendChild(g)}if(t.forEach(g=>{let h=document.createElement("th");h.textContent=g.label,g.width&&(h.style.width=g.width),p.appendChild(h)}),i||a||o||c.length>0){let g=document.createElement("th");g.textContent="Aksi",g.style.width="120px",p.appendChild(g)}m.appendChild(p),l.appendChild(m);let d=document.createElement("tbody");return e.forEach(g=>{let h=document.createElement("tr");if(r){let b=document.createElement("td");b.style.textAlign="center",b.style.width="40px";let u=document.createElement("input");u.type="checkbox",u.className="row-checkbox",u.value=g.id,u.checked=r.selectedIds.has(g.id),u.addEventListener("change",()=>{if(u.checked)r.selectedIds.add(g.id);else{r.selectedIds.delete(g.id);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1)}r.onToggle()}),b.appendChild(u),h.appendChild(b)}if(t.forEach(b=>{let u=document.createElement("td");if(b.render){let y=b.render(g[b.key],g);y instanceof HTMLElement?u.appendChild(y):u.innerHTML=y||""}else u.textContent=g[b.key]!==null&&g[b.key]!==void 0&&g[b.key]!==""?g[b.key]:"";b.nowrap&&(u.style.whiteSpace="nowrap"),h.appendChild(u)}),i||a||o||c.length>0){let b=document.createElement("td");b.className="actions-cell";let u=document.createElement("div");if(u.className="btn-group",o){let y=document.createElement("button");y.className="btn btn-xs btn-ghost",y.innerHTML="\u{1F441}",y.title="Lihat",y.addEventListener("click",()=>o(g)),u.appendChild(y)}if(i){let y=document.createElement("button");y.className="btn btn-xs btn-secondary",y.innerHTML="\u270F\uFE0F",y.title="Edit",y.addEventListener("click",()=>i(g)),u.appendChild(y)}c.forEach(y=>{let S=document.createElement("button");S.className=`btn btn-xs ${y.class||"btn-ghost"}`,S.innerHTML=y.icon||y.label,S.title=y.label,S.addEventListener("click",()=>y.handler(g)),u.appendChild(S)}),b.appendChild(u),h.appendChild(b)}d.appendChild(h)}),l.appendChild(d),n.appendChild(l),n}function wt({page:t,pages:e,total:i,limit:a,onPage:o}){if(e<=1)return null;let c=document.createElement("div");c.className="pagination";let s=document.createElement("span");s.className="pagination-info",s.textContent=`Total: ${i} data`,c.appendChild(s);let r=document.createElement("div");r.className="pagination-btns";let n=(p,d,g=!1,h=!1)=>{let b=document.createElement("button");b.className=`btn btn-sm ${h?"btn-primary":"btn-ghost"} pagination-btn`,b.textContent=p,b.disabled=g,b.addEventListener("click",()=>o(d)),r.appendChild(b)};n("\xAB",1,t===1),n("\u2039",t-1,t===1);let l=Math.max(1,t-2),m=Math.min(e,t+2);for(let p=l;p<=m;p++)n(p,p,!1,p===t);return n("\u203A",t+1,t===e),n("\xBB",e,t===e),c.appendChild(r),c}Be();function tt(t){return t.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="html")return e.html||"";if(e.type==="row")return`<div class="form-row">${tt(e.fields)}</div>`;let i=e.required?"required":"",a=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",o="";switch(e.type){case"textarea":o=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${i} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let s=(e.options||[]).map(p=>{let d=typeof p=="object"?p.value:p,g=typeof p=="object"?p.label:p,h=e.value==d?"selected":"";return`<option value="${d}" ${h}>${g}</option>`}).join("");o=`<select name="${e.name}" class="form-control" ${i}><option value="">-- Pilih ${e.label||""} --</option>${s}</select>`;break;case"combobox":let r=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,n=(e.options||[]).map(p=>{let d=typeof p=="object"?p.value:p,g=typeof p=="object"?p.label||p.value||"":p||"";return(g==="undefined"||g==="[object Object]"||g==="null")&&(g=""),g?`<option value="${g}"></option>`:""}).join(""),l=e.value||"";if(e.value){let p=(e.options||[]).find(d=>(typeof d=="object"?d.value:d)==e.value);if(p){let d=typeof p=="object"?p.label||p.value||"":p||"";d&&d!=="undefined"&&d!=="[object Object]"&&d!=="null"&&(l=d)}}o=`
          <input type="text" name="${e.name}" list="${r}" class="form-control" value="${l}" placeholder="Pilih atau ketik baru..." ${i} autocomplete="off">
          <datalist id="${r}">${n}</datalist>
        `;break;case"checkbox":o=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":let m=window.parseFlexibleDate&&e.value?window.parseFlexibleDate(e.value):e.value||"";o=`<input type="date" name="${e.name}" class="form-control" value="${m}" ${i}>`;break;case"number":o=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${i}>`;break;case"email":o=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i}>`;break;case"url":o=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${i}>`;break;default:o=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${i} autocomplete="off">`}let c=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${a}${o}${c}</div>`}).join("")}function _t(t){let e={},i=new FormData(t);for(let[a,o]of i.entries())e[a]=o===""?null:o;return t.querySelectorAll("input[type=checkbox]").forEach(a=>{a.checked||(e[a.name]=null)}),e}function xt(t,e){e&&Object.entries(e).forEach(([i,a])=>{let o=t.querySelector(`[name="${i}"]`);o&&(o.hasAttribute("list")||(o.type==="checkbox"?o.checked=!!a:o.type==="date"&&a&&window.parseFlexibleDate?o.value=window.parseFlexibleDate(a):o.value=a??""))})}L();var oe={},Le={on(t,e){oe[t]||(oe[t]=new Set),oe[t].add(e)},off(t,e){oe[t]&&oe[t].delete(e)},emit(t,e){oe[t]&&oe[t].forEach(i=>{try{i(e)}catch(a){console.warn("[calendarBus] Handler error:",a)}})},clear(){Object.keys(oe).forEach(t=>delete oe[t])}},ma=new Set(["schedule","cleaning","cleaning_reports","inspection","inspection_reports","fogging","fogging_reports","reliever","relievers","contract","contracts","issue","issues","training","one_on_one","sp","sp_data","mutasi","basecamp","basecamp_reports","supply"]);function ge(t){if(!t){Le.emit("data:changed",{module:"unknown"});return}let e=String(t).toLowerCase().replace(/^\/api\//,"").replace(/^reports\//,"");Le.emit("data:changed",{module:e,relevant:ma.has(e)})}function $({container:t,title:e,icon:i,apiPath:a,columns:o,formFields:c,filterFields:s,defaultFilters:r={},itemLabel:n="Data",canCreate:l=!0,canEdit:m=!0,canDelete:p=!0,onBeforeSubmit:d,onAfterLoad:g,onDataLoaded:h,extraActions:b=[],initialSearch:u="",exportOptions:y=null,bulkDelete:S=!1,paginationMode:k="server"}){let w=1,f={...r};u&&(f.search=u);let C=new Set;t.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${i} ${e}</h1>
      <div class="page-actions">
        ${l?`<button class="btn btn-primary" id="btn-create">+ Tambah ${n}</button>`:""}
      </div>
    </div>

    ${S?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${y?nt(y.moduleName):""}

    ${s&&s.length>0?`
    <div class="filter-bar card">
      <div class="filter-bar-inner">
        ${s.map(x=>x.type==="search"?`<div class="filter-search"><input type="search" class="form-control" placeholder="${x.placeholder||"Cari..."}" id="filter-search" value="${f.search||""}"></div>`:x.type==="select"||x.type==="combobox"?`<select class="form-control filter-select" name="${x.name}" id="filter-${x.name}"><option value="">-- ${x.label} --</option>${(x.options||[]).map(_=>`<option value="${typeof _=="object"?_.value:_}" ${f[x.name]===(typeof _=="object"?_.value:_)?"selected":""}>${typeof _=="object"?_.label:_}</option>`).join("")}</select>`:"").join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
      </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function P(){if(!document.getElementById("bulk-toolbar"))return;let _=document.getElementById("bulk-count"),D=document.getElementById("btn-bulk-delete"),G=document.getElementById("btn-bulk-cancel");_.textContent=`${C.size} item dipilih`,C.size>0?(D.disabled=!1,G.disabled=!1):(D.disabled=!0,G.disabled=!0);let U=document.getElementById("select-all-checkbox");if(U){let Y=document.querySelectorAll(".row-checkbox");if(Y.length>0){let F=[...Y].every(me=>me.checked),q=[...Y].some(me=>me.checked);U.checked=F,U.indeterminate=q&&!F}else U.checked=!1,U.indeterminate=!1}}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{C.clear(),document.querySelectorAll(".row-checkbox").forEach(_=>_.checked=!1);let x=document.getElementById("select-all-checkbox");x&&(x.checked=!1),P()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(C.size===0)return;let x=[...C],_=document.createElement("div");_.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",_.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${x.length} ${n}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${x.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(_),_.querySelector("#bulk-cancel-btn").addEventListener("click",()=>_.remove()),_.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let D=_.querySelector("#bulk-confirm-btn");D.disabled=!0,D.textContent="Menghapus...";let G=await v(`${a}/bulk`,{method:"DELETE",body:JSON.stringify({ids:x})});_.remove(),G.ok?(V(`${x.length} ${n} berhasil dihapus.`),C.clear(),P(),ge(a),O()):W(G.data?.error||"Gagal menghapus data.")})});let T=document.getElementById("filter-search"),R;if(T?.addEventListener("input",x=>{clearTimeout(R),R=setTimeout(()=>{f.search=x.target.value,w=1,C.clear(),P(),O()},400)}),s?.forEach(x=>{(x.type==="select"||x.type==="combobox")&&document.getElementById(`filter-${x.name}`)?.addEventListener("change",_=>{f[x.name]=_.target.value,w=1,C.clear(),P(),O()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{f={},T&&(T.value=""),s?.forEach(x=>{let _=document.getElementById(`filter-${x.name}`);_&&(_.value="")}),w=1,C.clear(),P(),O()}),document.getElementById("btn-create")?.addEventListener("click",()=>J(null)),y){document.getElementById(`btn-export-${y.moduleName}`)?.addEventListener("click",async _=>{let D=_.target,G=D.innerHTML;D.innerHTML="\u23F3 Loading...",D.disabled=!0;try{await y.onExport()}catch{W("Gagal export data")}finally{D.innerHTML=G,D.disabled=!1}}),document.getElementById(`btn-template-${y.moduleName}`)?.addEventListener("click",()=>{y.onTemplate()});let x=document.getElementById(`input-import-${y.moduleName}`);x?.addEventListener("change",async _=>{let D=_.target.files[0];if(!D)return;let G=document.getElementById(`label-import-${y.moduleName}`),U=G?G.querySelector(".import-text"):null,Y=U?U.innerText:"";U&&(U.innerText="\u231B Memproses..."),G&&(G.style.pointerEvents="none"),x.disabled=!0;try{let F=await at(D);if(F.length===0)throw new Error("File kosong atau format salah");await y.onImport(F),V("Import berhasil!"),ge(a),O()}catch(F){W(F.message||"Gagal import data")}finally{U&&(U.innerText=Y),G&&(G.style.pointerEvents="auto"),x.disabled=!1,x.value=""}})}async function O(){P();let x=document.getElementById("table-container");if(!x)return;x.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let _=k==="client",D=_?1:w,G=_?$e:20,U=new URLSearchParams({page:D,limit:G,...Object.fromEntries(Object.entries(f).filter(([,z])=>z))}),Y=await v(`${a}?${U}`);if(!Y.ok){x.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${Y.data?.error||"Error"}</p></div>`;return}let F=Y.data?.data||Y.data||[],q=Y.data?.pagination,me=F.length;if(_){F=h(F);let z=F.length,ae=20,le=Math.ceil(z/ae);w>le&&le>0&&(w=le);let Q=(w-1)*ae,be=w*ae;F=F.slice(Q,be),q={page:w,limit:ae,total:z,pages:le}}!1,g&&g(F);let Me=St({columns:o,data:F,onEdit:m?z=>J(z):null,actions:b.map(z=>({...z,handler:ae=>z.handler(ae,O)})),emptyText:`Tidak ada ${String(n||"").toLowerCase()}`,bulkSelect:S?{selectedIds:C,onToggle:P}:null});x.innerHTML="",x.appendChild(Me);let Te=document.getElementById("pagination-container");if(Te&&(Te.innerHTML="",q&&q.pages>1)){let z=wt({page:q.page,pages:q.pages,total:q.total,limit:q.limit,onPage:ae=>{w=ae,O()}});z&&Te.appendChild(z)}}function A(x){let _=typeof c=="function"?c(x):c;return tt(_)}function J(x){let _=!!x,D=document.createElement("form");if(D.noValidate=!0,D.innerHTML=A(x),_){let U=typeof c=="function"?c(x):c;xt(D,x)}let{close:G}=se({title:_?`Edit ${n}`:`Tambah ${n}`,content:D,size:"lg",confirmText:_?"Simpan Perubahan":`Tambah ${n}`,onConfirm:async(U,Y)=>{if(!D.reportValidity())return;let F=U.querySelector(".modal-confirm");F.disabled=!0,F.textContent="Menyimpan...";let q=_t(D),me=typeof c=="function"?c(x):c,Me=async le=>{for(let Q of le)if(Q.type==="row")await Me(Q.fields);else if(Q.type==="combobox"&&q[Q.name]){let be=q[Q.name],Ke=(Q.options||[]).find(re=>{let ce=String(typeof re=="object"?re.value:re),la=String(typeof re=="object"?re.label:re);return ce===be||la===be});if(Ke)q[Q.name]=typeof Ke=="object"?Ke.value:Ke;else if(Q.createApi){let re={};re[Q.createApi.field]=be,Q.createApi.extra&&Object.assign(re,Q.createApi.extra);let ce=await v(Q.createApi.path,{method:"POST",body:JSON.stringify(re)});if(ce.ok&&ce.data?.id)q[Q.name]=ce.data.id;else if(ce.ok&&!ce.data?.id)q[Q.name]=be;else throw new Error(`Gagal membuat master data: ${ce.data?.error||"Unknown error"}`)}}};try{await Me(me)}catch(le){W(le.message),F.disabled=!1,F.textContent=_?"Simpan Perubahan":`Tambah ${n}`;return}d&&(q=await d(q,x));let Te=_?"PUT":"POST",z=_?`${a}/${x.id}`:a,ae=await v(z,{method:Te,body:JSON.stringify(q)});ae.ok?(V(_?`${n} berhasil diperbarui.`:`${n} berhasil ditambahkan.`),Y(),ge(a),O()):(W(ae.data?.error||"Gagal menyimpan data."),F.disabled=!1,F.textContent=_?"Simpan Perubahan":`Tambah ${n}`)}})}function te(x){et(`Hapus ${n} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let _=await v(`${a}/${x.id}`,{method:"DELETE"});_.ok?(V(`${n} berhasil dihapus.`),ge(a),O()):W(_.data?.error||"Gagal menghapus.")},`Hapus ${n}`)}return O(),O}I();I();var fe=null,qe=null;async function ve(t=!1){if(fe&&!t)return console.log("Employees Raw (Cache Hit)",fe.slice(0,5)),fe;let e=await v(`/api/employees?limit=${$e}&status=Aktif`);return fe=(e.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),console.log("Employees Raw",e.data?.data?.slice(0,5)),console.log("Employees Mapped (ID)",fe.slice(0,5)),fe}async function H(t=!1){let i=(await ve(t)).map(a=>({value:a.label,label:a.label}));return console.log("Employee Options",i.slice(0,5)),i}async function B(t=!1){return qe&&!t||(qe=((await v("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}))),qe}function N(t){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!t||t==="-"||String(t).trim()===""?"":`<span class="badge ${e[t]||"badge-neutral"}">${t}</span>`}function it(t){return t==null?'<span class="badge badge-neutral">-</span>':t<0?`<span class="badge badge-danger">Expired (${Math.abs(t)}h)</span>`:t<=14?`<span class="badge badge-danger">${t} hari</span>`:t<=30?`<span class="badge badge-warning">${t} hari</span>`:`<span class="badge badge-success">${t} hari</span>`}function ke(t){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[t]||"badge-neutral"}">${t||"-"}</span>`}function rt(t){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[t]||"badge-neutral"}">${t||"-"}</span>`}function ne(t){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[t]||"badge-neutral"}">${t||"-"}</span>`}L();function st(t,e){if(t.period!=="Q3")return!1;let i=String(t.status||"").toLowerCase();if(i!=="selesai"&&i!=="completed"&&i!=="done")return!1;let a=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?a.includes("inspeksi"):e==="gcdc"?a.includes("general cleaning")||a.includes("deep cleaning"):!1}I();L();function Ct(t,e){let i=String(t.status||"").toLowerCase();return e==="active"?i==="aktif":!1}I();L();function ot(t,e){if(String(t.status||"").toLowerCase()!=="aktif")return!1;if(e==="active")return!0;if(e==="expiring30"){if(!t.end_date)return!1;let a=new Date;a.setHours(0,0,0,0);let o=new Date(a);o.setDate(a.getDate()+30);let c=new Date(t.end_date);return c.setHours(0,0,0,0),c>=a&&c<=o}return!1}I();L();function Tt(t,e){let i=String(t.status||"").toLowerCase();return e==="open"?i==="open":!1}I();function $t(t,e){let i=String(t.status||"").toLowerCase();return e==="pending"?i==="pending":!1}var pe={};function Ae(t){if(pe[t]){try{pe[t].destroy()}catch{}delete pe[t]}}function ua(){Object.keys(pe).forEach(Ae)}var ie=(t,e=0)=>{let i=Number(t);return isNaN(i)||t===null||t===void 0?e:i},Se=(t,e="\u2014")=>{if(t==null||t==="")return e;let i=String(t).trim();return i===""||i==="[object Object]"?e:i};var Dt=t=>{if(!t||typeof t!="string")return"";try{let[e,i]=t.split("-");return new Date(Number(e),Number(i)-1).toLocaleDateString("id-ID",{month:"short",year:"2-digit"})}catch{return t}};function It(t,e,i=900){if(!t)return;let a=Math.max(0,Math.round(ie(e)));if(a===0){t.textContent="0";return}let o=Date.now(),c=()=>{let s=Math.min((Date.now()-o)/i,1),r=1-Math.pow(1-s,3);t.textContent=Math.round(r*a).toLocaleString("id-ID"),s<1?requestAnimationFrame(c):t.textContent=a.toLocaleString("id-ID")};requestAnimationFrame(c)}var ga={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},ba=t=>{let e=Se(t,"\u2014");return`<span class="status-pill ${ga[e]||"pill-neutral"}">${e}</span>`};var ee={family:"Inter",size:11},we="#94A3B8",He="#F1F5F9",lt=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],ha=()=>window.innerWidth<768;function pt(t={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:ha()?"bottom":"top",labels:{font:ee,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:ee,titleFont:{...ee,weight:"700"}}},scales:{x:{grid:{color:He},ticks:{font:ee,color:we,maxRotation:0}},y:{grid:{color:He},ticks:{font:ee,color:we},beginAtZero:!0}},...t}}var ya=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join("");function ct(t=3){return Array(t).fill(0).map((e,i)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<t-1?"6px":"0"};border-radius:6px"></div>`).join("")}async function Z(t,e,i=8e3){try{let a=new AbortController,o=setTimeout(()=>a.abort(),i),c=await v(t,{signal:a.signal}).catch(()=>null);if(clearTimeout(o),!c||!c.ok)return e;let s=c.data;return s?s.data!==void 0?s.data??e:s:e}catch{return e}}function va(){["skel-donut","skel-trend","skel-insp","skel-contract"].forEach(a=>{let o=document.getElementById(a);o&&(o.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract"].forEach(a=>{let o=document.getElementById(a);if(o&&o.style.display==="none"){o.style.display="block";let c=o.parentElement;if(c&&!c.querySelector(".chart-empty")){let s=document.createElement("div");s.className="chart-empty",s.textContent="Belum ada data",o.style.display="none",c.appendChild(s)}}});let t=document.getElementById("kpi-row");t&&t.querySelector(".skeleton")&&Bt({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Lt({}),["table-contracts","table-issues"].forEach(a=>{let o=document.getElementById(a);o&&o.querySelector(".skeleton")&&(o.innerHTML='<div class="chart-empty">Belum ada data</div>')});let i=document.getElementById("activity-log");i&&i.querySelector(".skeleton")&&(i.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function Pt(t){ua(),t._dashRefresh&&clearInterval(t._dashRefresh),t._skelTimeout&&clearTimeout(t._skelTimeout);let e=new Date().toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});t.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">


      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${ya()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${fa()}</div>

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
        <!-- Kontrak Akan Habis -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Kontrak Akan Habis</div>
            <a href="#/contracts" class="chart-link">Lihat Data</a>
          </div>
          <div class="chart-canvas-wrap" style="height:140px;position:relative;margin-top:10px">
            <div id="skel-contract-mini" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-contract-mini" style="display:none"></canvas>
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
          <div id="widget-agenda" class="dash-table-wrap" style="height:160px;overflow-y:auto;overflow-x:hidden">${ct(3)}</div>
        </div>
        <!-- KPI Kebersihan -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">KPI Kebersihan</div>
          </div>
          <div id="widget-kpi-kebersihan" style="margin-top:0px">${ct(4)}</div>
        </div>
        <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:160px;overflow-y:auto">${ct(3)}</div>
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
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>dt(t)),t._skelTimeout=setTimeout(()=>va(),5e3),await dt(t),t._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?dt(t):clearInterval(t._dashRefresh)},6e4)}async function dt(t){t._skelTimeout&&(clearTimeout(t._skelTimeout),t._skelTimeout=null);let[e,i,a,o,c,s,r,n,l,m,p,d]=await Promise.all([Z("/api/dashboard/kpi",{},8e3),Z("/api/dashboard/issues-trend",{},8e3),Z("/api/dashboard/issues-summary",{},8e3),Z("/api/dashboard/inspection-bar",{},8e3),Z("/api/dashboard/stats",{},8e3),Z("/api/dashboard/calendar",[],8e3),Z("/api/schedule?limit=10000",{data:[]},8e3),Z("/api/employees?limit=10000",{data:[]},8e3),Z("/api/contracts?limit=10000",{data:[]},8e3),Z("/api/issues?limit=10000",{data:[]},8e3),Z("/api/one_on_one?limit=10000",{data:[]},8e3),Z("/api/dashboard/contracts-chart",{},8e3)]);if(e){let h=Array.isArray(r?.data)?r.data:Array.isArray(r)?r:[],b=Array.isArray(n?.data)?n.data:Array.isArray(n)?n:[],u=Array.isArray(l?.data)?l.data:Array.isArray(l)?l:[],y=Array.isArray(m?.data)?m.data:Array.isArray(m)?m:[],S=Array.isArray(p?.data)?p.data:Array.isArray(p)?p:[];e.employees&&(e.employees.current=b.filter(k=>Ct(k,"active")).length),e.contracts&&(e.contracts.current=u.filter(k=>ot(k,"active")).length),e.expiring30&&(e.expiring30={current:u.filter(k=>ot(k,"expiring30")).length}),e.issues&&(e.issues.current=y.filter(k=>Tt(k,"open")).length),e.one_on_one&&(e.one_on_one.current=S.filter(k=>$t(k,"pending")).length),e.inspection_month&&(e.inspection_month.current=h.filter(k=>st(k,"inspeksi")).length),e.cleaning_month&&(e.cleaning_month.current=h.filter(k=>st(k,"gcdc")).length)}try{Bt(e)}catch(h){console.warn("KPI render:",h)}try{Lt(e)}catch(h){console.warn("MiniStats render:",h)}try{ka(Array.isArray(a?.by_category)?a.by_category:[])}catch(h){console.warn("Donut render:",h),de("skel-donut","chart-donut")}try{Sa(i)}catch(h){console.warn("Trend render:",h),de("skel-trend","chart-trend")}try{Et(o)}catch(h){console.warn("InspBar render:",h),de("skel-insp","chart-insp")}try{let h=Array.isArray(c)?c:Array.isArray(c?.recent_issues)?c.recent_issues:[];_a(h)}catch(h){console.warn("IssuesTable render:",h)}try{wa(d)}catch(h){console.warn("ContractsTable render:",h)}try{xa(Array.isArray(s)?s:[])}catch(h){console.warn("Agenda render:",h)}try{let h=Array.isArray(r?.data)?r.data:Array.isArray(r)?r:[],b=Array.isArray(m?.data)?m.data:Array.isArray(m)?m:[];Ca(h,b,o)}catch(h){console.warn("KPI Kebersihan render:",h)}try{Ta()}catch(h){console.warn("Quick Actions render:",h)}let g=document.getElementById("insp-month-filter");if(g){let h=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];g.innerHTML='<option value="">12 Bulan Terakhir</option>';let b=new Date().getFullYear();h.forEach((u,y)=>{let S=`${b}-${String(y+1).padStart(2,"0")}`,k=document.createElement("option");k.value=S,k.textContent=u+" "+b,g.appendChild(k)}),g.addEventListener("change",async u=>{let y=u.target.value,S=document.getElementById("skel-insp"),k=document.getElementById("chart-insp");S&&(S.style.display="block"),k&&(k.style.display="none");let w=y?`?month=${y}`:"",f=await v("/api/dashboard/inspection-bar"+w);if(f&&f.ok){let C=f.data?.data!==void 0?f.data.data:f.data;Et(C)}else de("skel-insp","chart-insp")})}}function Bt(t){let e=document.getElementById("kpi-row");if(!e)return;t=t||{};let i=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees?dash_filter=active",color:"kpi-blue",key:"employees",trendPct:"+2%",trendColor:"#10B981",points:"0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts?dash_filter=active",color:"kpi-green",key:"contracts",trendPct:"+1%",trendColor:"#10B981",points:"0,15 20,18 40,10 60,12 80,5 100,2"},{icon:"\u23F3",label:"Kontrak Habis 30 Hari",sub:"Akan segera berakhir",href:"#/contracts?dash_filter=expiring30",color:"kpi-warn",key:"expiring30",trendPct:"+25%",trendColor:"#F59E0B",points:"0,25 20,22 40,24 60,15 80,18 100,5"},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues?dash_filter=open",color:"kpi-red",key:"issues",trendPct:"0%",trendColor:"#EF4444",points:"0,20 20,18 40,22 60,19 80,21 100,20"},{icon:"\u{1F4AC}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one?dash_filter=pending",color:"kpi-purple",key:"one_on_one",trendPct:"+8%",trendColor:"#10B981",points:"0,25 20,15 40,18 60,8 80,10 100,2"}];e.innerHTML=i.map(a=>{let o=ie(t[a.key]?.current,0);return`
      <a href="${a.href}" class="kpi-card ${a.color}" style="text-decoration:none;padding:10px 12px">
        <div style="display:flex; gap:10px; align-items:center;">
          <div class="kpi-icon-wrap" style="width:40px;height:40px;border-radius:10px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${a.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${o}" style="font-size:1.6rem; font-weight:800; line-height:1; color:var(--text-1)">${o}</div>
            <div class="kpi-label" style="font-size:0.75rem; font-weight:700; color:var(--text-2); margin-top:6px">${a.label}</div>
            <div class="kpi-subtitle" style="font-size:0.65rem; color:var(--text-3); margin-top:2px; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis">${a.sub}</div>
          </div>
        </div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(a=>{It(a,parseInt(a.dataset.target)||0)})}function Lt(t){let e=document.getElementById("mini-stats-row");if(!e)return;t=t||{};let i=[{icon:"\u{1F4C5}",label:"Jadwal",val:t.schedule?.current,href:"#/timeline",color:"mini-blue"},{icon:"\u{1F393}",label:"Training",val:t.training_month?.current,href:"#/training",color:"mini-gray"},{icon:"\u{1F504}",label:"Reliefer",val:t.reliever_total?.current,href:"#/relievers?dash_filter=reliever",color:"mini-teal"},{icon:"\u{1F50D}",label:"Inspeksi",val:t.inspection_month?.current,href:"#/timeline?dash_filter=inspeksi",color:"mini-blue"},{icon:"\u{1F9F9}",label:"GCDC",val:t.cleaning_month?.current,href:"#/timeline?dash_filter=gcdc",color:"mini-green"},{icon:"\u{1F4A8}",label:"Fogging",val:t.fogging_month?.current,href:"#/reports/fogging",color:"mini-purple"},{icon:"\u{1F3E2}",label:"Cabang",val:t.branches?.current,href:"#/branches",color:"mini-teal"}];e.innerHTML=i.map(a=>`
    <a href="${a.href}" class="mini-stat ${a.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${a.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${ie(a.val)}">0</div>
        <div class="mini-stat-text">${a.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(a=>It(a,parseInt(a.dataset.target)||0,700))}function ka(t){de("skel-donut","chart-donut");let e=document.getElementById("chart-donut"),i=document.getElementById("donut-legend");if(!e||!i)return;Ae("donut");let a=(t||[]).filter(n=>ie(n.count)>0);if(!a.length){je(e,"Belum ada data permasalahan");return}let o=a.map(n=>`${Se(n.category,"Lainnya")}`),c=a.map(n=>ie(n.count)),s=c.reduce((n,l)=>n+l,0);i.innerHTML=a.map((n,l)=>{let m=lt[l%lt.length],p=s>0?Math.round(n.count/s*100):0;return`
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${m}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${n.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${p}%)</span></div>
          <div class="donut-legend-label">${o[l]}</div>
        </div>
      </div>
    `}).join("");let r={id:"centerText",beforeDraw:function(n){let l=n.width,m=n.height,p=n.ctx;p.restore();let d=(m/80).toFixed(2);p.font="bold "+d+"em Inter",p.textBaseline="middle",p.fillStyle="#1E293B";let g=s.toString(),h=Math.round((l-p.measureText(g).width)/2),b=m/2;p.fillText(g,h,b-10),p.font="600 "+(d*.35).toFixed(2)+"em Inter",p.fillStyle="#64748B";let u="Total",y=Math.round((l-p.measureText(u).width)/2);p.fillText(u,y,b+15),p.save()}};pe.donut=new Chart(e,{type:"doughnut",data:{labels:o,datasets:[{data:c,backgroundColor:lt,borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{display:!1},tooltip:{bodyFont:ee,titleFont:{...ee,weight:"700"},callbacks:{label:n=>` ${n.label}: ${n.parsed} kasus`}}},cutout:"75%"},plugins:[r]})}function Sa(t){de("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;Ae("trend"),t=t||{};let i=(t.labels||[]).map(Dt),a=(t.open||[]).map(c=>ie(c)),o=(t.closed||[]).map(c=>ie(c));if(!i.length){je(e,"Belum ada data trend");return}pe.trend=new Chart(e,{type:"line",data:{labels:i,datasets:[{label:"Open",data:a,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:o,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.08)",fill:!0,tension:.4,pointRadius:2,pointHoverRadius:4,pointBackgroundColor:"#10B981",borderWidth:2}]},options:pt({plugins:{legend:{display:!1}}})})}function Et(t){de("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;let i=e.parentElement;if(i){let s=i.querySelector(".chart-empty");s&&s.remove()}Ae("inspBar"),e.style.display="block",t=t||{};let a=t.labels||[],o=(t.fc||[]).map(s=>ie(s)),c=(t.spv||[]).map(s=>ie(s));if(!a.length){je(e,"Belum ada data inspeksi");return}pe.inspBar=new Chart(e,{type:"bar",data:{labels:a,datasets:[{label:"Skor FC",data:o,backgroundColor:"rgba(37,99,235,0.85)",borderRadius:6,borderSkipped:!1,barPercentage:.5,categoryPercentage:.7},{label:"Skor SPV",data:c,backgroundColor:"rgba(16,185,129,0.85)",borderRadius:6,borderSkipped:!1,barPercentage:.5,categoryPercentage:.7}]},options:pt({maintainAspectRatio:!1,plugins:{legend:{position:"top",labels:{font:{...ee,size:13,weight:"bold"},padding:20,usePointStyle:!0,pointStyle:"circle"}},tooltip:{padding:12,cornerRadius:8,backgroundColor:"rgba(15,23,42,0.9)",titleFont:{size:14,family:"'Inter', sans-serif"},bodyFont:{size:13,family:"'Inter', sans-serif"}}},scales:{x:{grid:{display:!1},ticks:{font:ee,color:we,maxRotation:45,minRotation:30}},y:{grid:{color:He,borderDash:[4,4]},ticks:{font:ee,color:we,stepSize:20},min:0,max:100}}})})}function wa(t){de("skel-contract-mini","chart-contract-mini");let e=document.getElementById("chart-contract-mini");if(!e)return;Ae("contractMiniBar"),t=t||{};let i=t.labels?t.labels.map(Dt):[],a=t.counts||[];if(!i.length){je(e,"Belum ada data kontrak");return}let c=e.getContext("2d").createLinearGradient(0,0,0,200);c.addColorStop(0,"#60A5FA"),c.addColorStop(1,"#2563EB"),pe.contractMiniBar=new Chart(e,{type:"bar",data:{labels:i,datasets:[{label:"Kontrak Habis",data:a,backgroundColor:c,borderRadius:4,borderSkipped:!1,barPercentage:.6,categoryPercentage:.7}]},options:pt({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:ee,color:we,maxRotation:0}},y:{grid:{color:He,borderDash:[4,4],drawBorder:!1},ticks:{font:ee,color:we,precision:0,maxTicksLimit:5},min:0}},animation:{y:{duration:1e3,easing:"easeOutQuart"}}})})}function _a(t){let e=document.getElementById("table-issues");if(!e)return;let i=(t||[]).slice(0,8);if(!i.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <div class="dash-list">
      ${i.map(a=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${ba(a.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;line-height:1.3">${Se(a.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${Se(a.branch_name)}</div>
          </div>
        </div>
      `).join("")}
    </div>`}function xa(t){let e=document.getElementById("widget-agenda");if(!e)return;let i=new Date,a=`${i.getFullYear()}-${String(i.getMonth()+1).padStart(2,"0")}-${String(i.getDate()).padStart(2,"0")}`,c=(t||[]).filter(s=>(s.event_date||"").startsWith(a)).slice(0,10);if(!c.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada agenda hari ini</div>';return}e.innerHTML=`
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${c.map(s=>{let r="#3B82F6",n="#EFF6FF",l="Agenda",m=(s.title||"").toLowerCase();return m.includes("inspeksi")?(r="#10B981",n="#ECFDF5",l="Inspeksi"):m.includes("cleaning")||m.includes("gcdc")?(r="#3B82F6",n="#EFF6FF",l="Cleaning"):m.includes("reliefer")?(r="#F59E0B",n="#FFFBEB",l="Reliefer"):m.includes("fogging")&&(r="#8B5CF6",n="#F5F3FF",l="Fogging"),`
        <div style="display:flex;gap:12px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px;white-space:nowrap">${new Date(s.event_date).toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit"})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${r};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1;min-width:0">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${Se(s.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${Se(s.branch_name)}</div>
          </div>
          <div style="flex-shrink:0;font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${n};color:${r}">${l}</div>
        </div>
      `}).join("")}
    </div>
  `}function Ca(t,e,i){let a=document.getElementById("widget-kpi-kebersihan");if(!a)return;t=t||[],e=e||[],i=i||{};let o=0;if(i.fc&&i.spv&&i.fc.length>0){let f=0,C=0;i.fc.forEach(P=>{f+=ie(P),C++}),i.spv.forEach(P=>{f+=ie(P),C++}),o=C>0?Math.round(f/C):0}let c=o>0?`${o}%`:"N/A",s=e.filter(f=>f.status==="Selesai"||f.status==="Closed"||f.status==="Done").length,r=e.length,n=r>0?Math.round(s/r*100):100,l=t.filter(f=>(f.title||"").toLowerCase().includes("cleaning")),m=l.filter(f=>f.status==="Selesai"||f.status==="Done").length,p=l.length>0?Math.round(m/l.length*100):100,d=t.filter(f=>(f.title||"").toLowerCase().includes("gcdc")),g=d.filter(f=>f.status==="Selesai"||f.status==="Done").length,h=d.length>0?Math.round(g/d.length*100):100,u=e.filter(f=>(f.category||"").toLowerCase().includes("kebersihan")).length,y=t.filter(f=>(f.title||"").toLowerCase().includes("fogging")),S=y.filter(f=>f.status==="Selesai"||f.status==="Done").length,k=y.length>0?Math.round(S/y.length*100):100,w=[{label:"Kebersihan Area",val:c,target:"Target 95%",icon:"\u{1F9F9}",bg:"#ECFDF5",color:"#10B981"},{label:"Penyelesaian Complaint",val:`${n}%`,target:"Target 100%",icon:"\u23F1\uFE0F",bg:"#ECFDF5",color:"#10B981"},{label:"Kepatuhan Jadwal Cleaning",val:`${p}%`,target:"Target 100%",icon:"\u23F1\uFE0F",bg:"#EFF6FF",color:"#3B82F6"},{label:"Kepatuhan GCDC",val:`${h}%`,target:"Target 100%",icon:"\u{1F9F9}",bg:"#EFF6FF",color:"#3B82F6"},{label:"Complaint Cleaning (\u226410)",val:`${u}`,target:"Target \u226410",icon:"\u{1F4DD}",bg:"#F5F3FF",color:"#8B5CF6"},{label:"Pelaksanaan Fogging",val:`${k}%`,target:"Target 100%",icon:"\u{1F4A8}",bg:"#F5F3FF",color:"#8B5CF6"}];a.innerHTML=`
    <div style="display:grid;grid-template-columns:1fr;gap:12px;padding-bottom:12px;height:160px;overflow-y:auto;padding-right:8px;">
      ${w.map(f=>{let C=f.val.includes("%")?parseInt(f.val):Math.min(100,parseInt(f.val)*10);return`
        <div class="prog-item">
          <div class="prog-header">
            <div class="prog-title">
              <div class="prog-title-icon" style="background:${f.bg};color:${f.color}">${f.icon}</div>
              ${f.label}
            </div>
            <div class="prog-val">${f.val}</div>
          </div>
          <span class="prog-target">${f.target}</span>
          <div class="prog-bar-bg">
            <div class="prog-bar-fill" style="width:${C}%;background:${f.color}"></div>
          </div>
        </div>
      `}).join("")}
    </div>
  `}function Ta(){let t=document.getElementById("quick-actions");if(!t)return;let e=[{label:"Buat Permasalahan",icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>',bg:"#3B82F6",href:"#/issues"},{label:"Permintaan Barang",icon:"\u{1F4E6}",bg:"#10B981",href:"#/reports/supply"},{label:"One on One Baru",icon:"\u{1F465}",bg:"#6366F1",href:"#/one-on-one"},{label:"Input Kegiatan",icon:"\u{1F4CB}",bg:"#8B5CF6",href:"#/timeline"},{label:"Buat Checklist",icon:"\u{1F4DD}",bg:"#0EA5E9",href:"#/checklist"},{label:"Laporan Basecamp",icon:"\u{1F4CA}",bg:"#14B8A6",href:"#/reports/basecamp"},{label:"Kalender",icon:"\u{1F4C5}",bg:"#8B5CF6",href:"#/calendar"}];t.innerHTML=e.map(i=>`
    <a href="${i.href}" class="action-btn">
      <div class="action-icon" style="background:${i.bg}">${i.icon}</div>
      ${i.label}
    </a>
  `).join("")}function de(t,e){let i=document.getElementById(t),a=document.getElementById(e);i&&(i.style.display="none",i.style.position=""),a&&(a.style.display="block")}function je(t,e="Belum ada data"){if(!t)return;t.style.display="none";let i=t.parentElement;if(!i)return;if(!i.querySelector(".chart-empty")){let o=document.createElement("div");o.className="chart-empty",o.textContent=e,i.appendChild(o)}}I();async function At(t){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),i=document.getElementById("login-error"),a=document.getElementById("login-btn"),o=document.getElementById("toggle-password"),c=document.getElementById("login-password");o?.addEventListener("click",()=>{let s=c.type==="text";c.type=s?"password":"text",o.style.color=s?"":"var(--primary)"}),e?.addEventListener("submit",async s=>{s.preventDefault(),i.style.display="none";let r=e.username.value.trim(),n=e.password.value;if(!r||!n){i.textContent="Username dan password wajib diisi.",i.style.display="block";return}a.querySelector(".btn-text").style.display="none",a.querySelector(".btn-spinner").style.display="",a.disabled=!0;try{let l=await v("/api/auth/login",{method:"POST",body:JSON.stringify({username:r,password:n})});l.ok&&l.data.success?(Xe(l.data.data.token),Ie(l.data.data.user),V("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(i.textContent=l.data.error||"Username atau password salah.",i.style.display="block",a.classList.add("shake"),setTimeout(()=>a.classList.remove("shake"),600))}catch{i.textContent="Gagal terhubung ke server. Periksa koneksi internet.",i.style.display="block"}finally{a.querySelector(".btn-text").style.display="",a.querySelector(".btn-spinner").style.display="none",a.disabled=!1}})}I();L();async function $a(){return await B()}async function Nt(t,e){let i=e?e.get("dash_filter"):null,a=await $a();$({container:t,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,paginationMode:"client",defaultFilters:{status:i==="active"?"Aktif":""},onDataLoaded:c=>c,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:c=>ke(c)},{key:"phone",label:"No. HP",render:c=>c?`<a href="tel:${c}">${c}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:c=>window.formatDate(c)},{key:"status",label:"Status",render:c=>N(c)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:a},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:c=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:c?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:c?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:a,value:c?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:c?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:c?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:c?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:c?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let c=await v(`/api/employees${window.location.search?window.location.search+"&":"?"}limit=10000`);if(c.ok){let s=c.data.data.map(r=>({"Nama Lengkap":r.full_name,Cabang:r.branch_name||"",Divisi:r.division||"","No. HP":r.phone||"","Tgl Masuk":r.join_date||"",Status:r.status||""}));E(s,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif"},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif"}],"Template_Import_Karyawan")},onImport:async c=>{let s=l=>{if(!l)return null;let m=String(l||"").toLowerCase(),p=a.find(d=>String(d.label||"").toLowerCase()===m);return p?p.value:null},r=c.map(l=>({full_name:String(l["Nama Lengkap"]||"").trim(),branch_id:s(String(l.Cabang||"").trim()),division:String(l.Divisi||"").trim()||"FACILITY CARE",phone:String(l["No. HP"]||"").trim(),join_date:String(l["Tgl Masuk"]||"").trim(),status:String(l.Status||"").trim(),notes:String(l.Catatan||"").trim()})).filter(l=>l.full_name),n=await v("/api/employees/import",{method:"POST",body:JSON.stringify(r)});if(!n.ok)throw new Error(n.data?.error||"Import gagal")}}})}I();L();var ut=[],Ot=[];async function Ea(){ut=await B(),Ot=await ve()}var mt=async t=>{let e=[],i=1;for(;;){let o=await(await Promise.resolve().then(()=>(I(),he))).apiFetch(`${t}${t.includes("?")?"&":"?"}limit=100&page=${i}`);if(!o.ok)break;let c=o.data?.data||o.data||[],s=Array.isArray(c)?c:[];if(e=e.concat(s),s.length<100||o.data?.pagination&&i>=o.data.pagination.pages)break;i++}return e};async function Ue(t,e){await Ea(),$({container:t,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",paginationMode:"client",defaultFilters:{},onDataLoaded:a=>a,columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:a=>ke(a)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:a=>window.formatDate(a)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:a=>!a||String(a).startsWith("2099")?"Tetap / PKWTT":window.formatDate(a)},{key:"days_remaining",label:"Sisa Kontrak",render:(a,o)=>o.end_date&&String(o.end_date).startsWith("2099")?'<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>':it(a)},{key:"status",label:"Status",render:a=>N(a)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:ut},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],onBeforeSubmit:a=>(a.end_date||(a.end_date="2099-12-31"),a),onAfterLoad:()=>{if(!document.getElementById("btn-find-missing")){let a=document.createElement("button");a.id="btn-find-missing",a.className="btn btn-ghost",a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.style.marginLeft="8px",a.style.color="#EF4444",a.style.border="1px solid currentColor",a.onclick=async()=>{a.innerHTML="\u231B Mencari...",a.disabled=!0;try{let[c,s]=await Promise.all([mt("/api/employees?status=Aktif"),mt("/api/contracts")]);if(c.length>0){let r=s.filter(p=>p.status==="Aktif"),n=new Set(r.map(p=>p.employee_id)),l=c.filter(p=>!n.has(p.id)),m=`<p style="margin-bottom:12px">Data yang terbaca: <b>${c.length}</b> Karyawan Aktif, dan <b>${r.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${l.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;l.forEach(p=>{let d=s.filter(h=>h.employee_id===p.id),g='<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';if(d.length>0){let h=d[0];g=`Pernah ada kontrak (Status: <b style="color:#EF4444">${h.status}</b>, Selesai: ${window.formatDate(h.end_date)})`}m+=`<li style="margin-bottom:8px"><b>${p.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${p.branch_name||"-"} | ${g}</span></li>`}),m+="</ul>",Promise.resolve().then(()=>(Be(),kt)).then(p=>p.createModal({title:"Karyawan Tanpa Kontrak Aktif",content:m,cancelText:"Tutup"}))}}catch(c){console.error(c)}a.innerHTML="\u{1F50D} Cek Selisih Karyawan",a.disabled=!1};let o=document.querySelector(".page-actions");o&&o.appendChild(a)}},formFields:a=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"combobox",required:!0,options:Ot,value:a?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:ut,value:a?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:a?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:a?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",value:a?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",value:a?.end_date&&!String(a.end_date).startsWith("2099")?a.end_date:""}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:a?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:a?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let a=await v(`/api/contracts${window.location.search?window.location.search+"&":"?"}limit=10000`);if(a.ok){let o=a.data.data.map(c=>({"Nama Lengkap":c.employee_name,Cabang:c.branch_name||"","Div / Bagian":c.division||"","Tanggal Mulai":c.start_date||"","Tanggal Selesai":c.end_date&&String(c.end_date).startsWith("2099")?"":c.end_date||"","Sisa Kontrak":c.end_date&&String(c.end_date).startsWith("2099")?"Tetap":c.days_remaining!==null&&c.days_remaining!==void 0?`${c.days_remaining} Hari`:"",Status:c.status||""}));E(o,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async a=>{let[o,c]=await Promise.all([v("/api/branches?limit=10000"),mt("/api/employees")]),s=o.data?.data||[],r=c||[];console.log(`Total employee yang berhasil dimuat dari database : ${r.length}`),r.length>0&&(console.log("Contoh 5 employee pertama:"),r.slice(0,5).forEach((u,y)=>{console.log(`${y+1}. ID: ${u.id}, Name: ${u.full_name}, Status: ${u.status}`)}));let n=u=>{if(!u)return null;let y=String(u||"").replace(/\s+/g," ").toLowerCase().trim(),S=s.find(k=>String(k.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===y||String(k.code||"").replace(/\s+/g," ").toLowerCase().trim()===y||String(k.name||"").replace(/\s+/g," ").toLowerCase().trim()===y);return S?S.id:null},l=(u,y)=>{if(console.log("------------------------------------------------"),console.log(`Row Excel : ${y}`),console.log(`Nama dari Excel : "${u}"`),!u)return console.log("Alasan gagal mapping : Nama kosong"),null;let S=String(u||"").replace(/\s+/g," ").toLowerCase().trim();console.log(`Nama setelah normalisasi : "${S}"`),console.log(`Jumlah employee di database : ${r.length}`);let k=r.find(w=>String(w.full_name||"").replace(/\s+/g," ").toLowerCase().trim()===S);return k?(console.log("Employee ditemukan atau tidak : Ditemukan"),console.log(`Employee ID jika ditemukan : ${k.id}`),k.id):(console.log("Employee ditemukan atau tidak : TIDAK Ditemukan"),console.log("Alasan gagal mapping : Tidak ada kecocokan full_name setelah normalisasi"),null)},m=u=>{if(!u)return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let y=String(u).trim();if(/^\d{4,5}(\.\d+)?$/.test(y)){let k=Math.floor(Number(y));if(k>2e4&&k<99999){let w=new Date(Date.UTC(1899,11,30)+k*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);let S=y.split(/[\/\-\.]/);if(S.length===3){let[k,w,f]=S.map(C=>C.trim());if(k.length===4&&w.length<=2&&f.length<=2)return`${k}-${w.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&w.length<=2&&k.length<=2)return`${f}-${w.padStart(2,"0")}-${k.padStart(2,"0")}`}return y},p=a.map((u,y)=>{let S=y+2,k=String(u["Nama Lengkap"]||"").trim(),w=u["Tanggal Mulai"],f=m(w);if(!f){let T=a.__worksheet,R=a.__headers||[],O=R.indexOf("Tanggal Mulai"),A="N/A",J="N/A",te="N/A";if(O!==-1&&T&&window.XLSX){let _=window.XLSX.utils.encode_cell({c:O,r:S-1});te=_;let D=T[_];D?(A=D.t||"undefined",J=D.w||"undefined"):A="CELL KOSONG/TIDAK ADA DI WORKSHEET"}let x="Unknown";w==null||w===""?x="Kondisi IF: Nilai murni undefined, null, atau string kosong dari parsed JSON.":w instanceof Date&&isNaN(w.getTime())?x="Kondisi IF: Nilai adalah object Date namun invalid (isNaN).":x="Kondisi IF: Tidak lolos Regex YYYY-MM-DD maupun konversi serial number Excel.",console.log("=========================="),console.log("[DEBUG] DATE PARSING FAILED"),console.log("=========================="),console.log(`Excel Row Number : ${S}`),console.log(`Employee Name : ${k}`),console.log(`Column Header Used : "Tanggal Mulai" (Index: ${O})`),console.log(`Raw Cell Value : "${w}"`),console.log(`JavaScript Type : ${typeof w}`),console.log(`SheetJS Cell Type : ${A}`),console.log(`SheetJS Formatted Value : "${J}"`),console.log(`Value After Trim : "${String(w||"").trim()}"`),console.log(`Value After Date Parser : "${f}"`),console.log(`Is Empty : ${!w}`),console.log(`Is Invalid Date : ${w instanceof Date?isNaN(w.getTime()):"Bukan JS Date Object"}`),console.log(`Reason : ${x}`),console.log(`Workbook Sheet : ${T?"Ada":"Tidak Ditemukan"}`),console.log(`Excel Cell Address : ${te}`),console.log(`
--- Seluruh Kolom Pada Baris Ini (Mencegah Column Shift) ---`),console.log(JSON.stringify(u,null,2)),console.log(`
--- Daftar Seluruh Header Yang Terbaca ---`),console.log(JSON.stringify(R)),console.log(`==========================
`)}let C=l(k,S),P=null;return C?f||(P="Tanggal Mulai kosong atau tidak berformat tanggal"):P="Karyawan tidak ditemukan di Database",{isValid:!!(C&&f),invalidReason:P,rowNum:S,data:{employee_id:C,branch_id:n(String(u.Cabang||"").trim()),division:String(u["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:f,end_date:m(u["Tanggal Selesai"])||"2099-12-31",status:String(u.Status||"").trim(),_rawName:k}}}),d=[],g=[];if(p.forEach(u=>{u.isValid?d.push(u.data):g.push({rowNum:u.rowNum,name:u.data._rawName,reason:u.invalidReason})}),console.log(`Split Validation - Valid: ${d.length}, Invalid: ${g.length}`),d.length===0){let u=`SEMUA BARIS GAGAL IMPORT!

Total Excel: ${a.length}
Valid: 0
Invalid: ${g.length}

Daftar Kegagalan (Contoh):
`;g.slice(0,10).forEach(y=>{u+=`- Row ${y.rowNum} | Nama: ${y.name} | Alasan: ${y.reason}
`}),g.length>10&&(u+=`- ... dan ${g.length-10} lainnya.
`),alert(u);return}let h=await v("/api/contracts/import",{method:"POST",body:JSON.stringify(d)}),b=`IMPORT SUMMARY
======================
`;b+=`Total Baris Excel : ${a.length}
`,b+=`Baris Valid       : ${d.length}
`,b+=`Baris Invalid     : ${g.length}

`,h&&h.data&&h.data.metrics?(b+=`Berhasil INSERT   : ${h.data.metrics.inserted}
`,b+=`Berhasil UPDATE   : ${h.data.metrics.updated}
`):b+=`Berhasil diproses : ${d.length}
`,g.length>0&&(b+=`
DAFTAR DATA DILEWATI:
`,g.forEach(u=>{b+=`- Row ${u.rowNum} | ${u.name} | ${u.reason}
`})),alert(b),typeof Ue=="function"&&Ue()}}})}I();L();var gt=[],Ne=[];function Da(t){if(!Array.isArray(t))return"Q3";let e=["Q4","Q3","Q2","Q1"];for(let i of e)if(t.some(a=>a.period===i))return i;return"Q3"}function Ia(t,e){if(t.period!=="Q3")return!1;let i=String(t.status||"").toLowerCase();if(i!=="selesai"&&i!=="completed"&&i!=="done")return!1;let a=String(t.activity_type||"").toLowerCase();return e==="inspeksi"?a.includes("inspeksi"):e==="gcdc"?a.includes("general cleaning")||a.includes("deep cleaning"):!1}async function Ft(t,e){gt=await B();let i=await H();Ne=["Ade Surahman","Berlin Ariansyah","Mizwar","Fajar","Ade","Berlin"];let a=l=>l&&!Ne.find(m=>(typeof m=="object"?m.value:m)===l)?[...Ne,l]:Ne,o=await v(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`),c=l=>{if(!l||l==="-"||String(l).trim()==="")return"";let m=String(l).split("-");return m.length===3&&m[0].length===4?`${m[2]}-${m[1]}-${m[0]}`:l},s=o.data?.data||[],r=Da(s),n=e?e.get("dash_filter"):null;$({container:t,title:"Jadwal Kegiatan",icon:"\u{1F4C5}",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",paginationMode:"client",defaultFilters:n?{period:"Q3"}:{},onDataLoaded:l=>(n&&(l=l.filter(m=>Ia(m,n))),l.sort((m,p)=>{let d=m.target_date?new Date(m.target_date).getTime():m.opening_date?new Date(m.opening_date).getTime():0;return(p.target_date?new Date(p.target_date).getTime():p.opening_date?new Date(p.opening_date).getTime():0)-d})),columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:l=>rt(l)},{key:"period",label:"Periode",render:l=>ne(l)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:l=>c(l)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:l=>c(l)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:l=>c(l)},{key:"status",label:"Status",render:l=>N(l)}],filterFields:[{type:"combobox",name:"branch_id",label:"Cabang",options:gt},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"combobox",name:"pic",label:"PIC",options:Ne}],formFields:l=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:gt,value:l?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene","General Cleaning","Deep Cleaning","Fogging"],value:l?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:l?.period},{name:"pic",label:"PIC",type:"combobox",options:a(l?.pic),value:l?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:l?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:l?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:l?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:l?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let l=await v(`/api/schedule${window.location.search?window.location.search+"&":"?"}limit=10000`);if(l.ok){let m=l.data.data.map(p=>({ID:p.id||"",Cabang:p.branch_name||"",Kegiatan:p.activity_type||"",Periode:p.period||"",PIC:p.pic||"","Tgl Opening":p.opening_date||"","Tgl Target":p.target_date||"","Tgl Selesai":p.completion_date||"",Status:p.status||""}));E(m,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{ID:"",Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done"}],"Template_Import_Jadwal")},onImport:async l=>{let p=(await v("/api/branches?all=1")).data?.data||[],d=u=>{if(!u)return null;let y=String(u||"").toLowerCase(),S=p.find(k=>String(k.full_name||"").toLowerCase()===y||String(k.code||"").toLowerCase()===y||String(k.name||"").toLowerCase()===y);return S?S.id:null},g=u=>{if(u==null||u==="")return"";if(u instanceof Date&&!isNaN(u.getTime()))return u.toISOString().slice(0,10);let y=String(u).trim();if(y===""||y==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);if(/^\d{4,5}$/.test(y)){let k=Number(y);if(k>2e4&&k<99999){let w=new Date(Date.UTC(1899,11,30)+k*864e5);return isNaN(w.getTime())?"":w.toISOString().slice(0,10)}}let S=y.split(/[\/\-\.]/);if(S.length===3){let[k,w,f]=S.map(C=>C.trim());if(k.length===4&&w.length<=2&&f.length<=2)return`${k}-${w.padStart(2,"0")}-${f.padStart(2,"0")}`;if(f.length===4&&w.length<=2&&k.length<=2)return`${f}-${w.padStart(2,"0")}-${k.padStart(2,"0")}`}return y},h=l.map(u=>({id:u.ID||null,branch_id:d(String(u.Cabang||"").trim()),activity_type:String(u.Kegiatan||"").trim(),period:String(u.Periode||"").trim(),pic:String(u.PIC||u.Pic||"").trim(),opening_date:g(u["Tgl Opening"]||u["Tanggal Opening"]||u["Tgl Openir"]),target_date:g(u["Tgl Target"]||u["Tanggal Target"]),completion_date:g(u["Tgl Selesai"]||u["Tanggal Selesai"]),status:String(u.Status||"").trim(),notes:String(u.Catatan||u.Keterangan||"").trim()})).filter(u=>u.activity_type&&u.period),b=await v("/api/schedule/import",{method:"POST",body:JSON.stringify(h)});if(!b.ok)throw new Error(b.data?.error||"Import gagal")}}})}I();L();var bt=[],Je=[];async function Mt(t,e){let i=e?e.get("dash_filter"):null;bt=await B(),Je=await H();let a=s=>s&&!Je.find(r=>r.value===s)?[...Je,{value:s,label:s}]:Je,o=new Date().getFullYear(),c=Array.from({length:5},(s,r)=>String(o-r));$({container:t,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",paginationMode:"client",defaultFilters:{status:i==="open"?"Open":""},onDataLoaded:s=>s,columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:s=>`<span class="badge badge-secondary">${s}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:s=>`<span title="${s}">${s?.length>50?s.slice(0,50)+"\u2026":s}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:s=>`<span title="${s||""}">${s?.length>40?s.slice(0,40)+"\u2026":s||"-"}</span>`},{key:"status",label:"Status",render:s=>N(s)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:s=>window.formatDate(s)},{key:"day_count",label:"Hari",render:s=>s??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:bt},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:c}],formFields:s=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:s?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:bt,value:s?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:s?.category},{name:"source",label:"Sumber Laporan",type:"combobox",options:["SPV","AM","RCP","Perawat","FC","Berlin","Ade","Pattrel","Dentrel"],value:s?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:s?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:a(s?.employee_name),value:s?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"combobox",options:a(s?.fc_specialist),value:s?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:s?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:s?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let s=await v(`/api/issues${window.location.search?window.location.search+"&":"?"}limit=10000`);if(s.ok){let r=s.data.data.map(n=>({Tanggal:n.report_date||"",Cabang:n.branch_name||"",Kategori:n.category||"",Sumber:n.source||"",Keluhan:n.complaint||"","Nama FC":n.employee_name||"","FC Spesialis":n.fc_specialist||"",Solusi:n.solution||"","Tgl Selesai":n.completion_date||"",Status:n.status||""}));E(r,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async s=>{let n=(await v("/api/branches?all=1")).data?.data||[],l=d=>{if(!d)return null;let g=String(d||"").toLowerCase(),h=n.find(b=>String(b.full_name||"").toLowerCase()===g||String(b.code||"").toLowerCase()===g||String(b.name||"").toLowerCase()===g);return h?h.id:null},m=s.map(d=>({branch_id:l(String(d.Cabang||"").trim()),report_date:String(d.Tanggal||"").trim(),category:String(d.Kategori||"").trim(),source:String(d.Sumber||"").trim(),complaint:String(d.Keluhan||"").trim(),employee_name:String(d["Nama FC"]||"").trim(),fc_specialist:String(d["FC Spesialis"]||"").trim(),solution:String(d.Solusi||"").trim(),completion_date:String(d["Tgl Selesai"]||"").trim(),status:String(d.Status||"").trim()})).filter(d=>d.report_date&&d.complaint&&d.category),p=await v("/api/issues/import",{method:"POST",body:JSON.stringify(m)});if(!p.ok)throw new Error(p.data?.error||"Import gagal")}}})}I();var _e=[];async function Kt(t,e){let i=e?e.get("dash_filter"):null;_e=await B();let a=await H(),o=["Ade Surahman","Berlin Ariansyah","Mizwar","Fajar","Ade","Berlin"],c=r=>r&&!a.find(n=>n.value===r)?[...a,{value:r,label:r}]:a,s=r=>r&&!o.find(n=>(typeof n=="object"?n.value:n)===r)?[...o,r]:o;$({container:t,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",paginationMode:"client",defaultFilters:{status:i==="pending"?"Open":""},onDataLoaded:r=>r,columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:r=>`<span title="${r||""}">${r?.length>50?r.slice(0,50)+"\u2026":r||"-"}</span>`},{key:"solution",label:"Solusi",render:r=>`<span title="${r||""}">${r?.length>40?r.slice(0,40)+"\u2026":r||"-"}</span>`},{key:"status",label:"Status",render:r=>N(r)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:r=>window.formatDate(r)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:r=>r?`<a href="${r}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"combobox",name:"branch_id",label:"Cabang",options:_e},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async r=>{let n=new URLSearchParams(r||{}).toString(),l=await v(`/api/one-on-one?limit=10000&${n}`);if(l.ok){let m=l.data.data.map(d=>({Tanggal:d.meeting_date||"",Cabang:d.branch_name||"","Nama Karyawan":d.employee_name||"",PIC:d.pic||"",Masalah:d.problem||"",Solusi:d.solution||"",Status:d.status||"","Tgl Selesai":d.completion_date||"",Dokumen:d.document_link||""})),{downloadExcel:p}=await Promise.resolve().then(()=>(L(),X));p(m,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let r=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(L(),X));n(r,"Template_Import_OneOnOne")},onImport:async r=>{let n=d=>{if(!d)return null;let g=String(d||"").toLowerCase(),h=_e.find(b=>String(b.label||"").toLowerCase()===g);return h?h.value:null},l=d=>{if(!d)return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let g=String(d).trim();if(/^\d{4,5}$/.test(g)){let b=Number(g);if(b>2e4&&b<99999){let u=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);let h=g.split(/[\/\-\.]/);if(h.length===3){let[b,u,y]=h.map(S=>S.trim());if(b.length===4&&u.length<=2&&y.length<=2)return`${b}-${u.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&u.length<=2&&b.length<=2)return`${y}-${u.padStart(2,"0")}-${b.padStart(2,"0")}`}return g},m=r.map(d=>({meeting_date:l(d.Tanggal),employee_name:String(d["Nama Karyawan"]||"").trim(),branch_id:n(String(d.Cabang||"").trim()),pic:String(d.PIC||"").trim(),problem:String(d.Masalah||"").trim(),solution:String(d.Solusi||"").trim(),status:String(d.Status||"").trim(),completion_date:l(d["Tgl Selesai"]),document_link:String(d.Dokumen||"").trim()})).filter(d=>d.meeting_date&&d.employee_name&&d.branch_id),p=await v("/api/one-on-one/import",{method:"POST",body:JSON.stringify(m)});if(!p.ok)throw new Error(p.data?.error||"Import gagal")}},formFields:r=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:r?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:r?.branch_id&&!_e.find(n=>n.value==r.branch_id)?[..._e,{value:r.branch_id,label:r.branch_name||r.branch_id}]:_e,createApi:{path:"/api/branches",field:"full_name"},value:r?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"combobox",required:!0,options:c(r?.employee_name),value:r?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:s(r?.pic),createApi:{path:"/api/pic",field:"name"},value:r?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:r?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:r?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:r?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:r?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:r?.document_link}]})}I();async function Rt(t){let e=await B(),i=await H(),a=["Ade Surahman","Berlin Ariansyah","Mizwar","Fajar","Ade","Berlin"],o=r=>r&&!i.find(n=>n.value===r)?[...i,{value:r,label:r}]:i,c=r=>r&&!a.find(n=>(typeof n=="object"?n.value:n)===r)?[...a,r]:a,s=Array.from({length:5},(r,n)=>String(new Date().getFullYear()-n));$({container:t,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:r=>{try{let n=JSON.parse(r);return Array.isArray(n)?n.join(", "):r||"-"}catch{return r||"-"}}},{key:"score",label:"Nilai",render:r=>r!=null?`<strong>${r}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:r=>r?`<a href="${r}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:s}],exportOptions:{moduleName:"training",onExport:async r=>{let n=new URLSearchParams(r||{}).toString(),l=await v(`/api/training?limit=10000&${n}`);if(l.ok){let m=l.data.data.map(d=>{let g=d.participants||"";try{let h=JSON.parse(g);g=Array.isArray(h)?h.join(", "):g}catch{}return{Tanggal:d.training_date||"",Batch:d.batch||"",Materi:d.subject||"",Cabang:d.branch_name||"",Trainer:d.trainer||"",Peserta:g,Nilai:d.score!==null&&d.score!==void 0?d.score:"",Dokumen:d.document_link||""}}),{downloadExcel:p}=await Promise.resolve().then(()=>(L(),X));p(m,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let r=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:n}=await Promise.resolve().then(()=>(L(),X));n(r,"Template_Import_Training")},onImport:async r=>{let n=d=>{if(!d)return null;let g=String(d||"").toLowerCase(),h=e.find(b=>String(b.label||"").toLowerCase()===g);return h?h.value:null},l=d=>{if(!d)return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let g=String(d).trim();if(/^\d{4,5}$/.test(g)){let b=Number(g);if(b>2e4&&b<99999){let u=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(u.getTime())?"":u.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(g))return g.slice(0,10);let h=g.split(/[\/\-\.]/);if(h.length===3){let[b,u,y]=h.map(S=>S.trim());if(b.length===4&&u.length<=2&&y.length<=2)return`${b}-${u.padStart(2,"0")}-${y.padStart(2,"0")}`;if(y.length===4&&u.length<=2&&b.length<=2)return`${y}-${u.padStart(2,"0")}-${b.padStart(2,"0")}`}return g},m=r.map(d=>({training_date:l(d.Tanggal),batch:String(d.Batch||"").trim(),subject:String(d.Materi||"").trim(),branch_id:n(String(d.Cabang||"").trim()),trainer:String(d.Trainer||"").trim(),participants:String(d.Peserta||"").trim(),score:d.Nilai?Number(d.Nilai):null,document_link:String(d.Dokumen||"").trim()})).filter(d=>d.training_date&&d.subject&&d.branch_id),p=await v("/api/training/import",{method:"POST",body:JSON.stringify(m)});if(!p.ok)throw new Error(p.data?.error||"Import gagal")}},formFields:r=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:r?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:r?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:r?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:e,value:r?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:c(r?.trainer),value:r?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let n=JSON.parse(r?.participants);return Array.isArray(n)?n.join(", "):r?.participants||""}catch{return r?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:r?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:r?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],onBeforeSubmit:async r=>(r.participants&&(r.participants=JSON.stringify(r.participants.split(",").map(n=>n.trim()).filter(Boolean))),r)})}I();L();async function qt(t,e){window.__RELIEVER_BUILD__="V3",console.log("RELIEVER PAGE LOADED");let i=await B(),a=await H(),o=e?e.get("dash_filter"):null;console.log("RAW",await ve()),console.log("OPTIONS",a);let c=n=>n&&!a.find(l=>l.value===n)?[...a,{value:n,label:n}]:a,s=["Agung Septiadi","Wasrikin","IQBAL AL BANNA","Muhammad Tri Ismandanu"],r=n=>n&&!s.includes(n)?[...s,n]:s;$({container:t,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",paginationMode:"client",onDataLoaded:n=>{if(o==="reliever"){let l=new Date,m=l.getFullYear(),p=String(l.getMonth()+1).padStart(2,"0");return n.filter(d=>{if(String(d.status||"").toLowerCase()!=="done")return!1;let g=d.backup_date||"";if(g.includes("/")){let h=g.split("/");if(h.length===3&&(h[2].length===4?h[2]:`20${h[2]}`)==m&&h[1].padStart(2,"0")==p)return!0}else if(g.includes("-")&&g.startsWith(`${m}-${p}`))return!0;return!1})}return n},columns:[{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"Nama Facility care"},{key:"period",label:"Periode",render:n=>ne(n)},{key:"reliever_name",label:"Relifer"},{key:"backup_date",label:"Tanggal Back Up",nowrap:!0,render:n=>window.formatDate(n)},{key:"completion_date",label:"Tanggal Selesai",nowrap:!0,render:n=>window.formatDate(n)},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:n=>n?`<span class="badge badge-info">${n}</span>`:"-"},{key:"status",label:"Status",render:n=>N(n)}],filterFields:[{type:"search",placeholder:"Cari reliefer / FC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:n=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,value:n?.branch_id},{name:"period",label:"Periode",type:"combobox",options:["Q1","Q2","Q3","Q4"],value:n?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"Nama Facility care",type:"combobox",options:c(n?.original_fc_name),value:n?.original_fc_name},{name:"reliever_name",label:"Relifer",type:"combobox",required:!0,options:r(n?.reliever_name),value:n?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Back Up",type:"date",required:!0,value:n?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:n?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"combobox",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:n?.reason},{name:"shift",label:"Shift",type:"combobox",options:["Pagi","Siang","Full Shift","Middle"],value:n?.shift}]},{name:"status",label:"Status",type:"combobox",required:!0,options:["Pending","Done","Tidak Datang"],value:n?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let n=await v(`/api/relievers${window.location.search?window.location.search+"&":"?"}limit=10000`);if(n.ok){let l=n.data.data.map(m=>({Cabang:m.branch_name||"","Nama Facility care":m.original_fc_name||"",Periode:m.period||"",Relifer:m.reliever_name||"","Tanggal Back Up":m.backup_date||"","Tanggal Selesai":m.completion_date||"",Keterangan:m.reason||"",Shift:m.shift||"",Status:m.status||""}));l.length===0&&l.push({Cabang:"","Nama Facility care":"",Periode:"",Relifer:"","Tanggal Back Up":"","Tanggal Selesai":"",Keterangan:"",Shift:"",Status:""}),E(l,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Q1",Relifer:"Andi","Tanggal Back Up":"2024-03-10","Tanggal Selesai":"2024-03-10",Keterangan:"Sakit",Shift:"Pagi",Status:"Done"}],"Template_Import_Reliefer")},onImport:async n=>{let m=(await v("/api/branches?all=1")).data?.data||[],p=h=>{if(!h)return null;let b=String(h||"").toLowerCase(),u=m.find(y=>String(y.full_name||"").toLowerCase()===b||String(y.code||"").toLowerCase()===b||String(y.name||"").toLowerCase()===b);return u?u.id:null},d=n.map(h=>({branch_name:String(h.Cabang||"").trim(),backup_date:String(h["Tanggal Back Up"]||h["Tanggal Backup"]||"").trim(),original_fc_name:String(h["Nama Facility care"]||h["FC Digantikan"]||"").trim(),reliever_name:String(h.Relifer||h.Reliefer||"").trim(),period:String(h.Periode||"").trim(),reason:String(h.Keterangan||"").trim(),shift:String(h.Shift||"").trim(),completion_date:String(h["Tanggal Selesai"]||"").trim(),status:String(h.Status||"").trim()})).filter(h=>h.reliever_name&&h.backup_date),g=await v("/api/import/relievers",{method:"POST",body:JSON.stringify({rows:d})});if(!g.ok)throw new Error(g.data?.error||"Import gagal")}}})}I();L();async function Ht(t){let e=await B(),i=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));$({container:t,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>ne(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>N(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),c=await v(`/api/reports/inspection?limit=10000&${o}`);if(c.ok){let s=c.data.data.map(r=>({Cabang:r.branch_name||"",Periode:r.period||"",Tanggal:r.inspection_date||"","Point FC":r.fc_score!==null&&r.fc_score!==void 0?r.fc_score:"","Point SPV":r.spv_score!==null&&r.spv_score!==void 0?r.spv_score:"",Status:r.status||"","Link Dokumen":r.document_link||""}));E(s,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Inspeksi")},onImport:async a=>{let o=n=>{if(!n)return null;let l=String(n||"").toLowerCase(),m=e.find(p=>String(p.label||"").toLowerCase()===l);return m?m.value:null},c=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let l=String(n).trim();if(l===""||l==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);if(/^\d{4,5}$/.test(l)){let p=Number(l);if(p>2e4&&p<99999){let d=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let m=l.split(/[\/\-\.]/);if(m.length===3){let[p,d,g]=m.map(h=>h.trim());if(p.length===4&&d.length<=2&&g.length<=2)return`${p}-${d.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&d.length<=2&&p.length<=2)return`${g}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`}return l},s=a.map(n=>({branch_id:o(String(n.Cabang||"").trim()),period:String(n.Periode||"").trim(),inspection_date:c(n.Tanggal),fc_score:n["Point FC"]!==void 0&&n["Point FC"]!==""?Number(n["Point FC"]):null,spv_score:n["Point SPV"]!==void 0&&n["Point SPV"]!==""?Number(n["Point SPV"]):null,status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.inspection_date),r=await v("/api/reports/inspection/import",{method:"POST",body:JSON.stringify(s)});if(!r.ok)throw new Error(r.data?.error||"Import gagal")}}})}I();L();async function jt(t){let e=await B(),i=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));$({container:t,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>ne(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>N(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),c=await v(`/api/reports/cleaning?limit=10000&${o}`);if(c.ok){let s=c.data.data.map(r=>({Cabang:r.branch_name||"",Jenis:r.activity_type||"",Periode:r.period||"",Tanggal:r.activity_date||"",Status:r.status||"","Link Dokumen":r.document_link||""}));E(s,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_GCDC")},onImport:async a=>{let o=n=>{if(!n)return null;let l=String(n||"").toLowerCase(),m=e.find(p=>String(p.label||"").toLowerCase()===l);return m?m.value:null},c=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let l=String(n).trim();if(l===""||l==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);if(/^\d{4,5}$/.test(l)){let p=Number(l);if(p>2e4&&p<99999){let d=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let m=l.split(/[\/\-\.]/);if(m.length===3){let[p,d,g]=m.map(h=>h.trim());if(p.length===4&&d.length<=2&&g.length<=2)return`${p}-${d.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&d.length<=2&&p.length<=2)return`${g}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`}return l},s=a.map(n=>({branch_id:o(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"").trim(),period:String(n.Periode||"").trim(),activity_date:c(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.activity_type&&n.period&&n.activity_date),r=await v("/api/reports/cleaning/import",{method:"POST",body:JSON.stringify(s)});if(!r.ok)throw new Error(r.data?.error||"Import gagal")}}})}I();L();async function Ut(t){let e=await B(),i=Array.from({length:4},(a,o)=>String(new Date().getFullYear()-o));$({container:t,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge badge-warning">${a}</span>`},{key:"period",label:"Periode",render:a=>ne(a)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:a=>window.formatDate(a)},{key:"status",label:"Status",render:a=>N(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:a=>a||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:a?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),c=await v(`/api/reports/fogging?limit=10000&${o}`);if(c.ok){let s=c.data.data.map(r=>({Cabang:r.branch_name||"",Jenis:r.activity_type||"Fogging",Periode:r.period||"",Tanggal:r.activity_date||"",Status:r.status||"","Link Dokumen":r.document_link||""}));E(s,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/..."}],"Template_Import_Fogging")},onImport:async a=>{let o=n=>{if(!n)return null;let l=String(n||"").toLowerCase(),m=e.find(p=>String(p.label||"").toLowerCase()===l);return m?m.value:null},c=n=>{if(n==null||n==="")return"";if(n instanceof Date&&!isNaN(n.getTime()))return n.toISOString().slice(0,10);let l=String(n).trim();if(l===""||l==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);if(/^\d{4,5}$/.test(l)){let p=Number(l);if(p>2e4&&p<99999){let d=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let m=l.split(/[\/\-\.]/);if(m.length===3){let[p,d,g]=m.map(h=>h.trim());if(p.length===4&&d.length<=2&&g.length<=2)return`${p}-${d.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&d.length<=2&&p.length<=2)return`${g}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`}return l},s=a.map(n=>({branch_id:o(String(n.Cabang||"").trim()),activity_type:String(n.Jenis||n.Kegiatan||"Fogging").trim(),period:String(n.Periode||"").trim(),activity_date:c(n.Tanggal),status:String(n.Status||"").trim(),document_link:String(n["Link Dokumen"]||"").trim(),notes:String(n.Catatan||n.Keterangan||"").trim()})).filter(n=>n.branch_id&&n.period&&n.activity_date),r=await v("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(s)});if(!r.ok)throw new Error(r.data?.error||"Import gagal")}}})}I();L();async function Jt(t){let e=await B(),i=await H(),a=i,o=s=>s&&!i.find(r=>r.value===s)?[...i,{value:s,label:s}]:i,c=s=>s&&!a.find(r=>r.value===s)?[...a,{value:s,label:s}]:a;$({container:t,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:s=>`<span title="${s||""}">${s?.length>60?s.slice(0,60)+"\u2026":s||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:s=>window.formatDate(s)},{key:"status",label:"Status",render:s=>N(s)},{key:"notes",label:"Keterangan",render:s=>s?.length>40?s.slice(0,40)+"\u2026":s||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan / PIC..."},{type:"combobox",name:"branch_id",label:"Cabang",options:e},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]}],formFields:s=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:e,value:s?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:c(s?.pic),value:s?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:s?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:s?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:s?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:s?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:s?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async s=>{let r=new URLSearchParams(s||{}).toString(),n=await v(`/api/reports/basecamp?limit=10000&${r}`);if(n.ok){let l=n.data.data.map(m=>({"Tgl Info":m.info_date||"",Cabang:m.branch_name||"",Permasalahan:m.problem||"",PIC:m.pic||"","Tgl Done":m.done_date||"",Status:m.status||"",Keterangan:m.notes||""}));E(l,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async s=>{let r=p=>{if(!p)return null;let d=String(p||"").toLowerCase(),g=e.find(h=>String(h.label||"").toLowerCase()===d);return g?g.value:null},n=p=>{if(p==null||p==="")return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let d=String(p).trim();if(d===""||d==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);if(/^\d{4,5}$/.test(d)){let h=Number(d);if(h>2e4&&h<99999){let b=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}let g=d.split(/[\/\-\.]/);if(g.length===3){let[h,b,u]=g.map(y=>y.trim());if(h.length===4&&b.length<=2&&u.length<=2)return`${h}-${b.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&b.length<=2&&h.length<=2)return`${u}-${b.padStart(2,"0")}-${h.padStart(2,"0")}`}return d},l=s.map(p=>({info_date:n(p["Tgl Info"]||p["Tanggal Info"]),branch_id:r(String(p.Cabang||"").trim()),problem:String(p.Permasalahan||"").trim(),pic:String(p.PIC||"").trim(),done_date:n(p["Tgl Done"]||p["Tanggal Done"]),status:String(p.Status||"").trim(),notes:String(p.Keterangan||p.Catatan||"").trim()})).filter(p=>p.info_date&&p.branch_id&&p.problem),m=await v("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(l)});if(!m.ok)throw new Error(m.data?.error||"Import gagal")}}})}async function Gt(t){$({container:t,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(I(),he)),o=await a(`/api/sop?limit=10000&${i}`);if(o.ok){let c=o.data.data.map(r=>({"Nama SOP":r.name||"",Kategori:r.category||"",Dokumen:r.document_link||"",Catatan:r.notes||r.description||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(L(),X));s(c,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(L(),X));i(e,"Template_Import_SOP")},onImport:async e=>{let i=e.map(c=>({name:String(c["Nama SOP"]||"").trim(),category:String(c.Kategori||"").trim(),document_link:String(c.Dokumen||"").trim(),description:String(c.Catatan||"").trim()})).filter(c=>c.name),{apiFetch:a}=await Promise.resolve().then(()=>(I(),he)),o=await a("/api/sop/import",{method:"POST",body:JSON.stringify(i)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function Qt(t){$({container:t,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),{apiFetch:a}=await Promise.resolve().then(()=>(I(),he)),o=await a(`/api/checklist?limit=10000&${i}`);if(o.ok){let c=o.data.data.map(r=>({"Nama Checklist":r.name||"",Kategori:r.category||"",Dokumen:r.document_link||"",Deskripsi:r.description||""})),{downloadExcel:s}=await Promise.resolve().then(()=>(L(),X));s(c,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:i}=await Promise.resolve().then(()=>(L(),X));i(e,"Template_Import_Checklist")},onImport:async e=>{let i=e.map(c=>({name:String(c["Nama Checklist"]||"").trim(),category:String(c.Kategori||"").trim(),document_link:String(c.Dokumen||"").trim(),description:String(c.Deskripsi||"").trim()})).filter(c=>c.name),{apiFetch:a}=await Promise.resolve().then(()=>(I(),he)),o=await a("/api/checklist/import",{method:"POST",body:JSON.stringify(i)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}I();Be();L();async function ht(t,e="forms"){if(e==="supply")return Ba(t);Pa(t)}function Pa(t){$({container:t,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}]})}async function Ba(t){let i=((await v("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));$({container:t,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,o)=>o.branch_name_ref||o.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let o=JSON.parse(a);return Array.isArray(o)?o.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let o=JSON.parse(a);return Array.isArray(o)?o.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>N(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:a=>{let o=a?.tools_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}let c=a?.chemical_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:a?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:a?.branch_id&&!i.find(s=>s.value==a.branch_id)?[...i,{value:a.branch_id,label:a.branch_name||a.branch_id}]:i,createApi:{path:"/api/branches",field:"full_name"},value:a?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:o},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:a?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:c},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:a?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:a?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:a?.status||""},{name:"processed_by",label:"Diproses Oleh",value:a?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async a=>{let o=new URLSearchParams(a||{}).toString(),c=await v(`/api/reports/supply?limit=10000&${o}`);if(c.ok){let s=c.data.data.map(r=>{let n=r.tools_items;try{n=Array.isArray(JSON.parse(n))?JSON.parse(n).join(", "):n}catch{}let l=r.chemical_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}return{Waktu:r.submitted_at||"",Pengirim:r.submitter_name||"",Cabang:r.branch_name_ref||r.branch_name||"","Alat/Barang":n||"",Chemical:l||"",Catatan:r.additional_notes||"",Status:r.status||"","Diproses Oleh":r.processed_by||""}});E(s,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async a=>{let c=(await v("/api/branches?all=1")).data?.data||[],s=m=>{if(!m)return null;let p=String(m||"").toLowerCase(),d=c.find(g=>String(g.full_name||"").toLowerCase()===p||String(g.code||"").toLowerCase()===p||String(g.name||"").toLowerCase()===p);return d?d.id:null},r=m=>{if(m==null||m==="")return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let p=String(m).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let g=Number(p);if(g>2e4&&g<99999){let h=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(h.getTime())?"":h.toISOString().slice(0,10)}}let d=p.split(/[\/\-\.]/);if(d.length===3){let[g,h,b]=d.map(u=>u.trim());if(g.length===4&&h.length<=2&&b.length<=2)return`${g}-${h.padStart(2,"0")}-${b.padStart(2,"0")}`;if(b.length===4&&h.length<=2&&g.length<=2)return`${b}-${h.padStart(2,"0")}-${g.padStart(2,"0")}`}return p},n=a.map(m=>({submitted_at:r(m.Waktu||m.Tanggal),submitter_name:String(m.Pengirim||"").trim(),branch_id:s(String(m.Cabang||"").trim()),tools_items:String(m["Alat/Barang"]||m.Alat||"").trim(),chemical_items:String(m.Chemical||"").trim(),additional_notes:String(m.Catatan||m.Keterangan||"").trim(),status:String(m.Status||"").trim(),processed_by:String(m["Diproses Oleh"]||m.PIC||"").trim()})).filter(m=>m.submitted_at&&m.submitter_name&&m.branch_id),l=await v("/api/reports/supply/import",{method:"POST",body:JSON.stringify(n)});if(!l.ok)throw new Error(l.data?.error||"Import gagal")}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,o)=>{let c=se({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(s,r)=>{let n=s.querySelector("#supply-status").value,l=s.querySelector("#supply-processed-by").value;(await v(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:n,processed_by:l})})).ok?(V("Status diperbarui."),r(),o()):W("Gagal update status.")}})}}]})}I();L();async function Vt(t){let e=ue();if(!e||!["superadmin","admin"].includes(e.role)){t.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}$({container:t,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:i=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[i]||"badge-neutral"}">${i}</span>`},{key:"is_active",label:"Status",render:i=>i?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:i=>i?new Date(i).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:i=>{let a=!!i;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:i?.full_name},{name:"username",label:"Username",required:!a,placeholder:"username",value:i?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!a,placeholder:"email@contoh.com",value:i?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:i?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:a?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!a,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:a?i?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let i=await v(`/api/users${window.location.search?window.location.search+"&":"?"}limit=10000`);if(i.ok){let a=i.data.data.map(o=>({"Nama Lengkap":o.full_name||"",Username:o.username||"",Email:o.email||"",Role:o.role||"",Status:o.is_active?"Aktif":"Nonaktif"}));E(a,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async i=>{let a=i.map(c=>({full_name:String(c["Nama Lengkap"]||"").trim(),username:String(c.Username||"").trim(),email:String(c.Email||"").trim(),role:String(c.Role||"").trim()||"viewer",password:String(c.Password||"").trim()})).filter(c=>c.username&&c.password&&c.email&&c.full_name),o=await v("/api/users/import",{method:"POST",body:JSON.stringify(a)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}}})}I();L();async function zt(t){$({container:t,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await v(`/api/branches${window.location.search?window.location.search+"&":"?"}limit=10000`);if(e.ok)E(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{E([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let i=e.map(o=>({code:String(o["Kode Cabang"]||"").trim(),name:String(o["Nama Pendek"]||"").trim(),full_name:String(o["Nama Lengkap"]||"").trim(),city:String(o.Kota||"").trim()})).filter(o=>o.code&&o.name),a=await v("/api/branches/import",{method:"POST",body:JSON.stringify(i)});if(!a.ok)throw new Error(a.data?.error||"Import gagal")}}})}I();async function Wt(t){let e=new Date,i=[],a=!1,o=null,c=null;t.innerHTML=`
    <div class="page-header">
      <div class="header-content">
        <h1 class="page-title">\u{1F4C5} Kalender Kegiatan</h1>
        <p class="page-subtitle">Terintegrasi real-time dengan seluruh modul \u2014 update otomatis setiap ada perubahan data.</p>
      </div>
      <div class="page-actions">
        <span id="cal-sync-status" class="cal-sync-badge sync-live">\u{1F7E2} Real-Time Live</span>
        <button class="btn btn-ghost btn-sm" id="cal-refresh-btn" title="Refresh manual">\u27F3 Refresh</button>
      </div>
    </div>
    <div class="card">
      <div class="card-header calendar-nav">
        <button class="btn btn-ghost btn-sm" id="cal-prev">\u2039 Prev</button>
        <span class="calendar-month-label" id="cal-month-label"></span>
        <button class="btn btn-ghost btn-sm" id="cal-next">Next \u203A</button>
        <div class="calendar-filters" style="display: flex; flex-wrap: wrap; gap: 10px;">
          <label class="filter-check"><input type="checkbox" value="schedule"        checked class="cal-filter"> \u{1F5D3} Jadwal</label>
          <label class="filter-check"><input type="checkbox" value="reliever"        checked class="cal-filter"> \u{1F504} Reliefer</label>
          <label class="filter-check"><input type="checkbox" value="cleaning"        checked class="cal-filter"> \u{1F9F9} Cleaning</label>
          <label class="filter-check"><input type="checkbox" value="inspection"      checked class="cal-filter"> \u{1F50E} Inspeksi</label>
          <label class="filter-check"><input type="checkbox" value="fogging"         checked class="cal-filter"> \u{1F4A8} Fogging</label>
          <label class="filter-check"><input type="checkbox" value="contract_expiry" checked class="cal-filter"> \u{1F4CB} Kontrak Habis</label>
        </div>
      </div>
      <div class="card-body p-0">
        <div id="calendar-grid" style="min-height:400px"></div>
      </div>
    </div>
    <!-- Event detail sidebar -->
    <div id="cal-event-list" class="cal-event-sidebar" style="display:none">
      <div class="cal-event-header">
        <span id="cal-event-date"></span>
        <button class="btn btn-ghost btn-sm" id="cal-event-close">&times;</button>
      </div>
      <div id="cal-event-items"></div>
    </div>
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),m()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),m()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.getElementById("cal-refresh-btn").addEventListener("click",()=>{l(),m(!0)}),document.querySelectorAll(".cal-filter").forEach(d=>d.addEventListener("change",()=>m()));let s=d=>{a||(clearTimeout(o),o=setTimeout(()=>{a||(l(),n().then(()=>p()))},300))};Le.on("data:changed",s),c=setInterval(()=>{a||n().then(()=>p())},6e4);let r=()=>{a=!0,clearTimeout(o),clearInterval(c),Le.off("data:changed",s)};async function n(){try{let d=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`,g=await v(`/api/dashboard/calendar?month=${d}`);a||(i=g.data?.data||[])}catch(d){console.warn("[Calendar] Failed to load events:",d)}}function l(){let d=document.getElementById("cal-sync-status");d&&(d.textContent="\u{1F504} Memuat...",d.className="cal-sync-badge sync-loading",setTimeout(()=>{a||(d.textContent="\u{1F7E2} Real-Time Live",d.className="cal-sync-badge sync-live")},1200))}async function m(d=!1){let g=document.getElementById("calendar-grid");g&&((d||i.length===0)&&(g.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
        ${Array(35).fill('<div style="background:var(--bg-2);min-height:70px;border-radius:4px;animation:pulse 1.5s ease-in-out infinite;"></div>').join("")}
      </div>`),await n(),p())}function p(){if(a)return;let d=document.getElementById("calendar-grid");if(d)try{let g=e.getFullYear(),h=e.getMonth(),b=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),u=document.getElementById("cal-month-label");u&&(u.textContent=b);let y=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(A=>A.value)),S=i.filter(A=>y.has(A.type)),k={};S.forEach(A=>{let J=(A.event_date||"").slice(0,10);k[J]||(k[J]=[]),k[J].push(A)});let w=new Date(g,h,1).getDay(),f=new Date(g,h+1,0).getDate(),C=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],P=new Date().toISOString().slice(0,10),T='<div class="calendar-grid">';C.forEach(A=>{T+=`<div class="cal-day-header">${A}</div>`});for(let A=0;A<w;A++)T+='<div class="cal-cell cal-cell-empty"></div>';for(let A=1;A<=f;A++){let J=`${g}-${String(h+1).padStart(2,"0")}-${String(A).padStart(2,"0")}`,te=k[J]||[],x=J===P;T+=`
          <div class="cal-cell ${x?"cal-today":""} ${te.length?"cal-has-events":""}"
               data-date="${J}" tabindex="0" role="button" aria-label="${J}">
            <div class="cal-day-num ${x?"today-num":""}">${A}</div>
            <div class="cal-events-preview">
              ${te.slice(0,3).map(_=>`
                <div class="cal-event-dot cal-color-${_.color||"gray"}" title="${Ge(_.title||_.type)}">
                  <span class="cal-event-dot-label">${La(_.title||_.branch_name||_.type,18)}</span>
                </div>
              `).join("")}
              ${te.length>3?`<div class="cal-more">+${te.length-3} lagi</div>`:""}
            </div>
          </div>`}let O=(w+f)%7;if(O!==0)for(let A=0;A<7-O;A++)T+='<div class="cal-cell cal-cell-empty"></div>';T+="</div>",d.innerHTML=T,d.querySelectorAll(".cal-cell[data-date]").forEach(A=>{A.addEventListener("click",()=>{let J=A.dataset.date,te=k[J]||[];if(!te.length)return;let x=document.getElementById("cal-event-list"),_=new Date(J+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=_,document.getElementById("cal-event-items").innerHTML=te.map(D=>`
            <div class="cal-event-item cal-color-border-${D.color||"gray"}">
              <div class="cal-event-type">${Aa(D.type)}</div>
              <div class="cal-event-title">${Ge(D.title||"-")}</div>
              <div class="cal-event-branch">${Ge(D.branch_name||"")}</div>
              ${D.status?`<div class="cal-event-status">${Ge(D.status)}</div>`:""}
              ${D.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${D.days_remaining} hari</div>`:""}
            </div>
          `).join(""),x.style.display="block"})})}catch(g){console.error("[Calendar] Render error:",g)}}return await m(!0),r}function La(t,e){return t?t.length>e?t.slice(0,e)+"\u2026":t:""}function Ge(t){return t?String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function Aa(t){return{schedule:"\u{1F5D3} Jadwal",reliever:"\u{1F504} Reliefer",cleaning:"\u{1F9F9} Cleaning",inspection:"\u{1F50E} Inspeksi",fogging:"\u{1F4A8} Fogging",contract_expiry:"\u{1F4CB} Kontrak Habis",issue:"\u26A0\uFE0F Permasalahan",training:"\u{1F4DA} Training",one_on_one:"\u{1F4AC} One on One",basecamp:"\u{1F4DD} Basecamp",supply:"\u{1F4E6} Permintaan"}[t]||t}I();async function Yt(t){let e=ue(),i=(e?.full_name||e?.username||"U")[0].toUpperCase(),o={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";t.innerHTML=`
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
  `;let c=localStorage.getItem("fm_token"),s=document.getElementById("session-info");if(c&&s)try{let r=JSON.parse(atob(c.split(".")[1])),n=new Date(r.exp*1e3);s.textContent=`Berakhir: ${n.toLocaleString("id-ID")}`}catch{s.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async r=>{r.preventDefault();let n=document.getElementById("pwd-error"),l=document.getElementById("pwd-success"),m=document.getElementById("btn-save-pwd");n.style.display="none",l.style.display="none";let p=r.target,d=p.current_password.value,g=p.new_password.value,h=p.confirm_password.value;if(g!==h){n.textContent="\u274C Konfirmasi password tidak cocok.",n.style.display="block";return}if(g.length<6){n.textContent="\u274C Password baru minimal 6 karakter.",n.style.display="block";return}m.disabled=!0,m.textContent="\u23F3 Menyimpan...";let b=await v("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:d,new_password:g})});m.disabled=!1,m.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',b.ok?(l.textContent="\u2705 Password berhasil diubah.",l.style.display="block",p.reset(),V("Password berhasil diubah.")):(n.textContent=b.data?.error||"Gagal mengubah password.",n.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}I();var Qe={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function j(t){if(t==null||t==="")return null;if(t instanceof Date)return isNaN(t.getTime())?null:t.toISOString().slice(0,10);let e=String(t).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let o=Number(e);if(o>2e4&&o<99999){let c=new Date(Date.UTC(1899,11,30)+o*864e5);return isNaN(c.getTime())?null:c.toISOString().slice(0,10)}}let i=e.split(/[\/\-\.]/);if(i.length===3){let[o,c,s]=i.map(m=>m.trim()),r=Number(o),n=Number(c),l=Number(s);if(o.length===4&&r>1900)return`${o}-${c.padStart(2,"0")}-${s.padStart(2,"0")}`;if(s.length===4&&l>1900)return r>12?`${s}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`:n>12?`${s}-${o.padStart(2,"0")}-${c.padStart(2,"0")}`:`${s}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`;if(s.length===2&&!isNaN(l)){let m=l>=50?`19${s}`:`20${s}`;return r>12?`${m}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`:`${m}-${c.padStart(2,"0")}-${o.padStart(2,"0")}`}}let a=new Date(e);return isNaN(a.getTime())?null:a.toISOString().slice(0,10)}function Xt(t){return Object.values(t).every(e=>e==null||String(e).trim()==="")}var Na={validation:{required:[],map:t=>({cabang:t.CABANG,pic:t.PIC,kegiatan:t.KEGIATAN,quartal:t.QUARTAL,masa_pkwt:t["MASA PKWT"],pic_pelapor:t["PIC PELAPOR"],kontrak:t.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:t=>({name:t["Nama SOP"],category:t.Kategori||"Umum",document_link:t["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({full_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",phone:t["No. Hp"],join_date:j(t["Tanggal Masuk"]),status:t.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:t=>({employee_name:t["Nama Lengkap"],branch_name:t.Cabang,division:t["Div / Bagian"]||"FACILITY CARE",start_date:j(t["Tanggal Mulai"]),end_date:j(t["Tanggal Selesai"]),contract_type:t["Tipe Kontrak"]||"",pkwt_number:t.PKWT||"",status:t.Status||"",notes:t.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:t=>({report_date:j(t["Tanggal Info"]),branch_name:t.Cabang,category:t.Kategori,source:t["Sumber Laporan"],complaint:t.Keluhan,employee_name:t["Nama FC"],fc_specialist:t["FC Spesialis"],solution:t.Solusi,status:t.Status||"",completion_date:j(t["Tanggal Selesai"])})},one_on_one:{required:[],map:t=>({meeting_date:j(t.Tanggal),branch_name:t.Cabang,employee_name:t["Nama Karyawan"],pic:t.Pic,problem:t.Masalah,solution:t.Solusi,status:t.Status||"",completion_date:j(t["Tanggal Selesai"]),document_link:t["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:t=>({branch_name:t.Cabang,activity_type:t.Kegiatan,period:t.Periode,pic:t.Pic||t.PIC,opening_date:j(t["Tanggal Opening"]||t["Tgl Opening"]),target_date:j(t["Tanggal Target"]||t["Tgl Target"]),completion_date:j(t["Tanggal Selesai"]||t["Tgl Selesai"]),status:t.Status||"",notes:t.Keterangan||t.Catatan})},inspection:{required:[],map:t=>({inspection_date:j(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",fc_score:t["Point FC SP"]!==void 0&&t["Point FC SP"]!==null?parseFloat(String(t["Point FC SP"]).replace(",",".")):null,spv_score:t["Point SPV"]!==void 0&&t["Point SPV"]!==null?parseFloat(String(t["Point SPV"]).replace(",",".")):null,document_link:t.Link,notes:""})},cleaning:{required:[],map:t=>({activity_date:j(t.Tanggal),branch_name:t.Cabang,activity_type:t["Jenis Kegiatan"]||"General Cleaning",period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},fogging:{required:[],map:t=>({activity_date:j(t.Tanggal),branch_name:t.Cabang,period:t.Periode,status:t.Status||"",document_link:t.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:t=>({info_date:j(t["Tgl Info"]),branch_name:t.Cabang,problem:t.Permasalahan,pic:t.PIC,done_date:j(t["Tgl Done"]),status:t.Status||"",notes:t.Ket})},relievers:{required:[],map:t=>({branch_name:t.Cabang,original_fc_name:t["Nama Facility care"],period:t.Periode,reliever_name:t.Relifer,backup_date:j(t["Tanggal Back Up"]),completion_date:j(t["Tanggal Selesai"]),reason:t.Keterangan,shift:t.Shift,status:t.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:t=>({training_date:j(t.Tanggal),batch:t.Batch,subject:t.Materi,participants:t.Peserta,branch_name:t.Cabang,trainer:t.Trainer,score:t.Nilai!==void 0&&t.Nilai!==null?parseFloat(String(t.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:t=>({name:t["Master Checklist"],category:"Umum",document_link:t["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:t=>({name:t["Master Form"],category:"Umum",document_link:t["Link Document"],description:""})},supply:{required:[],map:t=>({submitted_at:j(t.Timestamp),submitter_name:t["Nama Lengkap"],branch_name:t["Kebutuhan Untuk Cabang"],tools_items:t["Alat - Alat / Barang"],tools_quantity:t["Jumlah Permintaan Alat / Barang"],chemical_items:t.Chemical,chemical_quantity:t["Jumlah Permintaan Chemical"],additional_notes:t["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:t.Status||""})}};function Oa(t,e){let i=Qe[t];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let a=Na[i.module];if(!a)return{valid:[],errors:[],mapped:[],skipped:!0};let o=[],c=[],s=[];return e.filter(n=>!Xt(n)).forEach((n,l)=>{let m=e.indexOf(n)+2,p=[];a.required.forEach(({key:g,label:h})=>{let b=n[g];if(b==null||String(b).trim()===""){let u=Object.keys(n).filter(y=>y.trim()).join(", ");p.push({column:h,originalValue:b||"",reason:`Kolom "${h}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${u.slice(0,120)}`})}});let d=a.map(n);p.length>0?c.push({row:m,data:d,raw:n,errors:p}):(o.push(n),s.push(d))}),{valid:o,errors:c,mapped:s}}function Zt(t){let e=[];return t.SheetNames.forEach(i=>{let a=Qe[i];if(!a)return;let o=t.Sheets[i],c=window.XLSX.utils.sheet_to_json(o,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),s=Oa(i,c),r=c.filter(n=>!Xt(n));e.push({sheetName:i,module:a.module,label:a.label,total:r.length,valid:s.mapped.length,errorCount:s.errors.length,errors:s.errors,mapped:s.mapped,skipped:!1})}),e}function ea(){let t=window.XLSX,e=t.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu","NAMA KARYAWAN":"Budi Santoso",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02",Day:"1"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06",Day:"1","Link Document":""}],"Jadwal Reliefer":[{Cabang:"001. Pondok Bambu","Nama Facility care":"Budi Santoso",Periode:"Januari",Relifer:"Agung Septiadi","Tanggal Back Up":"2024-03-01","Tanggal Selesai":"2024-03-02",Keterangan:"Cuti",Shift:"Pagi",Status:"Done"}]}).forEach(([a,o])=>{t.utils.book_append_sheet(e,t.utils.json_to_sheet(o),a)}),t.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function ta(t){let e=window.XLSX,i=e.utils.book_new(),a=!1;return t.forEach(o=>{if(!o.errors||o.errors.length===0)return;a=!0;let c=o.errors.map(r=>({"No. Baris":r.row,"Kolom Gagal":(r.errors||[]).map(n=>n.column||n).join("; "),"Alasan Error":(r.errors||[]).map(n=>n.reason||n).join("; "),...Object.fromEntries(Object.entries(r.data||{}).map(([n,l])=>[n,l??""]))})),s=e.utils.json_to_sheet(c);e.utils.book_append_sheet(i,s,o.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),a?(e.writeFile(i,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var Fa=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function aa(t){t.innerHTML=`
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
              ${Object.entries(Qe).map(([b,{label:u}])=>`<span class="import-sheet-tag">\u{1F4C4} ${b} \u2192 ${u}</span>`).join("")}
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
  `;let e=null,i=null,a=0,o={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function c(b){Object.entries(o).forEach(([u,y])=>{y.style.display=u===b?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let b=document.getElementById("btn-backup-db");b.disabled=!0,b.textContent="\u23F3 Memproses Backup...";try{let u=await v("/api/import/backup");if(u.ok){let y=new Blob([JSON.stringify(u.data,null,2)],{type:"application/json"}),S=URL.createObjectURL(y),k=document.createElement("a");k.href=S,k.download=`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.json`,document.body.appendChild(k),k.click(),document.body.removeChild(k),URL.revokeObjectURL(S),V("Backup berhasil diunduh!")}else W("Gagal memproses backup: "+(u.data?.error||"Unknown error"))}catch(u){W("Gagal memproses backup: "+u.message)}finally{b.disabled=!1,b.textContent="\u{1F4E6} Backup Database"}});let s=document.getElementById("btn-sync-google");s&&s.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let b=s.innerHTML;s.innerHTML='<span class="spinner"></span> Menyinkronkan...',s.disabled=!0;try{let u=await v("/api/sync/google-sheets",{method:"POST"});u.ok?alert("Sinkronisasi Berhasil: "+(u.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(u.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{s.innerHTML=b,s.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{ea(),V("Template Excel berhasil didownload!")});let r=document.getElementById("file-input"),n=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",b=>{b.stopPropagation(),r.click()}),r.addEventListener("change",b=>{b.target.files[0]&&l(b.target.files[0])}),n.addEventListener("dragover",b=>{b.preventDefault(),n.classList.add("drag-over")}),n.addEventListener("dragleave",()=>n.classList.remove("drag-over")),n.addEventListener("drop",b=>{b.preventDefault(),n.classList.remove("drag-over");let u=b.dataTransfer.files[0];u&&u.name.match(/\.xlsx?$/i)?l(u):W("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,r.value="",document.getElementById("file-info").style.display="none",n.style.display="",c("upload")});async function l(b){e=b,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${b.name} (${(b.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",n.style.display="none",await m(b)}async function m(b){c("validating");let u=document.getElementById("validation-status"),y=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");u.textContent="Membaca file Excel...",y.style.width="20%",await Oe(200);let S=await b.arrayBuffer(),k=window.XLSX.read(S,{type:"array",cellDates:!0});u.textContent=`Memvalidasi ${k.SheetNames.length} sheet...`,y.style.width="50%",await Oe(100),i=Zt(k),y.style.width="100%",u.textContent="Validasi selesai!",await Oe(300),p()}catch(S){c("upload"),W("Gagal memproses file: "+S.message),document.getElementById("file-info").style.display="flex",n.style.display="none"}}function p(){c("preview");let b=i.filter(T=>!T.skipped).length,u=i.reduce((T,R)=>T+R.total,0),y=i.reduce((T,R)=>T+R.valid,0),S=i.reduce((T,R)=>T+R.errorCount,0),k=u>0?Math.round(y/u*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${b} sheet</span>
      <span class="badge badge-secondary">${u} baris</span>
      <span class="badge badge-success">${y} valid (${k}%)</span>
      ${S>0?`<span class="badge badge-danger">${S} error</span>`:""}
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
          ${i.map((T,R)=>`
            <tr class="${T.errorCount>0?"row-error":T.skipped?"row-skipped":"row-ok"}">
              <td><strong>${T.sheetName}</strong></td>
              <td>${T.label}</td>
              <td style="text-align:center">${T.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${T.valid}</span></td>
              <td style="text-align:center">${T.errorCount>0?`<span class="badge badge-danger">${T.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${T.skipped?'<span class="badge badge-neutral">Dilewati</span>':T.errorCount>0&&T.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':T.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':T.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${T.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${R}">\u{1F50D} ${T.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,w.querySelectorAll(".btn-detail-error").forEach(T=>{T.addEventListener("click",()=>{let R=i[Number(T.dataset.idx)];d(R)})});let f=document.getElementById("error-detail-section"),C=document.getElementById("error-detail-container");C.innerHTML="",f.style.display="none";let P=document.getElementById("btn-start-import");y===0?(P.disabled=!0,P.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(P.disabled=!1,S>0?(P.innerHTML=`\u{1F680} Import ${y} Data Valid (${S} dilewati)`,P.title="Baris error akan dilewati, baris valid tetap diimport"):P.innerHTML=`\u{1F680} Mulai Import ${y} Data`)}function d(b){let u=document.getElementById("error-detail-section"),y=document.getElementById("error-detail-container");u.style.display="";let S=b.errors.slice(0,100).map(k=>(Array.isArray(k.errors)?k.errors:[]).map(f=>{let C=typeof f=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${k.row}</span></td>
            <td><strong>${C?f.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${C&&f.originalValue!==void 0?f.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${C?f.reason:f}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${C&&f.aliases?`Gunakan salah satu nama kolom:<br><em>${f.aliases}</em>`:C&&f.hint?f.hint:""}
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
            <tbody>${S||'<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${b.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,u.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{c("upload"),document.getElementById("file-info").style.display="none",n.style.display="",e=null,r.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!i)return;ta(i)?V("Log error berhasil didownload."):V("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let b=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";g(b)});async function g(b){c("importing"),a=Date.now();let u=[];Fa.forEach(f=>{let C=i?.find(P=>P.module===f&&P.mapped?.length>0);C&&u.push(C)});let y=document.getElementById("import-steps-list");y.innerHTML=u.map(f=>`
      <div class="import-step-item" id="step-item-${f.module}">
        <span class="step-item-icon" id="step-icon-${f.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${f.label} <span class="step-item-count">(${f.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${f.module}"></span>
      </div>
    `).join("");let S=document.getElementById("import-bar"),k=document.getElementById("import-current-status"),w={totalSheets:u.length,totalRows:u.reduce((f,C)=>f+C.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let f=0;f<u.length;f++){let C=u[f],P=document.getElementById(`step-icon-${C.module}`),T=document.getElementById(`step-status-${C.module}`);P.textContent="\u{1F504}",T.textContent="Mengimport...",k.textContent=`Mengimport ${C.label}...`,S.style.width=`${Math.round(f/u.length*100)}%`;try{let R=await v(`/api/import/${C.module}`,{method:"POST",body:JSON.stringify({rows:C.mapped,onDuplicate:b})});if(R.ok){let O=R.data;w.inserted+=O.inserted||0,w.skipped+=O.skipped||0,w.moduleResults.push({label:C.label,inserted:O.inserted||0,skipped:O.skipped||0,status:"ok"}),P.textContent="\u2705",T.innerHTML=`<span class="badge badge-success">${O.inserted||0} berhasil</span>${O.skipped>0?` <span class="badge badge-neutral">${O.skipped} skip</span>`:""}`}else w.failed++,w.moduleResults.push({label:C.label,inserted:0,skipped:0,status:"error",error:R.data?.error}),P.textContent="\u274C",T.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(R){w.failed++,w.moduleResults.push({label:C.label,inserted:0,skipped:0,status:"error",error:R.message}),P.textContent="\u274C",T.innerHTML='<span class="badge badge-danger">Gagal</span>'}await Oe(150)}S.style.width="100%",k.textContent="Selesai!",await Oe(400),ge("schedule"),h(w)}function h(b){c("summary");let u=((Date.now()-a)/1e3).toFixed(1),y=b.failed===0;document.getElementById("summary-status-icon").innerHTML=`
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
        <div class="stat-value">${u}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${b.moduleResults.map(S=>`
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,i=null,r.value="",document.getElementById("file-info").style.display="none",n.style.display="",c("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function Oe(t){return new Promise(e=>setTimeout(e,t))}I();var Ve=[],na=[];async function ia(t){Ve=await B(),na=await H(),$({container:t,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:e=>e?`<span class="badge badge-info">${e}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:e=>`<span class="badge badge-warning">${e||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"branch_id",label:"Cabang",options:Ve}],exportOptions:{moduleName:"sp_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await v(`/api/sp?limit=10000&${i}`);if(a.ok){let o=a.data.data.map(s=>({"Nama Karyawan":s.employee_name||"",Divisi:s.division||"",Cabang:s.branch_name||"","Tanggal Sp":s.tanggal||"","Akhir Sp":s.akhir_sp||"","Jenis Sp":s.sp_type||"","Link Document / Foto":s.document_link||""})),{downloadExcel:c}=await Promise.resolve().then(()=>(L(),X));c(o,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(L(),X));i(e,"Template_Import_SP")},onImport:async e=>{let i=s=>{if(!s)return null;let r=String(s||"").toLowerCase(),n=Ve.find(l=>String(l.label||"").toLowerCase()===r);return n?n.value:null},a=s=>{if(!s)return"";if(s instanceof Date&&!isNaN(s.getTime()))return s.toISOString().slice(0,10);let r=String(s).trim();if(/^\d{4,5}$/.test(r)){let l=Number(r);if(l>2e4&&l<99999){let m=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(r))return r.slice(0,10);let n=r.split(/[\/\-\.]/);if(n.length===3){let[l,m,p]=n.map(d=>d.trim());if(l.length===4&&m.length<=2&&p.length<=2)return`${l}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&m.length<=2&&l.length<=2)return`${p}-${m.padStart(2,"0")}-${l.padStart(2,"0")}`}return r},o=e.map(s=>({employee_name:String(s["Nama Karyawan"]||"").trim(),division:String(s.Divisi||"").trim(),branch_id:i(String(s.Cabang||"").trim()),tanggal:a(s["Tanggal Sp"]),akhir_sp:a(s["Akhir Sp"]),sp_type:String(s["Jenis Sp"]||"").trim(),document_link:String(s["Link Document / Foto"]||"").trim()})).filter(s=>s.employee_name&&s.branch_id),c=await v("/api/sp/import",{method:"POST",body:JSON.stringify(o)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}},formFields:[{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:na},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"combobox",name:"branch_id",label:"Cabang",required:!0,options:Ve,createApi:{path:"/api/branches",field:"full_name"}},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}I();var xe=[],ra=[];async function sa(t){xe=await B(),ra=await H(),$({container:t,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:e=>e?new Date(e).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:e=>`<span class="badge ${e==="Selesai"?"badge-success":"badge-warning"}">${e||"-"}</span>`},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",options:xe},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",options:xe}],exportOptions:{moduleName:"mutasi_data",onExport:async e=>{let i=new URLSearchParams(e||{}).toString(),a=await v(`/api/mutasi?limit=10000&${i}`);if(a.ok){let o=a.data.data.map(s=>({Tanggal:s.tanggal||"","Nama Karyawan":s.employee_name||"","Cabang Asal":s.from_branch_name||"","Cabang Tujuan":s.to_branch_name||"",Status:s.status||"",Dokumen:s.document_link||""})),{downloadExcel:c}=await Promise.resolve().then(()=>(L(),X));c(o,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(L(),X));i(e,"Template_Import_Mutasi")},onImport:async e=>{let i=s=>{if(!s)return null;let r=String(s||"").toLowerCase(),n=xe.find(l=>String(l.label||"").toLowerCase()===r);return n?n.value:null},a=s=>{if(!s)return"";if(s instanceof Date&&!isNaN(s.getTime()))return s.toISOString().slice(0,10);let r=String(s).trim();if(/^\d{4,5}$/.test(r)){let l=Number(r);if(l>2e4&&l<99999){let m=new Date(Date.UTC(1899,11,30)+l*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(r))return r.slice(0,10);let n=r.split(/[\/\-\.]/);if(n.length===3){let[l,m,p]=n.map(d=>d.trim());if(l.length===4&&m.length<=2&&p.length<=2)return`${l}-${m.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&m.length<=2&&l.length<=2)return`${p}-${m.padStart(2,"0")}-${l.padStart(2,"0")}`}return r},o=e.map(s=>({tanggal:a(s.Tanggal),employee_name:String(s["Nama Karyawan"]||"").trim(),from_branch_id:i(String(s["Cabang Asal"]||"").trim()),to_branch_id:i(String(s["Cabang Tujuan"]||"").trim()),status:String(s.Status||"").trim(),document_link:String(s.Dokumen||"").trim()})).filter(s=>s.tanggal&&s.employee_name&&s.from_branch_id&&s.to_branch_id),c=await v("/api/mutasi/import",{method:"POST",body:JSON.stringify(o)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:ra},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:xe,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:xe,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}window.parseFlexibleDate=t=>{if(!t||t==="-")return"";if(t=String(t).trim(),/^\d{5}$/.test(t)){let i=Math.floor(Number(t)-25569);return new Date(i*86400*1e3).toISOString().split("T")[0]}if(t.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)){let i=t.split(/[\/\-]/);return`${i[2]}-${i[1]}-${i[0]}`}let e=t.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);if(e){let i=e[1],a=parseInt(e[2],10),o=parseInt(e[3],10);if(a>12&&o<=12)return`${i}-${e[3]}-${e[2]}`}return t.split("T")[0]};window.formatDate=t=>{let e=window.parseFlexibleDate(t);if(!e)return"";let i=e.split("-");if(i.length===3&&i[0].length===4){let a=["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"],o=parseInt(i[2],10),c=a[parseInt(i[1],10)-1];return`${o} ${c} ${i[0]}`}return e};function K(t){return async e=>{if(!Ee()){ye("/login");return}return t(e)}}var Fe=null;function Ma(){Fe&&clearInterval(Fe);let t=()=>{let e=new Date,i=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),a=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),o=document.getElementById("header-clock-time"),c=document.getElementById("header-clock-date");o&&(o.textContent=i),c&&(c.textContent=a)};t(),Fe=setInterval(t,1e3)}async function Ka(){try{let t=await v("/api/dashboard/kpi");if(!t.ok)return;let e=t.data?.data||t.data||{},i=(a,o)=>{let c=document.getElementById(a);c&&(c.textContent=o>0?o:"",c.style.display=o>0?"inline-flex":"none")};i("badge-issues",e.issues?.current||0),i("badge-contracts",e.expiring30?.current||0),i("badge-oo1",e.one_on_one?.current||0),i("badge-schedule",e.schedule?.current||0),i("badge-supply",e.supply?.current||0)}catch{}}var Ce=[];async function Ra(){try{let t=await v("/api/dashboard/notifications");if(!t.ok)return;Ce=t.data?.data||t.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=Ce.length>0?"block":"none",e.textContent=Ce.length)}catch{}}function qa(){if(!Ce.length){se({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,i)=>i()});return}let t=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${Ce.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;se({title:`Notifikasi (${Ce.length})`,content:t,confirmText:"Tutup",onConfirm:(e,i)=>i()})}function oa(){let t=ue(),e=(t?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
  `;let i=document.getElementById("sidebar"),a=document.getElementById("sidebar-overlay"),o=document.getElementById("topbar-menu-btn"),c=document.getElementById("sidebar-close"),s=()=>{i.classList.add("open"),a.classList.add("show")},r=()=>{i.classList.remove("open"),a.classList.remove("show")};o?.addEventListener("click",s),c?.addEventListener("click",r),a?.addEventListener("click",r),document.querySelectorAll(".nav-item").forEach(l=>l.addEventListener("click",r));function n(){let l=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(d=>{let g=d.dataset.route;d.classList.toggle("active",l===g||g!=="/dashboard"&&l.startsWith(g))});let m=document.getElementById("topbar-title"),p=document.querySelector(".nav-item.active .nav-label");m&&p&&(m.textContent=p.textContent)}window.addEventListener("hashchange",n),n(),Ma(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await v("/api/auth/logout",{method:"POST"}),De(),Fe&&clearInterval(Fe),ye("/login")}),Ka(),Ra(),document.getElementById("btn-notif")?.addEventListener("click",l=>{l.preventDefault(),qa()})}async function Ha(){M("/login",({main:e})=>At(e)),M("/dashboard",K(({main:e})=>Pt(e))),M("/calendar",K(({main:e})=>Wt(e))),M("/employees",K(({main:e,params:i})=>Nt(e,i))),M("/contracts",K(({main:e,params:i})=>Ue(e,i))),M("/sp",K(({main:e})=>ia(e))),M("/mutasi",K(({main:e})=>sa(e))),M("/timeline",K(({main:e,params:i})=>Ft(e,i))),M("/issues",K(({main:e,params:i})=>Mt(e,i))),M("/one-on-one",K(({main:e,params:i})=>Kt(e,i))),M("/training",K(({main:e})=>Rt(e))),M("/relievers",K(({main:e,params:i})=>qt(e,i))),M("/reports/inspection",K(({main:e})=>Ht(e))),M("/reports/cleaning",K(({main:e})=>jt(e))),M("/reports/fogging",K(({main:e})=>Ut(e))),M("/reports/basecamp",K(({main:e})=>Jt(e))),M("/reports/supply",K(({main:e})=>ht(e,"supply"))),M("/sop",K(({main:e})=>Gt(e))),M("/checklist",K(({main:e})=>Qt(e))),M("/forms",K(({main:e})=>ht(e))),M("/users",K(({main:e})=>Vt(e))),M("/branches",K(({main:e})=>zt(e))),M("/profile",K(({main:e})=>Yt(e))),M("/settings/import",K(({main:e})=>aa(e)));let t=Ee();if(!t&&window.location.hash!=="#/login"&&ye("/login"),t){let e=await v("/api/auth/me");e.ok?(Ie(e.data.data),oa()):(De(),ye("/login"))}window.addEventListener("fm:login",()=>{oa(),ye("/dashboard")}),ft()}Ha();
