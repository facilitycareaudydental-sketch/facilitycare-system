var Oe=["localhost","127.0.0.1","0.0.0.0","::1"],le=window.__FM_CONFIG?.API_BASE_URL||(Oe.includes(window.location.hostname)?"http://127.0.0.1:8787":void 0);if(!le)throw new Error("Missing API_BASE_URL configuration. Set window.__FM_CONFIG.API_BASE_URL for production.");var Re=le;function N(){return localStorage.getItem("fm_token")}function re(n){localStorage.setItem("fm_token",n)}function K(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function I(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function H(n){localStorage.setItem("fm_user",JSON.stringify(n))}async function p(n,e={}){let s=N(),t={"Content-Type":"application/json",...s?{Authorization:`Bearer ${s}`}:{},...e.headers||{}},a=await fetch(`${Re}${n}`,{...e,headers:t}),l=await a.json();return a.status===401&&(K(),window.location.hash="#/login"),{ok:a.ok,status:a.status,data:l}}var W={},U=null;function w(n,e){W[n]=e}function M(n){window.location.hash=n}function ie(){async function n(){let e=window.location.hash.replace("#","")||"/dashboard",[s,...t]=e.split("?"),a=W[s];if(!a){for(let[d,i]of Object.entries(W))if(d.endsWith("/*")&&s.startsWith(d.slice(0,-2))){a=i;break}}U&&(U(),U=null);let l=document.getElementById("main-content");if(l)if(l.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>',a){let d=new URLSearchParams(t.join("?")),i=s.split("/").filter(Boolean),r=await a({path:s,params:d,segments:i,main:l});r&&(U=r)}else l.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>'}window.addEventListener("hashchange",n),n()}var q;function Ne(){return q||(q=document.createElement("div"),q.id="toast-container",document.body.appendChild(q)),q}function oe(n,e="info",s=3500){let t=Ne(),a=document.createElement("div");a.className=`toast toast-${e}`;let l={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};a.innerHTML=`<span class="toast-icon">${l[e]||"\u2139"}</span><span class="toast-msg">${n}</span>`,t.appendChild(a),requestAnimationFrame(()=>a.classList.add("show")),setTimeout(()=>{a.classList.remove("show"),setTimeout(()=>a.remove(),350)},s)}var B=n=>oe(n,"success"),O=n=>oe(n,"error");function f(n){return`<span class="badge ${{Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"}[n]||"badge-neutral"}">${n||"-"}</span>`}function Q(n){return n==null?'<span class="badge badge-neutral">-</span>':n<0?`<span class="badge badge-danger">Expired (${Math.abs(n)}h)</span>`:n<=14?`<span class="badge badge-danger">${n} hari</span>`:n<=30?`<span class="badge badge-warning">${n} hari</span>`:`<span class="badge badge-success">${n} hari</span>`}function j(n){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[n]||"badge-neutral"}">${n||"-"}</span>`}function J(n){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[n]||"badge-neutral"}">${n||"-"}</span>`}function E(n){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[n]||"badge-neutral"}">${n||"-"}</span>`}async function ce(n){n.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">Dashboard</h1>
      <div class="page-actions">
        <span class="text-muted" id="dash-date"></span>
      </div>
    </div>
    <div id="dash-content">
      <div class="loading-spinner"><div class="spinner"></div></div>
    </div>
  `,document.getElementById("dash-date").textContent=new Date().toLocaleDateString("id-ID",{weekday:"long",year:"numeric",month:"long",day:"numeric"});let[e,s]=await Promise.all([p("/api/dashboard/stats"),p("/api/dashboard/issues-summary")]),t=e.data?.data||{},a=s.data?.data||{};document.getElementById("dash-content").innerHTML=`
    <!-- Stat cards -->
    <div class="stats-grid">
      <div class="stat-card stat-card-blue">
        <div class="stat-icon">\u{1F465}</div>
        <div class="stat-body">
          <div class="stat-value">${t.total_employees||0}</div>
          <div class="stat-label">Karyawan Aktif</div>
        </div>
      </div>
      <div class="stat-card stat-card-green">
        <div class="stat-icon">\u{1F4CB}</div>
        <div class="stat-body">
          <div class="stat-value">${t.active_contracts||0}</div>
          <div class="stat-label">Kontrak Aktif</div>
        </div>
      </div>
      <div class="stat-card stat-card-yellow">
        <div class="stat-icon">\u23F0</div>
        <div class="stat-body">
          <div class="stat-value">${t.expiring_contracts||0}</div>
          <div class="stat-label">Kontrak Habis 30 Hari</div>
        </div>
      </div>
      <div class="stat-card stat-card-red">
        <div class="stat-icon">\u26A0\uFE0F</div>
        <div class="stat-body">
          <div class="stat-value">${t.open_issues||0}</div>
          <div class="stat-label">Permasalahan Open</div>
        </div>
      </div>
      <div class="stat-card stat-card-purple">
        <div class="stat-icon">\u{1F91D}</div>
        <div class="stat-body">
          <div class="stat-value">${t.open_one_on_one||0}</div>
          <div class="stat-label">One on One Pending</div>
        </div>
      </div>
      <div class="stat-card stat-card-teal">
        <div class="stat-icon">\u{1F5D3}\uFE0F</div>
        <div class="stat-body">
          <div class="stat-value">${t.pending_schedule||0}</div>
          <div class="stat-label">Jadwal Pending</div>
        </div>
      </div>
      <div class="stat-card stat-card-orange">
        <div class="stat-icon">\u{1F4E6}</div>
        <div class="stat-body">
          <div class="stat-value">${t.pending_supply_requests||0}</div>
          <div class="stat-label">Permintaan Barang</div>
        </div>
      </div>
      <div class="stat-card stat-card-gray">
        <div class="stat-icon">\u{1F3E2}</div>
        <div class="stat-body">
          <div class="stat-value">${t.total_branches||0}</div>
          <div class="stat-label">Total Cabang</div>
        </div>
      </div>
    </div>

    <!-- Charts row -->
    <div class="dashboard-grid">
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">Permasalahan per Kategori</h3>
        </div>
        <div class="card-body">
          <div id="chart-issues-category"></div>
        </div>
      </div>
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">Tren Permasalahan (12 Bulan)</h3>
        </div>
        <div class="card-body">
          <div id="chart-issues-trend"></div>
        </div>
      </div>
    </div>

    <!-- Tables row -->
    <div class="dashboard-grid">
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">\u23F0 Kontrak Akan Habis</h3>
          <a href="#/contracts" class="btn btn-sm btn-ghost">Lihat Semua</a>
        </div>
        <div class="card-body p-0">
          <div class="table-wrapper">
            <table class="data-table">
              <thead><tr><th>Nama</th><th>Cabang</th><th>Sisa</th></tr></thead>
              <tbody id="expiring-tbody"></tbody>
            </table>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">\u26A0\uFE0F Permasalahan Terbaru</h3>
          <a href="#/issues" class="btn btn-sm btn-ghost">Lihat Semua</a>
        </div>
        <div class="card-body p-0">
          <div class="table-wrapper">
            <table class="data-table">
              <thead><tr><th>Tanggal</th><th>Cabang</th><th>Status</th></tr></thead>
              <tbody id="issues-tbody"></tbody>
            </table>
          </div>
        </div>
      </div>
    </div>

    <!-- Upcoming schedule -->
    <div class="card">
      <div class="card-header">
        <h3 class="card-title">\u{1F5D3}\uFE0F Jadwal Mendatang</h3>
        <a href="#/schedule" class="btn btn-sm btn-ghost">Lihat Semua</a>
      </div>
      <div class="card-body p-0">
        <div class="table-wrapper">
          <table class="data-table">
            <thead><tr><th>Cabang</th><th>Kegiatan</th><th>Periode</th><th>Target</th><th>PIC</th><th>Status</th></tr></thead>
            <tbody id="schedule-tbody"></tbody>
          </table>
        </div>
      </div>
    </div>
  `;let l=document.getElementById("expiring-tbody");t.expiring_contracts===0&&(l.innerHTML='<tr><td colspan="3" class="text-center text-muted">Tidak ada kontrak yang akan habis</td></tr>'),(e.data?.data?.expiring_contracts||[]).forEach(c=>{l.innerHTML+=`
      <tr>
        <td>${c.employee_name}</td>
        <td>${c.branch_name||"-"}</td>
        <td>${Q(c.days_remaining)}</td>
      </tr>`});let d=document.getElementById("issues-tbody"),i=e.data?.data?.recent_issues||[];i.length===0&&(d.innerHTML='<tr><td colspan="3" class="text-center text-muted">Tidak ada permasalahan</td></tr>'),i.forEach(c=>{d.innerHTML+=`
      <tr>
        <td>${c.report_date||"-"}</td>
        <td>${c.branch_name||"-"}</td>
        <td>${f(c.status)}</td>
      </tr>`});let r=document.getElementById("schedule-tbody"),u=e.data?.data?.upcoming_schedule||[];u.length===0&&(r.innerHTML='<tr><td colspan="6" class="text-center text-muted">Tidak ada jadwal mendatang</td></tr>'),u.forEach(c=>{r.innerHTML+=`
      <tr>
        <td>${c.branch_name||"-"}</td>
        <td>${J(c.activity_type)}</td>
        <td><span class="badge badge-info">${c.period}</span></td>
        <td>${c.target_date||"-"}</td>
        <td>${c.pic||"-"}</td>
        <td>${f(c.status)}</td>
      </tr>`}),Ke("chart-issues-category",a.by_category||[],"category","count","Permasalahan"),He("chart-issues-trend",a.by_month||[],"month","count")}function Ke(n,e,s,t,a){let l=document.getElementById(n);if(!l||!e.length){l&&(l.innerHTML='<p class="text-center text-muted">Tidak ada data</p>');return}let d=Math.max(...e.map(i=>i[t]));l.innerHTML=`
    <div class="bar-chart">
      ${e.slice(0,8).map(i=>`
        <div class="bar-item">
          <div class="bar-label">${i[s]}</div>
          <div class="bar-track">
            <div class="bar-fill" style="width:${d>0?i[t]/d*100:0}%">
              <span class="bar-value">${i[t]}</span>
            </div>
          </div>
        </div>
      `).join("")}
    </div>
  `}function He(n,e,s,t){let a=document.getElementById(n);if(!a||!e.length){a&&(a.innerHTML='<p class="text-center text-muted">Tidak ada data</p>');return}let l=Math.max(...e.map(c=>c[t]),1),d=400,i=120,r=30,u=e.map((c,k)=>{let m=r+k/(e.length-1||1)*(d-r*2),S=i-r-c[t]/l*(i-r*2);return`${m},${S}`}).join(" ");a.innerHTML=`
    <div class="line-chart-wrap">
      <svg viewBox="0 0 ${d} ${i}" class="line-chart-svg" role="img" aria-label="Tren permasalahan">
        <polyline points="${u}" class="line-path" fill="none"/>
        ${e.map((c,k)=>{let m=r+k/(e.length-1||1)*(d-r*2),S=i-r-c[t]/l*(i-r*2);return`<circle cx="${m}" cy="${S}" r="4" class="line-dot"/>
                  <text x="${m}" y="${i-5}" text-anchor="middle" class="chart-x-label">${(c[s]||"").slice(5)}</text>`}).join("")}
      </svg>
    </div>
  `}async function de(n){document.getElementById("app").innerHTML=`
    <div class="login-page">
      <div class="login-card">
        <div class="login-header">
          <div class="login-logo">\u{1F3E5}</div>
          <h1 class="login-title">FM<strong>Ops</strong></h1>
          <p class="login-subtitle">Facility Management Operations</p>
        </div>
        <form class="login-form" id="login-form" novalidate>
          <div class="form-group">
            <label class="form-label">Username / Email</label>
            <input type="text" name="username" class="form-control" placeholder="username atau email" required autofocus autocomplete="username">
          </div>
          <div class="form-group">
            <label class="form-label">Password</label>
            <div class="input-with-icon">
              <input type="password" name="password" class="form-control" placeholder="\u2022\u2022\u2022\u2022\u2022\u2022\u2022\u2022" required autocomplete="current-password" id="login-password">
              <button type="button" class="input-icon-btn" id="toggle-password" aria-label="Toggle password">\u{1F441}</button>
            </div>
          </div>
          <div id="login-error" class="alert alert-danger" style="display:none"></div>
          <button type="submit" class="btn btn-primary btn-full btn-lg" id="login-btn">
            <span class="btn-text">Masuk</span>
            <span class="btn-spinner" style="display:none">\u23F3 Loading...</span>
          </button>
        </form>
        <div class="login-footer">
          <a href="#/form/chemical" class="link-subtle">\u{1F4E6} Form Permintaan Barang (Tanpa Login)</a>
        </div>
      </div>
    </div>
  `;let e=document.getElementById("login-form"),s=document.getElementById("login-error"),t=document.getElementById("login-btn"),a=document.getElementById("toggle-password"),l=document.getElementById("login-password");a?.addEventListener("click",()=>{l.type=l.type==="password"?"text":"password"}),e?.addEventListener("submit",async d=>{d.preventDefault(),s.style.display="none";let i=e.username.value.trim(),r=e.password.value;if(!i||!r){s.textContent="Username dan password wajib diisi.",s.style.display="block";return}t.querySelector(".btn-text").style.display="none",t.querySelector(".btn-spinner").style.display="",t.disabled=!0;try{let u=await p("/api/auth/login",{method:"POST",body:JSON.stringify({username:i,password:r})});u.ok&&u.data.success?(re(u.data.data.token),H(u.data.data.user),B("Login berhasil! Selamat datang."),window.dispatchEvent(new Event("fm:login"))):(s.textContent=u.data.error||"Username atau password salah.",s.style.display="block")}catch{s.textContent="Gagal terhubung ke server. Periksa koneksi internet.",s.style.display="block"}finally{t.querySelector(".btn-text").style.display="",t.querySelector(".btn-spinner").style.display="none",t.disabled=!1}})}function pe({columns:n,data:e,onEdit:s,onDelete:t,onView:a,actions:l=[],emptyText:d="Tidak ada data"}){let i=document.createElement("div");if(i.className="table-wrapper",!e||e.length===0)return i.innerHTML=`<div class="empty-state"><p>${d}</p></div>`,i;let r=document.createElement("table");r.className="data-table";let u=document.createElement("thead"),c=document.createElement("tr");if(n.forEach(m=>{let S=document.createElement("th");S.textContent=m.label,m.width&&(S.style.width=m.width),c.appendChild(S)}),s||t||a||l.length>0){let m=document.createElement("th");m.textContent="Aksi",m.style.width="120px",c.appendChild(m)}u.appendChild(c),r.appendChild(u);let k=document.createElement("tbody");return e.forEach(m=>{let S=document.createElement("tr");if(n.forEach($=>{let _=document.createElement("td");if($.render){let o=$.render(m[$.key],m);o instanceof HTMLElement?_.appendChild(o):_.innerHTML=o||""}else _.textContent=m[$.key]!==null&&m[$.key]!==void 0?m[$.key]:"-";$.nowrap&&(_.style.whiteSpace="nowrap"),S.appendChild(_)}),s||t||a||l.length>0){let $=document.createElement("td");$.className="actions-cell";let _=document.createElement("div");if(_.className="btn-group",a){let o=document.createElement("button");o.className="btn btn-xs btn-ghost",o.innerHTML="\u{1F441}",o.title="Lihat",o.addEventListener("click",()=>a(m)),_.appendChild(o)}if(s){let o=document.createElement("button");o.className="btn btn-xs btn-secondary",o.innerHTML="\u270F\uFE0F",o.title="Edit",o.addEventListener("click",()=>s(m)),_.appendChild(o)}if(l.forEach(o=>{let v=document.createElement("button");v.className=`btn btn-xs ${o.class||"btn-ghost"}`,v.innerHTML=o.icon||o.label,v.title=o.label,v.addEventListener("click",()=>o.handler(m)),_.appendChild(v)}),t){let o=document.createElement("button");o.className="btn btn-xs btn-danger",o.innerHTML="\u{1F5D1}\uFE0F",o.title="Hapus",o.addEventListener("click",()=>t(m)),_.appendChild(o)}$.appendChild(_),S.appendChild($)}k.appendChild(S)}),r.appendChild(k),i.appendChild(r),i}function ue({page:n,pages:e,total:s,limit:t,onPage:a}){if(e<=1)return null;let l=document.createElement("div");l.className="pagination";let d=document.createElement("span");d.className="pagination-info",d.textContent=`Total: ${s} data`,l.appendChild(d);let i=document.createElement("div");i.className="pagination-btns";let r=(k,m,S=!1,$=!1)=>{let _=document.createElement("button");_.className=`btn btn-sm ${$?"btn-primary":"btn-ghost"} pagination-btn`,_.textContent=k,_.disabled=S,_.addEventListener("click",()=>a(m)),i.appendChild(_)};r("\xAB",1,n===1),r("\u2039",n-1,n===1);let u=Math.max(1,n-2),c=Math.min(e,n+2);for(let k=u;k<=c;k++)r(k,k,!1,k===n);return r("\u203A",n+1,n===e),r("\xBB",e,n===e),l.appendChild(i),l}function R({title:n,content:e,onConfirm:s,onCancel:t,confirmText:a="Simpan",cancelText:l="Batal",size:d="md",confirmClass:i="btn-primary"}){let r={sm:"400px",md:"560px",lg:"720px",xl:"900px"},u=document.createElement("div");u.className="modal-overlay",u.innerHTML=`
    <div class="modal" style="max-width:${r[d]||r.md}">
      <div class="modal-header">
        <h3 class="modal-title">${n}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${l}</button>
        ${s?`<button class="btn ${i} modal-confirm">${a}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&u.querySelector(".modal-body").appendChild(e);let c=()=>{u.classList.remove("show"),setTimeout(()=>u.remove(),250)};return u.querySelector(".modal-close").addEventListener("click",()=>{t&&t(),c()}),u.querySelector(".modal-cancel").addEventListener("click",()=>{t&&t(),c()}),s&&u.querySelector(".modal-confirm").addEventListener("click",()=>s(u,c)),u.addEventListener("click",k=>{k.target===u&&(t&&t(),c())}),document.body.appendChild(u),requestAnimationFrame(()=>u.classList.add("show")),{overlay:u,close:c}}function me(n,e,s="Konfirmasi"){return R({title:s,content:`<p>${n}</p>`,onConfirm:(t,a)=>{e(),a()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}function z(n){return n.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="row")return`<div class="form-row">${z(e.fields)}</div>`;let s=e.required?"required":"",t=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",a="";switch(e.type){case"textarea":a=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${s} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let d=(e.options||[]).map(i=>{let r=typeof i=="object"?i.value:i,u=typeof i=="object"?i.label:i,c=e.value==r?"selected":"";return`<option value="${r}" ${c}>${u}</option>`}).join("");a=`<select name="${e.name}" class="form-control" ${s}><option value="">-- Pilih ${e.label||""} --</option>${d}</select>`;break;case"checkbox":a=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":a=`<input type="date" name="${e.name}" class="form-control" value="${e.value||""}" ${s}>`;break;case"number":a=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${s}>`;break;case"email":a=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${s}>`;break;case"url":a=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${s}>`;break;default:a=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${s}>`}let l=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${t}${a}${l}</div>`}).join("")}function be(n){let e={},s=new FormData(n);for(let[t,a]of s.entries())e[t]=a===""?null:a;return n.querySelectorAll("input[type=checkbox]").forEach(t=>{t.checked||(e[t.name]=null)}),e}function ge(n,e){e&&Object.entries(e).forEach(([s,t])=>{let a=n.querySelector(`[name="${s}"]`);a&&(a.type==="checkbox"?a.checked=!!t:a.value=t??"")})}function g({container:n,title:e,icon:s,apiPath:t,columns:a,formFields:l,filterFields:d,defaultFilters:i={},itemLabel:r="Data",canCreate:u=!0,canEdit:c=!0,canDelete:k=!0,onBeforeSubmit:m,onAfterLoad:S,extraActions:$=[],initialSearch:_=""}){let o=1,v={...i};_&&(v.search=_),n.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${s} ${e}</h1>
      <div class="page-actions">
        ${u?`<button class="btn btn-primary" id="btn-create">+ Tambah ${r}</button>`:""}
      </div>
    </div>

    ${d&&d.length>0?`
    <div class="filter-bar card">
      <div class="filter-bar-inner">
        ${d.map(b=>b.type==="search"?`<div class="filter-search"><input type="search" class="form-control" placeholder="${b.placeholder||"Cari..."}" id="filter-search" value="${v.search||""}"></div>`:b.type==="select"?`<select class="form-control filter-select" name="${b.name}" id="filter-${b.name}"><option value="">-- ${b.label} --</option>${(b.options||[]).map(y=>`<option value="${typeof y=="object"?y.value:y}" ${v[b.name]===(typeof y=="object"?y.value:y)?"selected":""}>${typeof y=="object"?y.label:y}</option>`).join("")}</select>`:"").join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
      </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;let x=document.getElementById("filter-search"),F;x?.addEventListener("input",b=>{clearTimeout(F),F=setTimeout(()=>{v.search=b.target.value,o=1,h()},400)}),d?.forEach(b=>{b.type==="select"&&document.getElementById(`filter-${b.name}`)?.addEventListener("change",y=>{v[b.name]=y.target.value,o=1,h()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{v={...i},x&&(x.value=""),d?.forEach(b=>{let y=document.getElementById(`filter-${b.name}`);y&&(y.value="")}),o=1,h()}),document.getElementById("btn-create")?.addEventListener("click",()=>ne(null));async function h(){let b=document.getElementById("table-container");if(!b)return;b.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let y=new URLSearchParams({page:o,limit:20,...Object.fromEntries(Object.entries(v).filter(([,P])=>P))}),T=await p(`${t}?${y}`);if(!T.ok){b.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${T.data?.error||"Error"}</p></div>`;return}let Y=T.data?.data||[],L=T.data?.pagination;S&&S(Y);let G=pe({columns:a,data:Y,onEdit:c?P=>ne(P):null,onDelete:k?P=>Ae(P):null,actions:$.map(P=>({...P,handler:A=>P.handler(A,h)})),emptyText:`Tidak ada ${r.toLowerCase()}`});b.innerHTML="",b.appendChild(G);let D=document.getElementById("pagination-container");if(D&&(D.innerHTML="",L&&L.pages>1)){let P=ue({page:L.page,pages:L.pages,total:L.total,limit:L.limit,onPage:A=>{o=A,h()}});P&&D.appendChild(P)}}function Me(b){let y=typeof l=="function"?l(b):l;return z(y)}function ne(b){let y=!!b,T=document.createElement("form");if(T.noValidate=!0,T.innerHTML=Me(b),y){let L=typeof l=="function"?l(b):l;ge(T,b)}let{close:Y}=R({title:y?`Edit ${r}`:`Tambah ${r}`,content:T,size:"lg",confirmText:y?"Simpan Perubahan":`Tambah ${r}`,onConfirm:async(L,G)=>{let D=L.querySelector(".modal-confirm");D.disabled=!0,D.textContent="Menyimpan...";let P=be(T);m&&(P=await m(P,b));let A=y?"PUT":"POST",qe=y?`${t}/${b.id}`:t,se=await p(qe,{method:A,body:JSON.stringify(P)});se.ok?(B(y?`${r} berhasil diperbarui.`:`${r} berhasil ditambahkan.`),G(),h()):(O(se.data?.error||"Gagal menyimpan data."),D.disabled=!1,D.textContent=y?"Simpan Perubahan":`Tambah ${r}`)}})}function Ae(b){me(`Hapus ${r} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let y=await p(`${t}/${b.id}`,{method:"DELETE"});y.ok?(B(`${r} berhasil dihapus.`),h()):O(y.data?.error||"Gagal menghapus.")},`Hapus ${r}`)}return h(),h}var V=[];async function Ue(){V=((await p("/api/branches?all=1")).data?.data||[]).map(e=>({value:e.id,label:e.full_name}))}async function ye(n){await Ue(),g({container:n,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:e=>j(e)},{key:"phone",label:"No. HP",render:e=>e?`<a href="tel:${e}">${e}</a>`:"-"},{key:"join_date",label:"Tgl Masuk"},{key:"status",label:"Status",render:e=>f(e)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:V},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:e=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:e?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:e?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:V,value:e?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:e?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:e?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:e?.status||"Aktif"}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:e?.notes}]})}var X=[],ve=[];async function Qe(){let[n,e]=await Promise.all([p("/api/branches?all=1"),p("/api/employees?limit=500&status=Aktif")]);X=(n.data?.data||[]).map(s=>({value:s.id,label:s.full_name})),ve=(e.data?.data||[]).map(s=>({value:s.id,label:s.full_name}))}async function he(n){await Qe(),g({container:n,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",itemLabel:"Kontrak",columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:e=>j(e)},{key:"start_date",label:"Tgl Mulai",nowrap:!0},{key:"end_date",label:"Tgl Selesai",nowrap:!0},{key:"days_remaining",label:"Sisa",render:e=>Q(e)},{key:"contract_type",label:"Tipe Kontrak"},{key:"pkwt_number",label:"PKWT"},{key:"status",label:"Status",render:e=>f(e)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:X},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],formFields:e=>[{type:"row",fields:[{name:"employee_id",label:"Karyawan",type:"select",required:!0,options:ve,value:e?.employee_id},{name:"branch_id",label:"Cabang",type:"select",options:X,value:e?.branch_id}]},{type:"row",fields:[{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:e?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif"],value:e?.status||"Aktif"}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",required:!0,value:e?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",required:!0,value:e?.end_date}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:e?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:e?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:e?.notes}]})}var Z=[],ee=[];async function fe(n){let[e,s]=await Promise.all([p("/api/branches?all=1"),p("/api/pic")]);Z=(e.data?.data||[]).map(t=>({value:t.id,label:t.full_name})),ee=(s.data?.data||[]).map(t=>t.name),g({container:n,title:"Jadwal Kegiatan",icon:"\u{1F5D3}\uFE0F",apiPath:"/api/schedule",itemLabel:"Jadwal",columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:t=>J(t)},{key:"period",label:"Periode",render:t=>E(t)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0},{key:"target_date",label:"Tgl Target",nowrap:!0},{key:"completion_date",label:"Tgl Selesai",nowrap:!0},{key:"status",label:"Status",render:t=>f(t)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:Z},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:ee}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:Z,value:t?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"],value:t?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period},{name:"pic",label:"PIC",type:"select",options:ee,value:t?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:t?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:t?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:t?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:t?.status||"Pending"}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}]})}var ae=[];async function ke(n){let[e]=await Promise.all([p("/api/branches?all=1")]);ae=(e.data?.data||[]).map(a=>({value:a.id,label:a.full_name}));let s=new Date().getFullYear(),t=Array.from({length:5},(a,l)=>String(s-l));g({container:n,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",itemLabel:"Permasalahan",columns:[{key:"report_date",label:"Tanggal",nowrap:!0},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:a=>`<span class="badge badge-secondary">${a}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:a=>`<span title="${a}">${a?.length>50?a.slice(0,50)+"\u2026":a}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:a=>`<span title="${a||""}">${a?.length>40?a.slice(0,40)+"\u2026":a||"-"}</span>`},{key:"status",label:"Status",render:a=>f(a)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0},{key:"day_count",label:"Hari",render:a=>a??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:ae},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:t}],formFields:a=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:a?.report_date},{name:"branch_id",label:"Cabang",type:"select",required:!0,options:ae,value:a?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:a?.category},{name:"source",label:"Sumber Laporan",type:"select",options:["SPV","AM","Berlin","Ade","Miswar","Pattrel","Perawat","FC","RCP","Lainnya"],value:a?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:a?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",placeholder:"Nama yang bermasalah",value:a?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"select",options:["Fajar","Miswar","Ade","Berlin","Pattrel","Lainnya"],value:a?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:a?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:a?.status||"Open"},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:a?.completion_date}]}]})}async function _e(n){let s=((await p("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name}));g({container:n,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",itemLabel:"One on One",columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:t=>`<span title="${t||""}">${t?.length>50?t.slice(0,50)+"\u2026":t||"-"}</span>`},{key:"solution",label:"Solusi",render:t=>`<span title="${t||""}">${t?.length>40?t.slice(0,40)+"\u2026":t||"-"}</span>`},{key:"status",label:"Status",render:t=>f(t)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:s},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],formFields:t=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:t?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:s,value:t?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",required:!0,placeholder:"Nama karyawan",value:t?.employee_name},{name:"pic",label:"PIC",type:"select",options:["Berlin","Ade","Miswar","Fajar","Pattrel","SPV","AM"],value:t?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:t?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:t?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:t?.status||"Open"},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:t?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:t?.document_link}]})}async function we(n){let s=((await p("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name})),t=Array.from({length:5},(a,l)=>String(new Date().getFullYear()-l));g({container:n,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a||"-"}catch{return a||"-"}}},{key:"score",label:"Nilai",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:t}],formFields:a=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:a?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:a?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:a?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:s,value:a?.branch_id},{name:"trainer",label:"Trainer",placeholder:"Nama trainer",value:a?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let l=JSON.parse(a?.participants);return Array.isArray(l)?l.join(", "):a?.participants||""}catch{return a?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:a?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:a?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}],onBeforeSubmit:async a=>(a.participants&&(a.participants=JSON.stringify(a.participants.split(",").map(l=>l.trim()).filter(Boolean))),a)})}async function $e(n){let s=((await p("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name}));g({container:n,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",itemLabel:"Reliefer",columns:[{key:"backup_date",label:"Tanggal Backup",nowrap:!0},{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"FC Digantikan"},{key:"period",label:"Periode",render:t=>E(t)},{key:"reliever_name",label:"Reliefer"},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:t=>t?`<span class="badge badge-info">${t}</span>`:"-"},{key:"status",label:"Status",render:t=>f(t)}],filterFields:[{type:"search",placeholder:"Cari reliefer / FC..."},{type:"select",name:"branch_id",label:"Cabang",options:s},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:s,value:t?.branch_id},{name:"period",label:"Periode",type:"select",options:["Q1","Q2","Q3","Q4"],value:t?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"FC yang Digantikan",placeholder:"Nama FC / BELUM ADA FC",value:t?.original_fc_name},{name:"reliever_name",label:"Nama Reliefer",required:!0,placeholder:"Nama yang menggantikan",value:t?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Backup",type:"date",required:!0,value:t?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:t?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"select",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:t?.reason},{name:"shift",label:"Shift",type:"select",options:["Pagi","Siang","Full Shift","Middle"],value:t?.shift}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done","Tidak Datang"],value:t?.status||"Pending"}]})}async function Ce(n){let s=((await p("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name})),t=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));g({container:n,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:a=>E(a)},{key:"inspection_date",label:"Tanggal",nowrap:!0},{key:"fc_score",label:"Point FC",render:a=>a!=null?`<strong class="${a>=80?"text-success":a>=60?"text-warning":"text-danger"}">${a}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:a=>a!=null?`<strong>${a}</strong>`:"-"},{key:"status",label:"Status",render:a=>f(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:s},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:t}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:s,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:a?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||"Pending"}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:a?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:a?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}]})}async function Se(n){let s=((await p("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name})),t=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));g({container:n,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge ${a==="Deep Cleaning"?"badge-purple":"badge-success"}">${a}</span>`},{key:"period",label:"Periode",render:a=>E(a)},{key:"activity_date",label:"Tanggal",nowrap:!0},{key:"status",label:"Status",render:a=>f(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:s},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:t}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:s,value:a?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:a?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:a?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||"Pending"},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}]})}async function Pe(n){let s=((await p("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name})),t=Array.from({length:4},(a,l)=>String(new Date().getFullYear()-l));g({container:n,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:a=>`<span class="badge badge-warning">${a}</span>`},{key:"period",label:"Periode",render:a=>E(a)},{key:"activity_date",label:"Tanggal",nowrap:!0},{key:"status",label:"Status",render:a=>f(a)},{key:"document_link",label:"Dokumen",render:a=>a?`<a href="${a}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:a=>a||"-"}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:s},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:t}],formFields:a=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:s,value:a?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:a?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:a?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:a?.status||"Pending"}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:a?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:a?.notes}]})}async function xe(n){let s=((await p("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name}));g({container:n,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:t=>`<span title="${t||""}">${t?.length>60?t.slice(0,60)+"\u2026":t||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0},{key:"status",label:"Status",render:t=>f(t)},{key:"notes",label:"Keterangan",render:t=>t?.length>40?t.slice(0,40)+"\u2026":t||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan / PIC..."},{type:"select",name:"branch_id",label:"Cabang",options:s},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:s,value:t?.branch_id},{name:"pic",label:"PIC",type:"select",options:["Berlin","Ade","Miswar","Fajar","Pattrel","Dentrel"],value:t?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:t?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:t?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:t?.done_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:t?.status||"Open"},{name:"notes",label:"Keterangan",type:"textarea",rows:2,value:t?.notes}]})}async function Te(n){g({container:n,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function Ee(n){g({container:n,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}async function te(n,e="forms"){if(e==="supply")return Je(n);je(n)}function je(n){g({container:n,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}]})}async function Je(n){let s=((await p("/api/branches?all=1")).data?.data||[]).map(a=>({value:a.id,label:a.full_name}));g({container:n,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",itemLabel:"Permintaan",canCreate:!1,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:a=>a?new Date(a).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(a,l)=>l.branch_name_ref||l.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a}catch{return a||"-"}}},{key:"chemical_items",label:"Chemical",render:a=>{try{let l=JSON.parse(a);return Array.isArray(l)?l.join(", "):a}catch{return a||"-"}}},{key:"additional_notes",label:"Catatan",render:a=>a?.length>40?a.slice(0,40)+"\u2026":a||"-"},{key:"status",label:"Status",render:a=>f(a)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(a,l)=>{let d=R({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(i,r)=>{let u=i.querySelector("#supply-status").value,c=i.querySelector("#supply-processed-by").value;(await p(`/api/reports/supply/${a.id}`,{method:"PUT",body:JSON.stringify({status:u,processed_by:c})})).ok?(B("Status diperbarui."),r(),l()):O("Gagal update status.")}})}}]});let t=n.querySelector(".page-header .page-actions");if(t){let a=document.createElement("a");a.href="#/form/chemical",a.className="btn btn-ghost btn-sm",a.textContent="\u{1F517} Form Publik",t.prepend(a)}}async function Le(n){let e=I();if(!e||!["superadmin","admin"].includes(e.role)){n.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}g({container:n,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:s=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[s]||"badge-neutral"}">${s}</span>`},{key:"is_active",label:"Status",render:s=>s?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:s=>s?new Date(s).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:s=>{let t=!!s;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:s?.full_name},{name:"username",label:"Username",required:!t,placeholder:"username",value:s?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!t,placeholder:"email@contoh.com",value:s?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:s?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:t?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!t,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:t?s?.is_active:1}]}]}})}async function Be(n){g({container:n,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}]})}async function De(n){let e=new Date,s=[];n.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">Kalender</h1>
    </div>
    <div class="card">
      <div class="card-header calendar-nav">
        <button class="btn btn-ghost btn-sm" id="cal-prev">\u2039 Prev</button>
        <span class="calendar-month-label" id="cal-month-label"></span>
        <button class="btn btn-ghost btn-sm" id="cal-next">Next \u203A</button>
        <div class="calendar-filters">
          <label class="filter-check"><input type="checkbox" value="schedule" checked class="cal-filter"> Jadwal</label>
          <label class="filter-check"><input type="checkbox" value="issue" checked class="cal-filter"> Permasalahan</label>
          <label class="filter-check"><input type="checkbox" value="reliever" checked class="cal-filter"> Reliefer</label>
          <label class="filter-check"><input type="checkbox" value="training" checked class="cal-filter"> Training</label>
          <label class="filter-check"><input type="checkbox" value="contract_expiry" checked class="cal-filter"> Kontrak Habis</label>
        </div>
      </div>
      <div class="card-body p-0">
        <div id="calendar-grid"></div>
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),a()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),a()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(l=>l.addEventListener("change",a));async function t(){let l=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;s=(await p(`/api/dashboard/calendar?month=${l}`)).data?.data||[]}async function a(){await t();let l=e.getFullYear(),d=e.getMonth(),i=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"});document.getElementById("cal-month-label").textContent=i;let r=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(o=>o.value)),u=s.filter(o=>r.has(o.type)),c={};u.forEach(o=>{let v=(o.event_date||"").slice(0,10);c[v]||(c[v]=[]),c[v].push(o)});let k=new Date(l,d,1).getDay(),m=new Date(l,d+1,0).getDate(),S=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],$='<div class="calendar-grid">';S.forEach(o=>{$+=`<div class="cal-day-header">${o}</div>`});for(let o=0;o<k;o++)$+='<div class="cal-cell cal-cell-empty"></div>';let _=new Date().toISOString().slice(0,10);for(let o=1;o<=m;o++){let v=`${l}-${String(d+1).padStart(2,"0")}-${String(o).padStart(2,"0")}`,x=c[v]||[],F=v===_;$+=`
        <div class="cal-cell ${F?"cal-today":""} ${x.length?"cal-has-events":""}" 
             data-date="${v}" tabindex="0" role="button" aria-label="${v}">
          <div class="cal-day-num ${F?"today-num":""}">${o}</div>
          <div class="cal-events-preview">
            ${x.slice(0,3).map(h=>`
              <div class="cal-event-dot cal-color-${h.color||"gray"}" title="${h.title||h.type}">
                <span class="cal-event-dot-label">${Ye(h.title||h.branch_name||h.type,18)}</span>
              </div>
            `).join("")}
            ${x.length>3?`<div class="cal-more">+${x.length-3} lagi</div>`:""}
          </div>
        </div>`}$+="</div>",document.getElementById("calendar-grid").innerHTML=$,document.querySelectorAll(".cal-cell[data-date]").forEach(o=>{o.addEventListener("click",()=>{let v=o.dataset.date,x=c[v]||[];if(!x.length)return;let F=document.getElementById("cal-event-list");document.getElementById("cal-event-date").textContent=new Date(v+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),document.getElementById("cal-event-items").innerHTML=x.map(h=>`
          <div class="cal-event-item cal-color-border-${h.color||"gray"}">
            <div class="cal-event-type">${Ge(h.type)}</div>
            <div class="cal-event-title">${h.title||"-"}</div>
            <div class="cal-event-branch">${h.branch_name||""}</div>
            ${h.status?`<div class="cal-event-status">${h.status}</div>`:""}
            ${h.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${h.days_remaining} hari</div>`:""}
          </div>
        `).join(""),F.style.display="block"})})}a()}function Ye(n,e){return n&&n.length>e?n.slice(0,e)+"\u2026":n||""}function Ge(n){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[n]||n}async function Fe(n){let e=I();n.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>
    <div class="profile-grid">
      <div class="card">
        <div class="card-header"><h3 class="card-title">Informasi Akun</h3></div>
        <div class="card-body">
          <div class="profile-avatar-lg">${(e?.full_name||"U")[0].toUpperCase()}</div>
          <div class="profile-info">
            <div class="profile-row"><span class="profile-label">Nama Lengkap</span><span>${e?.full_name||"-"}</span></div>
            <div class="profile-row"><span class="profile-label">Username</span><span>${e?.username||"-"}</span></div>
            <div class="profile-row"><span class="profile-label">Email</span><span>${e?.email||"-"}</span></div>
            <div class="profile-row"><span class="profile-label">Role</span><span class="badge badge-info">${e?.role||"-"}</span></div>
          </div>
        </div>
      </div>

      <div class="card">
        <div class="card-header"><h3 class="card-title">Ganti Password</h3></div>
        <div class="card-body">
          <form id="change-pwd-form" novalidate>
            <div class="form-group">
              <label class="form-label">Password Lama <span class="required">*</span></label>
              <input type="password" name="current_password" class="form-control" required placeholder="Password saat ini">
            </div>
            <div class="form-group">
              <label class="form-label">Password Baru <span class="required">*</span></label>
              <input type="password" name="new_password" class="form-control" required placeholder="Min. 6 karakter">
            </div>
            <div class="form-group">
              <label class="form-label">Konfirmasi Password Baru <span class="required">*</span></label>
              <input type="password" name="confirm_password" class="form-control" required placeholder="Ulangi password baru">
            </div>
            <div id="pwd-error" class="alert alert-danger" style="display:none"></div>
            <button type="submit" class="btn btn-primary">Simpan Password</button>
          </form>
        </div>
      </div>
    </div>
  `,document.getElementById("change-pwd-form")?.addEventListener("submit",async s=>{s.preventDefault();let t=document.getElementById("pwd-error");t.style.display="none";let a=s.target,l=a.current_password.value,d=a.new_password.value,i=a.confirm_password.value;if(d!==i){t.textContent="Konfirmasi password tidak cocok.",t.style.display="block";return}if(d.length<6){t.textContent="Password baru minimal 6 karakter.",t.style.display="block";return}let r=await p("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:l,new_password:d})});r.ok?(B("Password berhasil diubah."),a.reset()):(t.textContent=r.data?.error||"Gagal mengubah password.",t.style.display="block")})}function C(n){return async e=>{if(!N()){M("/login");return}return n(e)}}function Ie(){let n=I();document.getElementById("app").innerHTML=`
    <div class="app-layout">
      <!-- Sidebar -->
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-header">
          <div class="sidebar-logo">
            <span class="logo-icon">\u{1F3E5}</span>
            <span class="logo-text">FM<strong>Ops</strong></span>
          </div>
          <button class="sidebar-close" id="sidebar-close" aria-label="Close sidebar">&times;</button>
        </div>

        <nav class="sidebar-nav">
          <div class="nav-section">
            <span class="nav-section-label">Utama</span>
            <a href="#/dashboard" class="nav-item" data-route="/dashboard">
              <span class="nav-icon">\u{1F4CA}</span><span class="nav-label">Dashboard</span>
            </a>
            <a href="#/calendar" class="nav-item" data-route="/calendar">
              <span class="nav-icon">\u{1F4C5}</span><span class="nav-label">Kalender</span>
            </a>
          </div>

          <div class="nav-section">
            <span class="nav-section-label">SDM</span>
            <a href="#/employees" class="nav-item" data-route="/employees">
              <span class="nav-icon">\u{1F465}</span><span class="nav-label">Karyawan</span>
            </a>
            <a href="#/contracts" class="nav-item" data-route="/contracts">
              <span class="nav-icon">\u{1F4CB}</span><span class="nav-label">Kontrak</span>
            </a>
            <a href="#/relievers" class="nav-item" data-route="/relievers">
              <span class="nav-icon">\u{1F504}</span><span class="nav-label">Reliefer</span>
            </a>
          </div>

          <div class="nav-section">
            <span class="nav-section-label">Operasional</span>
            <a href="#/schedule" class="nav-item" data-route="/schedule">
              <span class="nav-icon">\u{1F5D3}\uFE0F</span><span class="nav-label">Jadwal Kegiatan</span>
            </a>
            <a href="#/issues" class="nav-item" data-route="/issues">
              <span class="nav-icon">\u26A0\uFE0F</span><span class="nav-label">Permasalahan</span>
            </a>
            <a href="#/one-on-one" class="nav-item" data-route="/one-on-one">
              <span class="nav-icon">\u{1F91D}</span><span class="nav-label">One on One</span>
            </a>
            <a href="#/training" class="nav-item" data-route="/training">
              <span class="nav-icon">\u{1F393}</span><span class="nav-label">Training</span>
            </a>
          </div>

          <div class="nav-section">
            <span class="nav-section-label">Laporan</span>
            <a href="#/reports/inspection" class="nav-item" data-route="/reports/inspection">
              <span class="nav-icon">\u{1F50D}</span><span class="nav-label">Laporan Inspeksi</span>
            </a>
            <a href="#/reports/cleaning" class="nav-item" data-route="/reports/cleaning">
              <span class="nav-icon">\u{1F9F9}</span><span class="nav-label">Laporan GC/DC</span>
            </a>
            <a href="#/reports/fogging" class="nav-item" data-route="/reports/fogging">
              <span class="nav-icon">\u{1F4A8}</span><span class="nav-label">Rekap Fogging</span>
            </a>
            <a href="#/reports/basecamp" class="nav-item" data-route="/reports/basecamp">
              <span class="nav-icon">\u{1F4DD}</span><span class="nav-label">Laporan Basecamp</span>
            </a>
            <a href="#/reports/supply" class="nav-item" data-route="/reports/supply">
              <span class="nav-icon">\u{1F4E6}</span><span class="nav-label">Permintaan Barang</span>
            </a>
          </div>

          <div class="nav-section">
            <span class="nav-section-label">Referensi</span>
            <a href="#/sop" class="nav-item" data-route="/sop">
              <span class="nav-icon">\u{1F4D6}</span><span class="nav-label">SOP</span>
            </a>
            <a href="#/checklist" class="nav-item" data-route="/checklist">
              <span class="nav-icon">\u2705</span><span class="nav-label">Master Checklist</span>
            </a>
            <a href="#/forms" class="nav-item" data-route="/forms">
              <span class="nav-icon">\u{1F4C4}</span><span class="nav-label">Master Form</span>
            </a>
          </div>

          ${n&&(n.role==="superadmin"||n.role==="admin")?`
          <div class="nav-section">
            <span class="nav-section-label">Admin</span>
            <a href="#/users" class="nav-item" data-route="/users">
              <span class="nav-icon">\u{1F510}</span><span class="nav-label">Manajemen User</span>
            </a>
            <a href="#/branches" class="nav-item" data-route="/branches">
              <span class="nav-icon">\u{1F3E2}</span><span class="nav-label">Cabang</span>
            </a>
          </div>`:""}
        </nav>

        <div class="sidebar-footer">
          <div class="user-info">
            <div class="user-avatar">${(n?.full_name||"U")[0].toUpperCase()}</div>
            <div class="user-details">
              <div class="user-name">${n?.full_name||"User"}</div>
              <div class="user-role">${n?.role||""}</div>
            </div>
          </div>
          <button class="btn btn-ghost btn-sm logout-btn" id="logout-btn">\u21A9 Keluar</button>
        </div>
      </aside>

      <!-- Overlay for mobile -->
      <div class="sidebar-overlay" id="sidebar-overlay"></div>

      <!-- Main content -->
      <div class="main-wrapper">
        <header class="topbar">
          <button class="topbar-menu-btn" id="topbar-menu-btn" aria-label="Toggle menu">\u2630</button>
          <div class="topbar-title" id="topbar-title">Dashboard</div>
          <div class="topbar-actions">
            <a href="#/profile" class="btn btn-ghost btn-sm topbar-profile">
              <span class="user-avatar-sm">${(n?.full_name||"U")[0].toUpperCase()}</span>
            </a>
          </div>
        </header>
        <main id="main-content" class="main-content"></main>
      </div>
    </div>
  `;let e=document.getElementById("sidebar"),s=document.getElementById("sidebar-overlay"),t=document.getElementById("topbar-menu-btn"),a=document.getElementById("sidebar-close"),l=()=>{e.classList.add("open"),s.classList.add("show")},d=()=>{e.classList.remove("open"),s.classList.remove("show")};t?.addEventListener("click",l),a?.addEventListener("click",d),s?.addEventListener("click",d);function i(){let r=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(k=>{let m=k.dataset.route;k.classList.toggle("active",r===m||m!=="/dashboard"&&r.startsWith(m))});let u=document.getElementById("topbar-title"),c=document.querySelector(".nav-item.active .nav-label");u&&c&&(u.textContent=c.textContent)}window.addEventListener("hashchange",i),i(),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await p("/api/auth/logout",{method:"POST"}),K(),M("/login")}),document.querySelectorAll(".nav-item").forEach(r=>{r.addEventListener("click",d)})}async function We(){w("/login",({main:e})=>de(e)),w("/dashboard",C(({main:e})=>ce(e))),w("/calendar",C(({main:e})=>De(e))),w("/employees",C(({main:e})=>ye(e))),w("/contracts",C(({main:e})=>he(e))),w("/schedule",C(({main:e})=>fe(e))),w("/issues",C(({main:e})=>ke(e))),w("/one-on-one",C(({main:e})=>_e(e))),w("/training",C(({main:e})=>we(e))),w("/relievers",C(({main:e})=>$e(e))),w("/reports/inspection",C(({main:e})=>Ce(e))),w("/reports/cleaning",C(({main:e})=>Se(e))),w("/reports/fogging",C(({main:e})=>Pe(e))),w("/reports/basecamp",C(({main:e})=>xe(e))),w("/reports/supply",C(({main:e})=>te(e,"supply"))),w("/sop",C(({main:e})=>Te(e))),w("/checklist",C(({main:e})=>Ee(e))),w("/forms",C(({main:e})=>te(e))),w("/users",C(({main:e})=>Le(e))),w("/branches",C(({main:e})=>Be(e))),w("/profile",C(({main:e})=>Fe(e))),w("/form/chemical",()=>{window.location.href="/form.html"});let n=N();if(!n&&window.location.hash!=="#/login"&&M("/login"),n){let e=await p("/api/auth/me");e.ok?(H(e.data.data),Ie()):(K(),M("/login"))}window.addEventListener("fm:login",()=>{Ie(),M("/dashboard")}),ie()}We();
