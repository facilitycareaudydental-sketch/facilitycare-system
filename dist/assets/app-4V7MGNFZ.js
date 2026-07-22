var Ha=Object.defineProperty;var Je=(a,e)=>()=>(a&&(e=a(a=0)),e);var Ge=(a,e)=>{for(var l in e)Ha(a,l,{get:e[l],enumerable:!0})};var me={};Ge(me,{API:()=>Qe,apiFetch:()=>f,clearToken:()=>de,getToken:()=>ce,getUser:()=>ie,setToken:()=>Ee,setUser:()=>pe});function ce(){return localStorage.getItem("fm_token")}function Ee(a){localStorage.setItem("fm_token",a)}function de(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ie(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function pe(a){localStorage.setItem("fm_user",JSON.stringify(a))}async function f(a,e={}){let l=ce(),n={"Content-Type":"application/json",...l?{Authorization:`Bearer ${l}`}:{},...e.headers||{}};try{let t=`cb=${Date.now()}`,o=a.includes("?")?"&":"?",u=`${Qe}${a}${o}${t}`,b=await fetch(u,{...e,headers:n}),s;try{let c=await b.text();try{s=JSON.parse(c)}catch{s={error:`Server Error (${b.status}): ${c.substring(0,80)}...`}}}catch{s={error:"Gagal membaca respon dari server"}}return b.status===401&&(de(),window.location.hash="#/login"),{ok:b.ok,status:b.status,data:s}}catch(t){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${t.message})`}}}}var Ua,Qe,P=Je(()=>{Ua="",Qe=Ua});var J={};Ge(J,{downloadExcel:()=>T,parseExcel:()=>Be,renderExcelButtons:()=>Ne});function Be(a){return new Promise((e,l)=>{let n=new FileReader;n.onload=t=>{try{let o=new Uint8Array(t.target.result),u=XLSX.read(o,{type:"array"}),b=u.SheetNames[0],s=u.Sheets[b],c=XLSX.utils.sheet_to_json(s,{defval:""});e(c)}catch(o){l(o)}},n.onerror=t=>l(t),n.readAsArrayBuffer(a)})}function T(a,e){try{let l=XLSX.utils.json_to_sheet(a),n=XLSX.utils.book_new();XLSX.utils.book_append_sheet(n,l,"Data"),XLSX.writeFile(n,`${e}.xlsx`)}catch(l){throw console.error("Error generating Excel file:",l),l}}function Ne(a){return`
    <div class="excel-actions" style="display:flex;gap:0.5rem;margin-bottom:1rem;">
      <button class="btn btn-secondary btn-sm" id="btn-export-${a}">
        \u{1F4E5} Export Excel
      </button>
      <button class="btn btn-secondary btn-sm" id="btn-template-${a}">
        \u{1F4C4} Download Template
      </button>
      <label class="btn btn-primary btn-sm" style="cursor:pointer;margin:0;">
        \u{1F4E4} Import Excel
        <input type="file" id="input-import-${a}" accept=".xlsx, .xls, .csv" style="display:none;">
      </label>
    </div>
  `}var L=Je(()=>{});P();var Pe={},ke=null;function B(a,e){Pe[a]=e}function se(a){window.location.hash=a}function Ve(){async function a(){let e=window.location.hash.replace("#","")||"/dashboard",[l,...n]=e.split("?"),t=Pe[l];if(!t){for(let[u,b]of Object.entries(Pe))if(u.endsWith("/*")&&l.startsWith(u.slice(0,-2))){t=b;break}}ke&&(ke(),ke=null);let o=document.getElementById("main-content");if(o&&(o.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),t){let u=new URLSearchParams(n.join("?")),b=l.split("/").filter(Boolean),s=await t({path:l,params:u,segments:b,main:o});s&&(ke=s)}else{let u=o||document.getElementById("app");u&&(u.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",a),a()}var ue;function Ja(){return ue||(ue=document.createElement("div"),ue.id="toast-container",document.body.appendChild(ue)),ue}function ze(a,e="info",l=3500){let n=Ja(),t=document.createElement("div");t.className=`toast toast-${e}`;let o={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};t.innerHTML=`<span class="toast-icon">${o[e]||"\u2139"}</span><span class="toast-msg">${a}</span>`,n.appendChild(t),requestAnimationFrame(()=>t.classList.add("show")),setTimeout(()=>{t.classList.remove("show"),setTimeout(()=>t.remove(),350)},l)}var K=a=>ze(a,"success"),j=a=>ze(a,"error");function ee({title:a,content:e,onConfirm:l,onCancel:n,confirmText:t="Simpan",cancelText:o="Batal",size:u="md",confirmClass:b="btn-primary"}){let s={sm:"400px",md:"560px",lg:"720px",xl:"900px"},c=document.createElement("div");c.className="modal-overlay",c.innerHTML=`
    <div class="modal" style="max-width:${s[u]||s.md}">
      <div class="modal-header">
        <h3 class="modal-title">${a}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${o}</button>
        ${l?`<button class="btn ${b} modal-confirm">${t}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&c.querySelector(".modal-body").appendChild(e);let i=()=>{c.classList.remove("show"),setTimeout(()=>c.remove(),250)};return c.querySelector(".modal-close").addEventListener("click",()=>{n&&n(),i()}),c.querySelector(".modal-cancel").addEventListener("click",()=>{n&&n(),i()}),l&&c.querySelector(".modal-confirm").addEventListener("click",()=>l(c,i)),c.addEventListener("click",r=>{r.target===c&&(n&&n(),i())}),document.body.appendChild(c),requestAnimationFrame(()=>c.classList.add("show")),{overlay:c,close:i}}function We(a,e,l="Konfirmasi"){return ee({title:l,content:`<p>${a}</p>`,onConfirm:(n,t)=>{e(),t()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}P();var ne={};function ge(a){if(ne[a]){try{ne[a].destroy()}catch{}delete ne[a]}}function Ga(){Object.keys(ne).forEach(ge)}var H=(a,e=0)=>{let l=Number(a);return isNaN(l)||a===null||a===void 0?e:l},U=(a,e="\u2014")=>{if(a==null||a==="")return e;let l=String(a).trim();return l===""||l==="[object Object]"?e:l},Xe=a=>{if(!a)return"\u2014";try{let e=new Date(a);return isNaN(e)?U(a):e.toLocaleDateString("id-ID",{day:"2-digit",month:"short",year:"numeric"})}catch{return U(a)}},Qa=a=>{if(!a)return"";try{let e=Date.now()-new Date(a).getTime();if(e<0)return"Baru saja";let l=Math.floor(e/6e4);if(l<1)return"Baru saja";if(l<60)return`${l} menit lalu`;let n=Math.floor(l/60);return n<24?`${n} jam lalu`:`${Math.floor(n/24)} hari lalu`}catch{return""}},Ze=a=>{if(!a||typeof a!="string")return"";try{let[e,l]=a.split("-");return new Date(Number(e),Number(l)-1).toLocaleDateString("id-ID",{month:"short",year:"2-digit"})}catch{return a}};function ea(a,e,l=900){if(!a)return;let n=Math.max(0,Math.round(H(e)));if(n===0){a.textContent="0";return}let t=Date.now(),o=()=>{let u=Math.min((Date.now()-t)/l,1),b=1-Math.pow(1-u,3);a.textContent=Math.round(b*n).toLocaleString("id-ID"),u<1?requestAnimationFrame(o):a.textContent=n.toLocaleString("id-ID")};requestAnimationFrame(o)}function Va(a,e){if(a=H(a),e=H(e),e===0)return"";let l=a-e,n=Math.abs(Math.round(l/e*100));return l>0?`<span class="kpi-trend up">\u25B2 ${n}%</span>`:l<0?`<span class="kpi-trend down">\u25BC ${n}%</span>`:'<span class="kpi-trend neutral">= Sama</span>'}var za={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},aa=a=>{let e=U(a,"\u2014");return`<span class="status-pill ${za[e]||"pill-neutral"}">${e}</span>`},Wa=a=>{let e=H(a,999);return e<=7?`<span class="days-badge days-critical">${e} hari</span>`:e<=14?`<span class="days-badge days-warning">${e} hari</span>`:e<=30?`<span class="days-badge days-soon">${e} hari</span>`:`<span class="days-badge days-ok">${e} hari</span>`},Ya={issue:{emoji:"\u26A0\uFE0F",dot:"dot-danger",label:"Permasalahan"},contract:{emoji:"\u{1F4C4}",dot:"dot-info",label:"Kontrak"},employee:{emoji:"\u{1F464}",dot:"dot-success",label:"Karyawan"},one_on_one:{emoji:"\u{1F91D}",dot:"dot-purple",label:"One on One"},training:{emoji:"\u{1F393}",dot:"dot-primary",label:"Training"},supply:{emoji:"\u{1F4E6}",dot:"dot-warning",label:"Permintaan Barang"},reliever:{emoji:"\u{1F504}",dot:"dot-teal",label:"Reliefer"},inspection:{emoji:"\u{1F50D}",dot:"dot-blue",label:"Laporan Inspeksi"}},Xa=a=>Ya[a]||{emoji:"\u{1F4CC}",dot:"dot-neutral",label:U(a,"Aktivitas")},Q={family:"Inter",size:11},re="#94A3B8",Se="#F1F5F9",Za=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],et=()=>window.innerWidth<768;function De(a={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:et()?"bottom":"top",labels:{font:Q,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:Q,titleFont:{...Q,weight:"700"}}},scales:{x:{grid:{color:Se},ticks:{font:Q,color:re,maxRotation:0}},y:{grid:{color:Se},ticks:{font:Q,color:re},beginAtZero:!0}},...a}}var at=()=>Array(5).fill(0).map(()=>`
  <div class="kpi-card" style="pointer-events:none">
    <div class="kpi-card-top"><div class="skeleton" style="width:44px;height:44px;border-radius:12px"></div></div>
    <div class="skeleton skeleton-text" style="width:55%;height:32px;margin:10px 0 6px"></div>
    <div class="skeleton skeleton-text" style="width:75%;height:12px;margin-bottom:4px"></div>
    <div class="skeleton skeleton-text" style="width:55%;height:11px"></div>
  </div>`).join(""),tt=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Ye(a=3){return Array(a).fill(0).map((e,l)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${l<a-1?"6px":"0"};border-radius:6px"></div>`).join("")}function nt(){return Array(5).fill(0).map(()=>`
    <div class="activity-item">
      <div class="skeleton" style="width:34px;height:34px;border-radius:10px;flex-shrink:0"></div>
      <div style="flex:1">
        <div class="skeleton skeleton-text" style="width:65%;height:13px;margin-bottom:5px"></div>
        <div class="skeleton skeleton-text" style="width:35%;height:11px"></div>
      </div>
    </div>`).join("")}async function ae(a,e,l=8e3){try{let n=new AbortController,t=setTimeout(()=>n.abort(),l),o=await f(a,{signal:n.signal}).catch(()=>null);if(clearTimeout(t),!o||!o.ok)return e;let u=o.data;return u?u.data!==void 0?u.data??e:u:e}catch{return e}}function it(){["skel-donut","skel-trend","skel-insp","skel-contract"].forEach(n=>{let t=document.getElementById(n);t&&(t.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract"].forEach(n=>{let t=document.getElementById(n);if(t&&t.style.display==="none"){t.style.display="block";let o=t.parentElement;if(o&&!o.querySelector(".chart-empty")){let u=document.createElement("div");u.className="chart-empty",u.textContent="Belum ada data",t.style.display="none",o.appendChild(u)}}});let a=document.getElementById("kpi-row");a&&a.querySelector(".skeleton")&&na({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&ia({}),["table-contracts","table-issues"].forEach(n=>{let t=document.getElementById(n);t&&t.querySelector(".skeleton")&&(t.innerHTML='<div class="chart-empty">Belum ada data</div>')});let l=document.getElementById("activity-log");l&&l.querySelector(".skeleton")&&(l.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function ta(a){Ga(),a._dashRefresh&&clearInterval(a._dashRefresh),a._skelTimeout&&clearTimeout(a._skelTimeout),a.innerHTML=`
    <div class="dashboard-wrap" id="dash-root">

      <div class="section-header">
        <h2 class="section-title">\u{1F4CA} Dashboard Operasional FCMS</h2>
        <div class="dash-refresh-info">
          <span id="dash-updated" class="dash-last-updated"></span>
          <button class="btn btn-ghost btn-sm" id="btn-dash-refresh" title="Refresh">
            <svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24">
              <polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/>
              <path d="M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15"/>
            </svg>
            Refresh
          </button>
        </div>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${at()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${tt()}</div>

      <!-- Charts Row 1 -->
      <div class="charts-row">
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">\u{1F369} Permasalahan per Kategori</div>
              <div class="chart-card-subtitle">Distribusi semua permasalahan</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:260px;position:relative">
            <div id="skel-donut" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-donut" style="display:none"></canvas>
          </div>
        </div>
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">\u{1F4C8} Trend Permasalahan 12 Bulan</div>
              <div class="chart-card-subtitle">Open vs Closed per bulan</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:260px;position:relative">
            <div id="skel-trend" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-trend" style="display:none"></canvas>
          </div>
        </div>
      </div>

      <!-- Charts Row 2 -->
      <div class="charts-row">
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">\u{1F50D} Rata-rata Skor Inspeksi per Cabang</div>
              <div class="chart-card-subtitle">6 bulan terakhir \u2014 FC vs SPV</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:280px;position:relative">
            <div id="skel-insp" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-insp" style="display:none"></canvas>
          </div>
        </div>
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">\u{1F4CB} Kontrak Berakhir 6 Bulan ke Depan</div>
              <div class="chart-card-subtitle">Jumlah kontrak aktif per bulan</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:280px;position:relative">
            <div id="skel-contract" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-contract" style="display:none"></canvas>
          </div>
        </div>
      </div>

      <!-- Tables -->
      <div class="tables-row">
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">\u23F0 Kontrak Akan Habis</div>
              <div class="chart-card-subtitle">30 hari ke depan</div>
            </div>
            <a href="#/contracts" class="btn btn-ghost btn-sm">Lihat Semua \u2192</a>
          </div>
          <div id="table-contracts" class="dash-table-wrap">${Ye(3)}</div>
        </div>
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">\u26A0\uFE0F Permasalahan Terbaru</div>
              <div class="chart-card-subtitle">Open dan In Progress</div>
            </div>
            <a href="#/issues" class="btn btn-ghost btn-sm">Lihat Semua \u2192</a>
          </div>
          <div id="table-issues" class="dash-table-wrap">${Ye(3)}</div>
        </div>
      </div>

      <!-- Activity Log -->
      <div class="chart-card">
        <div class="chart-card-header">
          <div>
            <div class="chart-card-title">\u{1F550} Aktivitas Terbaru</div>
            <div class="chart-card-subtitle">Update real-time dari semua modul</div>
          </div>
        </div>
        <div id="activity-log">${nt()}</div>
      </div>

    </div>
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>Ie(a)),a._skelTimeout=setTimeout(()=>it(),5e3),await Ie(a),a._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?Ie(a):clearInterval(a._dashRefresh)},6e4)}async function Ie(a){a._skelTimeout&&(clearTimeout(a._skelTimeout),a._skelTimeout=null);let[e,l,n,t,o,u,b,s]=await Promise.all([ae("/api/dashboard/kpi",{},8e3),ae("/api/dashboard/issues-trend",{},8e3),ae("/api/dashboard/contracts-chart",{},8e3),ae("/api/dashboard/issues-summary",{},8e3),ae("/api/dashboard/inspection-bar",{},8e3),ae("/api/dashboard/contracts-expiring",[],8e3),ae("/api/dashboard/stats",{},8e3),ae("/api/dashboard/activity-log",[],8e3)]);try{na(e)}catch(i){console.warn("KPI render:",i)}try{ia(e)}catch(i){console.warn("MiniStats render:",i)}try{st(Array.isArray(t?.by_category)?t.by_category:[])}catch(i){console.warn("Donut render:",i),te("skel-donut","chart-donut")}try{rt(l)}catch(i){console.warn("Trend render:",i),te("skel-trend","chart-trend")}try{lt(o)}catch(i){console.warn("InspBar render:",i),te("skel-insp","chart-insp")}try{ot(n)}catch(i){console.warn("ContractBar render:",i),te("skel-contract","chart-contract")}try{let i=Array.isArray(u)?u:[];ct(i)}catch(i){console.warn("ContractsTable render:",i)}try{let i=Array.isArray(b)?b:Array.isArray(b?.recent_issues)?b.recent_issues:[];dt(i)}catch(i){console.warn("IssuesTable render:",i)}try{pt(Array.isArray(s)?s:[])}catch(i){console.warn("ActivityLog render:",i)}let c=document.getElementById("dash-updated");c&&(c.textContent=`Diperbarui: ${new Date().toLocaleTimeString("id-ID")}`)}function na(a){let e=document.getElementById("kpi-row");if(!e)return;a=a||{};let l=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees",color:"kpi-blue",key:"employees"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts",color:"kpi-green",key:"contracts"},{icon:"\u23F0",label:"Kontrak Habis 30 Hari",sub:"",href:"#/contracts",color:"",key:"expiring30",warn:!0},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues",color:"",key:"issues",warnIfGT0:!0},{icon:"\u{1F91D}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one",color:"kpi-purple",key:"one_on_one"}];e.innerHTML=l.map(n=>{let t=H(a[n.key]?.current,0),o=a[n.key]?.prev,u=o!=null?Va(t,o):"",b=n.color||"",s=n.sub||"";return n.warn&&(b=t>0?"kpi-amber":"kpi-green",s=t>0?`\u26A0\uFE0F ${t} kontrak segera berakhir`:"\u2705 Semua kontrak aman"),n.warnIfGT0&&(b=t>0?"kpi-red":"kpi-green"),`
      <a href="${n.href}" class="kpi-card ${b}" style="text-decoration:none">
        <div class="kpi-card-top">
          <div class="kpi-icon-wrap"><span class="kpi-icon-emoji">${n.icon}</span></div>
          ${u}
        </div>
        <div class="kpi-value" data-target="${t}">0</div>
        <div class="kpi-label">${n.label}</div>
        <div class="kpi-subtitle">${s}</div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(n=>ea(n,parseInt(n.dataset.target)||0))}function ia(a){let e=document.getElementById("mini-stats-row");if(!e)return;a=a||{};let l=[{icon:"\u{1F4C5}",label:"Jadwal Pending",val:a.schedule?.current,href:"#/schedule",color:"mini-blue"},{icon:"\u{1F393}",label:"Training Bulan Ini",val:a.training_month?.current,href:"#/training",color:"mini-indigo"},{icon:"\u{1F4E6}",label:"Permintaan Barang",val:a.supply?.current,href:"#/reports/supply",color:"mini-orange"},{icon:"\u{1F50D}",label:"Inspeksi Bulan Ini",val:a.inspection_month?.current,href:"#/reports/inspection",color:"mini-teal"},{icon:"\u{1F9F9}",label:"GC/DC Bulan Ini",val:a.cleaning_month?.current,href:"#/reports/cleaning",color:"mini-green"},{icon:"\u{1F99F}",label:"Fogging Bulan Ini",val:a.fogging_month?.current,href:"#/reports/fogging",color:"mini-purple"},{icon:"\u{1F3E2}",label:"Total Cabang",val:a.branches?.current,href:"#/branches",color:"mini-gray"}];e.innerHTML=l.map(n=>`
    <a href="${n.href}" class="mini-stat ${n.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${n.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${H(n.val)}">0</div>
        <div class="mini-stat-label">${n.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(n=>ea(n,parseInt(n.dataset.target)||0,700))}function st(a){te("skel-donut","chart-donut");let e=document.getElementById("chart-donut");if(!e)return;ge("donut");let l=(a||[]).filter(o=>H(o.count)>0);if(!l.length){_e(e,"Belum ada data permasalahan");return}let n=l.map(o=>U(o.category,"Lainnya")),t=l.map(o=>H(o.count));ne.donut=new Chart(e,{type:"doughnut",data:{labels:n,datasets:[{data:t,backgroundColor:Za.slice(0,l.length),borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{position:window.innerWidth<768?"bottom":"right",labels:{font:Q,color:"#475569",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{bodyFont:Q,titleFont:{...Q,weight:"700"},callbacks:{label:o=>` ${o.label}: ${o.parsed} kasus`}}},cutout:"65%"}})}function rt(a){te("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;ge("trend"),a=a||{};let l=(a.labels||[]).map(Ze),n=(a.open||[]).map(o=>H(o)),t=(a.closed||[]).map(o=>H(o));if(!l.length){_e(e,"Belum ada data trend");return}ne.trend=new Chart(e,{type:"line",data:{labels:l,datasets:[{label:"Open",data:n,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:t,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:De({plugins:{legend:{position:"top"}}})})}function lt(a){te("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;ge("inspBar"),a=a||{};let l=a.labels||[],n=(a.fc||[]).map(o=>H(o)),t=(a.spv||[]).map(o=>H(o));if(!l.length){_e(e,"Belum ada data inspeksi");return}ne.inspBar=new Chart(e,{type:"bar",data:{labels:l,datasets:[{label:"Skor FC",data:n,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:t,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:De({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:Q,color:re,maxRotation:45,minRotation:30}},y:{grid:{color:Se},ticks:{font:Q,color:re},min:0,max:100}}})})}function ot(a){te("skel-contract","chart-contract");let e=document.getElementById("chart-contract");if(!e)return;ge("contractBar"),a=a||{};let l=(a.labels||[]).map(Ze),n=(a.counts||[]).map(o=>H(o));if(!l.length){_e(e,"Belum ada data kontrak");return}let t=n.map(o=>o>5?"rgba(239,68,68,.75)":o>2?"rgba(245,158,11,.75)":"rgba(37,99,235,.65)");ne.contractBar=new Chart(e,{type:"bar",data:{labels:l,datasets:[{label:"Kontrak Berakhir",data:n,backgroundColor:t,borderRadius:6,borderSkipped:!1}]},options:De({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:Q,color:re}},y:{grid:{color:Se},ticks:{font:Q,color:re,precision:0},beginAtZero:!0}}})})}function ct(a){let e=document.getElementById("table-contracts");if(!e)return;let l=(a||[]).filter(n=>H(n.days_remaining,999)<=30).slice(0,10);if(!l.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada kontrak yang habis dalam 30 hari</div>';return}e.innerHTML=`
    <table class="dash-table">
      <thead><tr>
        <th>#</th><th>Nama Karyawan</th><th>Cabang</th><th>Berakhir</th><th>Sisa</th><th>Status</th>
      </tr></thead>
      <tbody>${l.map((n,t)=>`
        <tr>
          <td class="td-num">${t+1}</td>
          <td><strong>${U(n.emp_name||n.employee_name)}</strong></td>
          <td class="td-branch">${U(n.branch_name)}</td>
          <td style="white-space:nowrap;font-size:.8rem">${Xe(n.end_date)}</td>
          <td>${Wa(n.days_remaining)}</td>
          <td>${aa(n.status)}</td>
        </tr>`).join("")}
      </tbody>
    </table>`}function dt(a){let e=document.getElementById("table-issues");if(!e)return;let l=(a||[]).slice(0,8);if(!l.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <table class="dash-table">
      <thead><tr>
        <th>Tanggal</th><th>Keluhan</th><th>Cabang</th><th>Kategori</th><th>Status</th>
      </tr></thead>
      <tbody>${l.map(n=>`
        <tr>
          <td style="white-space:nowrap;font-size:.78rem">${Xe(n.report_date)}</td>
          <td class="td-complaint" title="${U(n.complaint)}">${U(n.complaint)}</td>
          <td class="td-branch">${U(n.branch_name)}</td>
          <td><span class="category-tag">${U(n.category)}</span></td>
          <td>${aa(n.status)}</td>
        </tr>`).join("")}
      </tbody>
    </table>`}function pt(a){let e=document.getElementById("activity-log");if(!e)return;let l=(a||[]).slice(0,15);if(!l.length){e.innerHTML='<div class="chart-empty">Belum ada aktivitas tercatat</div>';return}e.innerHTML=`<div class="activity-list">${l.map(n=>{let t=Xa(n.type),o=U(n.label),u=n.branch?` \u2022 ${U(n.branch)}`:"",b=Qa(n.created_at);return`
      <div class="activity-item">
        <div class="activity-dot ${t.dot}">${t.emoji}</div>
        <div class="activity-body">
          <div class="activity-text"><strong>${t.label}</strong> \u2014 ${o}${u}</div>
          <div class="activity-time">${b||"\u2014"}</div>
        </div>
      </div>`}).join("")}</div>`}function te(a,e){let l=document.getElementById(a),n=document.getElementById(e);l&&(l.style.display="none",l.style.position=""),n&&(n.style.display="block")}function _e(a,e="Belum ada data"){if(!a)return;a.style.display="none";let l=a.parentElement;if(!l)return;if(!l.querySelector(".chart-empty")){let t=document.createElement("div");t.className="chart-empty",t.textContent=e,l.appendChild(t)}}P();async function sa(a){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),l=document.getElementById("login-error"),n=document.getElementById("login-btn"),t=document.getElementById("toggle-password"),o=document.getElementById("login-password");t?.addEventListener("click",()=>{let u=o.type==="text";o.type=u?"password":"text",t.style.color=u?"":"var(--primary)"}),e?.addEventListener("submit",async u=>{u.preventDefault(),l.style.display="none";let b=e.username.value.trim(),s=e.password.value;if(!b||!s){l.textContent="Username dan password wajib diisi.",l.style.display="block";return}n.querySelector(".btn-text").style.display="none",n.querySelector(".btn-spinner").style.display="",n.disabled=!0;try{let c=await f("/api/auth/login",{method:"POST",body:JSON.stringify({username:b,password:s})});c.ok&&c.data.success?(Ee(c.data.data.token),pe(c.data.data.user),K("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(l.textContent=c.data.error||"Username atau password salah.",l.style.display="block",n.classList.add("shake"),setTimeout(()=>n.classList.remove("shake"),600))}catch{l.textContent="Gagal terhubung ke server. Periksa koneksi internet.",l.style.display="block"}finally{n.querySelector(".btn-text").style.display="",n.querySelector(".btn-spinner").style.display="none",n.disabled=!1}})}P();function ra({columns:a,data:e,onEdit:l,onDelete:n,onView:t,actions:o=[],emptyText:u="Tidak ada data",bulkSelect:b=null}){let s=document.createElement("div");if(s.className="table-wrapper",!e||e.length===0)return s.innerHTML=`<div class="empty-state"><p>${u}</p></div>`,s;let c=document.createElement("table");c.className="data-table";let i=document.createElement("thead"),r=document.createElement("tr");if(b){let g=document.createElement("th");g.style.width="40px",g.style.textAlign="center";let m=document.createElement("input");m.type="checkbox",m.id="select-all-checkbox",m.title="Pilih semua",m.addEventListener("change",()=>{e.forEach(d=>{m.checked?b.selectedIds.add(d.id):b.selectedIds.delete(d.id)}),s.querySelectorAll(".row-checkbox").forEach(d=>d.checked=m.checked),b.onToggle()}),g.appendChild(m),r.appendChild(g)}if(a.forEach(g=>{let m=document.createElement("th");m.textContent=g.label,g.width&&(m.style.width=g.width),r.appendChild(m)}),l||n||t||o.length>0){let g=document.createElement("th");g.textContent="Aksi",g.style.width="120px",r.appendChild(g)}i.appendChild(r),c.appendChild(i);let y=document.createElement("tbody");return e.forEach(g=>{let m=document.createElement("tr");if(b){let d=document.createElement("td");d.style.textAlign="center",d.style.width="40px";let p=document.createElement("input");p.type="checkbox",p.className="row-checkbox",p.checked=b.selectedIds.has(g.id),p.addEventListener("change",()=>{if(p.checked)b.selectedIds.add(g.id);else{b.selectedIds.delete(g.id);let h=document.getElementById("select-all-checkbox");h&&(h.checked=!1)}b.onToggle()}),d.appendChild(p),m.appendChild(d)}if(a.forEach(d=>{let p=document.createElement("td");if(d.render){let h=d.render(g[d.key],g);h instanceof HTMLElement?p.appendChild(h):p.innerHTML=h||""}else p.textContent=g[d.key]!==null&&g[d.key]!==void 0&&g[d.key]!==""?g[d.key]:"";d.nowrap&&(p.style.whiteSpace="nowrap"),m.appendChild(p)}),l||n||t||o.length>0){let d=document.createElement("td");d.className="actions-cell";let p=document.createElement("div");if(p.className="btn-group",t){let h=document.createElement("button");h.className="btn btn-xs btn-ghost",h.innerHTML="\u{1F441}",h.title="Lihat",h.addEventListener("click",()=>t(g)),p.appendChild(h)}if(l){let h=document.createElement("button");h.className="btn btn-xs btn-secondary",h.innerHTML="\u270F\uFE0F",h.title="Edit",h.addEventListener("click",()=>l(g)),p.appendChild(h)}o.forEach(h=>{let v=document.createElement("button");v.className=`btn btn-xs ${h.class||"btn-ghost"}`,v.innerHTML=h.icon||h.label,v.title=h.label,v.addEventListener("click",()=>h.handler(g)),p.appendChild(v)}),d.appendChild(p),m.appendChild(d)}y.appendChild(m)}),c.appendChild(y),s.appendChild(c),s}function la({page:a,pages:e,total:l,limit:n,onPage:t}){if(e<=1)return null;let o=document.createElement("div");o.className="pagination";let u=document.createElement("span");u.className="pagination-info",u.textContent=`Total: ${l} data`,o.appendChild(u);let b=document.createElement("div");b.className="pagination-btns";let s=(r,y,g=!1,m=!1)=>{let d=document.createElement("button");d.className=`btn btn-sm ${m?"btn-primary":"btn-ghost"} pagination-btn`,d.textContent=r,d.disabled=g,d.addEventListener("click",()=>t(y)),b.appendChild(d)};s("\xAB",1,a===1),s("\u2039",a-1,a===1);let c=Math.max(1,a-2),i=Math.min(e,a+2);for(let r=c;r<=i;r++)s(r,r,!1,r===a);return s("\u203A",a+1,a===e),s("\xBB",e,a===e),o.appendChild(b),o}function Le(a){return a.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="row")return`<div class="form-row">${Le(e.fields)}</div>`;let l=e.required?"required":"",n=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",t="";switch(e.type){case"textarea":t=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${l} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let u=(e.options||[]).map(i=>{let r=typeof i=="object"?i.value:i,y=typeof i=="object"?i.label:i,g=e.value==r?"selected":"";return`<option value="${r}" ${g}>${y}</option>`}).join("");t=`<select name="${e.name}" class="form-control" ${l}><option value="">-- Pilih ${e.label||""} --</option>${u}</select>`;break;case"combobox":let b=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,s=(e.options||[]).map(i=>{let r=typeof i=="object"?i.value:i;return`<option value="${typeof i=="object"?i.label:i}"></option>`}).join(""),c=e.value||"";if(e.value){let i=(e.options||[]).find(r=>(typeof r=="object"?r.value:r)==e.value);i&&(c=typeof i=="object"?i.label:i)}t=`
          <input type="text" name="${e.name}" list="${b}" class="form-control" value="${c}" placeholder="Pilih atau ketik baru..." ${l} autocomplete="off">
          <datalist id="${b}">${s}</datalist>
        `;break;case"checkbox":t=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":t=`<input type="date" name="${e.name}" class="form-control" value="${e.value||""}" ${l}>`;break;case"number":t=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${l}>`;break;case"email":t=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${l}>`;break;case"url":t=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${l}>`;break;default:t=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${l} autocomplete="off">`}let o=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${n}${t}${o}</div>`}).join("")}function oa(a){let e={},l=new FormData(a);for(let[n,t]of l.entries())e[n]=t===""?null:t;return a.querySelectorAll("input[type=checkbox]").forEach(n=>{n.checked||(e[n.name]=null)}),e}function ca(a,e){e&&Object.entries(e).forEach(([l,n])=>{let t=a.querySelector(`[name="${l}"]`);t&&(t.type==="checkbox"?t.checked=!!n:t.value=n??"")})}L();function $({container:a,title:e,icon:l,apiPath:n,columns:t,formFields:o,filterFields:u,defaultFilters:b={},itemLabel:s="Data",canCreate:c=!0,canEdit:i=!0,canDelete:r=!0,onBeforeSubmit:y,onAfterLoad:g,extraActions:m=[],initialSearch:d="",exportOptions:p=null,bulkDelete:h=!1}){let v=1,S={...b};d&&(S.search=d);let k=new Set;a.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${l} ${e}</h1>
      <div class="page-actions">
        ${c?`<button class="btn btn-primary" id="btn-create">+ Tambah ${s}</button>`:""}
      </div>
    </div>

    ${h?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${p?Ne(p.moduleName):""}

    ${u&&u.length>0?`
    <div class="filter-bar card">
      <div class="filter-bar-inner">
        ${u.map(C=>C.type==="search"?`<div class="filter-search"><input type="search" class="form-control" placeholder="${C.placeholder||"Cari..."}" id="filter-search" value="${S.search||""}"></div>`:C.type==="select"?`<select class="form-control filter-select" name="${C.name}" id="filter-${C.name}"><option value="">-- ${C.label} --</option>${(C.options||[]).map(x=>`<option value="${typeof x=="object"?x.value:x}" ${S[C.name]===(typeof x=="object"?x.value:x)?"selected":""}>${typeof x=="object"?x.label:x}</option>`).join("")}</select>`:"").join("")}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
      </div>
    </div>`:""}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;function _(){if(!document.getElementById("bulk-toolbar"))return;let x=document.getElementById("bulk-count"),D=document.getElementById("btn-bulk-delete"),F=document.getElementById("btn-bulk-cancel");x.textContent=`${k.size} item dipilih`,k.size>0?(D.disabled=!1,F.disabled=!1):(D.disabled=!0,F.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{k.clear(),document.querySelectorAll(".row-checkbox").forEach(x=>x.checked=!1);let C=document.getElementById("select-all-checkbox");C&&(C.checked=!1),_()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(k.size===0)return;let C=[...k],x=document.createElement("div");x.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",x.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${C.length} ${s}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${C.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(x),x.querySelector("#bulk-cancel-btn").addEventListener("click",()=>x.remove()),x.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let D=x.querySelector("#bulk-confirm-btn");D.disabled=!0,D.textContent="Menghapus...";let F=await f(`${n}/bulk`,{method:"DELETE",body:JSON.stringify({ids:C})});x.remove(),F.ok?(K(`${C.length} ${s} berhasil dihapus.`),k.clear(),_(),w()):j(F.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),A;if(E?.addEventListener("input",C=>{clearTimeout(A),A=setTimeout(()=>{S.search=C.target.value,v=1,w()},400)}),u?.forEach(C=>{C.type==="select"&&document.getElementById(`filter-${C.name}`)?.addEventListener("change",x=>{S[C.name]=x.target.value,v=1,w()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{S={...b},E&&(E.value=""),u?.forEach(C=>{let x=document.getElementById(`filter-${C.name}`);x&&(x.value="")}),v=1,w()}),document.getElementById("btn-create")?.addEventListener("click",()=>V(null)),p){document.getElementById(`btn-export-${p.moduleName}`)?.addEventListener("click",async x=>{let D=x.target,F=D.innerHTML;D.innerHTML="\u23F3 Loading...",D.disabled=!0;try{await p.onExport()}catch{j("Gagal export data")}finally{D.innerHTML=F,D.disabled=!1}}),document.getElementById(`btn-template-${p.moduleName}`)?.addEventListener("click",()=>{p.onTemplate()});let C=document.getElementById(`input-import-${p.moduleName}`);C?.addEventListener("change",async x=>{let D=x.target.files[0];if(!D)return;let F=C.parentElement,G=F.innerHTML;F.innerHTML="\u23F3 Memproses...",F.style.pointerEvents="none",C.disabled=!0;try{let X=await Be(D);if(X.length===0)throw new Error("File kosong atau format salah");await p.onImport(X),K("Import berhasil!"),w()}catch(X){j(X.message||"Gagal import data")}finally{F.innerHTML=G,F.style.pointerEvents="auto",C.disabled=!1,C.value=""}})}async function w(){let C=document.getElementById("table-container");if(!C)return;C.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let x=new URLSearchParams({page:v,limit:20,...Object.fromEntries(Object.entries(S).filter(([,R])=>R))}),D=await f(`${n}?${x}`);if(!D.ok){C.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${D.data?.error||"Error"}</p></div>`;return}let F=D.data?.data||[],G=D.data?.pagination;g&&g(F);let X=ra({columns:t,data:F,onEdit:i?R=>V(R):null,actions:m.map(R=>({...R,handler:oe=>R.handler(oe,w)})),emptyText:`Tidak ada ${s.toLowerCase()}`,bulkSelect:h?{selectedIds:k,onToggle:_}:null});C.innerHTML="",C.appendChild(X);let z=document.getElementById("pagination-container");if(z&&(z.innerHTML="",G&&G.pages>1)){let R=la({page:G.page,pages:G.pages,total:G.total,limit:G.limit,onPage:oe=>{v=oe,w()}});R&&z.appendChild(R)}}function I(C){let x=typeof o=="function"?o(C):o;return Le(x)}function V(C){let x=!!C,D=document.createElement("form");if(D.noValidate=!0,D.innerHTML=I(C),x){let G=typeof o=="function"?o(C):o;ca(D,C)}let{close:F}=ee({title:x?`Edit ${s}`:`Tambah ${s}`,content:D,size:"lg",confirmText:x?"Simpan Perubahan":`Tambah ${s}`,onConfirm:async(G,X)=>{if(!D.reportValidity())return;let z=G.querySelector(".modal-confirm");z.disabled=!0,z.textContent="Menyimpan...";let R=oa(D),oe=typeof o=="function"?o(C):o,He=async $e=>{for(let q of $e)if(q.type==="row")await He(q.fields);else if(q.type==="combobox"&&R[q.name]){let fe=R[q.name],ve=(q.options||[]).find(W=>{let Z=String(typeof W=="object"?W.value:W),ja=String(typeof W=="object"?W.label:W);return Z===fe||ja===fe});if(ve)R[q.name]=typeof ve=="object"?ve.value:ve;else if(q.createApi){let W={};W[q.createApi.field]=fe,q.createApi.extra&&Object.assign(W,q.createApi.extra);let Z=await f(q.createApi.path,{method:"POST",body:JSON.stringify(W)});if(Z.ok&&Z.data?.id)R[q.name]=Z.data.id;else if(Z.ok&&!Z.data?.id)R[q.name]=fe;else throw new Error(`Gagal membuat master data: ${Z.data?.error||"Unknown error"}`)}}};try{await He(oe)}catch($e){j($e.message),z.disabled=!1,z.textContent=x?"Simpan Perubahan":`Tambah ${s}`;return}y&&(R=await y(R,C));let Ka=x?"PUT":"POST",qa=x?`${n}/${C.id}`:n,Ue=await f(qa,{method:Ka,body:JSON.stringify(R)});Ue.ok?(K(x?`${s} berhasil diperbarui.`:`${s} berhasil ditambahkan.`),X(),w()):(j(Ue.data?.error||"Gagal menyimpan data."),z.disabled=!1,z.textContent=x?"Simpan Perubahan":`Tambah ${s}`)}})}function Tt(C){We(`Hapus ${s} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let x=await f(`${n}/${C.id}`,{method:"DELETE"});x.ok?(K(`${s} berhasil dihapus.`),w()):j(x.data?.error||"Gagal menghapus.")},`Hapus ${s}`)}return w(),w}P();function O(a){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!a||a==="-"||String(a).trim()===""?"":`<span class="badge ${e[a]||"badge-neutral"}">${a}</span>`}function da(a){return a==null?'<span class="badge badge-neutral">-</span>':a<0?`<span class="badge badge-danger">Expired (${Math.abs(a)}h)</span>`:a<=14?`<span class="badge badge-danger">${a} hari</span>`:a<=30?`<span class="badge badge-warning">${a} hari</span>`:`<span class="badge badge-success">${a} hari</span>`}function we(a){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[a]||"badge-neutral"}">${a||"-"}</span>`}function pa(a){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[a]||"badge-neutral"}">${a||"-"}</span>`}function Y(a){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[a]||"badge-neutral"}">${a||"-"}</span>`}L();var Ae=[],Oe=[];async function mt(){Oe=(await f("/api/branches?all=1")).data?.data||[],Ae=Oe.map(e=>({value:e.id,label:e.full_name}))}async function ma(a){await mt(),$({container:a,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:e=>we(e)},{key:"phone",label:"No. HP",render:e=>e?`<a href="tel:${e}">${e}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:e=>window.formatDate(e)},{key:"status",label:"Status",render:e=>O(e)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:Ae},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:e=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:e?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:e?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:Ae,createApi:{path:"/api/branches",field:"full_name"},value:e?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:e?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:e?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:e?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:e?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let e=await f("/api/employees?limit=10000");if(e.ok){let l=e.data.data.map(n=>({"Nama Lengkap":n.full_name,Cabang:n.branch_name||"",Divisi:n.division||"","No. HP":n.phone||"","Tgl Masuk":n.join_date||"",Status:n.status||"",Catatan:n.notes||""}));T(l,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif",Catatan:""},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif",Catatan:""}],"Template_Import_Karyawan")},onImport:async e=>{let l=o=>{if(!o)return null;let u=o.toLowerCase(),b=Oe.find(s=>s.full_name.toLowerCase()===u||s.code.toLowerCase()===u||s.name.toLowerCase()===u);return b?b.id:null},n=e.map(o=>({full_name:String(o["Nama Lengkap"]||"").trim(),branch_id:l(String(o.Cabang||"").trim()),division:String(o.Divisi||"").trim()||"FACILITY CARE",phone:String(o["No. HP"]||"").trim(),join_date:String(o["Tgl Masuk"]||"").trim(),status:String(o.Status||"").trim(),notes:String(o.Catatan||"").trim()})).filter(o=>o.full_name),t=await f("/api/employees/import",{method:"POST",body:JSON.stringify(n)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}}})}P();L();var Me=[],ua=[];async function ut(){let[a,e]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000&status=Aktif")]);Me=(a.data?.data||[]).map(l=>({value:l.id,label:l.full_name})),ua=(e.data?.data||[]).map(l=>({value:l.id,label:l.full_name}))}async function ga(a){await ut(),$({container:a,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:e=>we(e)},{key:"start_date",label:"Tgl Mulai",nowrap:!0,render:e=>window.formatDate(e)},{key:"end_date",label:"Tgl Selesai",nowrap:!0,render:e=>window.formatDate(e)},{key:"days_remaining",label:"Sisa",render:e=>da(e)},{key:"contract_type",label:"Tipe Kontrak"},{key:"pkwt_number",label:"PKWT"},{key:"status",label:"Status",render:e=>O(e)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:Me},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],formFields:e=>[{type:"row",fields:[{name:"employee_id",label:"Karyawan",type:"combobox",required:!0,options:ua,createApi:{path:"/api/employees",field:"full_name",extra:{status:"Aktif"}},value:e?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:Me,createApi:{path:"/api/branches",field:"full_name"},value:e?.branch_id}]},{type:"row",fields:[{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:e?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif"],value:e?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",required:!0,value:e?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",required:!0,value:e?.end_date}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:e?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:e?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:e?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let e=await f("/api/contracts?limit=10000");if(e.ok){let l=e.data.data.map(n=>({"Nama Karyawan":n.employee_name,Cabang:n.branch_name||"",Divisi:n.division||"","Tgl Mulai":n.start_date||"","Tgl Selesai":n.end_date||"","Tipe Kontrak":n.contract_type||"",PKWT:n.pkwt_number||"",Status:n.status||"",Catatan:n.notes||""}));T(l,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Nama Karyawan":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","Tgl Mulai":"2024-01-01","Tgl Selesai":"2024-12-31","Tipe Kontrak":"KONTRAK 1 TAHUN",PKWT:"PKWT 1",Status:"Aktif",Catatan:""}],"Template_Import_Kontrak")},onImport:async e=>{let[l,n]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000")]),t=l.data?.data||[],o=n.data?.data||[],u=i=>{if(!i)return null;let r=i.toLowerCase(),y=t.find(g=>g.full_name.toLowerCase()===r||g.code.toLowerCase()===r||g.name.toLowerCase()===r);return y?y.id:null},b=i=>{if(!i)return null;let r=i.toLowerCase(),y=o.find(g=>g.full_name.toLowerCase()===r);return y?y.id:null},s=e.map(i=>({employee_id:b(String(i["Nama Karyawan"]||"").trim()),branch_id:u(String(i.Cabang||"").trim()),division:String(i.Divisi||"").trim()||"FACILITY CARE",start_date:String(i["Tgl Mulai"]||"").trim(),end_date:String(i["Tgl Selesai"]||"").trim(),contract_type:String(i["Tipe Kontrak"]||"").trim(),pkwt_number:String(i.PKWT||"").trim(),status:String(i.Status||"").trim(),notes:String(i.Catatan||"").trim()})).filter(i=>i.employee_id&&i.start_date&&i.end_date),c=await f("/api/contracts/import",{method:"POST",body:JSON.stringify(s)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}}})}P();L();var Re=[],Fe=[];async function ba(a){let[e,l,n]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]);Re=(e.data?.data||[]).map(s=>({value:s.id,label:s.full_name}));let t=(l.data?.data||[]).map(s=>({value:s.full_name,label:s.full_name}));Fe=[...(n.data?.data||[]).filter(s=>s.role==="FC Spesialis").map(s=>({value:s.name,label:s.name}))];let u=s=>s&&!t.find(c=>c.value===s)?[...t,{value:s,label:s}]:t,b=s=>{if(!s||s==="-"||String(s).trim()==="")return"";let c=String(s).split("-");return c.length===3&&c[0].length===4?`${c[2]}-${c[1]}-${c[0]}`:s};$({container:a,title:"Jadwal Kegiatan",icon:"\u{1F5D3}\uFE0F",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:s=>pa(s)},{key:"period",label:"Periode",render:s=>Y(s)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:s=>b(s)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:s=>b(s)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:s=>b(s)},{key:"status",label:"Status",render:s=>O(s)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:Re},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Fe}],formFields:s=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Re,createApi:{path:"/api/branches",field:"full_name"},value:s?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"],value:s?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:s?.period},{name:"pic",label:"PIC",type:"combobox",options:Fe,createApi:{path:"/api/pic",field:"name"},value:s?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:s?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:s?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:s?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:s?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let s=await f("/api/schedule?limit=10000");if(s.ok){let c=s.data.data.map(i=>({Cabang:i.branch_name||"",Kegiatan:i.activity_type||"",Periode:i.period||"",PIC:i.pic||"","Tgl Opening":i.opening_date||"","Tgl Target":i.target_date||"","Tgl Selesai":i.completion_date||"",Status:i.status||"",Catatan:i.notes||""}));T(c,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done",Catatan:""}],"Template_Import_Jadwal")},onImport:async s=>{let i=(await f("/api/branches?all=1")).data?.data||[],r=d=>{if(!d)return null;let p=d.toLowerCase(),h=i.find(v=>v.full_name.toLowerCase()===p||v.code.toLowerCase()===p||v.name.toLowerCase()===p);return h?h.id:null},y=d=>{if(d==null||d==="")return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let p=String(d).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let v=Number(p);if(v>2e4&&v<99999){let S=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}let h=p.split(/[\/\-\.]/);if(h.length===3){let[v,S,k]=h.map(_=>_.trim());if(v.length===4&&S.length<=2&&k.length<=2)return`${v}-${S.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&S.length<=2&&v.length<=2)return`${k}-${S.padStart(2,"0")}-${v.padStart(2,"0")}`}return p},g=s.map(d=>({branch_id:r(String(d.Cabang||"").trim()),activity_type:String(d.Kegiatan||"").trim(),period:String(d.Periode||"").trim(),pic:String(d.PIC||d.Pic||"").trim(),opening_date:y(d["Tgl Opening"]||d["Tanggal Opening"]),target_date:y(d["Tgl Target"]||d["Tanggal Target"]),completion_date:y(d["Tgl Selesai"]||d["Tanggal Selesai"]),status:String(d.Status||"").trim(),notes:String(d.Catatan||d.Keterangan||"").trim()})).filter(d=>d.activity_type&&d.period),m=await f("/api/schedule/import",{method:"POST",body:JSON.stringify(g)});if(!m.ok)throw new Error(m.data?.error||"Import gagal")}}})}P();L();var Ke=[],Ce=[];async function ha(a){let[e,l,n]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]);Ke=(e.data?.data||[]).map(r=>({value:r.id,label:r.full_name})),Ce=(l.data?.data||[]).map(r=>({value:r.full_name,label:r.full_name}));let t=(n.data?.data||[]).filter(r=>r.role==="FC Spesialis").map(r=>({value:r.name,label:r.name})),o=(n.data?.data||[]).filter(r=>r.role==="Pelapor").map(r=>({value:r.name,label:r.name})),u=r=>r&&!Ce.find(y=>y.value===r)?[...Ce,{value:r,label:r}]:Ce,b=r=>r&&!t.find(y=>y.value===r)?[...t,{value:r,label:r}]:t,s=r=>r&&!o.find(y=>y.value===r)?[...o,{value:r,label:r}]:o,c=new Date().getFullYear(),i=Array.from({length:5},(r,y)=>String(c-y));$({container:a,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:r=>`<span class="badge badge-secondary">${r}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:r=>`<span title="${r}">${r?.length>50?r.slice(0,50)+"\u2026":r}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:r=>`<span title="${r||""}">${r?.length>40?r.slice(0,40)+"\u2026":r||"-"}</span>`},{key:"status",label:"Status",render:r=>O(r)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:r=>window.formatDate(r)},{key:"day_count",label:"Hari",render:r=>r??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:Ke},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:r=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:r?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Ke,createApi:{path:"/api/branches",field:"full_name"},value:r?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:r?.category},{name:"source",label:"Sumber Laporan",type:"select",options:[...s(r?.source),{value:"Lainnya",label:"Lainnya"}],value:r?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:r?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:u(r?.employee_name),createApi:{path:"/api/employees",field:"full_name",extra:{status:"Aktif"}},value:r?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"select",options:b(r?.fc_specialist),value:r?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:r?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:r?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:r?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let r=await f("/api/issues?limit=10000");if(r.ok){let y=r.data.data.map(g=>({Tanggal:g.report_date||"",Cabang:g.branch_name||"",Kategori:g.category||"",Sumber:g.source||"",Keluhan:g.complaint||"","Nama FC":g.employee_name||"","FC Spesialis":g.fc_specialist||"",Solusi:g.solution||"","Tgl Selesai":g.completion_date||"",Status:g.status||""}));T(y,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async r=>{let g=(await f("/api/branches?all=1")).data?.data||[],m=h=>{if(!h)return null;let v=h.toLowerCase(),S=g.find(k=>k.full_name.toLowerCase()===v||k.code.toLowerCase()===v||k.name.toLowerCase()===v);return S?S.id:null},d=r.map(h=>({branch_id:m(String(h.Cabang||"").trim()),report_date:String(h.Tanggal||"").trim(),category:String(h.Kategori||"").trim(),source:String(h.Sumber||"").trim(),complaint:String(h.Keluhan||"").trim(),employee_name:String(h["Nama FC"]||"").trim(),fc_specialist:String(h["FC Spesialis"]||"").trim(),solution:String(h.Solusi||"").trim(),completion_date:String(h["Tgl Selesai"]||"").trim(),status:String(h.Status||"").trim()})).filter(h=>h.report_date&&h.complaint&&h.category),p=await f("/api/issues/import",{method:"POST",body:JSON.stringify(d)});if(!p.ok)throw new Error(p.data?.error||"Import gagal")}}})}P();async function ya(a){let[e,l,n]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]),t=(e.data?.data||[]).map(c=>({value:c.id,label:c.full_name})),o=(l.data?.data||[]).map(c=>({value:c.full_name,label:c.full_name})),u=(n.data?.data||[]).filter(c=>c.role==="FC Spesialis").map(c=>({value:c.name,label:c.name})),b=c=>c&&!o.find(i=>i.value===c)?[...o,{value:c,label:c}]:o,s=c=>c&&!u.find(i=>i.value===c)?[...u,{value:c,label:c}]:u;$({container:a,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:c=>window.formatDate(c)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:c=>`<span title="${c||""}">${c?.length>50?c.slice(0,50)+"\u2026":c||"-"}</span>`},{key:"solution",label:"Solusi",render:c=>`<span title="${c||""}">${c?.length>40?c.slice(0,40)+"\u2026":c||"-"}</span>`},{key:"status",label:"Status",render:c=>O(c)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:c=>window.formatDate(c)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:c=>c?`<a href="${c}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async c=>{let i=new URLSearchParams(c||{}).toString(),r=await f(`/api/one-on-one?limit=10000&${i}`);if(r.ok){let y=r.data.data.map(m=>({Tanggal:m.meeting_date||"",Cabang:m.branch_name||"","Nama Karyawan":m.employee_name||"",PIC:m.pic||"",Masalah:m.problem||"",Solusi:m.solution||"",Status:m.status||"","Tgl Selesai":m.completion_date||"",Dokumen:m.document_link||""})),{downloadExcel:g}=await Promise.resolve().then(()=>(L(),J));g(y,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let c=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(L(),J));i(c,"Template_Import_OneOnOne")},onImport:async c=>{let i=m=>{if(!m)return null;let d=m.toLowerCase(),p=e.data?.data.find(h=>h.full_name.toLowerCase()===d||h.code.toLowerCase()===d||h.name.toLowerCase()===d);return p?p.id:null},r=m=>{if(!m)return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let d=String(m).trim();if(/^\d{4,5}$/.test(d)){let h=Number(d);if(h>2e4&&h<99999){let v=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);let p=d.split(/[\/\-\.]/);if(p.length===3){let[h,v,S]=p.map(k=>k.trim());if(h.length===4&&v.length<=2&&S.length<=2)return`${h}-${v.padStart(2,"0")}-${S.padStart(2,"0")}`;if(S.length===4&&v.length<=2&&h.length<=2)return`${S}-${v.padStart(2,"0")}-${h.padStart(2,"0")}`}return d},y=c.map(m=>({meeting_date:r(m.Tanggal),employee_name:String(m["Nama Karyawan"]||"").trim(),branch_id:i(String(m.Cabang||"").trim()),pic:String(m.PIC||"").trim(),problem:String(m.Masalah||"").trim(),solution:String(m.Solusi||"").trim(),status:String(m.Status||"").trim(),completion_date:r(m["Tgl Selesai"]),document_link:String(m.Dokumen||"").trim()})).filter(m=>m.meeting_date&&m.employee_name&&m.branch_id),g=await f("/api/one-on-one/import",{method:"POST",body:JSON.stringify(y)});if(!g.ok)throw new Error(g.data?.error||"Import gagal")}},formFields:c=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:c?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:t,createApi:{path:"/api/branches",field:"full_name"},value:c?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"combobox",required:!0,options:b(c?.employee_name),createApi:{path:"/api/employees",field:"full_name",extra:{status:"Aktif"}},value:c?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:s(c?.pic),createApi:{path:"/api/pic",field:"name"},value:c?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:c?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:c?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:c?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:c?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:c?.document_link}]})}P();async function fa(a){let[e,l,n]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]),t=(e.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),o=(l.data?.data||[]).map(i=>({value:i.full_name,label:i.full_name})),u=(n.data?.data||[]).filter(i=>i.role==="FC Spesialis").map(i=>({value:i.name,label:i.name})),b=i=>i&&!o.find(r=>r.value===i)?[...o,{value:i,label:i}]:o,s=i=>i&&!u.find(r=>r.value===i)?[...u,{value:i,label:i}]:u,c=Array.from({length:5},(i,r)=>String(new Date().getFullYear()-r));$({container:a,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:i=>{try{let r=JSON.parse(i);return Array.isArray(r)?r.join(", "):i||"-"}catch{return i||"-"}}},{key:"score",label:"Nilai",render:i=>i!=null?`<strong>${i}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:i=>i?`<a href="${i}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:c}],exportOptions:{moduleName:"training",onExport:async i=>{let r=new URLSearchParams(i||{}).toString(),y=await f(`/api/training?limit=10000&${r}`);if(y.ok){let g=y.data.data.map(d=>{let p=d.participants||"";try{let h=JSON.parse(p);p=Array.isArray(h)?h.join(", "):p}catch{}return{Tanggal:d.training_date||"",Batch:d.batch||"",Materi:d.subject||"",Cabang:d.branch_name||"",Trainer:d.trainer||"",Peserta:p,Nilai:d.score!==null&&d.score!==void 0?d.score:"",Dokumen:d.document_link||""}}),{downloadExcel:m}=await Promise.resolve().then(()=>(L(),J));m(g,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let i=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(L(),J));r(i,"Template_Import_Training")},onImport:async i=>{let r=d=>{if(!d)return null;let p=d.toLowerCase(),h=e.data?.data.find(v=>v.full_name.toLowerCase()===p||v.code.toLowerCase()===p||v.name.toLowerCase()===p);return h?h.id:null},y=d=>{if(!d)return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let p=String(d).trim();if(/^\d{4,5}$/.test(p)){let v=Number(p);if(v>2e4&&v<99999){let S=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);let h=p.split(/[\/\-\.]/);if(h.length===3){let[v,S,k]=h.map(_=>_.trim());if(v.length===4&&S.length<=2&&k.length<=2)return`${v}-${S.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&S.length<=2&&v.length<=2)return`${k}-${S.padStart(2,"0")}-${v.padStart(2,"0")}`}return p},g=i.map(d=>({training_date:y(d.Tanggal),batch:String(d.Batch||"").trim(),subject:String(d.Materi||"").trim(),branch_id:r(String(d.Cabang||"").trim()),trainer:String(d.Trainer||"").trim(),participants:String(d.Peserta||"").trim(),score:d.Nilai?Number(d.Nilai):null,document_link:String(d.Dokumen||"").trim()})).filter(d=>d.training_date&&d.subject&&d.branch_id),m=await f("/api/training/import",{method:"POST",body:JSON.stringify(g)});if(!m.ok)throw new Error(m.data?.error||"Import gagal")}},formFields:i=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:i?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:i?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:i?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:t,createApi:{path:"/api/branches",field:"full_name"},value:i?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:s(i?.trainer),createApi:{path:"/api/pic",field:"name"},value:i?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let r=JSON.parse(i?.participants);return Array.isArray(r)?r.join(", "):i?.participants||""}catch{return i?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:i?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:i?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:i?.notes}],onBeforeSubmit:async i=>(i.participants&&(i.participants=JSON.stringify(i.participants.split(",").map(r=>r.trim()).filter(Boolean))),i)})}P();L();async function va(a){let[e,l]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000")]),n=(e.data?.data||[]).map(s=>({value:s.id,label:s.full_name})),t=(l.data?.data||[]).map(s=>({value:s.full_name,label:s.full_name})),o=s=>s&&!t.find(c=>c.value===s)?[...t,{value:s,label:s}]:t,u=["Krishna Aryaan Permana","Agung Septiadi","Indra Saputro","Wariskin","Iqbal"],b=s=>{let c=u.map(i=>({value:i,label:i}));return s&&!c.find(i=>i.value===s)?[...c,{value:s,label:s}]:c};$({container:a,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",columns:[{key:"backup_date",label:"Tanggal Backup",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"FC Digantikan"},{key:"period",label:"Periode",render:s=>Y(s)},{key:"reliever_name",label:"Reliefer"},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:s=>s?`<span class="badge badge-info">${s}</span>`:"-"},{key:"status",label:"Status",render:s=>O(s)}],filterFields:[{type:"search",placeholder:"Cari reliefer / FC..."},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:s=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:n,createApi:{path:"/api/branches",field:"full_name"},value:s?.branch_id},{name:"period",label:"Periode",type:"select",options:["Q1","Q2","Q3","Q4"],value:s?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"FC yang Digantikan",type:"combobox",options:[{value:"",label:"BELUM ADA FC"},...o(s?.original_fc_name)],createApi:{path:"/api/employees",field:"full_name",extra:{status:"Aktif"}},value:s?.original_fc_name},{name:"reliever_name",label:"Nama Reliefer",type:"combobox",required:!0,options:b(s?.reliever_name),createApi:{path:"/api/employees",field:"full_name",extra:{status:"Aktif"}},value:s?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Backup",type:"date",required:!0,value:s?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"select",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:s?.reason},{name:"shift",label:"Shift",type:"select",options:["Pagi","Siang","Full Shift","Middle"],value:s?.shift}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done","Tidak Datang"],value:s?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let s=await f("/api/relievers?limit=10000");if(s.ok){let c=s.data.data.map(i=>({"Tanggal Backup":i.backup_date||"",Cabang:i.branch_name||"","FC Digantikan":i.original_fc_name||"",Periode:i.period||"",Reliefer:i.reliever_name||"",Keterangan:i.reason||"",Shift:i.shift||"","Tanggal Selesai":i.completion_date||"",Status:i.status||""}));T(c,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Tanggal Backup":"2024-03-10",Cabang:"001. Pondok Bambu","FC Digantikan":"Budi Santoso",Periode:"Q1",Reliefer:"Andi",Keterangan:"Sakit",Shift:"Pagi","Tanggal Selesai":"2024-03-10",Status:"Done"}],"Template_Import_Reliefer")},onImport:async s=>{let i=(await f("/api/branches?all=1")).data?.data||[],r=m=>{if(!m)return null;let d=m.toLowerCase(),p=i.find(h=>h.full_name.toLowerCase()===d||h.code.toLowerCase()===d||h.name.toLowerCase()===d);return p?p.id:null},y=s.map(m=>({branch_id:r(String(m.Cabang||"").trim()),backup_date:String(m["Tanggal Backup"]||"").trim(),original_fc_name:String(m["FC Digantikan"]||"").trim(),reliever_name:String(m.Reliefer||"").trim(),period:String(m.Periode||"").trim(),reason:String(m.Keterangan||"").trim(),shift:String(m.Shift||"").trim(),completion_date:String(m["Tanggal Selesai"]||"").trim(),status:String(m.Status||"").trim()})).filter(m=>m.reliever_name&&m.backup_date),g=await f("/api/relievers/import",{method:"POST",body:JSON.stringify(y)});if(!g.ok)throw new Error(g.data?.error||"Import gagal")}}})}P();L();async function ka(a){let l=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name})),n=Array.from({length:4},(t,o)=>String(new Date().getFullYear()-o));$({container:a,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:t=>Y(t)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"fc_score",label:"Point FC",render:t=>t!=null?`<strong class="${t>=80?"text-success":t>=60?"text-warning":"text-danger"}">${t}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:t=>t!=null?`<strong>${t}</strong>`:"-"},{key:"status",label:"Status",render:t=>O(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"select",name:"branch_id",label:"Cabang",options:l},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:l,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:t?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:t?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:t?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async t=>{let o=new URLSearchParams(t||{}).toString(),u=await f(`/api/reports/inspection?limit=10000&${o}`);if(u.ok){let b=u.data.data.map(s=>({Cabang:s.branch_name||"",Periode:s.period||"",Tanggal:s.inspection_date||"","Point FC":s.fc_score!==null&&s.fc_score!==void 0?s.fc_score:"","Point SPV":s.spv_score!==null&&s.spv_score!==void 0?s.spv_score:"",Status:s.status||"","Link Dokumen":s.document_link||"",Catatan:s.notes||""}));T(b,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:"Semua area bersih"}],"Template_Import_Inspeksi")},onImport:async t=>{let u=(await f("/api/branches?all=1")).data?.data||[],b=r=>{if(!r)return null;let y=r.toLowerCase(),g=u.find(m=>m.full_name.toLowerCase()===y||m.code.toLowerCase()===y||m.name.toLowerCase()===y);return g?g.id:null},s=r=>{if(r==null||r==="")return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let y=String(r).trim();if(y===""||y==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);if(/^\d{4,5}$/.test(y)){let m=Number(y);if(m>2e4&&m<99999){let d=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let g=y.split(/[\/\-\.]/);if(g.length===3){let[m,d,p]=g.map(h=>h.trim());if(m.length===4&&d.length<=2&&p.length<=2)return`${m}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&d.length<=2&&m.length<=2)return`${p}-${d.padStart(2,"0")}-${m.padStart(2,"0")}`}return y},c=t.map(r=>({branch_id:b(String(r.Cabang||"").trim()),period:String(r.Periode||"").trim(),inspection_date:s(r.Tanggal),fc_score:r["Point FC"]!==void 0&&r["Point FC"]!==""?Number(r["Point FC"]):null,spv_score:r["Point SPV"]!==void 0&&r["Point SPV"]!==""?Number(r["Point SPV"]):null,status:String(r.Status||"").trim(),document_link:String(r["Link Dokumen"]||"").trim(),notes:String(r.Catatan||r.Keterangan||"").trim()})).filter(r=>r.branch_id&&r.period&&r.inspection_date),i=await f("/api/reports/inspection/import",{method:"POST",body:JSON.stringify(c)});if(!i.ok)throw new Error(i.data?.error||"Import gagal")}}})}P();L();async function Sa(a){let l=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name})),n=Array.from({length:4},(t,o)=>String(new Date().getFullYear()-o));$({container:a,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:t=>`<span class="badge ${t==="Deep Cleaning"?"badge-purple":"badge-success"}">${t}</span>`},{key:"period",label:"Periode",render:t=>Y(t)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"status",label:"Status",render:t=>O(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:l},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:l,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:t?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:t?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async t=>{let o=new URLSearchParams(t||{}).toString(),u=await f(`/api/reports/cleaning?limit=10000&${o}`);if(u.ok){let b=u.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||"",Catatan:s.notes||""}));T(b,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:"Pembersihan lantai"}],"Template_Import_GCDC")},onImport:async t=>{let u=(await f("/api/branches?all=1")).data?.data||[],b=r=>{if(!r)return null;let y=r.toLowerCase(),g=u.find(m=>m.full_name.toLowerCase()===y||m.code.toLowerCase()===y||m.name.toLowerCase()===y);return g?g.id:null},s=r=>{if(r==null||r==="")return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let y=String(r).trim();if(y===""||y==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);if(/^\d{4,5}$/.test(y)){let m=Number(y);if(m>2e4&&m<99999){let d=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let g=y.split(/[\/\-\.]/);if(g.length===3){let[m,d,p]=g.map(h=>h.trim());if(m.length===4&&d.length<=2&&p.length<=2)return`${m}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&d.length<=2&&m.length<=2)return`${p}-${d.padStart(2,"0")}-${m.padStart(2,"0")}`}return y},c=t.map(r=>({branch_id:b(String(r.Cabang||"").trim()),activity_type:String(r.Jenis||r.Kegiatan||"").trim(),period:String(r.Periode||"").trim(),activity_date:s(r.Tanggal),status:String(r.Status||"").trim(),document_link:String(r["Link Dokumen"]||"").trim(),notes:String(r.Catatan||r.Keterangan||"").trim()})).filter(r=>r.branch_id&&r.activity_type&&r.period&&r.activity_date),i=await f("/api/reports/cleaning/import",{method:"POST",body:JSON.stringify(c)});if(!i.ok)throw new Error(i.data?.error||"Import gagal")}}})}P();L();async function _a(a){let l=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name})),n=Array.from({length:4},(t,o)=>String(new Date().getFullYear()-o));$({container:a,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:t=>`<span class="badge badge-warning">${t}</span>`},{key:"period",label:"Periode",render:t=>Y(t)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"status",label:"Status",render:t=>O(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:t=>t||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:l},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:l,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:t?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async t=>{let o=new URLSearchParams(t||{}).toString(),u=await f(`/api/reports/fogging?limit=10000&${o}`);if(u.ok){let b=u.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"Fogging",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||"",Catatan:s.notes||""}));T(b,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:"Tuntas"}],"Template_Import_Fogging")},onImport:async t=>{let u=(await f("/api/branches?all=1")).data?.data||[],b=r=>{if(!r)return null;let y=r.toLowerCase(),g=u.find(m=>m.full_name.toLowerCase()===y||m.code.toLowerCase()===y||m.name.toLowerCase()===y);return g?g.id:null},s=r=>{if(r==null||r==="")return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let y=String(r).trim();if(y===""||y==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);if(/^\d{4,5}$/.test(y)){let m=Number(y);if(m>2e4&&m<99999){let d=new Date(Date.UTC(1899,11,30)+m*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let g=y.split(/[\/\-\.]/);if(g.length===3){let[m,d,p]=g.map(h=>h.trim());if(m.length===4&&d.length<=2&&p.length<=2)return`${m}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&d.length<=2&&m.length<=2)return`${p}-${d.padStart(2,"0")}-${m.padStart(2,"0")}`}return y},c=t.map(r=>({branch_id:b(String(r.Cabang||"").trim()),activity_type:String(r.Jenis||r.Kegiatan||"Fogging").trim(),period:String(r.Periode||"").trim(),activity_date:s(r.Tanggal),status:String(r.Status||"").trim(),document_link:String(r["Link Dokumen"]||"").trim(),notes:String(r.Catatan||r.Keterangan||"").trim()})).filter(r=>r.branch_id&&r.period&&r.activity_date),i=await f("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(c)});if(!i.ok)throw new Error(i.data?.error||"Import gagal")}}})}P();L();async function wa(a){let[e,l,n]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]),t=(e.data?.data||[]).map(c=>({value:c.id,label:c.full_name})),o=(l.data?.data||[]).map(c=>({value:c.full_name,label:c.full_name})),u=(n.data?.data||[]).filter(c=>c.role==="FC Spesialis").map(c=>({value:c.name,label:c.name})),b=c=>c&&!o.find(i=>i.value===c)?[...o,{value:c,label:c}]:o,s=c=>c&&!u.find(i=>i.value===c)?[...u,{value:c,label:c}]:u;$({container:a,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:c=>window.formatDate(c)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:c=>`<span title="${c||""}">${c?.length>60?c.slice(0,60)+"\u2026":c||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:c=>window.formatDate(c)},{key:"status",label:"Status",render:c=>O(c)},{key:"notes",label:"Keterangan",render:c=>c?.length>40?c.slice(0,40)+"\u2026":c||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan / PIC..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]}],formFields:c=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t,createApi:{path:"/api/branches",field:"full_name"},value:c?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:s(c?.pic),createApi:{path:"/api/pic",field:"name"},value:c?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:c?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:c?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:c?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:c?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:c?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async c=>{let i=new URLSearchParams(c||{}).toString(),r=await f(`/api/reports/basecamp?limit=10000&${i}`);if(r.ok){let y=r.data.data.map(g=>({"Tgl Info":g.info_date||"",Cabang:g.branch_name||"",Permasalahan:g.problem||"",PIC:g.pic||"","Tgl Done":g.done_date||"",Status:g.status||"",Keterangan:g.notes||""}));T(y,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async c=>{let r=(await f("/api/branches?all=1")).data?.data||[],y=p=>{if(!p)return null;let h=p.toLowerCase(),v=r.find(S=>S.full_name.toLowerCase()===h||S.code.toLowerCase()===h||S.name.toLowerCase()===h);return v?v.id:null},g=p=>{if(p==null||p==="")return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let h=String(p).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let S=Number(h);if(S>2e4&&S<99999){let k=new Date(Date.UTC(1899,11,30)+S*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}let v=h.split(/[\/\-\.]/);if(v.length===3){let[S,k,_]=v.map(E=>E.trim());if(S.length===4&&k.length<=2&&_.length<=2)return`${S}-${k.padStart(2,"0")}-${_.padStart(2,"0")}`;if(_.length===4&&k.length<=2&&S.length<=2)return`${_}-${k.padStart(2,"0")}-${S.padStart(2,"0")}`}return h},m=c.map(p=>({info_date:g(p["Tgl Info"]||p["Tanggal Info"]),branch_id:y(String(p.Cabang||"").trim()),problem:String(p.Permasalahan||"").trim(),pic:String(p.PIC||"").trim(),done_date:g(p["Tgl Done"]||p["Tanggal Done"]),status:String(p.Status||"").trim(),notes:String(p.Keterangan||p.Catatan||"").trim()})).filter(p=>p.info_date&&p.branch_id&&p.problem),d=await f("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(m)});if(!d.ok)throw new Error(d.data?.error||"Import gagal")}}})}async function Ca(a){$({container:a,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let l=new URLSearchParams(e||{}).toString(),{apiFetch:n}=await Promise.resolve().then(()=>(P(),me)),t=await n(`/api/sop?limit=10000&${l}`);if(t.ok){let o=t.data.data.map(b=>({"Nama SOP":b.name||"",Kategori:b.category||"",Dokumen:b.document_link||"",Catatan:b.notes||b.description||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(L(),J));u(o,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:l}=await Promise.resolve().then(()=>(L(),J));l(e,"Template_Import_SOP")},onImport:async e=>{let l=e.map(o=>({name:String(o["Nama SOP"]||"").trim(),category:String(o.Kategori||"").trim(),document_link:String(o.Dokumen||"").trim(),description:String(o.Catatan||"").trim()})).filter(o=>o.name),{apiFetch:n}=await Promise.resolve().then(()=>(P(),me)),t=await n("/api/sop/import",{method:"POST",body:JSON.stringify(l)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function xa(a){$({container:a,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let l=new URLSearchParams(e||{}).toString(),{apiFetch:n}=await Promise.resolve().then(()=>(P(),me)),t=await n(`/api/checklist?limit=10000&${l}`);if(t.ok){let o=t.data.data.map(b=>({"Nama Checklist":b.name||"",Kategori:b.category||"",Dokumen:b.document_link||"",Deskripsi:b.description||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(L(),J));u(o,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:l}=await Promise.resolve().then(()=>(L(),J));l(e,"Template_Import_Checklist")},onImport:async e=>{let l=e.map(o=>({name:String(o["Nama Checklist"]||"").trim(),category:String(o.Kategori||"").trim(),document_link:String(o.Dokumen||"").trim(),description:String(o.Deskripsi||"").trim()})).filter(o=>o.name),{apiFetch:n}=await Promise.resolve().then(()=>(P(),me)),t=await n("/api/checklist/import",{method:"POST",body:JSON.stringify(l)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}P();L();async function qe(a,e="forms"){if(e==="supply")return bt(a);gt(a)}function gt(a){$({container:a,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}]})}async function bt(a){let l=((await f("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name}));$({container:a,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:n=>n?new Date(n).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(n,t)=>t.branch_name_ref||t.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:n=>{try{let t=JSON.parse(n);return Array.isArray(t)?t.join(", "):n}catch{return n||"-"}}},{key:"chemical_items",label:"Chemical",render:n=>{try{let t=JSON.parse(n);return Array.isArray(t)?t.join(", "):n}catch{return n||"-"}}},{key:"additional_notes",label:"Catatan",render:n=>n?.length>40?n.slice(0,40)+"\u2026":n||"-"},{key:"status",label:"Status",render:n=>O(n)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:n=>{let t=n?.tools_items;try{t=Array.isArray(JSON.parse(t))?JSON.parse(t).join(", "):t}catch{}let o=n?.chemical_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:n?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:l,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:t},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:n?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:o},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:n?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:n?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:n?.status||""},{name:"processed_by",label:"Diproses Oleh",value:n?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async n=>{let t=new URLSearchParams(n||{}).toString(),o=await f(`/api/reports/supply?limit=10000&${t}`);if(o.ok){let u=o.data.data.map(b=>{let s=b.tools_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}let c=b.chemical_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}return{Waktu:b.submitted_at||"",Pengirim:b.submitter_name||"",Cabang:b.branch_name_ref||b.branch_name||"","Alat/Barang":s||"",Chemical:c||"",Catatan:b.additional_notes||"",Status:b.status||"","Diproses Oleh":b.processed_by||""}});T(u,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async n=>{let o=(await f("/api/branches?all=1")).data?.data||[],u=i=>{if(!i)return null;let r=i.toLowerCase(),y=o.find(g=>g.full_name.toLowerCase()===r||g.code.toLowerCase()===r||g.name.toLowerCase()===r);return y?y.id:null},b=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let r=String(i).trim();if(r===""||r==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(r))return r.slice(0,10);if(/^\d{4,5}$/.test(r)){let g=Number(r);if(g>2e4&&g<99999){let m=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}let y=r.split(/[\/\-\.]/);if(y.length===3){let[g,m,d]=y.map(p=>p.trim());if(g.length===4&&m.length<=2&&d.length<=2)return`${g}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&m.length<=2&&g.length<=2)return`${d}-${m.padStart(2,"0")}-${g.padStart(2,"0")}`}return r},s=n.map(i=>({submitted_at:b(i.Waktu||i.Tanggal),submitter_name:String(i.Pengirim||"").trim(),branch_id:u(String(i.Cabang||"").trim()),tools_items:String(i["Alat/Barang"]||i.Alat||"").trim(),chemical_items:String(i.Chemical||"").trim(),additional_notes:String(i.Catatan||i.Keterangan||"").trim(),status:String(i.Status||"").trim(),processed_by:String(i["Diproses Oleh"]||i.PIC||"").trim()})).filter(i=>i.submitted_at&&i.submitter_name&&i.branch_id),c=await f("/api/reports/supply/import",{method:"POST",body:JSON.stringify(s)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(n,t)=>{let o=ee({title:"Update Status Permintaan",content:`
              <div class="form-group">
                <label class="form-label">Status</label>
                <select class="form-control" id="supply-status">
                  <option value="Pending" ${n.status==="Pending"?"selected":""}>Pending</option>
                  <option value="Diproses" ${n.status==="Diproses"?"selected":""}>Diproses</option>
                  <option value="Selesai" ${n.status==="Selesai"?"selected":""}>Selesai</option>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label">Diproses Oleh</label>
                <input type="text" class="form-control" id="supply-processed-by" value="${n.processed_by||""}" placeholder="Nama">
              </div>
            `,onConfirm:async(u,b)=>{let s=u.querySelector("#supply-status").value,c=u.querySelector("#supply-processed-by").value;(await f(`/api/reports/supply/${n.id}`,{method:"PUT",body:JSON.stringify({status:s,processed_by:c})})).ok?(K("Status diperbarui."),b(),t()):j("Gagal update status.")}})}}]})}P();L();async function Ta(a){let e=ie();if(!e||!["superadmin","admin"].includes(e.role)){a.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}$({container:a,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:l=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[l]||"badge-neutral"}">${l}</span>`},{key:"is_active",label:"Status",render:l=>l?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:l=>l?new Date(l).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:l=>{let n=!!l;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:l?.full_name},{name:"username",label:"Username",required:!n,placeholder:"username",value:l?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!n,placeholder:"email@contoh.com",value:l?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:l?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:n?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!n,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:n?l?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let l=await f("/api/users?limit=10000");if(l.ok){let n=l.data.data.map(t=>({"Nama Lengkap":t.full_name||"",Username:t.username||"",Email:t.email||"",Role:t.role||"",Status:t.is_active?"Aktif":"Nonaktif"}));T(n,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async l=>{let n=l.map(o=>({full_name:String(o["Nama Lengkap"]||"").trim(),username:String(o.Username||"").trim(),email:String(o.Email||"").trim(),role:String(o.Role||"").trim()||"viewer",password:String(o.Password||"").trim()})).filter(o=>o.username&&o.password&&o.email&&o.full_name),t=await f("/api/users/import",{method:"POST",body:JSON.stringify(n)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}}})}P();L();async function $a(a){$({container:a,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await f("/api/branches?limit=10000");if(e.ok)T(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let l=e.map(t=>({code:String(t["Kode Cabang"]||"").trim(),name:String(t["Nama Pendek"]||"").trim(),full_name:String(t["Nama Lengkap"]||"").trim(),city:String(t.Kota||"").trim()})).filter(t=>t.code&&t.name),n=await f("/api/branches/import",{method:"POST",body:JSON.stringify(l)});if(!n.ok)throw new Error(n.data?.error||"Import gagal")}}})}P();async function Ea(a){let e=new Date,l=[];a.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),t()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),t()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(o=>o.addEventListener("change",t));async function n(){try{let o=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;l=(await f(`/api/dashboard/calendar?month=${o}`)).data?.data||[]}catch(o){console.warn("[Calendar] Failed to load events, rendering empty grid:",o),l=[]}}async function t(){let o=document.getElementById("calendar-grid");if(o){o.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await n();try{let u=e.getFullYear(),b=e.getMonth(),s=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),c=document.getElementById("cal-month-label");c&&(c.textContent=s);let i=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(k=>k.value)),r=l.filter(k=>i.has(k.type)),y={};r.forEach(k=>{let _=(k.event_date||"").slice(0,10);y[_]||(y[_]=[]),y[_].push(k)});let g=new Date(u,b,1).getDay(),m=new Date(u,b+1,0).getDate(),d=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],p=new Date().toISOString().slice(0,10),h='<div class="calendar-grid">';d.forEach(k=>{h+=`<div class="cal-day-header">${k}</div>`});for(let k=0;k<g;k++)h+='<div class="cal-cell cal-cell-empty"></div>';for(let k=1;k<=m;k++){let _=`${u}-${String(b+1).padStart(2,"0")}-${String(k).padStart(2,"0")}`,E=y[_]||[],A=_===p;h+=`
          <div class="cal-cell ${A?"cal-today":""} ${E.length?"cal-has-events":""}"
               data-date="${_}" tabindex="0" role="button" aria-label="${_}">
            <div class="cal-day-num ${A?"today-num":""}">${k}</div>
            <div class="cal-events-preview">
              ${E.slice(0,3).map(w=>`
                <div class="cal-event-dot cal-color-${w.color||"gray"}" title="${xe(w.title||w.type)}">
                  <span class="cal-event-dot-label">${ht(w.title||w.branch_name||w.type,18)}</span>
                </div>
              `).join("")}
              ${E.length>3?`<div class="cal-more">+${E.length-3} lagi</div>`:""}
            </div>
          </div>`}let S=(g+m)%7;if(S!==0)for(let k=0;k<7-S;k++)h+='<div class="cal-cell cal-cell-empty"></div>';h+="</div>",o.innerHTML=h,o.querySelectorAll(".cal-cell[data-date]").forEach(k=>{k.addEventListener("click",()=>{let _=k.dataset.date,E=y[_]||[];if(!E.length)return;let A=document.getElementById("cal-event-list"),w=new Date(_+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=w,document.getElementById("cal-event-items").innerHTML=E.map(I=>`
            <div class="cal-event-item cal-color-border-${I.color||"gray"}">
              <div class="cal-event-type">${yt(I.type)}</div>
              <div class="cal-event-title">${xe(I.title||"-")}</div>
              <div class="cal-event-branch">${xe(I.branch_name||"")}</div>
              ${I.status?`<div class="cal-event-status">${xe(I.status)}</div>`:""}
              ${I.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${I.days_remaining} hari</div>`:""}
            </div>
          `).join(""),A.style.display="block"})})}catch(u){console.error("[Calendar] Render error:",u),o&&(o.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}t()}function ht(a,e){return a?a.length>e?a.slice(0,e)+"\u2026":a:""}function xe(a){return a?String(a).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function yt(a){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[a]||a}P();async function Pa(a){let e=ie(),l=(e?.full_name||e?.username||"U")[0].toUpperCase(),t={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";a.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${t},${t}99)">
            ${l}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${e?.full_name||"\u2014"}</div>
            <div class="profile-username">@${e?.username||"\u2014"}</div>
            <span class="badge badge-info" style="background:${t}18;color:${t};margin-top:6px">
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
            <span class="info-value" style="color:${t};font-weight:700">${e?.role||"\u2014"}</span>
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
  `;let o=localStorage.getItem("fm_token"),u=document.getElementById("session-info");if(o&&u)try{let b=JSON.parse(atob(o.split(".")[1])),s=new Date(b.exp*1e3);u.textContent=`Berakhir: ${s.toLocaleString("id-ID")}`}catch{u.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async b=>{b.preventDefault();let s=document.getElementById("pwd-error"),c=document.getElementById("pwd-success"),i=document.getElementById("btn-save-pwd");s.style.display="none",c.style.display="none";let r=b.target,y=r.current_password.value,g=r.new_password.value,m=r.confirm_password.value;if(g!==m){s.textContent="\u274C Konfirmasi password tidak cocok.",s.style.display="block";return}if(g.length<6){s.textContent="\u274C Password baru minimal 6 karakter.",s.style.display="block";return}i.disabled=!0,i.textContent="\u23F3 Menyimpan...";let d=await f("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:y,new_password:g})});i.disabled=!1,i.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',d.ok?(c.textContent="\u2705 Password berhasil diubah.",c.style.display="block",r.reset(),K("Password berhasil diubah.")):(s.textContent=d.data?.error||"Gagal mengubah password.",s.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}P();var Te={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function M(a){if(a==null||a==="")return null;if(a instanceof Date)return isNaN(a.getTime())?null:a.toISOString().slice(0,10);let e=String(a).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let t=Number(e);if(t>2e4&&t<99999){let o=new Date(Date.UTC(1899,11,30)+t*864e5);return isNaN(o.getTime())?null:o.toISOString().slice(0,10)}}let l=e.split(/[\/\-\.]/);if(l.length===3){let[t,o,u]=l.map(i=>i.trim()),b=Number(t),s=Number(o),c=Number(u);if(t.length===4&&b>1900)return`${t}-${o.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&c>1900)return b>12?`${u}-${o.padStart(2,"0")}-${t.padStart(2,"0")}`:s>12?`${u}-${t.padStart(2,"0")}-${o.padStart(2,"0")}`:`${u}-${o.padStart(2,"0")}-${t.padStart(2,"0")}`;if(u.length===2&&!isNaN(c)){let i=c>=50?`19${u}`:`20${u}`;return b>12?`${i}-${o.padStart(2,"0")}-${t.padStart(2,"0")}`:`${i}-${o.padStart(2,"0")}-${t.padStart(2,"0")}`}}let n=new Date(e);return isNaN(n.getTime())?null:n.toISOString().slice(0,10)}function Ia(a){return Object.values(a).every(e=>e==null||String(e).trim()==="")}var ft={validation:{required:[],map:a=>({cabang:a.CABANG,pic:a.PIC,kegiatan:a.KEGIATAN,quartal:a.QUARTAL,masa_pkwt:a["MASA PKWT"],pic_pelapor:a["PIC PELAPOR"],kontrak:a.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:a=>({name:a["Nama SOP"],category:a.Kategori||"Umum",document_link:a["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:a=>({full_name:a["Nama Lengkap"],branch_name:a.Cabang,division:a["Div / Bagian"]||"FACILITY CARE",phone:a["No. Hp"],join_date:M(a["Tanggal Masuk"]),status:a.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:a=>({employee_name:a["Nama Lengkap"],branch_name:a.Cabang,division:a["Div / Bagian"]||"FACILITY CARE",start_date:M(a["Tanggal Mulai"]),end_date:M(a["Tanggal Selesai"]),contract_type:a["Tipe Kontrak"]||"",pkwt_number:a.PKWT||"",status:a.Status||"",notes:a.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:a=>({report_date:M(a["Tanggal Info"]),branch_name:a.Cabang,category:a.Kategori,source:a["Sumber Laporan"],complaint:a.Keluhan,employee_name:a["Nama FC"],fc_specialist:a["FC Spesialis"],solution:a.Solusi,status:a.Status||"",completion_date:M(a["Tanggal Selesai"])})},one_on_one:{required:[],map:a=>({meeting_date:M(a.Tanggal),branch_name:a.Cabang,employee_name:a["Nama Karyawan"],pic:a.Pic,problem:a.Masalah,solution:a.Solusi,status:a.Status||"",completion_date:M(a["Tanggal Selesai"]),document_link:a["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:a=>({branch_name:a.Cabang,activity_type:a.Kegiatan,period:a.Periode,pic:a.Pic||a.PIC,opening_date:M(a["Tanggal Opening"]||a["Tgl Opening"]),target_date:M(a["Tanggal Target"]||a["Tgl Target"]),completion_date:M(a["Tanggal Selesai"]||a["Tgl Selesai"]),status:a.Status||"",notes:a.Keterangan||a.Catatan})},inspection:{required:[],map:a=>({inspection_date:M(a.Tanggal),branch_name:a.Cabang,period:a.Periode,status:a.Status||"",fc_score:a["Point FC SP"]!==void 0&&a["Point FC SP"]!==null?parseFloat(String(a["Point FC SP"]).replace(",",".")):null,spv_score:a["Point SPV"]!==void 0&&a["Point SPV"]!==null?parseFloat(String(a["Point SPV"]).replace(",",".")):null,document_link:a.Link,notes:""})},cleaning:{required:[],map:a=>({activity_date:M(a.Tanggal),branch_name:a.Cabang,activity_type:a["Jenis Kegiatan"]||"General Cleaning",period:a.Periode,status:a.Status||"",document_link:a.Link,notes:""})},fogging:{required:[],map:a=>({activity_date:M(a.Tanggal),branch_name:a.Cabang,period:a.Periode,status:a.Status||"",document_link:a.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:a=>({info_date:M(a["Tgl Info"]),branch_name:a.Cabang,problem:a.Permasalahan,pic:a.PIC,done_date:M(a["Tgl Done"]),status:a.Status||"",notes:a.Ket})},relievers:{required:[],map:a=>({branch_name:a.Cabang,original_fc_name:a["Nama Facility care"],period:a.Periode,reliever_name:a.Relifer,backup_date:M(a["Tanggal Back Up"]),completion_date:M(a["Tanggal Selesai"]),reason:a.Keterangan,shift:a.Shift,status:a.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:a=>({training_date:M(a.Tanggal),batch:a.Batch,subject:a.Materi,participants:a.Peserta,branch_name:a.Cabang,trainer:a.Trainer,score:a.Nilai!==void 0&&a.Nilai!==null?parseFloat(String(a.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:a=>({name:a["Master Checklist"],category:"Umum",document_link:a["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:a=>({name:a["Master Form"],category:"Umum",document_link:a["Link Document"],description:""})},supply:{required:[],map:a=>({submitted_at:M(a.Timestamp),submitter_name:a["Nama Lengkap"],branch_name:a["Kebutuhan Untuk Cabang"],tools_items:a["Alat - Alat / Barang"],tools_quantity:a["Jumlah Permintaan Alat / Barang"],chemical_items:a.Chemical,chemical_quantity:a["Jumlah Permintaan Chemical"],additional_notes:a["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:a.Status||""})}};function vt(a,e){let l=Te[a];if(!l)return{valid:[],errors:[],mapped:[],skipped:!0};let n=ft[l.module];if(!n)return{valid:[],errors:[],mapped:[],skipped:!0};let t=[],o=[],u=[];return e.filter(s=>!Ia(s)).forEach((s,c)=>{let i=e.indexOf(s)+2,r=[];n.required.forEach(({key:g,label:m})=>{let d=s[g];if(d==null||String(d).trim()===""){let p=Object.keys(s).filter(h=>h.trim()).join(", ");r.push({column:m,originalValue:d||"",reason:`Kolom "${m}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${p.slice(0,120)}`})}});let y=n.map(s);r.length>0?o.push({row:i,data:y,raw:s,errors:r}):(t.push(s),u.push(y))}),{valid:t,errors:o,mapped:u}}function Da(a){let e=[];return a.SheetNames.forEach(l=>{let n=Te[l];if(!n)return;let t=a.Sheets[l],o=window.XLSX.utils.sheet_to_json(t,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),u=vt(l,o),b=o.filter(s=>!Ia(s));e.push({sheetName:l,module:n.module,label:n.label,total:b.length,valid:u.mapped.length,errorCount:u.errors.length,errors:u.errors,mapped:u.mapped,skipped:!1})}),e}function La(){let a=window.XLSX,e=a.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu","NAMA KARYAWAN":"Budi Santoso",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02",Day:"1"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06",Day:"1","Link Document":""}]}).forEach(([n,t])=>{a.utils.book_append_sheet(e,a.utils.json_to_sheet(t),n)}),a.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function Ba(a){let e=window.XLSX,l=e.utils.book_new(),n=!1;return a.forEach(t=>{if(!t.errors||t.errors.length===0)return;n=!0;let o=t.errors.map(b=>({"No. Baris":b.row,"Kolom Gagal":(b.errors||[]).map(s=>s.column||s).join("; "),"Alasan Error":(b.errors||[]).map(s=>s.reason||s).join("; "),...Object.fromEntries(Object.entries(b.data||{}).map(([s,c])=>[s,c??""]))})),u=e.utils.json_to_sheet(o);e.utils.book_append_sheet(l,u,t.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),n?(e.writeFile(l,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var kt=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function Na(a){a.innerHTML=`
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
              ${Object.entries(Te).map(([d,{label:p}])=>`<span class="import-sheet-tag">\u{1F4C4} ${d} \u2192 ${p}</span>`).join("")}
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
  `;let e=null,l=null,n=0,t={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function o(d){Object.entries(t).forEach(([p,h])=>{h.style.display=p===d?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let d=document.getElementById("btn-backup-db");d.disabled=!0,d.textContent="\u23F3 Memproses Backup...";try{let p=await f("/api/import/backup");if(p.ok){let h=new Blob([JSON.stringify(p.data,null,2)],{type:"application/json"}),v=URL.createObjectURL(h),S=document.createElement("a");S.href=v,S.download=`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.json`,document.body.appendChild(S),S.click(),document.body.removeChild(S),URL.revokeObjectURL(v),K("Backup berhasil diunduh!")}else j("Gagal memproses backup: "+(p.data?.error||"Unknown error"))}catch(p){j("Gagal memproses backup: "+p.message)}finally{d.disabled=!1,d.textContent="\u{1F4E6} Backup Database"}});let u=document.getElementById("btn-sync-google");u&&u.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let d=u.innerHTML;u.innerHTML='<span class="spinner"></span> Menyinkronkan...',u.disabled=!0;try{let p=await f("/api/sync/google-sheets",{method:"POST"});p.ok?alert("Sinkronisasi Berhasil: "+(p.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(p.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{u.innerHTML=d,u.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{La(),K("Template Excel berhasil didownload!")});let b=document.getElementById("file-input"),s=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",d=>{d.stopPropagation(),b.click()}),b.addEventListener("change",d=>{d.target.files[0]&&c(d.target.files[0])}),s.addEventListener("dragover",d=>{d.preventDefault(),s.classList.add("drag-over")}),s.addEventListener("dragleave",()=>s.classList.remove("drag-over")),s.addEventListener("drop",d=>{d.preventDefault(),s.classList.remove("drag-over");let p=d.dataTransfer.files[0];p&&p.name.match(/\.xlsx?$/i)?c(p):j("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,b.value="",document.getElementById("file-info").style.display="none",s.style.display="",o("upload")});async function c(d){e=d,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${d.name} (${(d.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",s.style.display="none",await i(d)}async function i(d){o("validating");let p=document.getElementById("validation-status"),h=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");p.textContent="Membaca file Excel...",h.style.width="20%",await be(200);let v=await d.arrayBuffer(),S=window.XLSX.read(v,{type:"array",cellDates:!0});p.textContent=`Memvalidasi ${S.SheetNames.length} sheet...`,h.style.width="50%",await be(100),l=Da(S),h.style.width="100%",p.textContent="Validasi selesai!",await be(300),r()}catch(v){o("upload"),j("Gagal memproses file: "+v.message),document.getElementById("file-info").style.display="flex",s.style.display="none"}}function r(){o("preview");let d=l.filter(w=>!w.skipped).length,p=l.reduce((w,I)=>w+I.total,0),h=l.reduce((w,I)=>w+I.valid,0),v=l.reduce((w,I)=>w+I.errorCount,0),S=p>0?Math.round(h/p*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${d} sheet</span>
      <span class="badge badge-secondary">${p} baris</span>
      <span class="badge badge-success">${h} valid (${S}%)</span>
      ${v>0?`<span class="badge badge-danger">${v} error</span>`:""}
    `;let k=document.getElementById("preview-table-container");k.innerHTML=`
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
          ${l.map((w,I)=>`
            <tr class="${w.errorCount>0?"row-error":w.skipped?"row-skipped":"row-ok"}">
              <td><strong>${w.sheetName}</strong></td>
              <td>${w.label}</td>
              <td style="text-align:center">${w.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${w.valid}</span></td>
              <td style="text-align:center">${w.errorCount>0?`<span class="badge badge-danger">${w.errorCount}</span>`:'<span class="text-muted">\u2013</span>'}</td>
              <td style="text-align:center">
                ${w.skipped?'<span class="badge badge-neutral">Dilewati</span>':w.errorCount>0&&w.valid===0?'<span class="badge badge-danger">\u274C 0 Valid</span>':w.errorCount>0?'<span class="badge badge-warning">\u26A0\uFE0F Sebagian</span>':w.valid===0?'<span class="badge badge-neutral">Kosong</span>':'<span class="badge badge-success">\u2705 Siap</span>'}
              </td>
              <td style="text-align:center">
                ${w.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${I}">\u{1F50D} ${w.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,k.querySelectorAll(".btn-detail-error").forEach(w=>{w.addEventListener("click",()=>{let I=l[Number(w.dataset.idx)];y(I)})});let _=document.getElementById("error-detail-section"),E=document.getElementById("error-detail-container");E.innerHTML="",_.style.display="none";let A=document.getElementById("btn-start-import");h===0?(A.disabled=!0,A.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(A.disabled=!1,v>0?(A.innerHTML=`\u{1F680} Import ${h} Data Valid (${v} dilewati)`,A.title="Baris error akan dilewati, baris valid tetap diimport"):A.innerHTML=`\u{1F680} Mulai Import ${h} Data`)}function y(d){let p=document.getElementById("error-detail-section"),h=document.getElementById("error-detail-container");p.style.display="";let v=d.errors.slice(0,100).map(S=>(Array.isArray(S.errors)?S.errors:[]).map(_=>{let E=typeof _=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${S.row}</span></td>
            <td><strong>${E?_.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${E&&_.originalValue!==void 0?_.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${E?_.reason:_}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${E&&_.aliases?`Gunakan salah satu nama kolom:<br><em>${_.aliases}</em>`:E&&_.hint?_.hint:""}
            </td>
          </tr>
        `}).join("")).join("");h.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${d.sheetName} \u2014 ${d.errorCount} baris error dari ${d.total} total
          ${d.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
        ${d.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,p.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{o("upload"),document.getElementById("file-info").style.display="none",s.style.display="",e=null,b.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!l)return;Ba(l)?K("Log error berhasil didownload."):K("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let d=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";g(d)});async function g(d){o("importing"),n=Date.now();let p=[];kt.forEach(_=>{let E=l?.find(A=>A.module===_&&A.mapped?.length>0);E&&p.push(E)});let h=document.getElementById("import-steps-list");h.innerHTML=p.map(_=>`
      <div class="import-step-item" id="step-item-${_.module}">
        <span class="step-item-icon" id="step-icon-${_.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${_.label} <span class="step-item-count">(${_.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${_.module}"></span>
      </div>
    `).join("");let v=document.getElementById("import-bar"),S=document.getElementById("import-current-status"),k={totalSheets:p.length,totalRows:p.reduce((_,E)=>_+E.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let _=0;_<p.length;_++){let E=p[_],A=document.getElementById(`step-icon-${E.module}`),w=document.getElementById(`step-status-${E.module}`);A.textContent="\u{1F504}",w.textContent="Mengimport...",S.textContent=`Mengimport ${E.label}...`,v.style.width=`${Math.round(_/p.length*100)}%`;try{let I=await f(`/api/import/${E.module}`,{method:"POST",body:JSON.stringify({rows:E.mapped,onDuplicate:d})});if(I.ok){let V=I.data;k.inserted+=V.inserted||0,k.skipped+=V.skipped||0,k.moduleResults.push({label:E.label,inserted:V.inserted||0,skipped:V.skipped||0,status:"ok"}),A.textContent="\u2705",w.innerHTML=`<span class="badge badge-success">${V.inserted||0} berhasil</span>${V.skipped>0?` <span class="badge badge-neutral">${V.skipped} skip</span>`:""}`}else k.failed++,k.moduleResults.push({label:E.label,inserted:0,skipped:0,status:"error",error:I.data?.error}),A.textContent="\u274C",w.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(I){k.failed++,k.moduleResults.push({label:E.label,inserted:0,skipped:0,status:"error",error:I.message}),A.textContent="\u274C",w.innerHTML='<span class="badge badge-danger">Gagal</span>'}await be(150)}v.style.width="100%",S.textContent="Selesai!",await be(400),m(k)}function m(d){o("summary");let p=((Date.now()-n)/1e3).toFixed(1),h=d.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${h?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${h?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${d.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${d.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${d.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${d.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${d.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${d.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
      <div class="summary-stat-card info">
        <div class="stat-value">${p}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `,document.getElementById("summary-module-results").innerHTML=`
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${d.moduleResults.map(v=>`
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,l=null,b.value="",document.getElementById("file-info").style.display="none",s.style.display="",o("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function be(a){return new Promise(e=>setTimeout(e,a))}P();var je=[],Aa=[];async function Oa(a){let[e,l]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000")]);je=(e.data?.data||[]).map(n=>({value:n.id,label:n.full_name})),Aa=(l.data?.data||[]).map(n=>({value:n.full_name,label:n.full_name})),$({container:a,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:n=>n?new Date(n).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"branch_name",label:"Cabang"},{key:"sp_type",label:"Jenis SP",render:n=>`<span class="badge badge-warning">${n||"-"}</span>`},{key:"status",label:"Status",render:n=>`<span class="badge ${n==="Aktif"?"badge-danger":"badge-success"}">${n||"-"}</span>`},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:je}],exportOptions:{moduleName:"sp_data",onExport:async n=>{let t=new URLSearchParams(n||{}).toString(),o=await f(`/api/sp?limit=10000&${t}`);if(o.ok){let u=o.data.data.map(s=>({Tanggal:s.tanggal||"","Nama Karyawan":s.employee_name||"",Cabang:s.branch_name||"","Jenis SP":s.sp_type||"",Status:s.status||"",Dokumen:s.document_link||""})),{downloadExcel:b}=await Promise.resolve().then(()=>(L(),J));b(u,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08","Nama Karyawan":"Budi Santoso",Cabang:"001. Pondok Bambu","Jenis SP":"SP 1",Status:"Aktif",Dokumen:"https://link.doc"}],{downloadExcel:t}=await Promise.resolve().then(()=>(L(),J));t(n,"Template_Import_SP")},onImport:async n=>{let o=(await f("/api/branches?all=1")).data?.data||[],u=i=>{if(!i)return null;let r=i.toLowerCase(),y=o.find(g=>g.full_name.toLowerCase()===r||g.code.toLowerCase()===r||g.name.toLowerCase()===r);return y?y.id:null},b=i=>{if(!i)return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let r=String(i).trim();if(/^\d{4,5}$/.test(r)){let g=Number(r);if(g>2e4&&g<99999){let m=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(r))return r.slice(0,10);let y=r.split(/[\/\-\.]/);if(y.length===3){let[g,m,d]=y.map(p=>p.trim());if(g.length===4&&m.length<=2&&d.length<=2)return`${g}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&m.length<=2&&g.length<=2)return`${d}-${m.padStart(2,"0")}-${g.padStart(2,"0")}`}return r},s=n.map(i=>({tanggal:b(i.Tanggal),employee_name:String(i["Nama Karyawan"]||"").trim(),branch_id:u(String(i.Cabang||"").trim()),sp_type:String(i["Jenis SP"]||"").trim(),status:String(i.Status||"").trim(),document_link:String(i.Dokumen||"").trim()})).filter(i=>i.tanggal&&i.employee_name&&i.branch_id),c=await f("/api/sp/import",{method:"POST",body:JSON.stringify(s)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:Aa,createApi:{path:"/api/employees",field:"full_name",extra:{status:"Aktif"}}},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:je},{type:"select",name:"sp_type",label:"Jenis Surat Peringatan",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"select",name:"status",label:"Status",required:!0,options:["Aktif","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}P();var he=[],Ma=[];async function Ra(a){let[e,l]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000")]);he=(e.data?.data||[]).map(n=>({value:n.id,label:n.full_name})),Ma=(l.data?.data||[]).map(n=>({value:n.full_name,label:n.full_name})),$({container:a,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:n=>n?new Date(n).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:n=>`<span class="badge ${n==="Selesai"?"badge-success":"badge-warning"}">${n||"-"}</span>`},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:he},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:he}],exportOptions:{moduleName:"mutasi_data",onExport:async n=>{let t=new URLSearchParams(n||{}).toString(),o=await f(`/api/mutasi?limit=10000&${t}`);if(o.ok){let u=o.data.data.map(s=>({Tanggal:s.tanggal||"","Nama Karyawan":s.employee_name||"","Cabang Asal":s.from_branch_name||"","Cabang Tujuan":s.to_branch_name||"",Status:s.status||"",Dokumen:s.document_link||""})),{downloadExcel:b}=await Promise.resolve().then(()=>(L(),J));b(u,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:t}=await Promise.resolve().then(()=>(L(),J));t(n,"Template_Import_Mutasi")},onImport:async n=>{let o=(await f("/api/branches?all=1")).data?.data||[],u=i=>{if(!i)return null;let r=i.toLowerCase(),y=o.find(g=>g.full_name.toLowerCase()===r||g.code.toLowerCase()===r||g.name.toLowerCase()===r);return y?y.id:null},b=i=>{if(!i)return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let r=String(i).trim();if(/^\d{4,5}$/.test(r)){let g=Number(r);if(g>2e4&&g<99999){let m=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(m.getTime())?"":m.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(r))return r.slice(0,10);let y=r.split(/[\/\-\.]/);if(y.length===3){let[g,m,d]=y.map(p=>p.trim());if(g.length===4&&m.length<=2&&d.length<=2)return`${g}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&m.length<=2&&g.length<=2)return`${d}-${m.padStart(2,"0")}-${g.padStart(2,"0")}`}return r},s=n.map(i=>({tanggal:b(i.Tanggal),employee_name:String(i["Nama Karyawan"]||"").trim(),from_branch_id:u(String(i["Cabang Asal"]||"").trim()),to_branch_id:u(String(i["Cabang Tujuan"]||"").trim()),status:String(i.Status||"").trim(),document_link:String(i.Dokumen||"").trim()})).filter(i=>i.tanggal&&i.employee_name&&i.from_branch_id&&i.to_branch_id),c=await f("/api/mutasi/import",{method:"POST",body:JSON.stringify(s)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"combobox",name:"employee_name",label:"Nama Karyawan",required:!0,options:Ma,createApi:{path:"/api/employees",field:"full_name",extra:{status:"Aktif"}}},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:he,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:he,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}window.formatDate=a=>{if(!a||a==="-")return"";let e=a.split("-");return e.length===3&&e[0].length===4?`${e[2]}-${e[1]}-${e[0]}`:a};function N(a){return async e=>{if(!ce()){se("/login");return}return a(e)}}var ye=null;function St(){ye&&clearInterval(ye);let a=()=>{let e=new Date,l=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),n=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),t=document.getElementById("header-clock-time"),o=document.getElementById("header-clock-date");t&&(t.textContent=l),o&&(o.textContent=n)};a(),ye=setInterval(a,1e3)}async function _t(){try{let a=await f("/api/dashboard/kpi");if(!a.ok)return;let e=a.data?.data||a.data||{},l=(n,t)=>{let o=document.getElementById(n);o&&(o.textContent=t>0?t:"",o.style.display=t>0?"inline-flex":"none")};l("badge-issues",e.issues?.current||0),l("badge-contracts",e.expiring30?.current||0),l("badge-oo1",e.one_on_one?.current||0),l("badge-schedule",e.schedule?.current||0),l("badge-supply",e.supply?.current||0)}catch{}}var le=[];async function wt(){try{let a=await f("/api/dashboard/notifications");if(!a.ok)return;le=a.data?.data||a.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=le.length>0?"block":"none",e.textContent=le.length)}catch{}}function Ct(){if(!le.length){ee({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,l)=>l()});return}let a=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${le.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;ee({title:`Notifikasi (${le.length})`,content:a,confirmText:"Tutup",onConfirm:(e,l)=>l()})}function Fa(){let a=ie(),e=(a?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
            <div class="sidebar-avatar">${e}</div>
            <div class="sidebar-user-info">
              <div class="sidebar-user-name">${a?.full_name||"User"}</div>
              <div class="sidebar-user-role">${a?.role||""}</div>
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
            <div class="topbar-page-title" id="topbar-title">Dashboard</div>
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
              <span class="topbar-avatar">${e}</span>
              <span class="topbar-user-name">${a?.full_name?.split(" ")[0]||"User"}</span>
            </a>
          </div>
        </header>

        <main id="main-content" class="main-content"></main>
      </div>
    </div>
  `;let l=document.getElementById("sidebar"),n=document.getElementById("sidebar-overlay"),t=document.getElementById("topbar-menu-btn"),o=document.getElementById("sidebar-close"),u=()=>{l.classList.add("open"),n.classList.add("show")},b=()=>{l.classList.remove("open"),n.classList.remove("show")};t?.addEventListener("click",u),o?.addEventListener("click",b),n?.addEventListener("click",b),document.querySelectorAll(".nav-item").forEach(c=>c.addEventListener("click",b));function s(){let c=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(y=>{let g=y.dataset.route;y.classList.toggle("active",c===g||g!=="/dashboard"&&c.startsWith(g))});let i=document.getElementById("topbar-title"),r=document.querySelector(".nav-item.active .nav-label");i&&r&&(i.textContent=r.textContent)}window.addEventListener("hashchange",s),s(),St(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await f("/api/auth/logout",{method:"POST"}),de(),ye&&clearInterval(ye),se("/login")}),_t(),wt(),document.getElementById("btn-notif")?.addEventListener("click",c=>{c.preventDefault(),Ct()})}async function xt(){B("/login",({main:e})=>sa(e)),B("/dashboard",N(({main:e})=>ta(e))),B("/calendar",N(({main:e})=>Ea(e))),B("/employees",N(({main:e})=>ma(e))),B("/contracts",N(({main:e})=>ga(e))),B("/sp",N(({main:e})=>Oa(e))),B("/mutasi",N(({main:e})=>Ra(e))),B("/timeline",N(({main:e})=>ba(e))),B("/issues",N(({main:e})=>ha(e))),B("/one-on-one",N(({main:e})=>ya(e))),B("/training",N(({main:e})=>fa(e))),B("/relievers",N(({main:e})=>va(e))),B("/reports/inspection",N(({main:e})=>ka(e))),B("/reports/cleaning",N(({main:e})=>Sa(e))),B("/reports/fogging",N(({main:e})=>_a(e))),B("/reports/basecamp",N(({main:e})=>wa(e))),B("/reports/supply",N(({main:e})=>qe(e,"supply"))),B("/sop",N(({main:e})=>Ca(e))),B("/checklist",N(({main:e})=>xa(e))),B("/forms",N(({main:e})=>qe(e))),B("/users",N(({main:e})=>Ta(e))),B("/branches",N(({main:e})=>$a(e))),B("/profile",N(({main:e})=>Pa(e))),B("/settings/import",N(({main:e})=>Na(e)));let a=ce();if(!a&&window.location.hash!=="#/login"&&se("/login"),a){let e=await f("/api/auth/me");e.ok?(pe(e.data.data),Fa()):(de(),se("/login"))}window.addEventListener("fm:login",()=>{Fa(),se("/dashboard")}),Ve()}xt();
