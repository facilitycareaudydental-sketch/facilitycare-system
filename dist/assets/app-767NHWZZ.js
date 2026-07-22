var qa=Object.defineProperty;var Je=(a,e)=>()=>(a&&(e=a(a=0)),e);var Ge=(a,e)=>{for(var l in e)qa(a,l,{get:e[l],enumerable:!0})};var me={};Ge(me,{API:()=>Qe,apiFetch:()=>f,clearToken:()=>de,getToken:()=>ce,getUser:()=>ie,setToken:()=>Ee,setUser:()=>pe});function ce(){return localStorage.getItem("fm_token")}function Ee(a){localStorage.setItem("fm_token",a)}function de(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ie(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function pe(a){localStorage.setItem("fm_user",JSON.stringify(a))}async function f(a,e={}){let l=ce(),n={"Content-Type":"application/json",...l?{Authorization:`Bearer ${l}`}:{},...e.headers||{}};try{let t=`cb=${Date.now()}`,o=a.includes("?")?"&":"?",u=`${Qe}${a}${o}${t}`,b=await fetch(u,{...e,headers:n}),s;try{let c=await b.text();try{s=JSON.parse(c)}catch{s={error:`Server Error (${b.status}): ${c.substring(0,80)}...`}}}catch{s={error:"Gagal membaca respon dari server"}}return b.status===401&&(de(),window.location.hash="#/login"),{ok:b.ok,status:b.status,data:s}}catch(t){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${t.message})`}}}}var ja,Qe,P=Je(()=>{ja="",Qe=ja});var J={};Ge(J,{downloadExcel:()=>T,parseExcel:()=>Be,renderExcelButtons:()=>Ne});function Be(a){return new Promise((e,l)=>{let n=new FileReader;n.onload=t=>{try{let o=new Uint8Array(t.target.result),u=XLSX.read(o,{type:"array"}),b=u.SheetNames[0],s=u.Sheets[b],c=XLSX.utils.sheet_to_json(s,{defval:""});e(c)}catch(o){l(o)}},n.onerror=t=>l(t),n.readAsArrayBuffer(a)})}function T(a,e){try{let l=XLSX.utils.json_to_sheet(a),n=XLSX.utils.book_new();XLSX.utils.book_append_sheet(n,l,"Data"),XLSX.writeFile(n,`${e}.xlsx`)}catch(l){throw console.error("Error generating Excel file:",l),l}}function Ne(a){return`
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
  `}var L=Je(()=>{});P();var Pe={},ke=null;function B(a,e){Pe[a]=e}function se(a){window.location.hash=a}function Ve(){async function a(){let e=window.location.hash.replace("#","")||"/dashboard",[l,...n]=e.split("?"),t=Pe[l];if(!t){for(let[u,b]of Object.entries(Pe))if(u.endsWith("/*")&&l.startsWith(u.slice(0,-2))){t=b;break}}ke&&(ke(),ke=null);let o=document.getElementById("main-content");if(o&&(o.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),t){let u=new URLSearchParams(n.join("?")),b=l.split("/").filter(Boolean),s=await t({path:l,params:u,segments:b,main:o});s&&(ke=s)}else{let u=o||document.getElementById("app");u&&(u.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",a),a()}var ue;function Ha(){return ue||(ue=document.createElement("div"),ue.id="toast-container",document.body.appendChild(ue)),ue}function ze(a,e="info",l=3500){let n=Ha(),t=document.createElement("div");t.className=`toast toast-${e}`;let o={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};t.innerHTML=`<span class="toast-icon">${o[e]||"\u2139"}</span><span class="toast-msg">${a}</span>`,n.appendChild(t),requestAnimationFrame(()=>t.classList.add("show")),setTimeout(()=>{t.classList.remove("show"),setTimeout(()=>t.remove(),350)},l)}var K=a=>ze(a,"success"),j=a=>ze(a,"error");function ee({title:a,content:e,onConfirm:l,onCancel:n,confirmText:t="Simpan",cancelText:o="Batal",size:u="md",confirmClass:b="btn-primary"}){let s={sm:"400px",md:"560px",lg:"720px",xl:"900px"},c=document.createElement("div");c.className="modal-overlay",c.innerHTML=`
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
  `,e instanceof HTMLElement&&c.querySelector(".modal-body").appendChild(e);let r=()=>{c.classList.remove("show"),setTimeout(()=>c.remove(),250)};return c.querySelector(".modal-close").addEventListener("click",()=>{n&&n(),r()}),c.querySelector(".modal-cancel").addEventListener("click",()=>{n&&n(),r()}),l&&c.querySelector(".modal-confirm").addEventListener("click",()=>l(c,r)),c.addEventListener("click",i=>{i.target===c&&(n&&n(),r())}),document.body.appendChild(c),requestAnimationFrame(()=>c.classList.add("show")),{overlay:c,close:r}}function Ye(a,e,l="Konfirmasi"){return ee({title:l,content:`<p>${a}</p>`,onConfirm:(n,t)=>{e(),t()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}P();var ne={};function ge(a){if(ne[a]){try{ne[a].destroy()}catch{}delete ne[a]}}function Ua(){Object.keys(ne).forEach(ge)}var H=(a,e=0)=>{let l=Number(a);return isNaN(l)||a===null||a===void 0?e:l},U=(a,e="\u2014")=>{if(a==null||a==="")return e;let l=String(a).trim();return l===""||l==="[object Object]"?e:l},We=a=>{if(!a)return"\u2014";try{let e=new Date(a);return isNaN(e)?U(a):e.toLocaleDateString("id-ID",{day:"2-digit",month:"short",year:"numeric"})}catch{return U(a)}},Ja=a=>{if(!a)return"";try{let e=Date.now()-new Date(a).getTime();if(e<0)return"Baru saja";let l=Math.floor(e/6e4);if(l<1)return"Baru saja";if(l<60)return`${l} menit lalu`;let n=Math.floor(l/60);return n<24?`${n} jam lalu`:`${Math.floor(n/24)} hari lalu`}catch{return""}},Ze=a=>{if(!a||typeof a!="string")return"";try{let[e,l]=a.split("-");return new Date(Number(e),Number(l)-1).toLocaleDateString("id-ID",{month:"short",year:"2-digit"})}catch{return a}};function ea(a,e,l=900){if(!a)return;let n=Math.max(0,Math.round(H(e)));if(n===0){a.textContent="0";return}let t=Date.now(),o=()=>{let u=Math.min((Date.now()-t)/l,1),b=1-Math.pow(1-u,3);a.textContent=Math.round(b*n).toLocaleString("id-ID"),u<1?requestAnimationFrame(o):a.textContent=n.toLocaleString("id-ID")};requestAnimationFrame(o)}function Ga(a,e){if(a=H(a),e=H(e),e===0)return"";let l=a-e,n=Math.abs(Math.round(l/e*100));return l>0?`<span class="kpi-trend up">\u25B2 ${n}%</span>`:l<0?`<span class="kpi-trend down">\u25BC ${n}%</span>`:'<span class="kpi-trend neutral">= Sama</span>'}var Qa={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},aa=a=>{let e=U(a,"\u2014");return`<span class="status-pill ${Qa[e]||"pill-neutral"}">${e}</span>`},Va=a=>{let e=H(a,999);return e<=7?`<span class="days-badge days-critical">${e} hari</span>`:e<=14?`<span class="days-badge days-warning">${e} hari</span>`:e<=30?`<span class="days-badge days-soon">${e} hari</span>`:`<span class="days-badge days-ok">${e} hari</span>`},za={issue:{emoji:"\u26A0\uFE0F",dot:"dot-danger",label:"Permasalahan"},contract:{emoji:"\u{1F4C4}",dot:"dot-info",label:"Kontrak"},employee:{emoji:"\u{1F464}",dot:"dot-success",label:"Karyawan"},one_on_one:{emoji:"\u{1F91D}",dot:"dot-purple",label:"One on One"},training:{emoji:"\u{1F393}",dot:"dot-primary",label:"Training"},supply:{emoji:"\u{1F4E6}",dot:"dot-warning",label:"Permintaan Barang"},reliever:{emoji:"\u{1F504}",dot:"dot-teal",label:"Reliefer"},inspection:{emoji:"\u{1F50D}",dot:"dot-blue",label:"Laporan Inspeksi"}},Ya=a=>za[a]||{emoji:"\u{1F4CC}",dot:"dot-neutral",label:U(a,"Aktivitas")},Q={family:"Inter",size:11},re="#94A3B8",Se="#F1F5F9",Xa=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],Wa=()=>window.innerWidth<768;function Ie(a={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:Wa()?"bottom":"top",labels:{font:Q,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:Q,titleFont:{...Q,weight:"700"}}},scales:{x:{grid:{color:Se},ticks:{font:Q,color:re,maxRotation:0}},y:{grid:{color:Se},ticks:{font:Q,color:re},beginAtZero:!0}},...a}}var Za=()=>Array(5).fill(0).map(()=>`
  <div class="kpi-card" style="pointer-events:none">
    <div class="kpi-card-top"><div class="skeleton" style="width:44px;height:44px;border-radius:12px"></div></div>
    <div class="skeleton skeleton-text" style="width:55%;height:32px;margin:10px 0 6px"></div>
    <div class="skeleton skeleton-text" style="width:75%;height:12px;margin-bottom:4px"></div>
    <div class="skeleton skeleton-text" style="width:55%;height:11px"></div>
  </div>`).join(""),et=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Xe(a=3){return Array(a).fill(0).map((e,l)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${l<a-1?"6px":"0"};border-radius:6px"></div>`).join("")}function at(){return Array(5).fill(0).map(()=>`
    <div class="activity-item">
      <div class="skeleton" style="width:34px;height:34px;border-radius:10px;flex-shrink:0"></div>
      <div style="flex:1">
        <div class="skeleton skeleton-text" style="width:65%;height:13px;margin-bottom:5px"></div>
        <div class="skeleton skeleton-text" style="width:35%;height:11px"></div>
      </div>
    </div>`).join("")}async function ae(a,e,l=8e3){try{let n=new AbortController,t=setTimeout(()=>n.abort(),l),o=await f(a,{signal:n.signal}).catch(()=>null);if(clearTimeout(t),!o||!o.ok)return e;let u=o.data;return u?u.data!==void 0?u.data??e:u:e}catch{return e}}function tt(){["skel-donut","skel-trend","skel-insp","skel-contract"].forEach(n=>{let t=document.getElementById(n);t&&(t.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract"].forEach(n=>{let t=document.getElementById(n);if(t&&t.style.display==="none"){t.style.display="block";let o=t.parentElement;if(o&&!o.querySelector(".chart-empty")){let u=document.createElement("div");u.className="chart-empty",u.textContent="Belum ada data",t.style.display="none",o.appendChild(u)}}});let a=document.getElementById("kpi-row");a&&a.querySelector(".skeleton")&&na({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&ia({}),["table-contracts","table-issues"].forEach(n=>{let t=document.getElementById(n);t&&t.querySelector(".skeleton")&&(t.innerHTML='<div class="chart-empty">Belum ada data</div>')});let l=document.getElementById("activity-log");l&&l.querySelector(".skeleton")&&(l.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function ta(a){Ua(),a._dashRefresh&&clearInterval(a._dashRefresh),a._skelTimeout&&clearTimeout(a._skelTimeout),a.innerHTML=`
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
      <div class="kpi-row" id="kpi-row">${Za()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${et()}</div>

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
          <div id="table-contracts" class="dash-table-wrap">${Xe(3)}</div>
        </div>
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">\u26A0\uFE0F Permasalahan Terbaru</div>
              <div class="chart-card-subtitle">Open dan In Progress</div>
            </div>
            <a href="#/issues" class="btn btn-ghost btn-sm">Lihat Semua \u2192</a>
          </div>
          <div id="table-issues" class="dash-table-wrap">${Xe(3)}</div>
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
        <div id="activity-log">${at()}</div>
      </div>

    </div>
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>De(a)),a._skelTimeout=setTimeout(()=>tt(),5e3),await De(a),a._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?De(a):clearInterval(a._dashRefresh)},6e4)}async function De(a){a._skelTimeout&&(clearTimeout(a._skelTimeout),a._skelTimeout=null);let[e,l,n,t,o,u,b,s]=await Promise.all([ae("/api/dashboard/kpi",{},8e3),ae("/api/dashboard/issues-trend",{},8e3),ae("/api/dashboard/contracts-chart",{},8e3),ae("/api/dashboard/issues-summary",{},8e3),ae("/api/dashboard/inspection-bar",{},8e3),ae("/api/dashboard/contracts-expiring",[],8e3),ae("/api/dashboard/stats",{},8e3),ae("/api/dashboard/activity-log",[],8e3)]);try{na(e)}catch(r){console.warn("KPI render:",r)}try{ia(e)}catch(r){console.warn("MiniStats render:",r)}try{nt(Array.isArray(t?.by_category)?t.by_category:[])}catch(r){console.warn("Donut render:",r),te("skel-donut","chart-donut")}try{it(l)}catch(r){console.warn("Trend render:",r),te("skel-trend","chart-trend")}try{st(o)}catch(r){console.warn("InspBar render:",r),te("skel-insp","chart-insp")}try{rt(n)}catch(r){console.warn("ContractBar render:",r),te("skel-contract","chart-contract")}try{let r=Array.isArray(u)?u:[];lt(r)}catch(r){console.warn("ContractsTable render:",r)}try{let r=Array.isArray(b)?b:Array.isArray(b?.recent_issues)?b.recent_issues:[];ot(r)}catch(r){console.warn("IssuesTable render:",r)}try{ct(Array.isArray(s)?s:[])}catch(r){console.warn("ActivityLog render:",r)}let c=document.getElementById("dash-updated");c&&(c.textContent=`Diperbarui: ${new Date().toLocaleTimeString("id-ID")}`)}function na(a){let e=document.getElementById("kpi-row");if(!e)return;a=a||{};let l=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees",color:"kpi-blue",key:"employees"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts",color:"kpi-green",key:"contracts"},{icon:"\u23F0",label:"Kontrak Habis 30 Hari",sub:"",href:"#/contracts",color:"",key:"expiring30",warn:!0},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues",color:"",key:"issues",warnIfGT0:!0},{icon:"\u{1F91D}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one",color:"kpi-purple",key:"one_on_one"}];e.innerHTML=l.map(n=>{let t=H(a[n.key]?.current,0),o=a[n.key]?.prev,u=o!=null?Ga(t,o):"",b=n.color||"",s=n.sub||"";return n.warn&&(b=t>0?"kpi-amber":"kpi-green",s=t>0?`\u26A0\uFE0F ${t} kontrak segera berakhir`:"\u2705 Semua kontrak aman"),n.warnIfGT0&&(b=t>0?"kpi-red":"kpi-green"),`
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
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(n=>ea(n,parseInt(n.dataset.target)||0,700))}function nt(a){te("skel-donut","chart-donut");let e=document.getElementById("chart-donut");if(!e)return;ge("donut");let l=(a||[]).filter(o=>H(o.count)>0);if(!l.length){_e(e,"Belum ada data permasalahan");return}let n=l.map(o=>U(o.category,"Lainnya")),t=l.map(o=>H(o.count));ne.donut=new Chart(e,{type:"doughnut",data:{labels:n,datasets:[{data:t,backgroundColor:Xa.slice(0,l.length),borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{position:window.innerWidth<768?"bottom":"right",labels:{font:Q,color:"#475569",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{bodyFont:Q,titleFont:{...Q,weight:"700"},callbacks:{label:o=>` ${o.label}: ${o.parsed} kasus`}}},cutout:"65%"}})}function it(a){te("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;ge("trend"),a=a||{};let l=(a.labels||[]).map(Ze),n=(a.open||[]).map(o=>H(o)),t=(a.closed||[]).map(o=>H(o));if(!l.length){_e(e,"Belum ada data trend");return}ne.trend=new Chart(e,{type:"line",data:{labels:l,datasets:[{label:"Open",data:n,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:t,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:Ie({plugins:{legend:{position:"top"}}})})}function st(a){te("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;ge("inspBar"),a=a||{};let l=a.labels||[],n=(a.fc||[]).map(o=>H(o)),t=(a.spv||[]).map(o=>H(o));if(!l.length){_e(e,"Belum ada data inspeksi");return}ne.inspBar=new Chart(e,{type:"bar",data:{labels:l,datasets:[{label:"Skor FC",data:n,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:t,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:Ie({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:Q,color:re,maxRotation:45,minRotation:30}},y:{grid:{color:Se},ticks:{font:Q,color:re},min:0,max:100}}})})}function rt(a){te("skel-contract","chart-contract");let e=document.getElementById("chart-contract");if(!e)return;ge("contractBar"),a=a||{};let l=(a.labels||[]).map(Ze),n=(a.counts||[]).map(o=>H(o));if(!l.length){_e(e,"Belum ada data kontrak");return}let t=n.map(o=>o>5?"rgba(239,68,68,.75)":o>2?"rgba(245,158,11,.75)":"rgba(37,99,235,.65)");ne.contractBar=new Chart(e,{type:"bar",data:{labels:l,datasets:[{label:"Kontrak Berakhir",data:n,backgroundColor:t,borderRadius:6,borderSkipped:!1}]},options:Ie({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:Q,color:re}},y:{grid:{color:Se},ticks:{font:Q,color:re,precision:0},beginAtZero:!0}}})})}function lt(a){let e=document.getElementById("table-contracts");if(!e)return;let l=(a||[]).filter(n=>H(n.days_remaining,999)<=30).slice(0,10);if(!l.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada kontrak yang habis dalam 30 hari</div>';return}e.innerHTML=`
    <table class="dash-table">
      <thead><tr>
        <th>#</th><th>Nama Karyawan</th><th>Cabang</th><th>Berakhir</th><th>Sisa</th><th>Status</th>
      </tr></thead>
      <tbody>${l.map((n,t)=>`
        <tr>
          <td class="td-num">${t+1}</td>
          <td><strong>${U(n.emp_name||n.employee_name)}</strong></td>
          <td class="td-branch">${U(n.branch_name)}</td>
          <td style="white-space:nowrap;font-size:.8rem">${We(n.end_date)}</td>
          <td>${Va(n.days_remaining)}</td>
          <td>${aa(n.status)}</td>
        </tr>`).join("")}
      </tbody>
    </table>`}function ot(a){let e=document.getElementById("table-issues");if(!e)return;let l=(a||[]).slice(0,8);if(!l.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <table class="dash-table">
      <thead><tr>
        <th>Tanggal</th><th>Keluhan</th><th>Cabang</th><th>Kategori</th><th>Status</th>
      </tr></thead>
      <tbody>${l.map(n=>`
        <tr>
          <td style="white-space:nowrap;font-size:.78rem">${We(n.report_date)}</td>
          <td class="td-complaint" title="${U(n.complaint)}">${U(n.complaint)}</td>
          <td class="td-branch">${U(n.branch_name)}</td>
          <td><span class="category-tag">${U(n.category)}</span></td>
          <td>${aa(n.status)}</td>
        </tr>`).join("")}
      </tbody>
    </table>`}function ct(a){let e=document.getElementById("activity-log");if(!e)return;let l=(a||[]).slice(0,15);if(!l.length){e.innerHTML='<div class="chart-empty">Belum ada aktivitas tercatat</div>';return}e.innerHTML=`<div class="activity-list">${l.map(n=>{let t=Ya(n.type),o=U(n.label),u=n.branch?` \u2022 ${U(n.branch)}`:"",b=Ja(n.created_at);return`
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
  `;let e=document.getElementById("login-form"),l=document.getElementById("login-error"),n=document.getElementById("login-btn"),t=document.getElementById("toggle-password"),o=document.getElementById("login-password");t?.addEventListener("click",()=>{let u=o.type==="text";o.type=u?"password":"text",t.style.color=u?"":"var(--primary)"}),e?.addEventListener("submit",async u=>{u.preventDefault(),l.style.display="none";let b=e.username.value.trim(),s=e.password.value;if(!b||!s){l.textContent="Username dan password wajib diisi.",l.style.display="block";return}n.querySelector(".btn-text").style.display="none",n.querySelector(".btn-spinner").style.display="",n.disabled=!0;try{let c=await f("/api/auth/login",{method:"POST",body:JSON.stringify({username:b,password:s})});c.ok&&c.data.success?(Ee(c.data.data.token),pe(c.data.data.user),K("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(l.textContent=c.data.error||"Username atau password salah.",l.style.display="block",n.classList.add("shake"),setTimeout(()=>n.classList.remove("shake"),600))}catch{l.textContent="Gagal terhubung ke server. Periksa koneksi internet.",l.style.display="block"}finally{n.querySelector(".btn-text").style.display="",n.querySelector(".btn-spinner").style.display="none",n.disabled=!1}})}P();function ra({columns:a,data:e,onEdit:l,onDelete:n,onView:t,actions:o=[],emptyText:u="Tidak ada data",bulkSelect:b=null}){let s=document.createElement("div");if(s.className="table-wrapper",!e||e.length===0)return s.innerHTML=`<div class="empty-state"><p>${u}</p></div>`,s;let c=document.createElement("table");c.className="data-table";let r=document.createElement("thead"),i=document.createElement("tr");if(b){let g=document.createElement("th");g.style.width="40px",g.style.textAlign="center";let p=document.createElement("input");p.type="checkbox",p.id="select-all-checkbox",p.title="Pilih semua",p.addEventListener("change",()=>{e.forEach(d=>{p.checked?b.selectedIds.add(d.id):b.selectedIds.delete(d.id)}),s.querySelectorAll(".row-checkbox").forEach(d=>d.checked=p.checked),b.onToggle()}),g.appendChild(p),i.appendChild(g)}if(a.forEach(g=>{let p=document.createElement("th");p.textContent=g.label,g.width&&(p.style.width=g.width),i.appendChild(p)}),l||n||t||o.length>0){let g=document.createElement("th");g.textContent="Aksi",g.style.width="120px",i.appendChild(g)}r.appendChild(i),c.appendChild(r);let y=document.createElement("tbody");return e.forEach(g=>{let p=document.createElement("tr");if(b){let d=document.createElement("td");d.style.textAlign="center",d.style.width="40px";let m=document.createElement("input");m.type="checkbox",m.className="row-checkbox",m.checked=b.selectedIds.has(g.id),m.addEventListener("change",()=>{if(m.checked)b.selectedIds.add(g.id);else{b.selectedIds.delete(g.id);let h=document.getElementById("select-all-checkbox");h&&(h.checked=!1)}b.onToggle()}),d.appendChild(m),p.appendChild(d)}if(a.forEach(d=>{let m=document.createElement("td");if(d.render){let h=d.render(g[d.key],g);h instanceof HTMLElement?m.appendChild(h):m.innerHTML=h||""}else m.textContent=g[d.key]!==null&&g[d.key]!==void 0&&g[d.key]!==""?g[d.key]:"";d.nowrap&&(m.style.whiteSpace="nowrap"),p.appendChild(m)}),l||n||t||o.length>0){let d=document.createElement("td");d.className="actions-cell";let m=document.createElement("div");if(m.className="btn-group",t){let h=document.createElement("button");h.className="btn btn-xs btn-ghost",h.innerHTML="\u{1F441}",h.title="Lihat",h.addEventListener("click",()=>t(g)),m.appendChild(h)}if(l){let h=document.createElement("button");h.className="btn btn-xs btn-secondary",h.innerHTML="\u270F\uFE0F",h.title="Edit",h.addEventListener("click",()=>l(g)),m.appendChild(h)}o.forEach(h=>{let v=document.createElement("button");v.className=`btn btn-xs ${h.class||"btn-ghost"}`,v.innerHTML=h.icon||h.label,v.title=h.label,v.addEventListener("click",()=>h.handler(g)),m.appendChild(v)}),d.appendChild(m),p.appendChild(d)}y.appendChild(p)}),c.appendChild(y),s.appendChild(c),s}function la({page:a,pages:e,total:l,limit:n,onPage:t}){if(e<=1)return null;let o=document.createElement("div");o.className="pagination";let u=document.createElement("span");u.className="pagination-info",u.textContent=`Total: ${l} data`,o.appendChild(u);let b=document.createElement("div");b.className="pagination-btns";let s=(i,y,g=!1,p=!1)=>{let d=document.createElement("button");d.className=`btn btn-sm ${p?"btn-primary":"btn-ghost"} pagination-btn`,d.textContent=i,d.disabled=g,d.addEventListener("click",()=>t(y)),b.appendChild(d)};s("\xAB",1,a===1),s("\u2039",a-1,a===1);let c=Math.max(1,a-2),r=Math.min(e,a+2);for(let i=c;i<=r;i++)s(i,i,!1,i===a);return s("\u203A",a+1,a===e),s("\xBB",e,a===e),o.appendChild(b),o}function Le(a){return a.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="row")return`<div class="form-row">${Le(e.fields)}</div>`;let l=e.required?"required":"",n=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",t="";switch(e.type){case"textarea":t=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${l} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let u=(e.options||[]).map(r=>{let i=typeof r=="object"?r.value:r,y=typeof r=="object"?r.label:r,g=e.value==i?"selected":"";return`<option value="${i}" ${g}>${y}</option>`}).join("");t=`<select name="${e.name}" class="form-control" ${l}><option value="">-- Pilih ${e.label||""} --</option>${u}</select>`;break;case"combobox":let b=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,s=(e.options||[]).map(r=>{let i=typeof r=="object"?r.value:r;return`<option value="${typeof r=="object"?r.label:r}"></option>`}).join(""),c=e.value||"";if(e.value){let r=(e.options||[]).find(i=>(typeof i=="object"?i.value:i)==e.value);r&&(c=typeof r=="object"?r.label:r)}t=`
          <input type="text" name="${e.name}" list="${b}" class="form-control" value="${c}" placeholder="Pilih atau ketik baru..." ${l} autocomplete="off">
          <datalist id="${b}">${s}</datalist>
        `;break;case"checkbox":t=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":t=`<input type="date" name="${e.name}" class="form-control" value="${e.value||""}" ${l}>`;break;case"number":t=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${l}>`;break;case"email":t=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${l}>`;break;case"url":t=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${l}>`;break;default:t=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${l} autocomplete="off">`}let o=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${n}${t}${o}</div>`}).join("")}function oa(a){let e={},l=new FormData(a);for(let[n,t]of l.entries())e[n]=t===""?null:t;return a.querySelectorAll("input[type=checkbox]").forEach(n=>{n.checked||(e[n.name]=null)}),e}function ca(a,e){e&&Object.entries(e).forEach(([l,n])=>{let t=a.querySelector(`[name="${l}"]`);t&&(t.type==="checkbox"?t.checked=!!n:t.value=n??"")})}L();function $({container:a,title:e,icon:l,apiPath:n,columns:t,formFields:o,filterFields:u,defaultFilters:b={},itemLabel:s="Data",canCreate:c=!0,canEdit:r=!0,canDelete:i=!0,onBeforeSubmit:y,onAfterLoad:g,extraActions:p=[],initialSearch:d="",exportOptions:m=null,bulkDelete:h=!1}){let v=1,S={...b};d&&(S.search=d);let k=new Set;a.innerHTML=`
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
    
    ${m?Ne(m.moduleName):""}

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
  `;function _(){if(!document.getElementById("bulk-toolbar"))return;let x=document.getElementById("bulk-count"),I=document.getElementById("btn-bulk-delete"),R=document.getElementById("btn-bulk-cancel");x.textContent=`${k.size} item dipilih`,k.size>0?(I.disabled=!1,R.disabled=!1):(I.disabled=!0,R.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{k.clear(),document.querySelectorAll(".row-checkbox").forEach(x=>x.checked=!1);let C=document.getElementById("select-all-checkbox");C&&(C.checked=!1),_()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(k.size===0)return;let C=[...k],x=document.createElement("div");x.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",x.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${C.length} ${s}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${C.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(x),x.querySelector("#bulk-cancel-btn").addEventListener("click",()=>x.remove()),x.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let I=x.querySelector("#bulk-confirm-btn");I.disabled=!0,I.textContent="Menghapus...";let R=await f(`${n}/bulk`,{method:"DELETE",body:JSON.stringify({ids:C})});x.remove(),R.ok?(K(`${C.length} ${s} berhasil dihapus.`),k.clear(),_(),w()):j(R.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),O;if(E?.addEventListener("input",C=>{clearTimeout(O),O=setTimeout(()=>{S.search=C.target.value,v=1,w()},400)}),u?.forEach(C=>{C.type==="select"&&document.getElementById(`filter-${C.name}`)?.addEventListener("change",x=>{S[C.name]=x.target.value,v=1,w()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{S={...b},E&&(E.value=""),u?.forEach(C=>{let x=document.getElementById(`filter-${C.name}`);x&&(x.value="")}),v=1,w()}),document.getElementById("btn-create")?.addEventListener("click",()=>V(null)),m){document.getElementById(`btn-export-${m.moduleName}`)?.addEventListener("click",async x=>{let I=x.target,R=I.innerHTML;I.innerHTML="\u23F3 Loading...",I.disabled=!0;try{await m.onExport()}catch{j("Gagal export data")}finally{I.innerHTML=R,I.disabled=!1}}),document.getElementById(`btn-template-${m.moduleName}`)?.addEventListener("click",()=>{m.onTemplate()});let C=document.getElementById(`input-import-${m.moduleName}`);C?.addEventListener("change",async x=>{let I=x.target.files[0];if(!I)return;let R=C.parentElement,G=R.innerHTML;R.innerHTML="\u23F3 Memproses...",R.style.pointerEvents="none",C.disabled=!0;try{let W=await Be(I);if(W.length===0)throw new Error("File kosong atau format salah");await m.onImport(W),K("Import berhasil!"),w()}catch(W){j(W.message||"Gagal import data")}finally{R.innerHTML=G,R.style.pointerEvents="auto",C.disabled=!1,C.value=""}})}async function w(){let C=document.getElementById("table-container");if(!C)return;C.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let x=new URLSearchParams({page:v,limit:20,...Object.fromEntries(Object.entries(S).filter(([,F])=>F))}),I=await f(`${n}?${x}`);if(!I.ok){C.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${I.data?.error||"Error"}</p></div>`;return}let R=I.data?.data||[],G=I.data?.pagination;g&&g(R);let W=ra({columns:t,data:R,onEdit:r?F=>V(F):null,actions:p.map(F=>({...F,handler:oe=>F.handler(oe,w)})),emptyText:`Tidak ada ${s.toLowerCase()}`,bulkSelect:h?{selectedIds:k,onToggle:_}:null});C.innerHTML="",C.appendChild(W);let z=document.getElementById("pagination-container");if(z&&(z.innerHTML="",G&&G.pages>1)){let F=la({page:G.page,pages:G.pages,total:G.total,limit:G.limit,onPage:oe=>{v=oe,w()}});F&&z.appendChild(F)}}function D(C){let x=typeof o=="function"?o(C):o;return Le(x)}function V(C){let x=!!C,I=document.createElement("form");if(I.noValidate=!0,I.innerHTML=D(C),x){let G=typeof o=="function"?o(C):o;ca(I,C)}let{close:R}=ee({title:x?`Edit ${s}`:`Tambah ${s}`,content:I,size:"lg",confirmText:x?"Simpan Perubahan":`Tambah ${s}`,onConfirm:async(G,W)=>{if(!I.reportValidity())return;let z=G.querySelector(".modal-confirm");z.disabled=!0,z.textContent="Menyimpan...";let F=oa(I),oe=typeof o=="function"?o(C):o,He=async $e=>{for(let q of $e)if(q.type==="row")await He(q.fields);else if(q.type==="combobox"&&F[q.name]){let fe=F[q.name],ve=(q.options||[]).find(Y=>{let Z=String(typeof Y=="object"?Y.value:Y),Ka=String(typeof Y=="object"?Y.label:Y);return Z===fe||Ka===fe});if(ve)F[q.name]=typeof ve=="object"?ve.value:ve;else if(q.createApi){let Y={};Y[q.createApi.field]=fe,q.createApi.extra&&Object.assign(Y,q.createApi.extra);let Z=await f(q.createApi.path,{method:"POST",body:JSON.stringify(Y)});if(Z.ok&&Z.data?.id)F[q.name]=Z.data.id;else if(Z.ok&&!Z.data?.id)F[q.name]=fe;else throw new Error(`Gagal membuat master data: ${Z.data?.error||"Unknown error"}`)}}};try{await He(oe)}catch($e){j($e.message),z.disabled=!1,z.textContent=x?"Simpan Perubahan":`Tambah ${s}`;return}y&&(F=await y(F,C));let Fa=x?"PUT":"POST",Ra=x?`${n}/${C.id}`:n,Ue=await f(Ra,{method:Fa,body:JSON.stringify(F)});Ue.ok?(K(x?`${s} berhasil diperbarui.`:`${s} berhasil ditambahkan.`),W(),w()):(j(Ue.data?.error||"Gagal menyimpan data."),z.disabled=!1,z.textContent=x?"Simpan Perubahan":`Tambah ${s}`)}})}function Tt(C){Ye(`Hapus ${s} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let x=await f(`${n}/${C.id}`,{method:"DELETE"});x.ok?(K(`${s} berhasil dihapus.`),w()):j(x.data?.error||"Gagal menghapus.")},`Hapus ${s}`)}return w(),w}P();function A(a){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!a||a==="-"||String(a).trim()===""?"":`<span class="badge ${e[a]||"badge-neutral"}">${a}</span>`}function da(a){return a==null?'<span class="badge badge-neutral">-</span>':a<0?`<span class="badge badge-danger">Expired (${Math.abs(a)}h)</span>`:a<=14?`<span class="badge badge-danger">${a} hari</span>`:a<=30?`<span class="badge badge-warning">${a} hari</span>`:`<span class="badge badge-success">${a} hari</span>`}function we(a){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[a]||"badge-neutral"}">${a||"-"}</span>`}function pa(a){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[a]||"badge-neutral"}">${a||"-"}</span>`}function X(a){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[a]||"badge-neutral"}">${a||"-"}</span>`}L();var Oe=[],Ae=[];async function dt(){Ae=(await f("/api/branches?all=1")).data?.data||[],Oe=Ae.map(e=>({value:e.id,label:e.full_name}))}async function ma(a){await dt(),$({container:a,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:e=>we(e)},{key:"phone",label:"No. HP",render:e=>e?`<a href="tel:${e}">${e}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:e=>window.formatDate(e)},{key:"status",label:"Status",render:e=>A(e)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:Oe},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:e=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:e?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:e?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:Oe,createApi:{path:"/api/branches",field:"full_name"},value:e?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:e?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:e?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:e?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:e?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let e=await f("/api/employees?limit=10000");if(e.ok){let l=e.data.data.map(n=>({"Nama Lengkap":n.full_name,Cabang:n.branch_name||"",Divisi:n.division||"","No. HP":n.phone||"","Tgl Masuk":n.join_date||"",Status:n.status||"",Catatan:n.notes||""}));T(l,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif",Catatan:""},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif",Catatan:""}],"Template_Import_Karyawan")},onImport:async e=>{let l=o=>{if(!o)return null;let u=o.toLowerCase(),b=Ae.find(s=>s.full_name.toLowerCase()===u||s.code.toLowerCase()===u||s.name.toLowerCase()===u);return b?b.id:null},n=e.map(o=>({full_name:String(o["Nama Lengkap"]||"").trim(),branch_id:l(String(o.Cabang||"").trim()),division:String(o.Divisi||"").trim()||"FACILITY CARE",phone:String(o["No. HP"]||"").trim(),join_date:String(o["Tgl Masuk"]||"").trim(),status:String(o.Status||"").trim(),notes:String(o.Catatan||"").trim()})).filter(o=>o.full_name),t=await f("/api/employees/import",{method:"POST",body:JSON.stringify(n)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}}})}P();L();var Me=[],ua=[];async function pt(){let[a,e]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000&status=Aktif")]);Me=(a.data?.data||[]).map(l=>({value:l.id,label:l.full_name})),ua=(e.data?.data||[]).map(l=>({value:l.id,label:l.full_name}))}async function ga(a){await pt(),$({container:a,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",columns:[{key:"employee_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Div / Bagian",render:e=>we(e)},{key:"start_date",label:"Tanggal Mulai",nowrap:!0,render:e=>window.formatDate(e)},{key:"end_date",label:"Tanggal Selesai",nowrap:!0,render:e=>window.formatDate(e)},{key:"days_remaining",label:"Sisa Kontrak",render:e=>da(e)},{key:"status",label:"Status",render:e=>A(e)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:Me},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],formFields:e=>[{type:"row",fields:[{name:"employee_id",label:"Nama Lengkap",type:"select",required:!0,options:ua,value:e?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:Me,createApi:{path:"/api/branches",field:"full_name"},value:e?.branch_id}]},{type:"row",fields:[{name:"division",label:"Div / Bagian",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:e?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif"],value:e?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",required:!0,value:e?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",required:!0,value:e?.end_date}]}],exportOptions:{moduleName:"contracts",onExport:async()=>{let e=await f("/api/contracts?limit=10000");if(e.ok){let l=e.data.data.map(n=>({"Nama Lengkap":n.employee_name,Cabang:n.branch_name||"","Div / Bagian":n.division||"","Tanggal Mulai":n.start_date||"","Tanggal Selesai":n.end_date||"","Sisa Kontrak":n.days_remaining!==null&&n.days_remaining!==void 0?`${n.days_remaining} Hari`:"",Status:n.status||""}));T(l,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365 Hari",Status:"Aktif"}],"Template_Import_Kontrak")},onImport:async e=>{let[l,n]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000")]),t=l.data?.data||[],o=n.data?.data||[],u=i=>{if(!i)return null;let y=i.toLowerCase(),g=t.find(p=>p.full_name.toLowerCase()===y||p.code.toLowerCase()===y||p.name.toLowerCase()===y);return g?g.id:null},b=i=>{if(!i)return null;let y=i.toLowerCase(),g=o.find(p=>p.full_name.toLowerCase()===y);return g?g.id:null},s=i=>{if(!i)return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let y=String(i).trim();if(/^\d{4,5}$/.test(y)){let p=Number(y);if(p>2e4&&p<99999){let d=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);let g=y.split(/[\/\-\.]/);if(g.length===3){let[p,d,m]=g.map(h=>h.trim());if(p.length===4&&d.length<=2&&m.length<=2)return`${p}-${d.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&d.length<=2&&p.length<=2)return`${m}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`}return y},c=e.map(i=>({employee_id:b(String(i["Nama Lengkap"]||"").trim()),branch_id:u(String(i.Cabang||"").trim()),division:String(i["Div / Bagian"]||"").trim()||"FACILITY CARE",start_date:s(i["Tanggal Mulai"]),end_date:s(i["Tanggal Selesai"]),status:String(i.Status||"").trim()})).filter(i=>i.employee_id&&i.start_date&&i.end_date),r=await f("/api/contracts/import",{method:"POST",body:JSON.stringify(c)});if(!r.ok)throw new Error(r.data?.error||"Import gagal")}}})}P();L();var Fe=[],Re=[];async function ba(a){let[e,l,n]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]);Fe=(e.data?.data||[]).map(s=>({value:s.id,label:s.full_name}));let t=(l.data?.data||[]).map(s=>({value:s.full_name,label:s.full_name}));Re=[...(n.data?.data||[]).filter(s=>s.role==="FC Spesialis").map(s=>({value:s.name,label:s.name}))];let u=s=>s&&!t.find(c=>c.value===s)?[...t,{value:s,label:s}]:t,b=s=>{if(!s||s==="-"||String(s).trim()==="")return"";let c=String(s).split("-");return c.length===3&&c[0].length===4?`${c[2]}-${c[1]}-${c[0]}`:s};$({container:a,title:"Jadwal Kegiatan",icon:"\u{1F5D3}\uFE0F",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:s=>pa(s)},{key:"period",label:"Periode",render:s=>X(s)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:s=>b(s)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:s=>b(s)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:s=>b(s)},{key:"status",label:"Status",render:s=>A(s)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:Fe},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Re}],formFields:s=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Fe,createApi:{path:"/api/branches",field:"full_name"},value:s?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"],value:s?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:s?.period},{name:"pic",label:"PIC",type:"combobox",options:Re,createApi:{path:"/api/pic",field:"name"},value:s?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:s?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:s?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:s?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:s?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let s=await f("/api/schedule?limit=10000");if(s.ok){let c=s.data.data.map(r=>({Cabang:r.branch_name||"",Kegiatan:r.activity_type||"",Periode:r.period||"",PIC:r.pic||"","Tgl Opening":r.opening_date||"","Tgl Target":r.target_date||"","Tgl Selesai":r.completion_date||"",Status:r.status||"",Catatan:r.notes||""}));T(c,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done",Catatan:""}],"Template_Import_Jadwal")},onImport:async s=>{let r=(await f("/api/branches?all=1")).data?.data||[],i=d=>{if(!d)return null;let m=d.toLowerCase(),h=r.find(v=>v.full_name.toLowerCase()===m||v.code.toLowerCase()===m||v.name.toLowerCase()===m);return h?h.id:null},y=d=>{if(d==null||d==="")return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let m=String(d).trim();if(m===""||m==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);if(/^\d{4,5}$/.test(m)){let v=Number(m);if(v>2e4&&v<99999){let S=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}let h=m.split(/[\/\-\.]/);if(h.length===3){let[v,S,k]=h.map(_=>_.trim());if(v.length===4&&S.length<=2&&k.length<=2)return`${v}-${S.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&S.length<=2&&v.length<=2)return`${k}-${S.padStart(2,"0")}-${v.padStart(2,"0")}`}return m},g=s.map(d=>({branch_id:i(String(d.Cabang||"").trim()),activity_type:String(d.Kegiatan||"").trim(),period:String(d.Periode||"").trim(),pic:String(d.PIC||d.Pic||"").trim(),opening_date:y(d["Tgl Opening"]||d["Tanggal Opening"]),target_date:y(d["Tgl Target"]||d["Tanggal Target"]),completion_date:y(d["Tgl Selesai"]||d["Tanggal Selesai"]),status:String(d.Status||"").trim(),notes:String(d.Catatan||d.Keterangan||"").trim()})).filter(d=>d.activity_type&&d.period),p=await f("/api/schedule/import",{method:"POST",body:JSON.stringify(g)});if(!p.ok)throw new Error(p.data?.error||"Import gagal")}}})}P();L();var Ke=[],Ce=[];async function ha(a){let[e,l,n]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]);Ke=(e.data?.data||[]).map(i=>({value:i.id,label:i.full_name})),Ce=(l.data?.data||[]).map(i=>({value:i.full_name,label:i.full_name}));let t=(n.data?.data||[]).filter(i=>i.role==="FC Spesialis").map(i=>({value:i.name,label:i.name})),o=(n.data?.data||[]).filter(i=>i.role==="Pelapor").map(i=>({value:i.name,label:i.name})),u=i=>i&&!Ce.find(y=>y.value===i)?[...Ce,{value:i,label:i}]:Ce,b=i=>i&&!t.find(y=>y.value===i)?[...t,{value:i,label:i}]:t,s=i=>i&&!o.find(y=>y.value===i)?[...o,{value:i,label:i}]:o,c=new Date().getFullYear(),r=Array.from({length:5},(i,y)=>String(c-y));$({container:a,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:i=>window.formatDate(i)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:i=>`<span class="badge badge-secondary">${i}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:i=>`<span title="${i}">${i?.length>50?i.slice(0,50)+"\u2026":i}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:i=>`<span title="${i||""}">${i?.length>40?i.slice(0,40)+"\u2026":i||"-"}</span>`},{key:"status",label:"Status",render:i=>A(i)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:i=>window.formatDate(i)},{key:"day_count",label:"Hari",render:i=>i??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:Ke},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:r}],formFields:i=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:i?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Ke,createApi:{path:"/api/branches",field:"full_name"},value:i?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:i?.category},{name:"source",label:"Sumber Laporan",type:"select",options:[...s(i?.source),{value:"Lainnya",label:"Lainnya"}],value:i?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:i?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"select",options:u(i?.employee_name),value:i?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"select",options:b(i?.fc_specialist),value:i?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:i?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:i?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:i?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let i=await f("/api/issues?limit=10000");if(i.ok){let y=i.data.data.map(g=>({Tanggal:g.report_date||"",Cabang:g.branch_name||"",Kategori:g.category||"",Sumber:g.source||"",Keluhan:g.complaint||"","Nama FC":g.employee_name||"","FC Spesialis":g.fc_specialist||"",Solusi:g.solution||"","Tgl Selesai":g.completion_date||"",Status:g.status||""}));T(y,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async i=>{let g=(await f("/api/branches?all=1")).data?.data||[],p=h=>{if(!h)return null;let v=h.toLowerCase(),S=g.find(k=>k.full_name.toLowerCase()===v||k.code.toLowerCase()===v||k.name.toLowerCase()===v);return S?S.id:null},d=i.map(h=>({branch_id:p(String(h.Cabang||"").trim()),report_date:String(h.Tanggal||"").trim(),category:String(h.Kategori||"").trim(),source:String(h.Sumber||"").trim(),complaint:String(h.Keluhan||"").trim(),employee_name:String(h["Nama FC"]||"").trim(),fc_specialist:String(h["FC Spesialis"]||"").trim(),solution:String(h.Solusi||"").trim(),completion_date:String(h["Tgl Selesai"]||"").trim(),status:String(h.Status||"").trim()})).filter(h=>h.report_date&&h.complaint&&h.category),m=await f("/api/issues/import",{method:"POST",body:JSON.stringify(d)});if(!m.ok)throw new Error(m.data?.error||"Import gagal")}}})}P();async function ya(a){let[e,l,n]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]),t=(e.data?.data||[]).map(c=>({value:c.id,label:c.full_name})),o=(l.data?.data||[]).map(c=>({value:c.full_name,label:c.full_name})),u=(n.data?.data||[]).filter(c=>c.role==="FC Spesialis").map(c=>({value:c.name,label:c.name})),b=c=>c&&!o.find(r=>r.value===c)?[...o,{value:c,label:c}]:o,s=c=>c&&!u.find(r=>r.value===c)?[...u,{value:c,label:c}]:u;$({container:a,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:c=>window.formatDate(c)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:c=>`<span title="${c||""}">${c?.length>50?c.slice(0,50)+"\u2026":c||"-"}</span>`},{key:"solution",label:"Solusi",render:c=>`<span title="${c||""}">${c?.length>40?c.slice(0,40)+"\u2026":c||"-"}</span>`},{key:"status",label:"Status",render:c=>A(c)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:c=>window.formatDate(c)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:c=>c?`<a href="${c}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async c=>{let r=new URLSearchParams(c||{}).toString(),i=await f(`/api/one-on-one?limit=10000&${r}`);if(i.ok){let y=i.data.data.map(p=>({Tanggal:p.meeting_date||"",Cabang:p.branch_name||"","Nama Karyawan":p.employee_name||"",PIC:p.pic||"",Masalah:p.problem||"",Solusi:p.solution||"",Status:p.status||"","Tgl Selesai":p.completion_date||"",Dokumen:p.document_link||""})),{downloadExcel:g}=await Promise.resolve().then(()=>(L(),J));g(y,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let c=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(L(),J));r(c,"Template_Import_OneOnOne")},onImport:async c=>{let r=p=>{if(!p)return null;let d=p.toLowerCase(),m=e.data?.data.find(h=>h.full_name.toLowerCase()===d||h.code.toLowerCase()===d||h.name.toLowerCase()===d);return m?m.id:null},i=p=>{if(!p)return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let d=String(p).trim();if(/^\d{4,5}$/.test(d)){let h=Number(d);if(h>2e4&&h<99999){let v=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);let m=d.split(/[\/\-\.]/);if(m.length===3){let[h,v,S]=m.map(k=>k.trim());if(h.length===4&&v.length<=2&&S.length<=2)return`${h}-${v.padStart(2,"0")}-${S.padStart(2,"0")}`;if(S.length===4&&v.length<=2&&h.length<=2)return`${S}-${v.padStart(2,"0")}-${h.padStart(2,"0")}`}return d},y=c.map(p=>({meeting_date:i(p.Tanggal),employee_name:String(p["Nama Karyawan"]||"").trim(),branch_id:r(String(p.Cabang||"").trim()),pic:String(p.PIC||"").trim(),problem:String(p.Masalah||"").trim(),solution:String(p.Solusi||"").trim(),status:String(p.Status||"").trim(),completion_date:i(p["Tgl Selesai"]),document_link:String(p.Dokumen||"").trim()})).filter(p=>p.meeting_date&&p.employee_name&&p.branch_id),g=await f("/api/one-on-one/import",{method:"POST",body:JSON.stringify(y)});if(!g.ok)throw new Error(g.data?.error||"Import gagal")}},formFields:c=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:c?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:t,createApi:{path:"/api/branches",field:"full_name"},value:c?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:b(c?.employee_name),value:c?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:s(c?.pic),createApi:{path:"/api/pic",field:"name"},value:c?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:c?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:c?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:c?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:c?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:c?.document_link}]})}P();async function fa(a){let[e,l,n]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]),t=(e.data?.data||[]).map(r=>({value:r.id,label:r.full_name})),o=(l.data?.data||[]).map(r=>({value:r.full_name,label:r.full_name})),u=(n.data?.data||[]).filter(r=>r.role==="FC Spesialis").map(r=>({value:r.name,label:r.name})),b=r=>r&&!o.find(i=>i.value===r)?[...o,{value:r,label:r}]:o,s=r=>r&&!u.find(i=>i.value===r)?[...u,{value:r,label:r}]:u,c=Array.from({length:5},(r,i)=>String(new Date().getFullYear()-i));$({container:a,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:r=>{try{let i=JSON.parse(r);return Array.isArray(i)?i.join(", "):r||"-"}catch{return r||"-"}}},{key:"score",label:"Nilai",render:r=>r!=null?`<strong>${r}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:r=>r?`<a href="${r}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:c}],exportOptions:{moduleName:"training",onExport:async r=>{let i=new URLSearchParams(r||{}).toString(),y=await f(`/api/training?limit=10000&${i}`);if(y.ok){let g=y.data.data.map(d=>{let m=d.participants||"";try{let h=JSON.parse(m);m=Array.isArray(h)?h.join(", "):m}catch{}return{Tanggal:d.training_date||"",Batch:d.batch||"",Materi:d.subject||"",Cabang:d.branch_name||"",Trainer:d.trainer||"",Peserta:m,Nilai:d.score!==null&&d.score!==void 0?d.score:"",Dokumen:d.document_link||""}}),{downloadExcel:p}=await Promise.resolve().then(()=>(L(),J));p(g,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let r=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(L(),J));i(r,"Template_Import_Training")},onImport:async r=>{let i=d=>{if(!d)return null;let m=d.toLowerCase(),h=e.data?.data.find(v=>v.full_name.toLowerCase()===m||v.code.toLowerCase()===m||v.name.toLowerCase()===m);return h?h.id:null},y=d=>{if(!d)return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let m=String(d).trim();if(/^\d{4,5}$/.test(m)){let v=Number(m);if(v>2e4&&v<99999){let S=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);let h=m.split(/[\/\-\.]/);if(h.length===3){let[v,S,k]=h.map(_=>_.trim());if(v.length===4&&S.length<=2&&k.length<=2)return`${v}-${S.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&S.length<=2&&v.length<=2)return`${k}-${S.padStart(2,"0")}-${v.padStart(2,"0")}`}return m},g=r.map(d=>({training_date:y(d.Tanggal),batch:String(d.Batch||"").trim(),subject:String(d.Materi||"").trim(),branch_id:i(String(d.Cabang||"").trim()),trainer:String(d.Trainer||"").trim(),participants:String(d.Peserta||"").trim(),score:d.Nilai?Number(d.Nilai):null,document_link:String(d.Dokumen||"").trim()})).filter(d=>d.training_date&&d.subject&&d.branch_id),p=await f("/api/training/import",{method:"POST",body:JSON.stringify(g)});if(!p.ok)throw new Error(p.data?.error||"Import gagal")}},formFields:r=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:r?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:r?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:r?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:t,createApi:{path:"/api/branches",field:"full_name"},value:r?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:s(r?.trainer),createApi:{path:"/api/pic",field:"name"},value:r?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let i=JSON.parse(r?.participants);return Array.isArray(i)?i.join(", "):r?.participants||""}catch{return r?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:r?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:r?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:r?.notes}],onBeforeSubmit:async r=>(r.participants&&(r.participants=JSON.stringify(r.participants.split(",").map(i=>i.trim()).filter(Boolean))),r)})}P();L();async function va(a){let[e,l]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000")]),n=(e.data?.data||[]).map(s=>({value:s.id,label:s.full_name})),t=(l.data?.data||[]).map(s=>({value:s.full_name,label:s.full_name})),o=s=>s&&!t.find(c=>c.value===s)?[...t,{value:s,label:s}]:t,u=["Krishna Aryaan Permana","Agung Septiadi","Indra Saputro","Wariskin","Iqbal"],b=s=>{let c=u.map(r=>({value:r,label:r}));return s&&!c.find(r=>r.value===s)?[...c,{value:s,label:s}]:c};$({container:a,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",columns:[{key:"backup_date",label:"Tanggal Backup",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"FC Digantikan"},{key:"period",label:"Periode",render:s=>X(s)},{key:"reliever_name",label:"Reliefer"},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:s=>s?`<span class="badge badge-info">${s}</span>`:"-"},{key:"status",label:"Status",render:s=>A(s)}],filterFields:[{type:"search",placeholder:"Cari reliefer / FC..."},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:s=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:n,createApi:{path:"/api/branches",field:"full_name"},value:s?.branch_id},{name:"period",label:"Periode",type:"select",options:["Q1","Q2","Q3","Q4"],value:s?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"FC yang Digantikan",type:"select",options:[{value:"",label:"BELUM ADA FC"},...o(s?.original_fc_name)],value:s?.original_fc_name},{name:"reliever_name",label:"Nama Reliefer",type:"select",required:!0,options:b(s?.reliever_name),value:s?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Backup",type:"date",required:!0,value:s?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"select",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:s?.reason},{name:"shift",label:"Shift",type:"select",options:["Pagi","Siang","Full Shift","Middle"],value:s?.shift}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done","Tidak Datang"],value:s?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let s=await f("/api/relievers?limit=10000");if(s.ok){let c=s.data.data.map(r=>({"Tanggal Backup":r.backup_date||"",Cabang:r.branch_name||"","FC Digantikan":r.original_fc_name||"",Periode:r.period||"",Reliefer:r.reliever_name||"",Keterangan:r.reason||"",Shift:r.shift||"","Tanggal Selesai":r.completion_date||"",Status:r.status||""}));T(c,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Tanggal Backup":"2024-03-10",Cabang:"001. Pondok Bambu","FC Digantikan":"Budi Santoso",Periode:"Q1",Reliefer:"Andi",Keterangan:"Sakit",Shift:"Pagi","Tanggal Selesai":"2024-03-10",Status:"Done"}],"Template_Import_Reliefer")},onImport:async s=>{let r=(await f("/api/branches?all=1")).data?.data||[],i=p=>{if(!p)return null;let d=p.toLowerCase(),m=r.find(h=>h.full_name.toLowerCase()===d||h.code.toLowerCase()===d||h.name.toLowerCase()===d);return m?m.id:null},y=s.map(p=>({branch_id:i(String(p.Cabang||"").trim()),backup_date:String(p["Tanggal Backup"]||"").trim(),original_fc_name:String(p["FC Digantikan"]||"").trim(),reliever_name:String(p.Reliefer||"").trim(),period:String(p.Periode||"").trim(),reason:String(p.Keterangan||"").trim(),shift:String(p.Shift||"").trim(),completion_date:String(p["Tanggal Selesai"]||"").trim(),status:String(p.Status||"").trim()})).filter(p=>p.reliever_name&&p.backup_date),g=await f("/api/relievers/import",{method:"POST",body:JSON.stringify(y)});if(!g.ok)throw new Error(g.data?.error||"Import gagal")}}})}P();L();async function ka(a){let l=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name})),n=Array.from({length:4},(t,o)=>String(new Date().getFullYear()-o));$({container:a,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:t=>X(t)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"fc_score",label:"Point FC",render:t=>t!=null?`<strong class="${t>=80?"text-success":t>=60?"text-warning":"text-danger"}">${t}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:t=>t!=null?`<strong>${t}</strong>`:"-"},{key:"status",label:"Status",render:t=>A(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"select",name:"branch_id",label:"Cabang",options:l},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:l,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:t?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:t?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:t?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async t=>{let o=new URLSearchParams(t||{}).toString(),u=await f(`/api/reports/inspection?limit=10000&${o}`);if(u.ok){let b=u.data.data.map(s=>({Cabang:s.branch_name||"",Periode:s.period||"",Tanggal:s.inspection_date||"","Point FC":s.fc_score!==null&&s.fc_score!==void 0?s.fc_score:"","Point SPV":s.spv_score!==null&&s.spv_score!==void 0?s.spv_score:"",Status:s.status||"","Link Dokumen":s.document_link||"",Catatan:s.notes||""}));T(b,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:"Semua area bersih"}],"Template_Import_Inspeksi")},onImport:async t=>{let u=(await f("/api/branches?all=1")).data?.data||[],b=i=>{if(!i)return null;let y=i.toLowerCase(),g=u.find(p=>p.full_name.toLowerCase()===y||p.code.toLowerCase()===y||p.name.toLowerCase()===y);return g?g.id:null},s=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let y=String(i).trim();if(y===""||y==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);if(/^\d{4,5}$/.test(y)){let p=Number(y);if(p>2e4&&p<99999){let d=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let g=y.split(/[\/\-\.]/);if(g.length===3){let[p,d,m]=g.map(h=>h.trim());if(p.length===4&&d.length<=2&&m.length<=2)return`${p}-${d.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&d.length<=2&&p.length<=2)return`${m}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`}return y},c=t.map(i=>({branch_id:b(String(i.Cabang||"").trim()),period:String(i.Periode||"").trim(),inspection_date:s(i.Tanggal),fc_score:i["Point FC"]!==void 0&&i["Point FC"]!==""?Number(i["Point FC"]):null,spv_score:i["Point SPV"]!==void 0&&i["Point SPV"]!==""?Number(i["Point SPV"]):null,status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.period&&i.inspection_date),r=await f("/api/reports/inspection/import",{method:"POST",body:JSON.stringify(c)});if(!r.ok)throw new Error(r.data?.error||"Import gagal")}}})}P();L();async function Sa(a){let l=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name})),n=Array.from({length:4},(t,o)=>String(new Date().getFullYear()-o));$({container:a,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:t=>`<span class="badge ${t==="Deep Cleaning"?"badge-purple":"badge-success"}">${t}</span>`},{key:"period",label:"Periode",render:t=>X(t)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"status",label:"Status",render:t=>A(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:l},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:l,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:t?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:t?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async t=>{let o=new URLSearchParams(t||{}).toString(),u=await f(`/api/reports/cleaning?limit=10000&${o}`);if(u.ok){let b=u.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||"",Catatan:s.notes||""}));T(b,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:"Pembersihan lantai"}],"Template_Import_GCDC")},onImport:async t=>{let u=(await f("/api/branches?all=1")).data?.data||[],b=i=>{if(!i)return null;let y=i.toLowerCase(),g=u.find(p=>p.full_name.toLowerCase()===y||p.code.toLowerCase()===y||p.name.toLowerCase()===y);return g?g.id:null},s=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let y=String(i).trim();if(y===""||y==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);if(/^\d{4,5}$/.test(y)){let p=Number(y);if(p>2e4&&p<99999){let d=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let g=y.split(/[\/\-\.]/);if(g.length===3){let[p,d,m]=g.map(h=>h.trim());if(p.length===4&&d.length<=2&&m.length<=2)return`${p}-${d.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&d.length<=2&&p.length<=2)return`${m}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`}return y},c=t.map(i=>({branch_id:b(String(i.Cabang||"").trim()),activity_type:String(i.Jenis||i.Kegiatan||"").trim(),period:String(i.Periode||"").trim(),activity_date:s(i.Tanggal),status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.activity_type&&i.period&&i.activity_date),r=await f("/api/reports/cleaning/import",{method:"POST",body:JSON.stringify(c)});if(!r.ok)throw new Error(r.data?.error||"Import gagal")}}})}P();L();async function _a(a){let l=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name})),n=Array.from({length:4},(t,o)=>String(new Date().getFullYear()-o));$({container:a,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:t=>`<span class="badge badge-warning">${t}</span>`},{key:"period",label:"Periode",render:t=>X(t)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"status",label:"Status",render:t=>A(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:t=>t||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:l},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:n}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:l,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:t?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async t=>{let o=new URLSearchParams(t||{}).toString(),u=await f(`/api/reports/fogging?limit=10000&${o}`);if(u.ok){let b=u.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"Fogging",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||"",Catatan:s.notes||""}));T(b,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:"Tuntas"}],"Template_Import_Fogging")},onImport:async t=>{let u=(await f("/api/branches?all=1")).data?.data||[],b=i=>{if(!i)return null;let y=i.toLowerCase(),g=u.find(p=>p.full_name.toLowerCase()===y||p.code.toLowerCase()===y||p.name.toLowerCase()===y);return g?g.id:null},s=i=>{if(i==null||i==="")return"";if(i instanceof Date&&!isNaN(i.getTime()))return i.toISOString().slice(0,10);let y=String(i).trim();if(y===""||y==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);if(/^\d{4,5}$/.test(y)){let p=Number(y);if(p>2e4&&p<99999){let d=new Date(Date.UTC(1899,11,30)+p*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let g=y.split(/[\/\-\.]/);if(g.length===3){let[p,d,m]=g.map(h=>h.trim());if(p.length===4&&d.length<=2&&m.length<=2)return`${p}-${d.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&d.length<=2&&p.length<=2)return`${m}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`}return y},c=t.map(i=>({branch_id:b(String(i.Cabang||"").trim()),activity_type:String(i.Jenis||i.Kegiatan||"Fogging").trim(),period:String(i.Periode||"").trim(),activity_date:s(i.Tanggal),status:String(i.Status||"").trim(),document_link:String(i["Link Dokumen"]||"").trim(),notes:String(i.Catatan||i.Keterangan||"").trim()})).filter(i=>i.branch_id&&i.period&&i.activity_date),r=await f("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(c)});if(!r.ok)throw new Error(r.data?.error||"Import gagal")}}})}P();L();async function wa(a){let[e,l,n]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]),t=(e.data?.data||[]).map(c=>({value:c.id,label:c.full_name})),o=(l.data?.data||[]).map(c=>({value:c.full_name,label:c.full_name})),u=(n.data?.data||[]).filter(c=>c.role==="FC Spesialis").map(c=>({value:c.name,label:c.name})),b=c=>c&&!o.find(r=>r.value===c)?[...o,{value:c,label:c}]:o,s=c=>c&&!u.find(r=>r.value===c)?[...u,{value:c,label:c}]:u;$({container:a,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:c=>window.formatDate(c)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:c=>`<span title="${c||""}">${c?.length>60?c.slice(0,60)+"\u2026":c||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:c=>window.formatDate(c)},{key:"status",label:"Status",render:c=>A(c)},{key:"notes",label:"Keterangan",render:c=>c?.length>40?c.slice(0,40)+"\u2026":c||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan / PIC..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]}],formFields:c=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t,createApi:{path:"/api/branches",field:"full_name"},value:c?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:s(c?.pic),createApi:{path:"/api/pic",field:"name"},value:c?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:c?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:c?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:c?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:c?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:c?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async c=>{let r=new URLSearchParams(c||{}).toString(),i=await f(`/api/reports/basecamp?limit=10000&${r}`);if(i.ok){let y=i.data.data.map(g=>({"Tgl Info":g.info_date||"",Cabang:g.branch_name||"",Permasalahan:g.problem||"",PIC:g.pic||"","Tgl Done":g.done_date||"",Status:g.status||"",Keterangan:g.notes||""}));T(y,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async c=>{let i=(await f("/api/branches?all=1")).data?.data||[],y=m=>{if(!m)return null;let h=m.toLowerCase(),v=i.find(S=>S.full_name.toLowerCase()===h||S.code.toLowerCase()===h||S.name.toLowerCase()===h);return v?v.id:null},g=m=>{if(m==null||m==="")return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let h=String(m).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let S=Number(h);if(S>2e4&&S<99999){let k=new Date(Date.UTC(1899,11,30)+S*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}let v=h.split(/[\/\-\.]/);if(v.length===3){let[S,k,_]=v.map(E=>E.trim());if(S.length===4&&k.length<=2&&_.length<=2)return`${S}-${k.padStart(2,"0")}-${_.padStart(2,"0")}`;if(_.length===4&&k.length<=2&&S.length<=2)return`${_}-${k.padStart(2,"0")}-${S.padStart(2,"0")}`}return h},p=c.map(m=>({info_date:g(m["Tgl Info"]||m["Tanggal Info"]),branch_id:y(String(m.Cabang||"").trim()),problem:String(m.Permasalahan||"").trim(),pic:String(m.PIC||"").trim(),done_date:g(m["Tgl Done"]||m["Tanggal Done"]),status:String(m.Status||"").trim(),notes:String(m.Keterangan||m.Catatan||"").trim()})).filter(m=>m.info_date&&m.branch_id&&m.problem),d=await f("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(p)});if(!d.ok)throw new Error(d.data?.error||"Import gagal")}}})}async function Ca(a){$({container:a,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let l=new URLSearchParams(e||{}).toString(),{apiFetch:n}=await Promise.resolve().then(()=>(P(),me)),t=await n(`/api/sop?limit=10000&${l}`);if(t.ok){let o=t.data.data.map(b=>({"Nama SOP":b.name||"",Kategori:b.category||"",Dokumen:b.document_link||"",Catatan:b.notes||b.description||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(L(),J));u(o,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:l}=await Promise.resolve().then(()=>(L(),J));l(e,"Template_Import_SOP")},onImport:async e=>{let l=e.map(o=>({name:String(o["Nama SOP"]||"").trim(),category:String(o.Kategori||"").trim(),document_link:String(o.Dokumen||"").trim(),description:String(o.Catatan||"").trim()})).filter(o=>o.name),{apiFetch:n}=await Promise.resolve().then(()=>(P(),me)),t=await n("/api/sop/import",{method:"POST",body:JSON.stringify(l)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function xa(a){$({container:a,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let l=new URLSearchParams(e||{}).toString(),{apiFetch:n}=await Promise.resolve().then(()=>(P(),me)),t=await n(`/api/checklist?limit=10000&${l}`);if(t.ok){let o=t.data.data.map(b=>({"Nama Checklist":b.name||"",Kategori:b.category||"",Dokumen:b.document_link||"",Deskripsi:b.description||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(L(),J));u(o,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:l}=await Promise.resolve().then(()=>(L(),J));l(e,"Template_Import_Checklist")},onImport:async e=>{let l=e.map(o=>({name:String(o["Nama Checklist"]||"").trim(),category:String(o.Kategori||"").trim(),document_link:String(o.Dokumen||"").trim(),description:String(o.Deskripsi||"").trim()})).filter(o=>o.name),{apiFetch:n}=await Promise.resolve().then(()=>(P(),me)),t=await n("/api/checklist/import",{method:"POST",body:JSON.stringify(l)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}P();L();async function qe(a,e="forms"){if(e==="supply")return ut(a);mt(a)}function mt(a){$({container:a,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}]})}async function ut(a){let l=((await f("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name}));$({container:a,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:n=>n?new Date(n).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(n,t)=>t.branch_name_ref||t.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:n=>{try{let t=JSON.parse(n);return Array.isArray(t)?t.join(", "):n}catch{return n||"-"}}},{key:"chemical_items",label:"Chemical",render:n=>{try{let t=JSON.parse(n);return Array.isArray(t)?t.join(", "):n}catch{return n||"-"}}},{key:"additional_notes",label:"Catatan",render:n=>n?.length>40?n.slice(0,40)+"\u2026":n||"-"},{key:"status",label:"Status",render:n=>A(n)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:n=>{let t=n?.tools_items;try{t=Array.isArray(JSON.parse(t))?JSON.parse(t).join(", "):t}catch{}let o=n?.chemical_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:n?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:l,createApi:{path:"/api/branches",field:"full_name"},value:n?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:t},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:n?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:o},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:n?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:n?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:n?.status||""},{name:"processed_by",label:"Diproses Oleh",value:n?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async n=>{let t=new URLSearchParams(n||{}).toString(),o=await f(`/api/reports/supply?limit=10000&${t}`);if(o.ok){let u=o.data.data.map(b=>{let s=b.tools_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}let c=b.chemical_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}return{Waktu:b.submitted_at||"",Pengirim:b.submitter_name||"",Cabang:b.branch_name_ref||b.branch_name||"","Alat/Barang":s||"",Chemical:c||"",Catatan:b.additional_notes||"",Status:b.status||"","Diproses Oleh":b.processed_by||""}});T(u,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async n=>{let o=(await f("/api/branches?all=1")).data?.data||[],u=r=>{if(!r)return null;let i=r.toLowerCase(),y=o.find(g=>g.full_name.toLowerCase()===i||g.code.toLowerCase()===i||g.name.toLowerCase()===i);return y?y.id:null},b=r=>{if(r==null||r==="")return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let i=String(r).trim();if(i===""||i==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(i))return i.slice(0,10);if(/^\d{4,5}$/.test(i)){let g=Number(i);if(g>2e4&&g<99999){let p=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}let y=i.split(/[\/\-\.]/);if(y.length===3){let[g,p,d]=y.map(m=>m.trim());if(g.length===4&&p.length<=2&&d.length<=2)return`${g}-${p.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&p.length<=2&&g.length<=2)return`${d}-${p.padStart(2,"0")}-${g.padStart(2,"0")}`}return i},s=n.map(r=>({submitted_at:b(r.Waktu||r.Tanggal),submitter_name:String(r.Pengirim||"").trim(),branch_id:u(String(r.Cabang||"").trim()),tools_items:String(r["Alat/Barang"]||r.Alat||"").trim(),chemical_items:String(r.Chemical||"").trim(),additional_notes:String(r.Catatan||r.Keterangan||"").trim(),status:String(r.Status||"").trim(),processed_by:String(r["Diproses Oleh"]||r.PIC||"").trim()})).filter(r=>r.submitted_at&&r.submitter_name&&r.branch_id),c=await f("/api/reports/supply/import",{method:"POST",body:JSON.stringify(s)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(n,t)=>{let o=ee({title:"Update Status Permintaan",content:`
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
    </div>`,await n();try{let u=e.getFullYear(),b=e.getMonth(),s=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),c=document.getElementById("cal-month-label");c&&(c.textContent=s);let r=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(k=>k.value)),i=l.filter(k=>r.has(k.type)),y={};i.forEach(k=>{let _=(k.event_date||"").slice(0,10);y[_]||(y[_]=[]),y[_].push(k)});let g=new Date(u,b,1).getDay(),p=new Date(u,b+1,0).getDate(),d=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],m=new Date().toISOString().slice(0,10),h='<div class="calendar-grid">';d.forEach(k=>{h+=`<div class="cal-day-header">${k}</div>`});for(let k=0;k<g;k++)h+='<div class="cal-cell cal-cell-empty"></div>';for(let k=1;k<=p;k++){let _=`${u}-${String(b+1).padStart(2,"0")}-${String(k).padStart(2,"0")}`,E=y[_]||[],O=_===m;h+=`
          <div class="cal-cell ${O?"cal-today":""} ${E.length?"cal-has-events":""}"
               data-date="${_}" tabindex="0" role="button" aria-label="${_}">
            <div class="cal-day-num ${O?"today-num":""}">${k}</div>
            <div class="cal-events-preview">
              ${E.slice(0,3).map(w=>`
                <div class="cal-event-dot cal-color-${w.color||"gray"}" title="${xe(w.title||w.type)}">
                  <span class="cal-event-dot-label">${gt(w.title||w.branch_name||w.type,18)}</span>
                </div>
              `).join("")}
              ${E.length>3?`<div class="cal-more">+${E.length-3} lagi</div>`:""}
            </div>
          </div>`}let S=(g+p)%7;if(S!==0)for(let k=0;k<7-S;k++)h+='<div class="cal-cell cal-cell-empty"></div>';h+="</div>",o.innerHTML=h,o.querySelectorAll(".cal-cell[data-date]").forEach(k=>{k.addEventListener("click",()=>{let _=k.dataset.date,E=y[_]||[];if(!E.length)return;let O=document.getElementById("cal-event-list"),w=new Date(_+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=w,document.getElementById("cal-event-items").innerHTML=E.map(D=>`
            <div class="cal-event-item cal-color-border-${D.color||"gray"}">
              <div class="cal-event-type">${bt(D.type)}</div>
              <div class="cal-event-title">${xe(D.title||"-")}</div>
              <div class="cal-event-branch">${xe(D.branch_name||"")}</div>
              ${D.status?`<div class="cal-event-status">${xe(D.status)}</div>`:""}
              ${D.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${D.days_remaining} hari</div>`:""}
            </div>
          `).join(""),O.style.display="block"})})}catch(u){console.error("[Calendar] Render error:",u),o&&(o.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}t()}function gt(a,e){return a?a.length>e?a.slice(0,e)+"\u2026":a:""}function xe(a){return a?String(a).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function bt(a){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[a]||a}P();async function Pa(a){let e=ie(),l=(e?.full_name||e?.username||"U")[0].toUpperCase(),t={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";a.innerHTML=`
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
  `;let o=localStorage.getItem("fm_token"),u=document.getElementById("session-info");if(o&&u)try{let b=JSON.parse(atob(o.split(".")[1])),s=new Date(b.exp*1e3);u.textContent=`Berakhir: ${s.toLocaleString("id-ID")}`}catch{u.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async b=>{b.preventDefault();let s=document.getElementById("pwd-error"),c=document.getElementById("pwd-success"),r=document.getElementById("btn-save-pwd");s.style.display="none",c.style.display="none";let i=b.target,y=i.current_password.value,g=i.new_password.value,p=i.confirm_password.value;if(g!==p){s.textContent="\u274C Konfirmasi password tidak cocok.",s.style.display="block";return}if(g.length<6){s.textContent="\u274C Password baru minimal 6 karakter.",s.style.display="block";return}r.disabled=!0,r.textContent="\u23F3 Menyimpan...";let d=await f("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:y,new_password:g})});r.disabled=!1,r.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',d.ok?(c.textContent="\u2705 Password berhasil diubah.",c.style.display="block",i.reset(),K("Password berhasil diubah.")):(s.textContent=d.data?.error||"Gagal mengubah password.",s.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}P();var Te={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function M(a){if(a==null||a==="")return null;if(a instanceof Date)return isNaN(a.getTime())?null:a.toISOString().slice(0,10);let e=String(a).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let t=Number(e);if(t>2e4&&t<99999){let o=new Date(Date.UTC(1899,11,30)+t*864e5);return isNaN(o.getTime())?null:o.toISOString().slice(0,10)}}let l=e.split(/[\/\-\.]/);if(l.length===3){let[t,o,u]=l.map(r=>r.trim()),b=Number(t),s=Number(o),c=Number(u);if(t.length===4&&b>1900)return`${t}-${o.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&c>1900)return b>12?`${u}-${o.padStart(2,"0")}-${t.padStart(2,"0")}`:s>12?`${u}-${t.padStart(2,"0")}-${o.padStart(2,"0")}`:`${u}-${o.padStart(2,"0")}-${t.padStart(2,"0")}`;if(u.length===2&&!isNaN(c)){let r=c>=50?`19${u}`:`20${u}`;return b>12?`${r}-${o.padStart(2,"0")}-${t.padStart(2,"0")}`:`${r}-${o.padStart(2,"0")}-${t.padStart(2,"0")}`}}let n=new Date(e);return isNaN(n.getTime())?null:n.toISOString().slice(0,10)}function Da(a){return Object.values(a).every(e=>e==null||String(e).trim()==="")}var ht={validation:{required:[],map:a=>({cabang:a.CABANG,pic:a.PIC,kegiatan:a.KEGIATAN,quartal:a.QUARTAL,masa_pkwt:a["MASA PKWT"],pic_pelapor:a["PIC PELAPOR"],kontrak:a.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:a=>({name:a["Nama SOP"],category:a.Kategori||"Umum",document_link:a["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:a=>({full_name:a["Nama Lengkap"],branch_name:a.Cabang,division:a["Div / Bagian"]||"FACILITY CARE",phone:a["No. Hp"],join_date:M(a["Tanggal Masuk"]),status:a.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:a=>({employee_name:a["Nama Lengkap"],branch_name:a.Cabang,division:a["Div / Bagian"]||"FACILITY CARE",start_date:M(a["Tanggal Mulai"]),end_date:M(a["Tanggal Selesai"]),contract_type:a["Tipe Kontrak"]||"",pkwt_number:a.PKWT||"",status:a.Status||"",notes:a.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:a=>({report_date:M(a["Tanggal Info"]),branch_name:a.Cabang,category:a.Kategori,source:a["Sumber Laporan"],complaint:a.Keluhan,employee_name:a["Nama FC"],fc_specialist:a["FC Spesialis"],solution:a.Solusi,status:a.Status||"",completion_date:M(a["Tanggal Selesai"])})},one_on_one:{required:[],map:a=>({meeting_date:M(a.Tanggal),branch_name:a.Cabang,employee_name:a["Nama Karyawan"],pic:a.Pic,problem:a.Masalah,solution:a.Solusi,status:a.Status||"",completion_date:M(a["Tanggal Selesai"]),document_link:a["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:a=>({branch_name:a.Cabang,activity_type:a.Kegiatan,period:a.Periode,pic:a.Pic||a.PIC,opening_date:M(a["Tanggal Opening"]||a["Tgl Opening"]),target_date:M(a["Tanggal Target"]||a["Tgl Target"]),completion_date:M(a["Tanggal Selesai"]||a["Tgl Selesai"]),status:a.Status||"",notes:a.Keterangan||a.Catatan})},inspection:{required:[],map:a=>({inspection_date:M(a.Tanggal),branch_name:a.Cabang,period:a.Periode,status:a.Status||"",fc_score:a["Point FC SP"]!==void 0&&a["Point FC SP"]!==null?parseFloat(String(a["Point FC SP"]).replace(",",".")):null,spv_score:a["Point SPV"]!==void 0&&a["Point SPV"]!==null?parseFloat(String(a["Point SPV"]).replace(",",".")):null,document_link:a.Link,notes:""})},cleaning:{required:[],map:a=>({activity_date:M(a.Tanggal),branch_name:a.Cabang,activity_type:a["Jenis Kegiatan"]||"General Cleaning",period:a.Periode,status:a.Status||"",document_link:a.Link,notes:""})},fogging:{required:[],map:a=>({activity_date:M(a.Tanggal),branch_name:a.Cabang,period:a.Periode,status:a.Status||"",document_link:a.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:a=>({info_date:M(a["Tgl Info"]),branch_name:a.Cabang,problem:a.Permasalahan,pic:a.PIC,done_date:M(a["Tgl Done"]),status:a.Status||"",notes:a.Ket})},relievers:{required:[],map:a=>({branch_name:a.Cabang,original_fc_name:a["Nama Facility care"],period:a.Periode,reliever_name:a.Relifer,backup_date:M(a["Tanggal Back Up"]),completion_date:M(a["Tanggal Selesai"]),reason:a.Keterangan,shift:a.Shift,status:a.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:a=>({training_date:M(a.Tanggal),batch:a.Batch,subject:a.Materi,participants:a.Peserta,branch_name:a.Cabang,trainer:a.Trainer,score:a.Nilai!==void 0&&a.Nilai!==null?parseFloat(String(a.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:a=>({name:a["Master Checklist"],category:"Umum",document_link:a["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:a=>({name:a["Master Form"],category:"Umum",document_link:a["Link Document"],description:""})},supply:{required:[],map:a=>({submitted_at:M(a.Timestamp),submitter_name:a["Nama Lengkap"],branch_name:a["Kebutuhan Untuk Cabang"],tools_items:a["Alat - Alat / Barang"],tools_quantity:a["Jumlah Permintaan Alat / Barang"],chemical_items:a.Chemical,chemical_quantity:a["Jumlah Permintaan Chemical"],additional_notes:a["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:a.Status||""})}};function yt(a,e){let l=Te[a];if(!l)return{valid:[],errors:[],mapped:[],skipped:!0};let n=ht[l.module];if(!n)return{valid:[],errors:[],mapped:[],skipped:!0};let t=[],o=[],u=[];return e.filter(s=>!Da(s)).forEach((s,c)=>{let r=e.indexOf(s)+2,i=[];n.required.forEach(({key:g,label:p})=>{let d=s[g];if(d==null||String(d).trim()===""){let m=Object.keys(s).filter(h=>h.trim()).join(", ");i.push({column:p,originalValue:d||"",reason:`Kolom "${p}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${m.slice(0,120)}`})}});let y=n.map(s);i.length>0?o.push({row:r,data:y,raw:s,errors:i}):(t.push(s),u.push(y))}),{valid:t,errors:o,mapped:u}}function Ia(a){let e=[];return a.SheetNames.forEach(l=>{let n=Te[l];if(!n)return;let t=a.Sheets[l],o=window.XLSX.utils.sheet_to_json(t,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),u=yt(l,o),b=o.filter(s=>!Da(s));e.push({sheetName:l,module:n.module,label:n.label,total:b.length,valid:u.mapped.length,errorCount:u.errors.length,errors:u.errors,mapped:u.mapped,skipped:!1})}),e}function La(){let a=window.XLSX,e=a.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu","NAMA KARYAWAN":"Budi Santoso",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02",Day:"1"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06",Day:"1","Link Document":""}]}).forEach(([n,t])=>{a.utils.book_append_sheet(e,a.utils.json_to_sheet(t),n)}),a.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function Ba(a){let e=window.XLSX,l=e.utils.book_new(),n=!1;return a.forEach(t=>{if(!t.errors||t.errors.length===0)return;n=!0;let o=t.errors.map(b=>({"No. Baris":b.row,"Kolom Gagal":(b.errors||[]).map(s=>s.column||s).join("; "),"Alasan Error":(b.errors||[]).map(s=>s.reason||s).join("; "),...Object.fromEntries(Object.entries(b.data||{}).map(([s,c])=>[s,c??""]))})),u=e.utils.json_to_sheet(o);e.utils.book_append_sheet(l,u,t.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),n?(e.writeFile(l,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var ft=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function Na(a){a.innerHTML=`
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
              ${Object.entries(Te).map(([d,{label:m}])=>`<span class="import-sheet-tag">\u{1F4C4} ${d} \u2192 ${m}</span>`).join("")}
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
  `;let e=null,l=null,n=0,t={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function o(d){Object.entries(t).forEach(([m,h])=>{h.style.display=m===d?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let d=document.getElementById("btn-backup-db");d.disabled=!0,d.textContent="\u23F3 Memproses Backup...";try{let m=await f("/api/import/backup");if(m.ok){let h=new Blob([JSON.stringify(m.data,null,2)],{type:"application/json"}),v=URL.createObjectURL(h),S=document.createElement("a");S.href=v,S.download=`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.json`,document.body.appendChild(S),S.click(),document.body.removeChild(S),URL.revokeObjectURL(v),K("Backup berhasil diunduh!")}else j("Gagal memproses backup: "+(m.data?.error||"Unknown error"))}catch(m){j("Gagal memproses backup: "+m.message)}finally{d.disabled=!1,d.textContent="\u{1F4E6} Backup Database"}});let u=document.getElementById("btn-sync-google");u&&u.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let d=u.innerHTML;u.innerHTML='<span class="spinner"></span> Menyinkronkan...',u.disabled=!0;try{let m=await f("/api/sync/google-sheets",{method:"POST"});m.ok?alert("Sinkronisasi Berhasil: "+(m.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(m.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{u.innerHTML=d,u.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{La(),K("Template Excel berhasil didownload!")});let b=document.getElementById("file-input"),s=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",d=>{d.stopPropagation(),b.click()}),b.addEventListener("change",d=>{d.target.files[0]&&c(d.target.files[0])}),s.addEventListener("dragover",d=>{d.preventDefault(),s.classList.add("drag-over")}),s.addEventListener("dragleave",()=>s.classList.remove("drag-over")),s.addEventListener("drop",d=>{d.preventDefault(),s.classList.remove("drag-over");let m=d.dataTransfer.files[0];m&&m.name.match(/\.xlsx?$/i)?c(m):j("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,b.value="",document.getElementById("file-info").style.display="none",s.style.display="",o("upload")});async function c(d){e=d,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${d.name} (${(d.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",s.style.display="none",await r(d)}async function r(d){o("validating");let m=document.getElementById("validation-status"),h=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");m.textContent="Membaca file Excel...",h.style.width="20%",await be(200);let v=await d.arrayBuffer(),S=window.XLSX.read(v,{type:"array",cellDates:!0});m.textContent=`Memvalidasi ${S.SheetNames.length} sheet...`,h.style.width="50%",await be(100),l=Ia(S),h.style.width="100%",m.textContent="Validasi selesai!",await be(300),i()}catch(v){o("upload"),j("Gagal memproses file: "+v.message),document.getElementById("file-info").style.display="flex",s.style.display="none"}}function i(){o("preview");let d=l.filter(w=>!w.skipped).length,m=l.reduce((w,D)=>w+D.total,0),h=l.reduce((w,D)=>w+D.valid,0),v=l.reduce((w,D)=>w+D.errorCount,0),S=m>0?Math.round(h/m*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${d} sheet</span>
      <span class="badge badge-secondary">${m} baris</span>
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
          ${l.map((w,D)=>`
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
                ${w.errorCount>0?`<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${D}">\u{1F50D} ${w.errorCount} Error</button>`:'<span class="text-muted">\u2013</span>'}
              </td>
            </tr>
          `).join("")}
        </tbody>
      </table>
    `,k.querySelectorAll(".btn-detail-error").forEach(w=>{w.addEventListener("click",()=>{let D=l[Number(w.dataset.idx)];y(D)})});let _=document.getElementById("error-detail-section"),E=document.getElementById("error-detail-container");E.innerHTML="",_.style.display="none";let O=document.getElementById("btn-start-import");h===0?(O.disabled=!0,O.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(O.disabled=!1,v>0?(O.innerHTML=`\u{1F680} Import ${h} Data Valid (${v} dilewati)`,O.title="Baris error akan dilewati, baris valid tetap diimport"):O.innerHTML=`\u{1F680} Mulai Import ${h} Data`)}function y(d){let m=document.getElementById("error-detail-section"),h=document.getElementById("error-detail-container");m.style.display="";let v=d.errors.slice(0,100).map(S=>(Array.isArray(S.errors)?S.errors:[]).map(_=>{let E=typeof _=="object";return`
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
    `,m.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{o("upload"),document.getElementById("file-info").style.display="none",s.style.display="",e=null,b.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!l)return;Ba(l)?K("Log error berhasil didownload."):K("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let d=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";g(d)});async function g(d){o("importing"),n=Date.now();let m=[];ft.forEach(_=>{let E=l?.find(O=>O.module===_&&O.mapped?.length>0);E&&m.push(E)});let h=document.getElementById("import-steps-list");h.innerHTML=m.map(_=>`
      <div class="import-step-item" id="step-item-${_.module}">
        <span class="step-item-icon" id="step-icon-${_.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${_.label} <span class="step-item-count">(${_.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${_.module}"></span>
      </div>
    `).join("");let v=document.getElementById("import-bar"),S=document.getElementById("import-current-status"),k={totalSheets:m.length,totalRows:m.reduce((_,E)=>_+E.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let _=0;_<m.length;_++){let E=m[_],O=document.getElementById(`step-icon-${E.module}`),w=document.getElementById(`step-status-${E.module}`);O.textContent="\u{1F504}",w.textContent="Mengimport...",S.textContent=`Mengimport ${E.label}...`,v.style.width=`${Math.round(_/m.length*100)}%`;try{let D=await f(`/api/import/${E.module}`,{method:"POST",body:JSON.stringify({rows:E.mapped,onDuplicate:d})});if(D.ok){let V=D.data;k.inserted+=V.inserted||0,k.skipped+=V.skipped||0,k.moduleResults.push({label:E.label,inserted:V.inserted||0,skipped:V.skipped||0,status:"ok"}),O.textContent="\u2705",w.innerHTML=`<span class="badge badge-success">${V.inserted||0} berhasil</span>${V.skipped>0?` <span class="badge badge-neutral">${V.skipped} skip</span>`:""}`}else k.failed++,k.moduleResults.push({label:E.label,inserted:0,skipped:0,status:"error",error:D.data?.error}),O.textContent="\u274C",w.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(D){k.failed++,k.moduleResults.push({label:E.label,inserted:0,skipped:0,status:"error",error:D.message}),O.textContent="\u274C",w.innerHTML='<span class="badge badge-danger">Gagal</span>'}await be(150)}v.style.width="100%",S.textContent="Selesai!",await be(400),p(k)}function p(d){o("summary");let m=((Date.now()-n)/1e3).toFixed(1),h=d.failed===0;document.getElementById("summary-status-icon").innerHTML=`
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
        <div class="stat-value">${m}s</div>
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,l=null,b.value="",document.getElementById("file-info").style.display="none",s.style.display="",o("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function be(a){return new Promise(e=>setTimeout(e,a))}P();var je=[],vt=[];async function Oa(a){let[e,l]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000")]);je=(e.data?.data||[]).map(n=>({value:n.id,label:n.full_name})),vt=(l.data?.data||[]).map(n=>({value:n.full_name,label:n.full_name})),$({container:a,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"division",label:"Divisi",render:n=>n?`<span class="badge badge-info">${n}</span>`:"-"},{key:"branch_name",label:"Cabang"},{key:"tanggal",label:"Tanggal Sp",render:n=>n?new Date(n).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"akhir_sp",label:"Akhir Sp",render:n=>n?new Date(n).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"sp_type",label:"Jenis Sp",render:n=>`<span class="badge badge-warning">${n||"-"}</span>`},{key:"document_link",label:"Link Document / Foto",render:n=>n?`<a href="${n}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:je}],exportOptions:{moduleName:"sp_data",onExport:async n=>{let t=new URLSearchParams(n||{}).toString(),o=await f(`/api/sp?limit=10000&${t}`);if(o.ok){let u=o.data.data.map(s=>({"Nama Karyawan":s.employee_name||"",Divisi:s.division||"",Cabang:s.branch_name||"","Tanggal Sp":s.tanggal||"","Akhir Sp":s.akhir_sp||"","Jenis Sp":s.sp_type||"","Link Document / Foto":s.document_link||""})),{downloadExcel:b}=await Promise.resolve().then(()=>(L(),J));b(u,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{"Nama Karyawan":"Budi Santoso",Divisi:"FACILITY CARE",Cabang:"001. Pondok Bambu","Tanggal Sp":"2026-01-08","Akhir Sp":"2026-07-08","Jenis Sp":"SP 1","Link Document / Foto":"https://link.doc"}],{downloadExcel:t}=await Promise.resolve().then(()=>(L(),J));t(n,"Template_Import_SP")},onImport:async n=>{let o=(await f("/api/branches?all=1")).data?.data||[],u=r=>{if(!r)return null;let i=r.toLowerCase(),y=o.find(g=>g.full_name.toLowerCase()===i||g.code.toLowerCase()===i||g.name.toLowerCase()===i);return y?y.id:null},b=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let i=String(r).trim();if(/^\d{4,5}$/.test(i)){let g=Number(i);if(g>2e4&&g<99999){let p=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(i))return i.slice(0,10);let y=i.split(/[\/\-\.]/);if(y.length===3){let[g,p,d]=y.map(m=>m.trim());if(g.length===4&&p.length<=2&&d.length<=2)return`${g}-${p.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&p.length<=2&&g.length<=2)return`${d}-${p.padStart(2,"0")}-${g.padStart(2,"0")}`}return i},s=n.map(r=>({employee_name:String(r["Nama Karyawan"]||"").trim(),division:String(r.Divisi||"").trim(),branch_id:u(String(r.Cabang||"").trim()),tanggal:b(r["Tanggal Sp"]),akhir_sp:b(r["Akhir Sp"]),sp_type:String(r["Jenis Sp"]||"").trim(),document_link:String(r["Link Document / Foto"]||"").trim()})).filter(r=>r.employee_name&&r.branch_id),c=await f("/api/sp/import",{method:"POST",body:JSON.stringify(s)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}},formFields:[{type:"text",name:"employee_name",label:"Nama Karyawan",required:!0},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"],required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:je},{type:"date",name:"tanggal",label:"Tanggal Sp",required:!0},{type:"date",name:"akhir_sp",label:"Akhir Sp",required:!0},{type:"select",name:"sp_type",label:"Jenis Sp",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"url",name:"document_link",label:"Link Document / Foto"}]})}P();var he=[],kt=[];async function Aa(a){let[e,l]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000")]);he=(e.data?.data||[]).map(n=>({value:n.id,label:n.full_name})),kt=(l.data?.data||[]).map(n=>({value:n.full_name,label:n.full_name})),$({container:a,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:n=>n?new Date(n).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:n=>`<span class="badge ${n==="Selesai"?"badge-success":"badge-warning"}">${n||"-"}</span>`},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:he},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:he}],exportOptions:{moduleName:"mutasi_data",onExport:async n=>{let t=new URLSearchParams(n||{}).toString(),o=await f(`/api/mutasi?limit=10000&${t}`);if(o.ok){let u=o.data.data.map(s=>({Tanggal:s.tanggal||"","Nama Karyawan":s.employee_name||"","Cabang Asal":s.from_branch_name||"","Cabang Tujuan":s.to_branch_name||"",Status:s.status||"",Dokumen:s.document_link||""})),{downloadExcel:b}=await Promise.resolve().then(()=>(L(),J));b(u,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:t}=await Promise.resolve().then(()=>(L(),J));t(n,"Template_Import_Mutasi")},onImport:async n=>{let o=(await f("/api/branches?all=1")).data?.data||[],u=r=>{if(!r)return null;let i=r.toLowerCase(),y=o.find(g=>g.full_name.toLowerCase()===i||g.code.toLowerCase()===i||g.name.toLowerCase()===i);return y?y.id:null},b=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let i=String(r).trim();if(/^\d{4,5}$/.test(i)){let g=Number(i);if(g>2e4&&g<99999){let p=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(p.getTime())?"":p.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(i))return i.slice(0,10);let y=i.split(/[\/\-\.]/);if(y.length===3){let[g,p,d]=y.map(m=>m.trim());if(g.length===4&&p.length<=2&&d.length<=2)return`${g}-${p.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&p.length<=2&&g.length<=2)return`${d}-${p.padStart(2,"0")}-${g.padStart(2,"0")}`}return i},s=n.map(r=>({tanggal:b(r.Tanggal),employee_name:String(r["Nama Karyawan"]||"").trim(),from_branch_id:u(String(r["Cabang Asal"]||"").trim()),to_branch_id:u(String(r["Cabang Tujuan"]||"").trim()),status:String(r.Status||"").trim(),document_link:String(r.Dokumen||"").trim()})).filter(r=>r.tanggal&&r.employee_name&&r.from_branch_id&&r.to_branch_id),c=await f("/api/mutasi/import",{method:"POST",body:JSON.stringify(s)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"text",name:"employee_name",label:"Nama Karyawan",required:!0},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:he,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:he,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}window.formatDate=a=>{if(!a||a==="-")return"";let e=a.split("-");return e.length===3&&e[0].length===4?`${e[2]}-${e[1]}-${e[0]}`:a};function N(a){return async e=>{if(!ce()){se("/login");return}return a(e)}}var ye=null;function St(){ye&&clearInterval(ye);let a=()=>{let e=new Date,l=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),n=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),t=document.getElementById("header-clock-time"),o=document.getElementById("header-clock-date");t&&(t.textContent=l),o&&(o.textContent=n)};a(),ye=setInterval(a,1e3)}async function _t(){try{let a=await f("/api/dashboard/kpi");if(!a.ok)return;let e=a.data?.data||a.data||{},l=(n,t)=>{let o=document.getElementById(n);o&&(o.textContent=t>0?t:"",o.style.display=t>0?"inline-flex":"none")};l("badge-issues",e.issues?.current||0),l("badge-contracts",e.expiring30?.current||0),l("badge-oo1",e.one_on_one?.current||0),l("badge-schedule",e.schedule?.current||0),l("badge-supply",e.supply?.current||0)}catch{}}var le=[];async function wt(){try{let a=await f("/api/dashboard/notifications");if(!a.ok)return;le=a.data?.data||a.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=le.length>0?"block":"none",e.textContent=le.length)}catch{}}function Ct(){if(!le.length){ee({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,l)=>l()});return}let a=`
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
  `;ee({title:`Notifikasi (${le.length})`,content:a,confirmText:"Tutup",onConfirm:(e,l)=>l()})}function Ma(){let a=ie(),e=(a?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
  `;let l=document.getElementById("sidebar"),n=document.getElementById("sidebar-overlay"),t=document.getElementById("topbar-menu-btn"),o=document.getElementById("sidebar-close"),u=()=>{l.classList.add("open"),n.classList.add("show")},b=()=>{l.classList.remove("open"),n.classList.remove("show")};t?.addEventListener("click",u),o?.addEventListener("click",b),n?.addEventListener("click",b),document.querySelectorAll(".nav-item").forEach(c=>c.addEventListener("click",b));function s(){let c=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(y=>{let g=y.dataset.route;y.classList.toggle("active",c===g||g!=="/dashboard"&&c.startsWith(g))});let r=document.getElementById("topbar-title"),i=document.querySelector(".nav-item.active .nav-label");r&&i&&(r.textContent=i.textContent)}window.addEventListener("hashchange",s),s(),St(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await f("/api/auth/logout",{method:"POST"}),de(),ye&&clearInterval(ye),se("/login")}),_t(),wt(),document.getElementById("btn-notif")?.addEventListener("click",c=>{c.preventDefault(),Ct()})}async function xt(){B("/login",({main:e})=>sa(e)),B("/dashboard",N(({main:e})=>ta(e))),B("/calendar",N(({main:e})=>Ea(e))),B("/employees",N(({main:e})=>ma(e))),B("/contracts",N(({main:e})=>ga(e))),B("/sp",N(({main:e})=>Oa(e))),B("/mutasi",N(({main:e})=>Aa(e))),B("/timeline",N(({main:e})=>ba(e))),B("/issues",N(({main:e})=>ha(e))),B("/one-on-one",N(({main:e})=>ya(e))),B("/training",N(({main:e})=>fa(e))),B("/relievers",N(({main:e})=>va(e))),B("/reports/inspection",N(({main:e})=>ka(e))),B("/reports/cleaning",N(({main:e})=>Sa(e))),B("/reports/fogging",N(({main:e})=>_a(e))),B("/reports/basecamp",N(({main:e})=>wa(e))),B("/reports/supply",N(({main:e})=>qe(e,"supply"))),B("/sop",N(({main:e})=>Ca(e))),B("/checklist",N(({main:e})=>xa(e))),B("/forms",N(({main:e})=>qe(e))),B("/users",N(({main:e})=>Ta(e))),B("/branches",N(({main:e})=>$a(e))),B("/profile",N(({main:e})=>Pa(e))),B("/settings/import",N(({main:e})=>Na(e)));let a=ce();if(!a&&window.location.hash!=="#/login"&&se("/login"),a){let e=await f("/api/auth/me");e.ok?(pe(e.data.data),Ma()):(de(),se("/login"))}window.addEventListener("fm:login",()=>{Ma(),se("/dashboard")}),Ve()}xt();
