var Da=Object.defineProperty;var Fe=(a,e)=>()=>(a&&(e=a(a=0)),e);var Re=(a,e)=>{for(var n in e)Da(a,n,{get:e[n],enumerable:!0})};var de={};Re(de,{API:()=>Ke,apiFetch:()=>f,clearToken:()=>le,getToken:()=>re,getUser:()=>ae,setToken:()=>_e,setUser:()=>oe});function re(){return localStorage.getItem("fm_token")}function _e(a){localStorage.setItem("fm_token",a)}function le(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ae(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function oe(a){localStorage.setItem("fm_user",JSON.stringify(a))}async function f(a,e={}){let n=re(),i={"Content-Type":"application/json",...n?{Authorization:`Bearer ${n}`}:{},...e.headers||{}};try{let t=`cb=${Date.now()}`,d=a.includes("?")?"&":"?",u=`${Ke}${a}${d}${t}`,p=await fetch(u,{...e,headers:i}),s;try{let r=await p.text();try{s=JSON.parse(r)}catch{s={error:`Server Error (${p.status}): ${r.substring(0,80)}...`}}}catch{s={error:"Gagal membaca respon dari server"}}return p.status===401&&(le(),window.location.hash="#/login"),{ok:p.ok,status:p.status,data:s}}catch(t){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${t.message})`}}}}var Ba,Ke,P=Fe(()=>{Ba="",Ke=Ba});var U={};Re(U,{downloadExcel:()=>T,parseExcel:()=>Ee,renderExcelButtons:()=>$e});function Ee(a){return new Promise((e,n)=>{let i=new FileReader;i.onload=t=>{try{let d=new Uint8Array(t.target.result),u=XLSX.read(d,{type:"array"}),p=u.SheetNames[0],s=u.Sheets[p],r=XLSX.utils.sheet_to_json(s,{defval:""});e(r)}catch(d){n(d)}},i.onerror=t=>n(t),i.readAsArrayBuffer(a)})}function T(a,e){try{let n=XLSX.utils.json_to_sheet(a),i=XLSX.utils.book_new();XLSX.utils.book_append_sheet(i,n,"Data"),XLSX.writeFile(i,`${e}.xlsx`)}catch(n){throw console.error("Error generating Excel file:",n),n}}function $e(a){return`
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
  `}var D=Fe(()=>{});P();var we={},be=null;function B(a,e){we[a]=e}function te(a){window.location.hash=a}function qe(){async function a(){let e=window.location.hash.replace("#","")||"/dashboard",[n,...i]=e.split("?"),t=we[n];if(!t){for(let[u,p]of Object.entries(we))if(u.endsWith("/*")&&n.startsWith(u.slice(0,-2))){t=p;break}}be&&(be(),be=null);let d=document.getElementById("main-content");if(d&&(d.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),t){let u=new URLSearchParams(i.join("?")),p=n.split("/").filter(Boolean),s=await t({path:n,params:u,segments:p,main:d});s&&(be=s)}else{let u=d||document.getElementById("app");u&&(u.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",a),a()}var ce;function Na(){return ce||(ce=document.createElement("div"),ce.id="toast-container",document.body.appendChild(ce)),ce}function He(a,e="info",n=3500){let i=Na(),t=document.createElement("div");t.className=`toast toast-${e}`;let d={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};t.innerHTML=`<span class="toast-icon">${d[e]||"\u2139"}</span><span class="toast-msg">${a}</span>`,i.appendChild(t),requestAnimationFrame(()=>t.classList.add("show")),setTimeout(()=>{t.classList.remove("show"),setTimeout(()=>t.remove(),350)},n)}var R=a=>He(a,"success"),H=a=>He(a,"error");function Y({title:a,content:e,onConfirm:n,onCancel:i,confirmText:t="Simpan",cancelText:d="Batal",size:u="md",confirmClass:p="btn-primary"}){let s={sm:"400px",md:"560px",lg:"720px",xl:"900px"},r=document.createElement("div");r.className="modal-overlay",r.innerHTML=`
    <div class="modal" style="max-width:${s[u]||s.md}">
      <div class="modal-header">
        <h3 class="modal-title">${a}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${d}</button>
        ${n?`<button class="btn ${p} modal-confirm">${t}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&r.querySelector(".modal-body").appendChild(e);let o=()=>{r.classList.remove("show"),setTimeout(()=>r.remove(),250)};return r.querySelector(".modal-close").addEventListener("click",()=>{i&&i(),o()}),r.querySelector(".modal-cancel").addEventListener("click",()=>{i&&i(),o()}),n&&r.querySelector(".modal-confirm").addEventListener("click",()=>n(r,o)),r.addEventListener("click",l=>{l.target===r&&(i&&i(),o())}),document.body.appendChild(r),requestAnimationFrame(()=>r.classList.add("show")),{overlay:r,close:o}}function je(a,e,n="Konfirmasi"){return Y({title:n,content:`<p>${a}</p>`,onConfirm:(i,t)=>{e(),t()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}P();var ee={};function pe(a){if(ee[a]){try{ee[a].destroy()}catch{}delete ee[a]}}function Oa(){Object.keys(ee).forEach(pe)}var K=(a,e=0)=>{let n=Number(a);return isNaN(n)||a===null||a===void 0?e:n},j=(a,e="\u2014")=>{if(a==null||a==="")return e;let n=String(a).trim();return n===""||n==="[object Object]"?e:n},Je=a=>{if(!a)return"\u2014";try{let e=new Date(a);return isNaN(e)?j(a):e.toLocaleDateString("id-ID",{day:"2-digit",month:"short",year:"numeric"})}catch{return j(a)}},Ma=a=>{if(!a)return"";try{let e=Date.now()-new Date(a).getTime();if(e<0)return"Baru saja";let n=Math.floor(e/6e4);if(n<1)return"Baru saja";if(n<60)return`${n} menit lalu`;let i=Math.floor(n/60);return i<24?`${i} jam lalu`:`${Math.floor(i/24)} hari lalu`}catch{return""}},Ge=a=>{if(!a||typeof a!="string")return"";try{let[e,n]=a.split("-");return new Date(Number(e),Number(n)-1).toLocaleDateString("id-ID",{month:"short",year:"2-digit"})}catch{return a}};function Qe(a,e,n=900){if(!a)return;let i=Math.max(0,Math.round(K(e)));if(i===0){a.textContent="0";return}let t=Date.now(),d=()=>{let u=Math.min((Date.now()-t)/n,1),p=1-Math.pow(1-u,3);a.textContent=Math.round(p*i).toLocaleString("id-ID"),u<1?requestAnimationFrame(d):a.textContent=i.toLocaleString("id-ID")};requestAnimationFrame(d)}function Aa(a,e){if(a=K(a),e=K(e),e===0)return"";let n=a-e,i=Math.abs(Math.round(n/e*100));return n>0?`<span class="kpi-trend up">\u25B2 ${i}%</span>`:n<0?`<span class="kpi-trend down">\u25BC ${i}%</span>`:'<span class="kpi-trend neutral">= Sama</span>'}var Fa={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},Ve=a=>{let e=j(a,"\u2014");return`<span class="status-pill ${Fa[e]||"pill-neutral"}">${e}</span>`},Ra=a=>{let e=K(a,999);return e<=7?`<span class="days-badge days-critical">${e} hari</span>`:e<=14?`<span class="days-badge days-warning">${e} hari</span>`:e<=30?`<span class="days-badge days-soon">${e} hari</span>`:`<span class="days-badge days-ok">${e} hari</span>`},Ka={issue:{emoji:"\u26A0\uFE0F",dot:"dot-danger",label:"Permasalahan"},contract:{emoji:"\u{1F4C4}",dot:"dot-info",label:"Kontrak"},employee:{emoji:"\u{1F464}",dot:"dot-success",label:"Karyawan"},one_on_one:{emoji:"\u{1F91D}",dot:"dot-purple",label:"One on One"},training:{emoji:"\u{1F393}",dot:"dot-primary",label:"Training"},supply:{emoji:"\u{1F4E6}",dot:"dot-warning",label:"Permintaan Barang"},reliever:{emoji:"\u{1F504}",dot:"dot-teal",label:"Reliefer"},inspection:{emoji:"\u{1F50D}",dot:"dot-blue",label:"Laporan Inspeksi"}},qa=a=>Ka[a]||{emoji:"\u{1F4CC}",dot:"dot-neutral",label:j(a,"Aktivitas")},G={family:"Inter",size:11},ne="#94A3B8",he="#F1F5F9",Ha=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],ja=()=>window.innerWidth<768;function xe(a={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:ja()?"bottom":"top",labels:{font:G,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:G,titleFont:{...G,weight:"700"}}},scales:{x:{grid:{color:he},ticks:{font:G,color:ne,maxRotation:0}},y:{grid:{color:he},ticks:{font:G,color:ne},beginAtZero:!0}},...a}}var Ua=()=>Array(5).fill(0).map(()=>`
  <div class="kpi-card" style="pointer-events:none">
    <div class="kpi-card-top"><div class="skeleton" style="width:44px;height:44px;border-radius:12px"></div></div>
    <div class="skeleton skeleton-text" style="width:55%;height:32px;margin:10px 0 6px"></div>
    <div class="skeleton skeleton-text" style="width:75%;height:12px;margin-bottom:4px"></div>
    <div class="skeleton skeleton-text" style="width:55%;height:11px"></div>
  </div>`).join(""),Ja=()=>Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join("");function Ue(a=3){return Array(a).fill(0).map((e,n)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${n<a-1?"6px":"0"};border-radius:6px"></div>`).join("")}function Ga(){return Array(5).fill(0).map(()=>`
    <div class="activity-item">
      <div class="skeleton" style="width:34px;height:34px;border-radius:10px;flex-shrink:0"></div>
      <div style="flex:1">
        <div class="skeleton skeleton-text" style="width:65%;height:13px;margin-bottom:5px"></div>
        <div class="skeleton skeleton-text" style="width:35%;height:11px"></div>
      </div>
    </div>`).join("")}async function X(a,e,n=8e3){try{let i=new AbortController,t=setTimeout(()=>i.abort(),n),d=await f(a,{signal:i.signal}).catch(()=>null);if(clearTimeout(t),!d||!d.ok)return e;let u=d.data;return u?u.data!==void 0?u.data??e:u:e}catch{return e}}function Qa(){["skel-donut","skel-trend","skel-insp","skel-contract"].forEach(i=>{let t=document.getElementById(i);t&&(t.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract"].forEach(i=>{let t=document.getElementById(i);if(t&&t.style.display==="none"){t.style.display="block";let d=t.parentElement;if(d&&!d.querySelector(".chart-empty")){let u=document.createElement("div");u.className="chart-empty",u.textContent="Belum ada data",t.style.display="none",d.appendChild(u)}}});let a=document.getElementById("kpi-row");a&&a.querySelector(".skeleton")&&We({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&Ye({}),["table-contracts","table-issues"].forEach(i=>{let t=document.getElementById(i);t&&t.querySelector(".skeleton")&&(t.innerHTML='<div class="chart-empty">Belum ada data</div>')});let n=document.getElementById("activity-log");n&&n.querySelector(".skeleton")&&(n.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function ze(a){Oa(),a._dashRefresh&&clearInterval(a._dashRefresh),a._skelTimeout&&clearTimeout(a._skelTimeout),a.innerHTML=`
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
      <div class="kpi-row" id="kpi-row">${Ua()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${Ja()}</div>

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
          <div id="table-contracts" class="dash-table-wrap">${Ue(3)}</div>
        </div>
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">\u26A0\uFE0F Permasalahan Terbaru</div>
              <div class="chart-card-subtitle">Open dan In Progress</div>
            </div>
            <a href="#/issues" class="btn btn-ghost btn-sm">Lihat Semua \u2192</a>
          </div>
          <div id="table-issues" class="dash-table-wrap">${Ue(3)}</div>
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
        <div id="activity-log">${Ga()}</div>
      </div>

    </div>
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>Ce(a)),a._skelTimeout=setTimeout(()=>Qa(),5e3),await Ce(a),a._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?Ce(a):clearInterval(a._dashRefresh)},6e4)}async function Ce(a){a._skelTimeout&&(clearTimeout(a._skelTimeout),a._skelTimeout=null);let[e,n,i,t,d,u,p,s]=await Promise.all([X("/api/dashboard/kpi",{},8e3),X("/api/dashboard/issues-trend",{},8e3),X("/api/dashboard/contracts-chart",{},8e3),X("/api/dashboard/issues-summary",{},8e3),X("/api/dashboard/inspection-bar",{},8e3),X("/api/dashboard/contracts-expiring",[],8e3),X("/api/dashboard/stats",{},8e3),X("/api/dashboard/activity-log",[],8e3)]);try{We(e)}catch(o){console.warn("KPI render:",o)}try{Ye(e)}catch(o){console.warn("MiniStats render:",o)}try{Va(Array.isArray(t?.by_category)?t.by_category:[])}catch(o){console.warn("Donut render:",o),Z("skel-donut","chart-donut")}try{za(n)}catch(o){console.warn("Trend render:",o),Z("skel-trend","chart-trend")}try{Wa(d)}catch(o){console.warn("InspBar render:",o),Z("skel-insp","chart-insp")}try{Ya(i)}catch(o){console.warn("ContractBar render:",o),Z("skel-contract","chart-contract")}try{let o=Array.isArray(u)?u:[];Xa(o)}catch(o){console.warn("ContractsTable render:",o)}try{let o=Array.isArray(p)?p:Array.isArray(p?.recent_issues)?p.recent_issues:[];Za(o)}catch(o){console.warn("IssuesTable render:",o)}try{et(Array.isArray(s)?s:[])}catch(o){console.warn("ActivityLog render:",o)}let r=document.getElementById("dash-updated");r&&(r.textContent=`Diperbarui: ${new Date().toLocaleTimeString("id-ID")}`)}function We(a){let e=document.getElementById("kpi-row");if(!e)return;a=a||{};let n=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees",color:"kpi-blue",key:"employees"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts",color:"kpi-green",key:"contracts"},{icon:"\u23F0",label:"Kontrak Habis 30 Hari",sub:"",href:"#/contracts",color:"",key:"expiring30",warn:!0},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues",color:"",key:"issues",warnIfGT0:!0},{icon:"\u{1F91D}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one",color:"kpi-purple",key:"one_on_one"}];e.innerHTML=n.map(i=>{let t=K(a[i.key]?.current,0),d=a[i.key]?.prev,u=d!=null?Aa(t,d):"",p=i.color||"",s=i.sub||"";return i.warn&&(p=t>0?"kpi-amber":"kpi-green",s=t>0?`\u26A0\uFE0F ${t} kontrak segera berakhir`:"\u2705 Semua kontrak aman"),i.warnIfGT0&&(p=t>0?"kpi-red":"kpi-green"),`
      <a href="${i.href}" class="kpi-card ${p}" style="text-decoration:none">
        <div class="kpi-card-top">
          <div class="kpi-icon-wrap"><span class="kpi-icon-emoji">${i.icon}</span></div>
          ${u}
        </div>
        <div class="kpi-value" data-target="${t}">0</div>
        <div class="kpi-label">${i.label}</div>
        <div class="kpi-subtitle">${s}</div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(i=>Qe(i,parseInt(i.dataset.target)||0))}function Ye(a){let e=document.getElementById("mini-stats-row");if(!e)return;a=a||{};let n=[{icon:"\u{1F4C5}",label:"Jadwal Pending",val:a.schedule?.current,href:"#/schedule",color:"mini-blue"},{icon:"\u{1F393}",label:"Training Bulan Ini",val:a.training_month?.current,href:"#/training",color:"mini-indigo"},{icon:"\u{1F4E6}",label:"Permintaan Barang",val:a.supply?.current,href:"#/reports/supply",color:"mini-orange"},{icon:"\u{1F50D}",label:"Inspeksi Bulan Ini",val:a.inspection_month?.current,href:"#/reports/inspection",color:"mini-teal"},{icon:"\u{1F9F9}",label:"GC/DC Bulan Ini",val:a.cleaning_month?.current,href:"#/reports/cleaning",color:"mini-green"},{icon:"\u{1F99F}",label:"Fogging Bulan Ini",val:a.fogging_month?.current,href:"#/reports/fogging",color:"mini-purple"},{icon:"\u{1F3E2}",label:"Total Cabang",val:a.branches?.current,href:"#/branches",color:"mini-gray"}];e.innerHTML=n.map(i=>`
    <a href="${i.href}" class="mini-stat ${i.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${i.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${K(i.val)}">0</div>
        <div class="mini-stat-label">${i.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(i=>Qe(i,parseInt(i.dataset.target)||0,700))}function Va(a){Z("skel-donut","chart-donut");let e=document.getElementById("chart-donut");if(!e)return;pe("donut");let n=(a||[]).filter(d=>K(d.count)>0);if(!n.length){ye(e,"Belum ada data permasalahan");return}let i=n.map(d=>j(d.category,"Lainnya")),t=n.map(d=>K(d.count));ee.donut=new Chart(e,{type:"doughnut",data:{labels:i,datasets:[{data:t,backgroundColor:Ha.slice(0,n.length),borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{position:window.innerWidth<768?"bottom":"right",labels:{font:G,color:"#475569",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{bodyFont:G,titleFont:{...G,weight:"700"},callbacks:{label:d=>` ${d.label}: ${d.parsed} kasus`}}},cutout:"65%"}})}function za(a){Z("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;pe("trend"),a=a||{};let n=(a.labels||[]).map(Ge),i=(a.open||[]).map(d=>K(d)),t=(a.closed||[]).map(d=>K(d));if(!n.length){ye(e,"Belum ada data trend");return}ee.trend=new Chart(e,{type:"line",data:{labels:n,datasets:[{label:"Open",data:i,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:t,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:xe({plugins:{legend:{position:"top"}}})})}function Wa(a){Z("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;pe("inspBar"),a=a||{};let n=a.labels||[],i=(a.fc||[]).map(d=>K(d)),t=(a.spv||[]).map(d=>K(d));if(!n.length){ye(e,"Belum ada data inspeksi");return}ee.inspBar=new Chart(e,{type:"bar",data:{labels:n,datasets:[{label:"Skor FC",data:i,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:t,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:xe({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:G,color:ne,maxRotation:45,minRotation:30}},y:{grid:{color:he},ticks:{font:G,color:ne},min:0,max:100}}})})}function Ya(a){Z("skel-contract","chart-contract");let e=document.getElementById("chart-contract");if(!e)return;pe("contractBar"),a=a||{};let n=(a.labels||[]).map(Ge),i=(a.counts||[]).map(d=>K(d));if(!n.length){ye(e,"Belum ada data kontrak");return}let t=i.map(d=>d>5?"rgba(239,68,68,.75)":d>2?"rgba(245,158,11,.75)":"rgba(37,99,235,.65)");ee.contractBar=new Chart(e,{type:"bar",data:{labels:n,datasets:[{label:"Kontrak Berakhir",data:i,backgroundColor:t,borderRadius:6,borderSkipped:!1}]},options:xe({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:G,color:ne}},y:{grid:{color:he},ticks:{font:G,color:ne,precision:0},beginAtZero:!0}}})})}function Xa(a){let e=document.getElementById("table-contracts");if(!e)return;let n=(a||[]).filter(i=>K(i.days_remaining,999)<=30).slice(0,10);if(!n.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada kontrak yang habis dalam 30 hari</div>';return}e.innerHTML=`
    <table class="dash-table">
      <thead><tr>
        <th>#</th><th>Nama Karyawan</th><th>Cabang</th><th>Berakhir</th><th>Sisa</th><th>Status</th>
      </tr></thead>
      <tbody>${n.map((i,t)=>`
        <tr>
          <td class="td-num">${t+1}</td>
          <td><strong>${j(i.emp_name||i.employee_name)}</strong></td>
          <td class="td-branch">${j(i.branch_name)}</td>
          <td style="white-space:nowrap;font-size:.8rem">${Je(i.end_date)}</td>
          <td>${Ra(i.days_remaining)}</td>
          <td>${Ve(i.status)}</td>
        </tr>`).join("")}
      </tbody>
    </table>`}function Za(a){let e=document.getElementById("table-issues");if(!e)return;let n=(a||[]).slice(0,8);if(!n.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <table class="dash-table">
      <thead><tr>
        <th>Tanggal</th><th>Keluhan</th><th>Cabang</th><th>Kategori</th><th>Status</th>
      </tr></thead>
      <tbody>${n.map(i=>`
        <tr>
          <td style="white-space:nowrap;font-size:.78rem">${Je(i.report_date)}</td>
          <td class="td-complaint" title="${j(i.complaint)}">${j(i.complaint)}</td>
          <td class="td-branch">${j(i.branch_name)}</td>
          <td><span class="category-tag">${j(i.category)}</span></td>
          <td>${Ve(i.status)}</td>
        </tr>`).join("")}
      </tbody>
    </table>`}function et(a){let e=document.getElementById("activity-log");if(!e)return;let n=(a||[]).slice(0,15);if(!n.length){e.innerHTML='<div class="chart-empty">Belum ada aktivitas tercatat</div>';return}e.innerHTML=`<div class="activity-list">${n.map(i=>{let t=qa(i.type),d=j(i.label),u=i.branch?` \u2022 ${j(i.branch)}`:"",p=Ma(i.created_at);return`
      <div class="activity-item">
        <div class="activity-dot ${t.dot}">${t.emoji}</div>
        <div class="activity-body">
          <div class="activity-text"><strong>${t.label}</strong> \u2014 ${d}${u}</div>
          <div class="activity-time">${p||"\u2014"}</div>
        </div>
      </div>`}).join("")}</div>`}function Z(a,e){let n=document.getElementById(a),i=document.getElementById(e);n&&(n.style.display="none",n.style.position=""),i&&(i.style.display="block")}function ye(a,e="Belum ada data"){if(!a)return;a.style.display="none";let n=a.parentElement;if(!n)return;if(!n.querySelector(".chart-empty")){let t=document.createElement("div");t.className="chart-empty",t.textContent=e,n.appendChild(t)}}P();async function Xe(a){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),n=document.getElementById("login-error"),i=document.getElementById("login-btn"),t=document.getElementById("toggle-password"),d=document.getElementById("login-password");t?.addEventListener("click",()=>{let u=d.type==="text";d.type=u?"password":"text",t.style.color=u?"":"var(--primary)"}),e?.addEventListener("submit",async u=>{u.preventDefault(),n.style.display="none";let p=e.username.value.trim(),s=e.password.value;if(!p||!s){n.textContent="Username dan password wajib diisi.",n.style.display="block";return}i.querySelector(".btn-text").style.display="none",i.querySelector(".btn-spinner").style.display="",i.disabled=!0;try{let r=await f("/api/auth/login",{method:"POST",body:JSON.stringify({username:p,password:s})});r.ok&&r.data.success?(_e(r.data.data.token),oe(r.data.data.user),R("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(n.textContent=r.data.error||"Username atau password salah.",n.style.display="block",i.classList.add("shake"),setTimeout(()=>i.classList.remove("shake"),600))}catch{n.textContent="Gagal terhubung ke server. Periksa koneksi internet.",n.style.display="block"}finally{i.querySelector(".btn-text").style.display="",i.querySelector(".btn-spinner").style.display="none",i.disabled=!1}})}P();function Ze({columns:a,data:e,onEdit:n,onDelete:i,onView:t,actions:d=[],emptyText:u="Tidak ada data",bulkSelect:p=null}){let s=document.createElement("div");if(s.className="table-wrapper",!e||e.length===0)return s.innerHTML=`<div class="empty-state"><p>${u}</p></div>`,s;let r=document.createElement("table");r.className="data-table";let o=document.createElement("thead"),l=document.createElement("tr");if(p){let b=document.createElement("th");b.style.width="40px",b.style.textAlign="center";let g=document.createElement("input");g.type="checkbox",g.id="select-all-checkbox",g.title="Pilih semua",g.addEventListener("change",()=>{e.forEach(c=>{g.checked?p.selectedIds.add(c.id):p.selectedIds.delete(c.id)}),s.querySelectorAll(".row-checkbox").forEach(c=>c.checked=g.checked),p.onToggle()}),b.appendChild(g),l.appendChild(b)}if(a.forEach(b=>{let g=document.createElement("th");g.textContent=b.label,b.width&&(g.style.width=b.width),l.appendChild(g)}),n||i||t||d.length>0){let b=document.createElement("th");b.textContent="Aksi",b.style.width="120px",l.appendChild(b)}o.appendChild(l),r.appendChild(o);let h=document.createElement("tbody");return e.forEach(b=>{let g=document.createElement("tr");if(p){let c=document.createElement("td");c.style.textAlign="center",c.style.width="40px";let m=document.createElement("input");m.type="checkbox",m.className="row-checkbox",m.checked=p.selectedIds.has(b.id),m.addEventListener("change",()=>{if(m.checked)p.selectedIds.add(b.id);else{p.selectedIds.delete(b.id);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1)}p.onToggle()}),c.appendChild(m),g.appendChild(c)}if(a.forEach(c=>{let m=document.createElement("td");if(c.render){let y=c.render(b[c.key],b);y instanceof HTMLElement?m.appendChild(y):m.innerHTML=y||""}else m.textContent=b[c.key]!==null&&b[c.key]!==void 0&&b[c.key]!==""?b[c.key]:"";c.nowrap&&(m.style.whiteSpace="nowrap"),g.appendChild(m)}),n||i||t||d.length>0){let c=document.createElement("td");c.className="actions-cell";let m=document.createElement("div");if(m.className="btn-group",t){let y=document.createElement("button");y.className="btn btn-xs btn-ghost",y.innerHTML="\u{1F441}",y.title="Lihat",y.addEventListener("click",()=>t(b)),m.appendChild(y)}if(n){let y=document.createElement("button");y.className="btn btn-xs btn-secondary",y.innerHTML="\u270F\uFE0F",y.title="Edit",y.addEventListener("click",()=>n(b)),m.appendChild(y)}d.forEach(y=>{let v=document.createElement("button");v.className=`btn btn-xs ${y.class||"btn-ghost"}`,v.innerHTML=y.icon||y.label,v.title=y.label,v.addEventListener("click",()=>y.handler(b)),m.appendChild(v)}),c.appendChild(m),g.appendChild(c)}h.appendChild(g)}),r.appendChild(h),s.appendChild(r),s}function ea({page:a,pages:e,total:n,limit:i,onPage:t}){if(e<=1)return null;let d=document.createElement("div");d.className="pagination";let u=document.createElement("span");u.className="pagination-info",u.textContent=`Total: ${n} data`,d.appendChild(u);let p=document.createElement("div");p.className="pagination-btns";let s=(l,h,b=!1,g=!1)=>{let c=document.createElement("button");c.className=`btn btn-sm ${g?"btn-primary":"btn-ghost"} pagination-btn`,c.textContent=l,c.disabled=b,c.addEventListener("click",()=>t(h)),p.appendChild(c)};s("\xAB",1,a===1),s("\u2039",a-1,a===1);let r=Math.max(1,a-2),o=Math.min(e,a+2);for(let l=r;l<=o;l++)s(l,l,!1,l===a);return s("\u203A",a+1,a===e),s("\xBB",e,a===e),d.appendChild(p),d}function Te(a){return a.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="row")return`<div class="form-row">${Te(e.fields)}</div>`;let n=e.required?"required":"",i=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",t="";switch(e.type){case"textarea":t=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${n} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let u=(e.options||[]).map(p=>{let s=typeof p=="object"?p.value:p,r=typeof p=="object"?p.label:p,o=e.value==s?"selected":"";return`<option value="${s}" ${o}>${r}</option>`}).join("");t=`<select name="${e.name}" class="form-control" ${n}><option value="">-- Pilih ${e.label||""} --</option>${u}</select>`;break;case"checkbox":t=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":t=`<input type="date" name="${e.name}" class="form-control" value="${e.value||""}" ${n}>`;break;case"number":t=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${n}>`;break;case"email":t=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${n}>`;break;case"url":t=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${n}>`;break;default:t=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${n} autocomplete="off">`}let d=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${i}${t}${d}</div>`}).join("")}function aa(a){let e={},n=new FormData(a);for(let[i,t]of n.entries())e[i]=t===""?null:t;return a.querySelectorAll("input[type=checkbox]").forEach(i=>{i.checked||(e[i.name]=null)}),e}function ta(a,e){e&&Object.entries(e).forEach(([n,i])=>{let t=a.querySelector(`[name="${n}"]`);t&&(t.type==="checkbox"?t.checked=!!i:t.value=i??"")})}D();function E({container:a,title:e,icon:n,apiPath:i,columns:t,formFields:d,filterFields:u,defaultFilters:p={},itemLabel:s="Data",canCreate:r=!0,canEdit:o=!0,canDelete:l=!0,onBeforeSubmit:h,onAfterLoad:b,extraActions:g=[],initialSearch:c="",exportOptions:m=null,bulkDelete:y=!1}){let v=1,S={...p};c&&(S.search=c);let k=new Set;a.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${n} ${e}</h1>
      <div class="page-actions">
        ${r?`<button class="btn btn-primary" id="btn-create">+ Tambah ${s}</button>`:""}
      </div>
    </div>

    ${y?`
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>\u{1F5D1}\uFE0F Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>`:""}
    
    ${m?$e(m.moduleName):""}

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
  `;function _(){if(!document.getElementById("bulk-toolbar"))return;let x=document.getElementById("bulk-count"),L=document.getElementById("btn-bulk-delete"),F=document.getElementById("btn-bulk-cancel");x.textContent=`${k.size} item dipilih`,k.size>0?(L.disabled=!1,F.disabled=!1):(L.disabled=!0,F.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{k.clear(),document.querySelectorAll(".row-checkbox").forEach(x=>x.checked=!1);let C=document.getElementById("select-all-checkbox");C&&(C.checked=!1),_()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(k.size===0)return;let C=[...k],x=document.createElement("div");x.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",x.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${C.length} ${s}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${C.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(x),x.querySelector("#bulk-cancel-btn").addEventListener("click",()=>x.remove()),x.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let L=x.querySelector("#bulk-confirm-btn");L.disabled=!0,L.textContent="Menghapus...";let F=await f(`${i}/bulk`,{method:"DELETE",body:JSON.stringify({ids:C})});x.remove(),F.ok?(R(`${C.length} ${s} berhasil dihapus.`),k.clear(),_(),w()):H(F.data?.error||"Gagal menghapus data.")})});let $=document.getElementById("filter-search"),O;if($?.addEventListener("input",C=>{clearTimeout(O),O=setTimeout(()=>{S.search=C.target.value,v=1,w()},400)}),u?.forEach(C=>{C.type==="select"&&document.getElementById(`filter-${C.name}`)?.addEventListener("change",x=>{S[C.name]=x.target.value,v=1,w()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{S={...p},$&&($.value=""),u?.forEach(C=>{let x=document.getElementById(`filter-${C.name}`);x&&(x.value="")}),v=1,w()}),document.getElementById("btn-create")?.addEventListener("click",()=>Q(null)),m){document.getElementById(`btn-export-${m.moduleName}`)?.addEventListener("click",async x=>{let L=x.target,F=L.innerHTML;L.innerHTML="\u23F3 Loading...",L.disabled=!0;try{await m.onExport()}catch{H("Gagal export data")}finally{L.innerHTML=F,L.disabled=!1}}),document.getElementById(`btn-template-${m.moduleName}`)?.addEventListener("click",()=>{m.onTemplate()});let C=document.getElementById(`input-import-${m.moduleName}`);C?.addEventListener("change",async x=>{let L=x.target.files[0];if(!L)return;let F=C.parentElement,J=F.innerHTML;F.innerHTML="\u23F3 Memproses...",F.style.pointerEvents="none",C.disabled=!0;try{let z=await Ee(L);if(z.length===0)throw new Error("File kosong atau format salah");await m.onImport(z),R("Import berhasil!"),w()}catch(z){H(z.message||"Gagal import data")}finally{F.innerHTML=J,F.style.pointerEvents="auto",C.disabled=!1,C.value=""}})}async function w(){let C=document.getElementById("table-container");if(!C)return;C.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let x=new URLSearchParams({page:v,limit:20,...Object.fromEntries(Object.entries(S).filter(([,q])=>q))}),L=await f(`${i}?${x}`);if(!L.ok){C.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${L.data?.error||"Error"}</p></div>`;return}let F=L.data?.data||[],J=L.data?.pagination;b&&b(F);let z=Ze({columns:t,data:F,onEdit:o?q=>Q(q):null,actions:g.map(q=>({...q,handler:se=>q.handler(se,w)})),emptyText:`Tidak ada ${s.toLowerCase()}`,bulkSelect:y?{selectedIds:k,onToggle:_}:null});C.innerHTML="",C.appendChild(z);let W=document.getElementById("pagination-container");if(W&&(W.innerHTML="",J&&J.pages>1)){let q=ea({page:J.page,pages:J.pages,total:J.total,limit:J.limit,onPage:se=>{v=se,w()}});q&&W.appendChild(q)}}function I(C){let x=typeof d=="function"?d(C):d;return Te(x)}function Q(C){let x=!!C,L=document.createElement("form");if(L.noValidate=!0,L.innerHTML=I(C),x){let J=typeof d=="function"?d(C):d;ta(L,C)}let{close:F}=Y({title:x?`Edit ${s}`:`Tambah ${s}`,content:L,size:"lg",confirmText:x?"Simpan Perubahan":`Tambah ${s}`,onConfirm:async(J,z)=>{if(!L.reportValidity())return;let W=J.querySelector(".modal-confirm");W.disabled=!0,W.textContent="Menyimpan...";let q=aa(L);h&&(q=await h(q,C));let se=x?"PUT":"POST",La=x?`${i}/${C.id}`:i,Ae=await f(La,{method:se,body:JSON.stringify(q)});Ae.ok?(R(x?`${s} berhasil diperbarui.`:`${s} berhasil ditambahkan.`),z(),w()):(H(Ae.data?.error||"Gagal menyimpan data."),W.disabled=!1,W.textContent=x?"Simpan Perubahan":`Tambah ${s}`)}})}function bt(C){je(`Hapus ${s} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let x=await f(`${i}/${C.id}`,{method:"DELETE"});x.ok?(R(`${s} berhasil dihapus.`),w()):H(x.data?.error||"Gagal menghapus.")},`Hapus ${s}`)}return w(),w}P();function M(a){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!a||a==="-"||String(a).trim()===""?"":`<span class="badge ${e[a]||"badge-neutral"}">${a}</span>`}function na(a){return a==null?'<span class="badge badge-neutral">-</span>':a<0?`<span class="badge badge-danger">Expired (${Math.abs(a)}h)</span>`:a<=14?`<span class="badge badge-danger">${a} hari</span>`:a<=30?`<span class="badge badge-warning">${a} hari</span>`:`<span class="badge badge-success">${a} hari</span>`}function fe(a){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[a]||"badge-neutral"}">${a||"-"}</span>`}function ia(a){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[a]||"badge-neutral"}">${a||"-"}</span>`}function V(a){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[a]||"badge-neutral"}">${a||"-"}</span>`}D();var Pe=[],Ie=[];async function at(){Ie=(await f("/api/branches?all=1")).data?.data||[],Pe=Ie.map(e=>({value:e.id,label:e.full_name}))}async function sa(a){await at(),E({container:a,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:e=>fe(e)},{key:"phone",label:"No. HP",render:e=>e?`<a href="tel:${e}">${e}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:e=>window.formatDate(e)},{key:"status",label:"Status",render:e=>M(e)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:Pe},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:e=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:e?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:e?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:Pe,value:e?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:e?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:e?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:e?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:e?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let e=await f("/api/employees?limit=10000");if(e.ok){let n=e.data.data.map(i=>({"Nama Lengkap":i.full_name,Cabang:i.branch_name||"",Divisi:i.division||"","No. HP":i.phone||"","Tgl Masuk":i.join_date||"",Status:i.status||"",Catatan:i.notes||""}));T(n,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif",Catatan:""},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif",Catatan:""}],"Template_Import_Karyawan")},onImport:async e=>{let n=d=>{if(!d)return null;let u=d.toLowerCase(),p=Ie.find(s=>s.full_name.toLowerCase()===u||s.code.toLowerCase()===u||s.name.toLowerCase()===u);return p?p.id:null},i=e.map(d=>({full_name:String(d["Nama Lengkap"]||"").trim(),branch_id:n(String(d.Cabang||"").trim()),division:String(d.Divisi||"").trim()||"FACILITY CARE",phone:String(d["No. HP"]||"").trim(),join_date:String(d["Tgl Masuk"]||"").trim(),status:String(d.Status||"").trim(),notes:String(d.Catatan||"").trim()})).filter(d=>d.full_name),t=await f("/api/employees/import",{method:"POST",body:JSON.stringify(i)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}}})}P();D();var Le=[],ra=[];async function tt(){let[a,e]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000&status=Aktif")]);Le=(a.data?.data||[]).map(n=>({value:n.id,label:n.full_name})),ra=(e.data?.data||[]).map(n=>({value:n.id,label:n.full_name}))}async function la(a){await tt(),E({container:a,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:e=>fe(e)},{key:"start_date",label:"Tgl Mulai",nowrap:!0,render:e=>window.formatDate(e)},{key:"end_date",label:"Tgl Selesai",nowrap:!0,render:e=>window.formatDate(e)},{key:"days_remaining",label:"Sisa",render:e=>na(e)},{key:"contract_type",label:"Tipe Kontrak"},{key:"pkwt_number",label:"PKWT"},{key:"status",label:"Status",render:e=>M(e)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:Le},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],formFields:e=>[{type:"row",fields:[{name:"employee_id",label:"Karyawan",type:"select",required:!0,options:ra,value:e?.employee_id},{name:"branch_id",label:"Cabang",type:"select",options:Le,value:e?.branch_id}]},{type:"row",fields:[{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:e?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif"],value:e?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",required:!0,value:e?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",required:!0,value:e?.end_date}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:e?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:e?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:e?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let e=await f("/api/contracts?limit=10000");if(e.ok){let n=e.data.data.map(i=>({"Nama Karyawan":i.employee_name,Cabang:i.branch_name||"",Divisi:i.division||"","Tgl Mulai":i.start_date||"","Tgl Selesai":i.end_date||"","Tipe Kontrak":i.contract_type||"",PKWT:i.pkwt_number||"",Status:i.status||"",Catatan:i.notes||""}));T(n,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Nama Karyawan":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","Tgl Mulai":"2024-01-01","Tgl Selesai":"2024-12-31","Tipe Kontrak":"KONTRAK 1 TAHUN",PKWT:"PKWT 1",Status:"Aktif",Catatan:""}],"Template_Import_Kontrak")},onImport:async e=>{let[n,i]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000")]),t=n.data?.data||[],d=i.data?.data||[],u=o=>{if(!o)return null;let l=o.toLowerCase(),h=t.find(b=>b.full_name.toLowerCase()===l||b.code.toLowerCase()===l||b.name.toLowerCase()===l);return h?h.id:null},p=o=>{if(!o)return null;let l=o.toLowerCase(),h=d.find(b=>b.full_name.toLowerCase()===l);return h?h.id:null},s=e.map(o=>({employee_id:p(String(o["Nama Karyawan"]||"").trim()),branch_id:u(String(o.Cabang||"").trim()),division:String(o.Divisi||"").trim()||"FACILITY CARE",start_date:String(o["Tgl Mulai"]||"").trim(),end_date:String(o["Tgl Selesai"]||"").trim(),contract_type:String(o["Tipe Kontrak"]||"").trim(),pkwt_number:String(o.PKWT||"").trim(),status:String(o.Status||"").trim(),notes:String(o.Catatan||"").trim()})).filter(o=>o.employee_id&&o.start_date&&o.end_date),r=await f("/api/contracts/import",{method:"POST",body:JSON.stringify(s)});if(!r.ok)throw new Error(r.data?.error||"Import gagal")}}})}P();D();var De=[],Be=[];async function oa(a){let[e,n,i]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]);De=(e.data?.data||[]).map(s=>({value:s.id,label:s.full_name}));let t=(n.data?.data||[]).map(s=>({value:s.full_name,label:s.full_name}));Be=[...(i.data?.data||[]).filter(s=>s.role==="FC Spesialis").map(s=>({value:s.name,label:s.name}))];let u=s=>s&&!t.find(r=>r.value===s)?[...t,{value:s,label:s}]:t,p=s=>{if(!s||s==="-"||String(s).trim()==="")return"";let r=String(s).split("-");return r.length===3&&r[0].length===4?`${r[2]}-${r[1]}-${r[0]}`:s};E({container:a,title:"Jadwal Kegiatan",icon:"\u{1F5D3}\uFE0F",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:s=>ia(s)},{key:"period",label:"Periode",render:s=>V(s)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:s=>p(s)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:s=>p(s)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:s=>p(s)},{key:"status",label:"Status",render:s=>M(s)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:De},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Be}],formFields:s=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:De,value:s?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"],value:s?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:s?.period},{name:"pic",label:"PIC",type:"select",options:Be,value:s?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:s?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:s?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:s?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:s?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let s=await f("/api/schedule?limit=10000");if(s.ok){let r=s.data.data.map(o=>({Cabang:o.branch_name||"",Kegiatan:o.activity_type||"",Periode:o.period||"",PIC:o.pic||"","Tgl Opening":o.opening_date||"","Tgl Target":o.target_date||"","Tgl Selesai":o.completion_date||"",Status:o.status||"",Catatan:o.notes||""}));T(r,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done",Catatan:""}],"Template_Import_Jadwal")},onImport:async s=>{let o=(await f("/api/branches?all=1")).data?.data||[],l=c=>{if(!c)return null;let m=c.toLowerCase(),y=o.find(v=>v.full_name.toLowerCase()===m||v.code.toLowerCase()===m||v.name.toLowerCase()===m);return y?y.id:null},h=c=>{if(c==null||c==="")return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let m=String(c).trim();if(m===""||m==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);if(/^\d{4,5}$/.test(m)){let v=Number(m);if(v>2e4&&v<99999){let S=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}let y=m.split(/[\/\-\.]/);if(y.length===3){let[v,S,k]=y.map(_=>_.trim());if(v.length===4&&S.length<=2&&k.length<=2)return`${v}-${S.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&S.length<=2&&v.length<=2)return`${k}-${S.padStart(2,"0")}-${v.padStart(2,"0")}`}return m},b=s.map(c=>({branch_id:l(String(c.Cabang||"").trim()),activity_type:String(c.Kegiatan||"").trim(),period:String(c.Periode||"").trim(),pic:String(c.PIC||c.Pic||"").trim(),opening_date:h(c["Tgl Opening"]||c["Tanggal Opening"]),target_date:h(c["Tgl Target"]||c["Tanggal Target"]),completion_date:h(c["Tgl Selesai"]||c["Tanggal Selesai"]),status:String(c.Status||"").trim(),notes:String(c.Catatan||c.Keterangan||"").trim()})).filter(c=>c.activity_type&&c.period),g=await f("/api/schedule/import",{method:"POST",body:JSON.stringify(b)});if(!g.ok)throw new Error(g.data?.error||"Import gagal")}}})}P();D();var Ne=[],ve=[];async function da(a){let[e,n,i]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]);Ne=(e.data?.data||[]).map(l=>({value:l.id,label:l.full_name})),ve=(n.data?.data||[]).map(l=>({value:l.full_name,label:l.full_name}));let t=(i.data?.data||[]).filter(l=>l.role==="FC Spesialis").map(l=>({value:l.name,label:l.name})),d=(i.data?.data||[]).filter(l=>l.role==="Pelapor").map(l=>({value:l.name,label:l.name})),u=l=>l&&!ve.find(h=>h.value===l)?[...ve,{value:l,label:l}]:ve,p=l=>l&&!t.find(h=>h.value===l)?[...t,{value:l,label:l}]:t,s=l=>l&&!d.find(h=>h.value===l)?[...d,{value:l,label:l}]:d,r=new Date().getFullYear(),o=Array.from({length:5},(l,h)=>String(r-h));E({container:a,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:l=>`<span class="badge badge-secondary">${l}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:l=>`<span title="${l}">${l?.length>50?l.slice(0,50)+"\u2026":l}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:l=>`<span title="${l||""}">${l?.length>40?l.slice(0,40)+"\u2026":l||"-"}</span>`},{key:"status",label:"Status",render:l=>M(l)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:l=>window.formatDate(l)},{key:"day_count",label:"Hari",render:l=>l??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:Ne},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:o}],formFields:l=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:l?.report_date},{name:"branch_id",label:"Cabang",type:"select",required:!0,options:Ne,value:l?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:l?.category},{name:"source",label:"Sumber Laporan",type:"select",options:[...s(l?.source),{value:"Lainnya",label:"Lainnya"}],value:l?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:l?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"select",options:u(l?.employee_name),value:l?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"select",options:p(l?.fc_specialist),value:l?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:l?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:l?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:l?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let l=await f("/api/issues?limit=10000");if(l.ok){let h=l.data.data.map(b=>({Tanggal:b.report_date||"",Cabang:b.branch_name||"",Kategori:b.category||"",Sumber:b.source||"",Keluhan:b.complaint||"","Nama FC":b.employee_name||"","FC Spesialis":b.fc_specialist||"",Solusi:b.solution||"","Tgl Selesai":b.completion_date||"",Status:b.status||""}));T(h,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async l=>{let b=(await f("/api/branches?all=1")).data?.data||[],g=y=>{if(!y)return null;let v=y.toLowerCase(),S=b.find(k=>k.full_name.toLowerCase()===v||k.code.toLowerCase()===v||k.name.toLowerCase()===v);return S?S.id:null},c=l.map(y=>({branch_id:g(String(y.Cabang||"").trim()),report_date:String(y.Tanggal||"").trim(),category:String(y.Kategori||"").trim(),source:String(y.Sumber||"").trim(),complaint:String(y.Keluhan||"").trim(),employee_name:String(y["Nama FC"]||"").trim(),fc_specialist:String(y["FC Spesialis"]||"").trim(),solution:String(y.Solusi||"").trim(),completion_date:String(y["Tgl Selesai"]||"").trim(),status:String(y.Status||"").trim()})).filter(y=>y.report_date&&y.complaint&&y.category),m=await f("/api/issues/import",{method:"POST",body:JSON.stringify(c)});if(!m.ok)throw new Error(m.data?.error||"Import gagal")}}})}P();async function ca(a){let[e,n,i]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]),t=(e.data?.data||[]).map(r=>({value:r.id,label:r.full_name})),d=(n.data?.data||[]).map(r=>({value:r.full_name,label:r.full_name})),u=(i.data?.data||[]).filter(r=>r.role==="FC Spesialis").map(r=>({value:r.name,label:r.name})),p=r=>r&&!d.find(o=>o.value===r)?[...d,{value:r,label:r}]:d,s=r=>r&&!u.find(o=>o.value===r)?[...u,{value:r,label:r}]:u;E({container:a,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:r=>`<span title="${r||""}">${r?.length>50?r.slice(0,50)+"\u2026":r||"-"}</span>`},{key:"solution",label:"Solusi",render:r=>`<span title="${r||""}">${r?.length>40?r.slice(0,40)+"\u2026":r||"-"}</span>`},{key:"status",label:"Status",render:r=>M(r)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:r=>window.formatDate(r)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:r=>r?`<a href="${r}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async r=>{let o=new URLSearchParams(r||{}).toString(),l=await f(`/api/one-on-one?limit=10000&${o}`);if(l.ok){let h=l.data.data.map(g=>({Tanggal:g.meeting_date||"",Cabang:g.branch_name||"","Nama Karyawan":g.employee_name||"",PIC:g.pic||"",Masalah:g.problem||"",Solusi:g.solution||"",Status:g.status||"","Tgl Selesai":g.completion_date||"",Dokumen:g.document_link||""})),{downloadExcel:b}=await Promise.resolve().then(()=>(D(),U));b(h,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let r=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:o}=await Promise.resolve().then(()=>(D(),U));o(r,"Template_Import_OneOnOne")},onImport:async r=>{let o=g=>{if(!g)return null;let c=g.toLowerCase(),m=e.data?.data.find(y=>y.full_name.toLowerCase()===c||y.code.toLowerCase()===c||y.name.toLowerCase()===c);return m?m.id:null},l=g=>{if(!g)return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let c=String(g).trim();if(/^\d{4,5}$/.test(c)){let y=Number(c);if(y>2e4&&y<99999){let v=new Date(Date.UTC(1899,11,30)+y*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(c))return c.slice(0,10);let m=c.split(/[\/\-\.]/);if(m.length===3){let[y,v,S]=m.map(k=>k.trim());if(y.length===4&&v.length<=2&&S.length<=2)return`${y}-${v.padStart(2,"0")}-${S.padStart(2,"0")}`;if(S.length===4&&v.length<=2&&y.length<=2)return`${S}-${v.padStart(2,"0")}-${y.padStart(2,"0")}`}return c},h=r.map(g=>({meeting_date:l(g.Tanggal),employee_name:String(g["Nama Karyawan"]||"").trim(),branch_id:o(String(g.Cabang||"").trim()),pic:String(g.PIC||"").trim(),problem:String(g.Masalah||"").trim(),solution:String(g.Solusi||"").trim(),status:String(g.Status||"").trim(),completion_date:l(g["Tgl Selesai"]),document_link:String(g.Dokumen||"").trim()})).filter(g=>g.meeting_date&&g.employee_name&&g.branch_id),b=await f("/api/one-on-one/import",{method:"POST",body:JSON.stringify(h)});if(!b.ok)throw new Error(b.data?.error||"Import gagal")}},formFields:r=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:r?.meeting_date},{name:"branch_id",label:"Cabang",type:"select",options:t,value:r?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"select",required:!0,options:p(r?.employee_name),value:r?.employee_name},{name:"pic",label:"PIC",type:"select",options:s(r?.pic),value:r?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:r?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:r?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:r?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:r?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:r?.document_link}]})}P();async function pa(a){let[e,n,i]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]),t=(e.data?.data||[]).map(o=>({value:o.id,label:o.full_name})),d=(n.data?.data||[]).map(o=>({value:o.full_name,label:o.full_name})),u=(i.data?.data||[]).filter(o=>o.role==="FC Spesialis").map(o=>({value:o.name,label:o.name})),p=o=>o&&!d.find(l=>l.value===o)?[...d,{value:o,label:o}]:d,s=o=>o&&!u.find(l=>l.value===o)?[...u,{value:o,label:o}]:u,r=Array.from({length:5},(o,l)=>String(new Date().getFullYear()-l));E({container:a,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:o=>window.formatDate(o)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:o=>{try{let l=JSON.parse(o);return Array.isArray(l)?l.join(", "):o||"-"}catch{return o||"-"}}},{key:"score",label:"Nilai",render:o=>o!=null?`<strong>${o}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:o=>o?`<a href="${o}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:r}],exportOptions:{moduleName:"training",onExport:async o=>{let l=new URLSearchParams(o||{}).toString(),h=await f(`/api/training?limit=10000&${l}`);if(h.ok){let b=h.data.data.map(c=>{let m=c.participants||"";try{let y=JSON.parse(m);m=Array.isArray(y)?y.join(", "):m}catch{}return{Tanggal:c.training_date||"",Batch:c.batch||"",Materi:c.subject||"",Cabang:c.branch_name||"",Trainer:c.trainer||"",Peserta:m,Nilai:c.score!==null&&c.score!==void 0?c.score:"",Dokumen:c.document_link||""}}),{downloadExcel:g}=await Promise.resolve().then(()=>(D(),U));g(b,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let o=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:l}=await Promise.resolve().then(()=>(D(),U));l(o,"Template_Import_Training")},onImport:async o=>{let l=c=>{if(!c)return null;let m=c.toLowerCase(),y=e.data?.data.find(v=>v.full_name.toLowerCase()===m||v.code.toLowerCase()===m||v.name.toLowerCase()===m);return y?y.id:null},h=c=>{if(!c)return"";if(c instanceof Date&&!isNaN(c.getTime()))return c.toISOString().slice(0,10);let m=String(c).trim();if(/^\d{4,5}$/.test(m)){let v=Number(m);if(v>2e4&&v<99999){let S=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(m))return m.slice(0,10);let y=m.split(/[\/\-\.]/);if(y.length===3){let[v,S,k]=y.map(_=>_.trim());if(v.length===4&&S.length<=2&&k.length<=2)return`${v}-${S.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&S.length<=2&&v.length<=2)return`${k}-${S.padStart(2,"0")}-${v.padStart(2,"0")}`}return m},b=o.map(c=>({training_date:h(c.Tanggal),batch:String(c.Batch||"").trim(),subject:String(c.Materi||"").trim(),branch_id:l(String(c.Cabang||"").trim()),trainer:String(c.Trainer||"").trim(),participants:String(c.Peserta||"").trim(),score:c.Nilai?Number(c.Nilai):null,document_link:String(c.Dokumen||"").trim()})).filter(c=>c.training_date&&c.subject&&c.branch_id),g=await f("/api/training/import",{method:"POST",body:JSON.stringify(b)});if(!g.ok)throw new Error(g.data?.error||"Import gagal")}},formFields:o=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:o?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:o?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:o?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",options:t,value:o?.branch_id},{name:"trainer",label:"Trainer",type:"select",options:s(o?.trainer),value:o?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let l=JSON.parse(o?.participants);return Array.isArray(l)?l.join(", "):o?.participants||""}catch{return o?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:o?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:o?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],onBeforeSubmit:async o=>(o.participants&&(o.participants=JSON.stringify(o.participants.split(",").map(l=>l.trim()).filter(Boolean))),o)})}P();D();async function ma(a){let[e,n]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000")]),i=(e.data?.data||[]).map(s=>({value:s.id,label:s.full_name})),t=(n.data?.data||[]).map(s=>({value:s.full_name,label:s.full_name})),d=s=>s&&!t.find(r=>r.value===s)?[...t,{value:s,label:s}]:t,u=["Krishna Aryaan Permana","Agung Septiadi","Indra Saputro","Wariskin","Iqbal"],p=s=>{let r=u.map(o=>({value:o,label:o}));return s&&!r.find(o=>o.value===s)?[...r,{value:s,label:s}]:r};E({container:a,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",columns:[{key:"backup_date",label:"Tanggal Backup",nowrap:!0,render:s=>window.formatDate(s)},{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"FC Digantikan"},{key:"period",label:"Periode",render:s=>V(s)},{key:"reliever_name",label:"Reliefer"},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:s=>s?`<span class="badge badge-info">${s}</span>`:"-"},{key:"status",label:"Status",render:s=>M(s)}],filterFields:[{type:"search",placeholder:"Cari reliefer / FC..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:s=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:i,value:s?.branch_id},{name:"period",label:"Periode",type:"select",options:["Q1","Q2","Q3","Q4"],value:s?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"FC yang Digantikan",type:"select",options:[{value:"",label:"BELUM ADA FC"},...d(s?.original_fc_name)],value:s?.original_fc_name},{name:"reliever_name",label:"Nama Reliefer",type:"select",required:!0,options:p(s?.reliever_name),value:s?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Backup",type:"date",required:!0,value:s?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:s?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"select",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:s?.reason},{name:"shift",label:"Shift",type:"select",options:["Pagi","Siang","Full Shift","Middle"],value:s?.shift}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done","Tidak Datang"],value:s?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let s=await f("/api/relievers?limit=10000");if(s.ok){let r=s.data.data.map(o=>({"Tanggal Backup":o.backup_date||"",Cabang:o.branch_name||"","FC Digantikan":o.original_fc_name||"",Periode:o.period||"",Reliefer:o.reliever_name||"",Keterangan:o.reason||"",Shift:o.shift||"","Tanggal Selesai":o.completion_date||"",Status:o.status||""}));T(r,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Tanggal Backup":"2024-03-10",Cabang:"001. Pondok Bambu","FC Digantikan":"Budi Santoso",Periode:"Q1",Reliefer:"Andi",Keterangan:"Sakit",Shift:"Pagi","Tanggal Selesai":"2024-03-10",Status:"Done"}],"Template_Import_Reliefer")},onImport:async s=>{let o=(await f("/api/branches?all=1")).data?.data||[],l=g=>{if(!g)return null;let c=g.toLowerCase(),m=o.find(y=>y.full_name.toLowerCase()===c||y.code.toLowerCase()===c||y.name.toLowerCase()===c);return m?m.id:null},h=s.map(g=>({branch_id:l(String(g.Cabang||"").trim()),backup_date:String(g["Tanggal Backup"]||"").trim(),original_fc_name:String(g["FC Digantikan"]||"").trim(),reliever_name:String(g.Reliefer||"").trim(),period:String(g.Periode||"").trim(),reason:String(g.Keterangan||"").trim(),shift:String(g.Shift||"").trim(),completion_date:String(g["Tanggal Selesai"]||"").trim(),status:String(g.Status||"").trim()})).filter(g=>g.reliever_name&&g.backup_date),b=await f("/api/relievers/import",{method:"POST",body:JSON.stringify(h)});if(!b.ok)throw new Error(b.data?.error||"Import gagal")}}})}P();D();async function ua(a){let n=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name})),i=Array.from({length:4},(t,d)=>String(new Date().getFullYear()-d));E({container:a,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:t=>V(t)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"fc_score",label:"Point FC",render:t=>t!=null?`<strong class="${t>=80?"text-success":t>=60?"text-warning":"text-danger"}">${t}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:t=>t!=null?`<strong>${t}</strong>`:"-"},{key:"status",label:"Status",render:t=>M(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:n,value:t?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:t?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:t?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:t?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async t=>{let d=new URLSearchParams(t||{}).toString(),u=await f(`/api/reports/inspection?limit=10000&${d}`);if(u.ok){let p=u.data.data.map(s=>({Cabang:s.branch_name||"",Periode:s.period||"",Tanggal:s.inspection_date||"","Point FC":s.fc_score!==null&&s.fc_score!==void 0?s.fc_score:"","Point SPV":s.spv_score!==null&&s.spv_score!==void 0?s.spv_score:"",Status:s.status||"","Link Dokumen":s.document_link||"",Catatan:s.notes||""}));T(p,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:"Semua area bersih"}],"Template_Import_Inspeksi")},onImport:async t=>{let u=(await f("/api/branches?all=1")).data?.data||[],p=l=>{if(!l)return null;let h=l.toLowerCase(),b=u.find(g=>g.full_name.toLowerCase()===h||g.code.toLowerCase()===h||g.name.toLowerCase()===h);return b?b.id:null},s=l=>{if(l==null||l==="")return"";if(l instanceof Date&&!isNaN(l.getTime()))return l.toISOString().slice(0,10);let h=String(l).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let g=Number(h);if(g>2e4&&g<99999){let c=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(c.getTime())?"":c.toISOString().slice(0,10)}}let b=h.split(/[\/\-\.]/);if(b.length===3){let[g,c,m]=b.map(y=>y.trim());if(g.length===4&&c.length<=2&&m.length<=2)return`${g}-${c.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&c.length<=2&&g.length<=2)return`${m}-${c.padStart(2,"0")}-${g.padStart(2,"0")}`}return h},r=t.map(l=>({branch_id:p(String(l.Cabang||"").trim()),period:String(l.Periode||"").trim(),inspection_date:s(l.Tanggal),fc_score:l["Point FC"]!==void 0&&l["Point FC"]!==""?Number(l["Point FC"]):null,spv_score:l["Point SPV"]!==void 0&&l["Point SPV"]!==""?Number(l["Point SPV"]):null,status:String(l.Status||"").trim(),document_link:String(l["Link Dokumen"]||"").trim(),notes:String(l.Catatan||l.Keterangan||"").trim()})).filter(l=>l.branch_id&&l.period&&l.inspection_date),o=await f("/api/reports/inspection/import",{method:"POST",body:JSON.stringify(r)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}}})}P();D();async function ga(a){let n=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name})),i=Array.from({length:4},(t,d)=>String(new Date().getFullYear()-d));E({container:a,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:t=>`<span class="badge ${t==="Deep Cleaning"?"badge-purple":"badge-success"}">${t}</span>`},{key:"period",label:"Periode",render:t=>V(t)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"status",label:"Status",render:t=>M(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:n,value:t?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:t?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:t?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async t=>{let d=new URLSearchParams(t||{}).toString(),u=await f(`/api/reports/cleaning?limit=10000&${d}`);if(u.ok){let p=u.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||"",Catatan:s.notes||""}));T(p,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:"Pembersihan lantai"}],"Template_Import_GCDC")},onImport:async t=>{let u=(await f("/api/branches?all=1")).data?.data||[],p=l=>{if(!l)return null;let h=l.toLowerCase(),b=u.find(g=>g.full_name.toLowerCase()===h||g.code.toLowerCase()===h||g.name.toLowerCase()===h);return b?b.id:null},s=l=>{if(l==null||l==="")return"";if(l instanceof Date&&!isNaN(l.getTime()))return l.toISOString().slice(0,10);let h=String(l).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let g=Number(h);if(g>2e4&&g<99999){let c=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(c.getTime())?"":c.toISOString().slice(0,10)}}let b=h.split(/[\/\-\.]/);if(b.length===3){let[g,c,m]=b.map(y=>y.trim());if(g.length===4&&c.length<=2&&m.length<=2)return`${g}-${c.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&c.length<=2&&g.length<=2)return`${m}-${c.padStart(2,"0")}-${g.padStart(2,"0")}`}return h},r=t.map(l=>({branch_id:p(String(l.Cabang||"").trim()),activity_type:String(l.Jenis||l.Kegiatan||"").trim(),period:String(l.Periode||"").trim(),activity_date:s(l.Tanggal),status:String(l.Status||"").trim(),document_link:String(l["Link Dokumen"]||"").trim(),notes:String(l.Catatan||l.Keterangan||"").trim()})).filter(l=>l.branch_id&&l.activity_type&&l.period&&l.activity_date),o=await f("/api/reports/cleaning/import",{method:"POST",body:JSON.stringify(r)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}}})}P();D();async function ba(a){let n=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name})),i=Array.from({length:4},(t,d)=>String(new Date().getFullYear()-d));E({container:a,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:t=>`<span class="badge badge-warning">${t}</span>`},{key:"period",label:"Periode",render:t=>V(t)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"status",label:"Status",render:t=>M(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:t=>t||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:n,value:t?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:t?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async t=>{let d=new URLSearchParams(t||{}).toString(),u=await f(`/api/reports/fogging?limit=10000&${d}`);if(u.ok){let p=u.data.data.map(s=>({Cabang:s.branch_name||"",Jenis:s.activity_type||"Fogging",Periode:s.period||"",Tanggal:s.activity_date||"",Status:s.status||"","Link Dokumen":s.document_link||"",Catatan:s.notes||""}));T(p,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:"Tuntas"}],"Template_Import_Fogging")},onImport:async t=>{let u=(await f("/api/branches?all=1")).data?.data||[],p=l=>{if(!l)return null;let h=l.toLowerCase(),b=u.find(g=>g.full_name.toLowerCase()===h||g.code.toLowerCase()===h||g.name.toLowerCase()===h);return b?b.id:null},s=l=>{if(l==null||l==="")return"";if(l instanceof Date&&!isNaN(l.getTime()))return l.toISOString().slice(0,10);let h=String(l).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let g=Number(h);if(g>2e4&&g<99999){let c=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(c.getTime())?"":c.toISOString().slice(0,10)}}let b=h.split(/[\/\-\.]/);if(b.length===3){let[g,c,m]=b.map(y=>y.trim());if(g.length===4&&c.length<=2&&m.length<=2)return`${g}-${c.padStart(2,"0")}-${m.padStart(2,"0")}`;if(m.length===4&&c.length<=2&&g.length<=2)return`${m}-${c.padStart(2,"0")}-${g.padStart(2,"0")}`}return h},r=t.map(l=>({branch_id:p(String(l.Cabang||"").trim()),activity_type:String(l.Jenis||l.Kegiatan||"Fogging").trim(),period:String(l.Periode||"").trim(),activity_date:s(l.Tanggal),status:String(l.Status||"").trim(),document_link:String(l["Link Dokumen"]||"").trim(),notes:String(l.Catatan||l.Keterangan||"").trim()})).filter(l=>l.branch_id&&l.period&&l.activity_date),o=await f("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(r)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}}})}P();D();async function ha(a){let[e,n,i]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]),t=(e.data?.data||[]).map(r=>({value:r.id,label:r.full_name})),d=(n.data?.data||[]).map(r=>({value:r.full_name,label:r.full_name})),u=(i.data?.data||[]).filter(r=>r.role==="FC Spesialis").map(r=>({value:r.name,label:r.name})),p=r=>r&&!d.find(o=>o.value===r)?[...d,{value:r,label:r}]:d,s=r=>r&&!u.find(o=>o.value===r)?[...u,{value:r,label:r}]:u;E({container:a,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:r=>`<span title="${r||""}">${r?.length>60?r.slice(0,60)+"\u2026":r||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:r=>window.formatDate(r)},{key:"status",label:"Status",render:r=>M(r)},{key:"notes",label:"Keterangan",render:r=>r?.length>40?r.slice(0,40)+"\u2026":r||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan / PIC..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]}],formFields:r=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"select",required:!0,options:t,value:r?.branch_id},{name:"pic",label:"PIC",type:"select",options:s(r?.pic),value:r?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:r?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:r?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:r?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:r?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:r?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async r=>{let o=new URLSearchParams(r||{}).toString(),l=await f(`/api/reports/basecamp?limit=10000&${o}`);if(l.ok){let h=l.data.data.map(b=>({"Tgl Info":b.info_date||"",Cabang:b.branch_name||"",Permasalahan:b.problem||"",PIC:b.pic||"","Tgl Done":b.done_date||"",Status:b.status||"",Keterangan:b.notes||""}));T(h,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async r=>{let l=(await f("/api/branches?all=1")).data?.data||[],h=m=>{if(!m)return null;let y=m.toLowerCase(),v=l.find(S=>S.full_name.toLowerCase()===y||S.code.toLowerCase()===y||S.name.toLowerCase()===y);return v?v.id:null},b=m=>{if(m==null||m==="")return"";if(m instanceof Date&&!isNaN(m.getTime()))return m.toISOString().slice(0,10);let y=String(m).trim();if(y===""||y==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);if(/^\d{4,5}$/.test(y)){let S=Number(y);if(S>2e4&&S<99999){let k=new Date(Date.UTC(1899,11,30)+S*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}let v=y.split(/[\/\-\.]/);if(v.length===3){let[S,k,_]=v.map($=>$.trim());if(S.length===4&&k.length<=2&&_.length<=2)return`${S}-${k.padStart(2,"0")}-${_.padStart(2,"0")}`;if(_.length===4&&k.length<=2&&S.length<=2)return`${_}-${k.padStart(2,"0")}-${S.padStart(2,"0")}`}return y},g=r.map(m=>({info_date:b(m["Tgl Info"]||m["Tanggal Info"]),branch_id:h(String(m.Cabang||"").trim()),problem:String(m.Permasalahan||"").trim(),pic:String(m.PIC||"").trim(),done_date:b(m["Tgl Done"]||m["Tanggal Done"]),status:String(m.Status||"").trim(),notes:String(m.Keterangan||m.Catatan||"").trim()})).filter(m=>m.info_date&&m.branch_id&&m.problem),c=await f("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(g)});if(!c.ok)throw new Error(c.data?.error||"Import gagal")}}})}async function ya(a){E({container:a,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let n=new URLSearchParams(e||{}).toString(),{apiFetch:i}=await Promise.resolve().then(()=>(P(),de)),t=await i(`/api/sop?limit=10000&${n}`);if(t.ok){let d=t.data.data.map(p=>({"Nama SOP":p.name||"",Kategori:p.category||"",Dokumen:p.document_link||"",Catatan:p.notes||p.description||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(D(),U));u(d,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:n}=await Promise.resolve().then(()=>(D(),U));n(e,"Template_Import_SOP")},onImport:async e=>{let n=e.map(d=>({name:String(d["Nama SOP"]||"").trim(),category:String(d.Kategori||"").trim(),document_link:String(d.Dokumen||"").trim(),description:String(d.Catatan||"").trim()})).filter(d=>d.name),{apiFetch:i}=await Promise.resolve().then(()=>(P(),de)),t=await i("/api/sop/import",{method:"POST",body:JSON.stringify(n)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function fa(a){E({container:a,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let n=new URLSearchParams(e||{}).toString(),{apiFetch:i}=await Promise.resolve().then(()=>(P(),de)),t=await i(`/api/checklist?limit=10000&${n}`);if(t.ok){let d=t.data.data.map(p=>({"Nama Checklist":p.name||"",Kategori:p.category||"",Dokumen:p.document_link||"",Deskripsi:p.description||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(D(),U));u(d,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:n}=await Promise.resolve().then(()=>(D(),U));n(e,"Template_Import_Checklist")},onImport:async e=>{let n=e.map(d=>({name:String(d["Nama Checklist"]||"").trim(),category:String(d.Kategori||"").trim(),document_link:String(d.Dokumen||"").trim(),description:String(d.Deskripsi||"").trim()})).filter(d=>d.name),{apiFetch:i}=await Promise.resolve().then(()=>(P(),de)),t=await i("/api/checklist/import",{method:"POST",body:JSON.stringify(n)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}P();D();async function Oe(a,e="forms"){if(e==="supply")return it(a);nt(a)}function nt(a){E({container:a,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}]})}async function it(a){let n=((await f("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}));E({container:a,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:i=>i?new Date(i).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(i,t)=>t.branch_name_ref||t.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:i=>{try{let t=JSON.parse(i);return Array.isArray(t)?t.join(", "):i}catch{return i||"-"}}},{key:"chemical_items",label:"Chemical",render:i=>{try{let t=JSON.parse(i);return Array.isArray(t)?t.join(", "):i}catch{return i||"-"}}},{key:"additional_notes",label:"Catatan",render:i=>i?.length>40?i.slice(0,40)+"\u2026":i||"-"},{key:"status",label:"Status",render:i=>M(i)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:i=>{let t=i?.tools_items;try{t=Array.isArray(JSON.parse(t))?JSON.parse(t).join(", "):t}catch{}let d=i?.chemical_items;try{d=Array.isArray(JSON.parse(d))?JSON.parse(d).join(", "):d}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:i?.submitter_name},{name:"branch_id",label:"Cabang",type:"select",options:n,value:i?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:t},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:i?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:d},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:i?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:i?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:i?.status||""},{name:"processed_by",label:"Diproses Oleh",value:i?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async i=>{let t=new URLSearchParams(i||{}).toString(),d=await f(`/api/reports/supply?limit=10000&${t}`);if(d.ok){let u=d.data.data.map(p=>{let s=p.tools_items;try{s=Array.isArray(JSON.parse(s))?JSON.parse(s).join(", "):s}catch{}let r=p.chemical_items;try{r=Array.isArray(JSON.parse(r))?JSON.parse(r).join(", "):r}catch{}return{Waktu:p.submitted_at||"",Pengirim:p.submitter_name||"",Cabang:p.branch_name_ref||p.branch_name||"","Alat/Barang":s||"",Chemical:r||"",Catatan:p.additional_notes||"",Status:p.status||"","Diproses Oleh":p.processed_by||""}});T(u,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async i=>{let d=(await f("/api/branches?all=1")).data?.data||[],u=o=>{if(!o)return null;let l=o.toLowerCase(),h=d.find(b=>b.full_name.toLowerCase()===l||b.code.toLowerCase()===l||b.name.toLowerCase()===l);return h?h.id:null},p=o=>{if(o==null||o==="")return"";if(o instanceof Date&&!isNaN(o.getTime()))return o.toISOString().slice(0,10);let l=String(o).trim();if(l===""||l==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(l))return l.slice(0,10);if(/^\d{4,5}$/.test(l)){let b=Number(l);if(b>2e4&&b<99999){let g=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}let h=l.split(/[\/\-\.]/);if(h.length===3){let[b,g,c]=h.map(m=>m.trim());if(b.length===4&&g.length<=2&&c.length<=2)return`${b}-${g.padStart(2,"0")}-${c.padStart(2,"0")}`;if(c.length===4&&g.length<=2&&b.length<=2)return`${c}-${g.padStart(2,"0")}-${b.padStart(2,"0")}`}return l},s=i.map(o=>({submitted_at:p(o.Waktu||o.Tanggal),submitter_name:String(o.Pengirim||"").trim(),branch_id:u(String(o.Cabang||"").trim()),tools_items:String(o["Alat/Barang"]||o.Alat||"").trim(),chemical_items:String(o.Chemical||"").trim(),additional_notes:String(o.Catatan||o.Keterangan||"").trim(),status:String(o.Status||"").trim(),processed_by:String(o["Diproses Oleh"]||o.PIC||"").trim()})).filter(o=>o.submitted_at&&o.submitter_name&&o.branch_id),r=await f("/api/reports/supply/import",{method:"POST",body:JSON.stringify(s)});if(!r.ok)throw new Error(r.data?.error||"Import gagal")}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(i,t)=>{let d=Y({title:"Update Status Permintaan",content:`
              <div class="form-group">
                <label class="form-label">Status</label>
                <select class="form-control" id="supply-status">
                  <option value="Pending" ${i.status==="Pending"?"selected":""}>Pending</option>
                  <option value="Diproses" ${i.status==="Diproses"?"selected":""}>Diproses</option>
                  <option value="Selesai" ${i.status==="Selesai"?"selected":""}>Selesai</option>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label">Diproses Oleh</label>
                <input type="text" class="form-control" id="supply-processed-by" value="${i.processed_by||""}" placeholder="Nama">
              </div>
            `,onConfirm:async(u,p)=>{let s=u.querySelector("#supply-status").value,r=u.querySelector("#supply-processed-by").value;(await f(`/api/reports/supply/${i.id}`,{method:"PUT",body:JSON.stringify({status:s,processed_by:r})})).ok?(R("Status diperbarui."),p(),t()):H("Gagal update status.")}})}}]})}P();D();async function va(a){let e=ae();if(!e||!["superadmin","admin"].includes(e.role)){a.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}E({container:a,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:n=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[n]||"badge-neutral"}">${n}</span>`},{key:"is_active",label:"Status",render:n=>n?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:n=>n?new Date(n).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:n=>{let i=!!n;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:n?.full_name},{name:"username",label:"Username",required:!i,placeholder:"username",value:n?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!i,placeholder:"email@contoh.com",value:n?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:n?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:i?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!i,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:i?n?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let n=await f("/api/users?limit=10000");if(n.ok){let i=n.data.data.map(t=>({"Nama Lengkap":t.full_name||"",Username:t.username||"",Email:t.email||"",Role:t.role||"",Status:t.is_active?"Aktif":"Nonaktif"}));T(i,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async n=>{let i=n.map(d=>({full_name:String(d["Nama Lengkap"]||"").trim(),username:String(d.Username||"").trim(),email:String(d.Email||"").trim(),role:String(d.Role||"").trim()||"viewer",password:String(d.Password||"").trim()})).filter(d=>d.username&&d.password&&d.email&&d.full_name),t=await f("/api/users/import",{method:"POST",body:JSON.stringify(i)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}}})}P();D();async function ka(a){E({container:a,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await f("/api/branches?limit=10000");if(e.ok)T(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let n=e.map(t=>({code:String(t["Kode Cabang"]||"").trim(),name:String(t["Nama Pendek"]||"").trim(),full_name:String(t["Nama Lengkap"]||"").trim(),city:String(t.Kota||"").trim()})).filter(t=>t.code&&t.name),i=await f("/api/branches/import",{method:"POST",body:JSON.stringify(n)});if(!i.ok)throw new Error(i.data?.error||"Import gagal")}}})}P();async function Sa(a){let e=new Date,n=[];a.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),t()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),t()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(d=>d.addEventListener("change",t));async function i(){try{let d=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;n=(await f(`/api/dashboard/calendar?month=${d}`)).data?.data||[]}catch(d){console.warn("[Calendar] Failed to load events, rendering empty grid:",d),n=[]}}async function t(){let d=document.getElementById("calendar-grid");if(d){d.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await i();try{let u=e.getFullYear(),p=e.getMonth(),s=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),r=document.getElementById("cal-month-label");r&&(r.textContent=s);let o=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(k=>k.value)),l=n.filter(k=>o.has(k.type)),h={};l.forEach(k=>{let _=(k.event_date||"").slice(0,10);h[_]||(h[_]=[]),h[_].push(k)});let b=new Date(u,p,1).getDay(),g=new Date(u,p+1,0).getDate(),c=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],m=new Date().toISOString().slice(0,10),y='<div class="calendar-grid">';c.forEach(k=>{y+=`<div class="cal-day-header">${k}</div>`});for(let k=0;k<b;k++)y+='<div class="cal-cell cal-cell-empty"></div>';for(let k=1;k<=g;k++){let _=`${u}-${String(p+1).padStart(2,"0")}-${String(k).padStart(2,"0")}`,$=h[_]||[],O=_===m;y+=`
          <div class="cal-cell ${O?"cal-today":""} ${$.length?"cal-has-events":""}"
               data-date="${_}" tabindex="0" role="button" aria-label="${_}">
            <div class="cal-day-num ${O?"today-num":""}">${k}</div>
            <div class="cal-events-preview">
              ${$.slice(0,3).map(w=>`
                <div class="cal-event-dot cal-color-${w.color||"gray"}" title="${ke(w.title||w.type)}">
                  <span class="cal-event-dot-label">${st(w.title||w.branch_name||w.type,18)}</span>
                </div>
              `).join("")}
              ${$.length>3?`<div class="cal-more">+${$.length-3} lagi</div>`:""}
            </div>
          </div>`}let S=(b+g)%7;if(S!==0)for(let k=0;k<7-S;k++)y+='<div class="cal-cell cal-cell-empty"></div>';y+="</div>",d.innerHTML=y,d.querySelectorAll(".cal-cell[data-date]").forEach(k=>{k.addEventListener("click",()=>{let _=k.dataset.date,$=h[_]||[];if(!$.length)return;let O=document.getElementById("cal-event-list"),w=new Date(_+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=w,document.getElementById("cal-event-items").innerHTML=$.map(I=>`
            <div class="cal-event-item cal-color-border-${I.color||"gray"}">
              <div class="cal-event-type">${rt(I.type)}</div>
              <div class="cal-event-title">${ke(I.title||"-")}</div>
              <div class="cal-event-branch">${ke(I.branch_name||"")}</div>
              ${I.status?`<div class="cal-event-status">${ke(I.status)}</div>`:""}
              ${I.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${I.days_remaining} hari</div>`:""}
            </div>
          `).join(""),O.style.display="block"})})}catch(u){console.error("[Calendar] Render error:",u),d&&(d.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}t()}function st(a,e){return a?a.length>e?a.slice(0,e)+"\u2026":a:""}function ke(a){return a?String(a).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function rt(a){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[a]||a}P();async function _a(a){let e=ae(),n=(e?.full_name||e?.username||"U")[0].toUpperCase(),t={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";a.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">\u{1F464} Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${t},${t}99)">
            ${n}
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
  `;let d=localStorage.getItem("fm_token"),u=document.getElementById("session-info");if(d&&u)try{let p=JSON.parse(atob(d.split(".")[1])),s=new Date(p.exp*1e3);u.textContent=`Berakhir: ${s.toLocaleString("id-ID")}`}catch{u.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async p=>{p.preventDefault();let s=document.getElementById("pwd-error"),r=document.getElementById("pwd-success"),o=document.getElementById("btn-save-pwd");s.style.display="none",r.style.display="none";let l=p.target,h=l.current_password.value,b=l.new_password.value,g=l.confirm_password.value;if(b!==g){s.textContent="\u274C Konfirmasi password tidak cocok.",s.style.display="block";return}if(b.length<6){s.textContent="\u274C Password baru minimal 6 karakter.",s.style.display="block";return}o.disabled=!0,o.textContent="\u23F3 Menyimpan...";let c=await f("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:h,new_password:b})});o.disabled=!1,o.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',c.ok?(r.textContent="\u2705 Password berhasil diubah.",r.style.display="block",l.reset(),R("Password berhasil diubah.")):(s.textContent=c.data?.error||"Gagal mengubah password.",s.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}P();var Se={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function A(a){if(a==null||a==="")return null;if(a instanceof Date)return isNaN(a.getTime())?null:a.toISOString().slice(0,10);let e=String(a).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let t=Number(e);if(t>2e4&&t<99999){let d=new Date(Date.UTC(1899,11,30)+t*864e5);return isNaN(d.getTime())?null:d.toISOString().slice(0,10)}}let n=e.split(/[\/\-\.]/);if(n.length===3){let[t,d,u]=n.map(o=>o.trim()),p=Number(t),s=Number(d),r=Number(u);if(t.length===4&&p>1900)return`${t}-${d.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&r>1900)return p>12?`${u}-${d.padStart(2,"0")}-${t.padStart(2,"0")}`:s>12?`${u}-${t.padStart(2,"0")}-${d.padStart(2,"0")}`:`${u}-${d.padStart(2,"0")}-${t.padStart(2,"0")}`;if(u.length===2&&!isNaN(r)){let o=r>=50?`19${u}`:`20${u}`;return p>12?`${o}-${d.padStart(2,"0")}-${t.padStart(2,"0")}`:`${o}-${d.padStart(2,"0")}-${t.padStart(2,"0")}`}}let i=new Date(e);return isNaN(i.getTime())?null:i.toISOString().slice(0,10)}function wa(a){return Object.values(a).every(e=>e==null||String(e).trim()==="")}var lt={validation:{required:[],map:a=>({cabang:a.CABANG,pic:a.PIC,kegiatan:a.KEGIATAN,quartal:a.QUARTAL,masa_pkwt:a["MASA PKWT"],pic_pelapor:a["PIC PELAPOR"],kontrak:a.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:a=>({name:a["Nama SOP"],category:a.Kategori||"Umum",document_link:a["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:a=>({full_name:a["Nama Lengkap"],branch_name:a.Cabang,division:a["Div / Bagian"]||"FACILITY CARE",phone:a["No. Hp"],join_date:A(a["Tanggal Masuk"]),status:a.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:a=>({employee_name:a["Nama Lengkap"],branch_name:a.Cabang,division:a["Div / Bagian"]||"FACILITY CARE",start_date:A(a["Tanggal Mulai"]),end_date:A(a["Tanggal Selesai"]),contract_type:a["Tipe Kontrak"]||"",pkwt_number:a.PKWT||"",status:a.Status||"",notes:a.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:a=>({report_date:A(a["Tanggal Info"]),branch_name:a.Cabang,category:a.Kategori,source:a["Sumber Laporan"],complaint:a.Keluhan,employee_name:a["Nama FC"],fc_specialist:a["FC Spesialis"],solution:a.Solusi,status:a.Status||"",completion_date:A(a["Tanggal Selesai"])})},one_on_one:{required:[],map:a=>({meeting_date:A(a.Tanggal),branch_name:a.Cabang,employee_name:a["Nama Karyawan"],pic:a.Pic,problem:a.Masalah,solution:a.Solusi,status:a.Status||"",completion_date:A(a["Tanggal Selesai"]),document_link:a["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:a=>({branch_name:a.Cabang,activity_type:a.Kegiatan,period:a.Periode,pic:a.Pic||a.PIC,opening_date:A(a["Tanggal Opening"]||a["Tgl Opening"]),target_date:A(a["Tanggal Target"]||a["Tgl Target"]),completion_date:A(a["Tanggal Selesai"]||a["Tgl Selesai"]),status:a.Status||"",notes:a.Keterangan||a.Catatan})},inspection:{required:[],map:a=>({inspection_date:A(a.Tanggal),branch_name:a.Cabang,period:a.Periode,status:a.Status||"",fc_score:a["Point FC SP"]!==void 0&&a["Point FC SP"]!==null?parseFloat(String(a["Point FC SP"]).replace(",",".")):null,spv_score:a["Point SPV"]!==void 0&&a["Point SPV"]!==null?parseFloat(String(a["Point SPV"]).replace(",",".")):null,document_link:a.Link,notes:""})},cleaning:{required:[],map:a=>({activity_date:A(a.Tanggal),branch_name:a.Cabang,activity_type:a["Jenis Kegiatan"]||"General Cleaning",period:a.Periode,status:a.Status||"",document_link:a.Link,notes:""})},fogging:{required:[],map:a=>({activity_date:A(a.Tanggal),branch_name:a.Cabang,period:a.Periode,status:a.Status||"",document_link:a.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:a=>({info_date:A(a["Tgl Info"]),branch_name:a.Cabang,problem:a.Permasalahan,pic:a.PIC,done_date:A(a["Tgl Done"]),status:a.Status||"",notes:a.Ket})},relievers:{required:[],map:a=>({branch_name:a.Cabang,original_fc_name:a["Nama Facility care"],period:a.Periode,reliever_name:a.Relifer,backup_date:A(a["Tanggal Back Up"]),completion_date:A(a["Tanggal Selesai"]),reason:a.Keterangan,shift:a.Shift,status:a.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:a=>({training_date:A(a.Tanggal),batch:a.Batch,subject:a.Materi,participants:a.Peserta,branch_name:a.Cabang,trainer:a.Trainer,score:a.Nilai!==void 0&&a.Nilai!==null?parseFloat(String(a.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:a=>({name:a["Master Checklist"],category:"Umum",document_link:a["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:a=>({name:a["Master Form"],category:"Umum",document_link:a["Link Document"],description:""})},supply:{required:[],map:a=>({submitted_at:A(a.Timestamp),submitter_name:a["Nama Lengkap"],branch_name:a["Kebutuhan Untuk Cabang"],tools_items:a["Alat - Alat / Barang"],tools_quantity:a["Jumlah Permintaan Alat / Barang"],chemical_items:a.Chemical,chemical_quantity:a["Jumlah Permintaan Chemical"],additional_notes:a["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:a.Status||""})}};function ot(a,e){let n=Se[a];if(!n)return{valid:[],errors:[],mapped:[],skipped:!0};let i=lt[n.module];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let t=[],d=[],u=[];return e.filter(s=>!wa(s)).forEach((s,r)=>{let o=e.indexOf(s)+2,l=[];i.required.forEach(({key:b,label:g})=>{let c=s[b];if(c==null||String(c).trim()===""){let m=Object.keys(s).filter(y=>y.trim()).join(", ");l.push({column:g,originalValue:c||"",reason:`Kolom "${g}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${m.slice(0,120)}`})}});let h=i.map(s);l.length>0?d.push({row:o,data:h,raw:s,errors:l}):(t.push(s),u.push(h))}),{valid:t,errors:d,mapped:u}}function Ca(a){let e=[];return a.SheetNames.forEach(n=>{let i=Se[n];if(!i)return;let t=a.Sheets[n],d=window.XLSX.utils.sheet_to_json(t,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),u=ot(n,d),p=d.filter(s=>!wa(s));e.push({sheetName:n,module:i.module,label:i.label,total:p.length,valid:u.mapped.length,errorCount:u.errors.length,errors:u.errors,mapped:u.mapped,skipped:!1})}),e}function xa(){let a=window.XLSX,e=a.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu","NAMA KARYAWAN":"Budi Santoso",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02",Day:"1"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06",Day:"1","Link Document":""}]}).forEach(([i,t])=>{a.utils.book_append_sheet(e,a.utils.json_to_sheet(t),i)}),a.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function Ta(a){let e=window.XLSX,n=e.utils.book_new(),i=!1;return a.forEach(t=>{if(!t.errors||t.errors.length===0)return;i=!0;let d=t.errors.map(p=>({"No. Baris":p.row,"Kolom Gagal":(p.errors||[]).map(s=>s.column||s).join("; "),"Alasan Error":(p.errors||[]).map(s=>s.reason||s).join("; "),...Object.fromEntries(Object.entries(p.data||{}).map(([s,r])=>[s,r??""]))})),u=e.utils.json_to_sheet(d);e.utils.book_append_sheet(n,u,t.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),i?(e.writeFile(n,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var dt=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function Ea(a){a.innerHTML=`
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
              ${Object.entries(Se).map(([c,{label:m}])=>`<span class="import-sheet-tag">\u{1F4C4} ${c} \u2192 ${m}</span>`).join("")}
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
  `;let e=null,n=null,i=0,t={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function d(c){Object.entries(t).forEach(([m,y])=>{y.style.display=m===c?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let c=document.getElementById("btn-backup-db");c.disabled=!0,c.textContent="\u23F3 Memproses Backup...";try{let m=await f("/api/import/backup");if(m.ok){let y=new Blob([JSON.stringify(m.data,null,2)],{type:"application/json"}),v=URL.createObjectURL(y),S=document.createElement("a");S.href=v,S.download=`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.json`,document.body.appendChild(S),S.click(),document.body.removeChild(S),URL.revokeObjectURL(v),R("Backup berhasil diunduh!")}else H("Gagal memproses backup: "+(m.data?.error||"Unknown error"))}catch(m){H("Gagal memproses backup: "+m.message)}finally{c.disabled=!1,c.textContent="\u{1F4E6} Backup Database"}});let u=document.getElementById("btn-sync-google");u&&u.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let c=u.innerHTML;u.innerHTML='<span class="spinner"></span> Menyinkronkan...',u.disabled=!0;try{let m=await f("/api/sync/google-sheets",{method:"POST"});m.ok?alert("Sinkronisasi Berhasil: "+(m.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(m.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{u.innerHTML=c,u.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{xa(),R("Template Excel berhasil didownload!")});let p=document.getElementById("file-input"),s=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",c=>{c.stopPropagation(),p.click()}),p.addEventListener("change",c=>{c.target.files[0]&&r(c.target.files[0])}),s.addEventListener("dragover",c=>{c.preventDefault(),s.classList.add("drag-over")}),s.addEventListener("dragleave",()=>s.classList.remove("drag-over")),s.addEventListener("drop",c=>{c.preventDefault(),s.classList.remove("drag-over");let m=c.dataTransfer.files[0];m&&m.name.match(/\.xlsx?$/i)?r(m):H("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,p.value="",document.getElementById("file-info").style.display="none",s.style.display="",d("upload")});async function r(c){e=c,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${c.name} (${(c.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",s.style.display="none",await o(c)}async function o(c){d("validating");let m=document.getElementById("validation-status"),y=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");m.textContent="Membaca file Excel...",y.style.width="20%",await me(200);let v=await c.arrayBuffer(),S=window.XLSX.read(v,{type:"array",cellDates:!0});m.textContent=`Memvalidasi ${S.SheetNames.length} sheet...`,y.style.width="50%",await me(100),n=Ca(S),y.style.width="100%",m.textContent="Validasi selesai!",await me(300),l()}catch(v){d("upload"),H("Gagal memproses file: "+v.message),document.getElementById("file-info").style.display="flex",s.style.display="none"}}function l(){d("preview");let c=n.filter(w=>!w.skipped).length,m=n.reduce((w,I)=>w+I.total,0),y=n.reduce((w,I)=>w+I.valid,0),v=n.reduce((w,I)=>w+I.errorCount,0),S=m>0?Math.round(y/m*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${c} sheet</span>
      <span class="badge badge-secondary">${m} baris</span>
      <span class="badge badge-success">${y} valid (${S}%)</span>
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
          ${n.map((w,I)=>`
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
    `,k.querySelectorAll(".btn-detail-error").forEach(w=>{w.addEventListener("click",()=>{let I=n[Number(w.dataset.idx)];h(I)})});let _=document.getElementById("error-detail-section"),$=document.getElementById("error-detail-container");$.innerHTML="",_.style.display="none";let O=document.getElementById("btn-start-import");y===0?(O.disabled=!0,O.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(O.disabled=!1,v>0?(O.innerHTML=`\u{1F680} Import ${y} Data Valid (${v} dilewati)`,O.title="Baris error akan dilewati, baris valid tetap diimport"):O.innerHTML=`\u{1F680} Mulai Import ${y} Data`)}function h(c){let m=document.getElementById("error-detail-section"),y=document.getElementById("error-detail-container");m.style.display="";let v=c.errors.slice(0,100).map(S=>(Array.isArray(S.errors)?S.errors:[]).map(_=>{let $=typeof _=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${S.row}</span></td>
            <td><strong>${$?_.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${$&&_.originalValue!==void 0?_.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${$?_.reason:_}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${$&&_.aliases?`Gunakan salah satu nama kolom:<br><em>${_.aliases}</em>`:$&&_.hint?_.hint:""}
            </td>
          </tr>
        `}).join("")).join("");y.innerHTML=`
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          \u{1F4C4} ${c.sheetName} \u2014 ${c.errorCount} baris error dari ${c.total} total
          ${c.errors.length>100?'<span style="font-weight:400">(menampilkan 100 pertama)</span>':""}
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
        ${c.errors.length>100?`
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>`:""}
      </div>
    `,m.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{d("upload"),document.getElementById("file-info").style.display="none",s.style.display="",e=null,p.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!n)return;Ta(n)?R("Log error berhasil didownload."):R("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let c=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";b(c)});async function b(c){d("importing"),i=Date.now();let m=[];dt.forEach(_=>{let $=n?.find(O=>O.module===_&&O.mapped?.length>0);$&&m.push($)});let y=document.getElementById("import-steps-list");y.innerHTML=m.map(_=>`
      <div class="import-step-item" id="step-item-${_.module}">
        <span class="step-item-icon" id="step-icon-${_.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${_.label} <span class="step-item-count">(${_.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${_.module}"></span>
      </div>
    `).join("");let v=document.getElementById("import-bar"),S=document.getElementById("import-current-status"),k={totalSheets:m.length,totalRows:m.reduce((_,$)=>_+$.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let _=0;_<m.length;_++){let $=m[_],O=document.getElementById(`step-icon-${$.module}`),w=document.getElementById(`step-status-${$.module}`);O.textContent="\u{1F504}",w.textContent="Mengimport...",S.textContent=`Mengimport ${$.label}...`,v.style.width=`${Math.round(_/m.length*100)}%`;try{let I=await f(`/api/import/${$.module}`,{method:"POST",body:JSON.stringify({rows:$.mapped,onDuplicate:c})});if(I.ok){let Q=I.data;k.inserted+=Q.inserted||0,k.skipped+=Q.skipped||0,k.moduleResults.push({label:$.label,inserted:Q.inserted||0,skipped:Q.skipped||0,status:"ok"}),O.textContent="\u2705",w.innerHTML=`<span class="badge badge-success">${Q.inserted||0} berhasil</span>${Q.skipped>0?` <span class="badge badge-neutral">${Q.skipped} skip</span>`:""}`}else k.failed++,k.moduleResults.push({label:$.label,inserted:0,skipped:0,status:"error",error:I.data?.error}),O.textContent="\u274C",w.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(I){k.failed++,k.moduleResults.push({label:$.label,inserted:0,skipped:0,status:"error",error:I.message}),O.textContent="\u274C",w.innerHTML='<span class="badge badge-danger">Gagal</span>'}await me(150)}v.style.width="100%",S.textContent="Selesai!",await me(400),g(k)}function g(c){d("summary");let m=((Date.now()-i)/1e3).toFixed(1),y=c.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${y?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${y?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
    `,document.getElementById("summary-stats").innerHTML=`
      <div class="summary-stat-card">
        <div class="stat-value">${c.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${c.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${c.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${c.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${c.failed>0?`<div class="summary-stat-card danger"><div class="stat-value">${c.failed}</div><div class="stat-label">Modul Gagal</div></div>`:""}
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
          ${c.moduleResults.map(v=>`
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,n=null,p.value="",document.getElementById("file-info").style.display="none",s.style.display="",d("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function me(a){return new Promise(e=>setTimeout(e,a))}P();var Me=[];async function $a(a){Me=((await f("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name})),E({container:a,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:n=>n?new Date(n).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"branch_name",label:"Cabang"},{key:"sp_type",label:"Jenis SP",render:n=>`<span class="badge badge-warning">${n||"-"}</span>`},{key:"status",label:"Status",render:n=>`<span class="badge ${n==="Aktif"?"badge-danger":"badge-success"}">${n||"-"}</span>`},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:Me}],exportOptions:{moduleName:"sp_data",onExport:async n=>{let i=new URLSearchParams(n||{}).toString(),t=await f(`/api/sp?limit=10000&${i}`);if(t.ok){let d=t.data.data.map(p=>({Tanggal:p.tanggal||"","Nama Karyawan":p.employee_name||"",Cabang:p.branch_name||"","Jenis SP":p.sp_type||"",Status:p.status||"",Dokumen:p.document_link||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(D(),U));u(d,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08","Nama Karyawan":"Budi Santoso",Cabang:"001. Pondok Bambu","Jenis SP":"SP 1",Status:"Aktif",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(D(),U));i(n,"Template_Import_SP")},onImport:async n=>{let t=(await f("/api/branches?all=1")).data?.data||[],d=r=>{if(!r)return null;let o=r.toLowerCase(),l=t.find(h=>h.full_name.toLowerCase()===o||h.code.toLowerCase()===o||h.name.toLowerCase()===o);return l?l.id:null},u=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let o=String(r).trim();if(/^\d{4,5}$/.test(o)){let h=Number(o);if(h>2e4&&h<99999){let b=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(o))return o.slice(0,10);let l=o.split(/[\/\-\.]/);if(l.length===3){let[h,b,g]=l.map(c=>c.trim());if(h.length===4&&b.length<=2&&g.length<=2)return`${h}-${b.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&b.length<=2&&h.length<=2)return`${g}-${b.padStart(2,"0")}-${h.padStart(2,"0")}`}return o},p=n.map(r=>({tanggal:u(r.Tanggal),employee_name:String(r["Nama Karyawan"]||"").trim(),branch_id:d(String(r.Cabang||"").trim()),sp_type:String(r["Jenis SP"]||"").trim(),status:String(r.Status||"").trim(),document_link:String(r.Dokumen||"").trim()})).filter(r=>r.tanggal&&r.employee_name&&r.branch_id),s=await f("/api/sp/import",{method:"POST",body:JSON.stringify(p)});if(!s.ok)throw new Error(s.data?.error||"Import gagal")}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"text",name:"employee_name",label:"Nama Karyawan",required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:Me},{type:"select",name:"sp_type",label:"Jenis Surat Peringatan",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"select",name:"status",label:"Status",required:!0,options:["Aktif","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}P();var ue=[];async function Pa(a){ue=((await f("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name})),E({container:a,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:n=>n?new Date(n).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:n=>`<span class="badge ${n==="Selesai"?"badge-success":"badge-warning"}">${n||"-"}</span>`},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:ue},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:ue}],exportOptions:{moduleName:"mutasi_data",onExport:async n=>{let i=new URLSearchParams(n||{}).toString(),t=await f(`/api/mutasi?limit=10000&${i}`);if(t.ok){let d=t.data.data.map(p=>({Tanggal:p.tanggal||"","Nama Karyawan":p.employee_name||"","Cabang Asal":p.from_branch_name||"","Cabang Tujuan":p.to_branch_name||"",Status:p.status||"",Dokumen:p.document_link||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(D(),U));u(d,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(D(),U));i(n,"Template_Import_Mutasi")},onImport:async n=>{let t=(await f("/api/branches?all=1")).data?.data||[],d=r=>{if(!r)return null;let o=r.toLowerCase(),l=t.find(h=>h.full_name.toLowerCase()===o||h.code.toLowerCase()===o||h.name.toLowerCase()===o);return l?l.id:null},u=r=>{if(!r)return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let o=String(r).trim();if(/^\d{4,5}$/.test(o)){let h=Number(o);if(h>2e4&&h<99999){let b=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(o))return o.slice(0,10);let l=o.split(/[\/\-\.]/);if(l.length===3){let[h,b,g]=l.map(c=>c.trim());if(h.length===4&&b.length<=2&&g.length<=2)return`${h}-${b.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&b.length<=2&&h.length<=2)return`${g}-${b.padStart(2,"0")}-${h.padStart(2,"0")}`}return o},p=n.map(r=>({tanggal:u(r.Tanggal),employee_name:String(r["Nama Karyawan"]||"").trim(),from_branch_id:d(String(r["Cabang Asal"]||"").trim()),to_branch_id:d(String(r["Cabang Tujuan"]||"").trim()),status:String(r.Status||"").trim(),document_link:String(r.Dokumen||"").trim()})).filter(r=>r.tanggal&&r.employee_name&&r.from_branch_id&&r.to_branch_id),s=await f("/api/mutasi/import",{method:"POST",body:JSON.stringify(p)});if(!s.ok)throw new Error(s.data?.error||"Import gagal")}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"text",name:"employee_name",label:"Nama Karyawan",required:!0},{type:"select",name:"from_branch_id",label:"Cabang Asal",required:!0,options:ue},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:ue},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}window.formatDate=a=>{if(!a||a==="-")return"";let e=a.split("-");return e.length===3&&e[0].length===4?`${e[2]}-${e[1]}-${e[0]}`:a};function N(a){return async e=>{if(!re()){te("/login");return}return a(e)}}var ge=null;function ct(){ge&&clearInterval(ge);let a=()=>{let e=new Date,n=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),i=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),t=document.getElementById("header-clock-time"),d=document.getElementById("header-clock-date");t&&(t.textContent=n),d&&(d.textContent=i)};a(),ge=setInterval(a,1e3)}async function pt(){try{let a=await f("/api/dashboard/kpi");if(!a.ok)return;let e=a.data?.data||a.data||{},n=(i,t)=>{let d=document.getElementById(i);d&&(d.textContent=t>0?t:"",d.style.display=t>0?"inline-flex":"none")};n("badge-issues",e.issues?.current||0),n("badge-contracts",e.expiring30?.current||0),n("badge-oo1",e.one_on_one?.current||0),n("badge-schedule",e.schedule?.current||0),n("badge-supply",e.supply?.current||0)}catch{}}var ie=[];async function mt(){try{let a=await f("/api/dashboard/notifications");if(!a.ok)return;ie=a.data?.data||a.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=ie.length>0?"block":"none",e.textContent=ie.length)}catch{}}function ut(){if(!ie.length){Y({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,n)=>n()});return}let a=`
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${ie.map(e=>`
        <div class="notif-item notif-severity-${e.severity||"info"}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"primary"}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${e.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>\u{1F4C5} ${e.date}</span>
            <span class="badge badge-${e.severity==="danger"?"danger":e.severity==="warning"?"warning":"info"}">${e.type.toUpperCase()}</span>
          </div>
        </div>
      `).join("")}
    </div>
  `;Y({title:`Notifikasi (${ie.length})`,content:a,confirmText:"Tutup",onConfirm:(e,n)=>n()})}function Ia(){let a=ae(),e=(a?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
  `;let n=document.getElementById("sidebar"),i=document.getElementById("sidebar-overlay"),t=document.getElementById("topbar-menu-btn"),d=document.getElementById("sidebar-close"),u=()=>{n.classList.add("open"),i.classList.add("show")},p=()=>{n.classList.remove("open"),i.classList.remove("show")};t?.addEventListener("click",u),d?.addEventListener("click",p),i?.addEventListener("click",p),document.querySelectorAll(".nav-item").forEach(r=>r.addEventListener("click",p));function s(){let r=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(h=>{let b=h.dataset.route;h.classList.toggle("active",r===b||b!=="/dashboard"&&r.startsWith(b))});let o=document.getElementById("topbar-title"),l=document.querySelector(".nav-item.active .nav-label");o&&l&&(o.textContent=l.textContent)}window.addEventListener("hashchange",s),s(),ct(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await f("/api/auth/logout",{method:"POST"}),le(),ge&&clearInterval(ge),te("/login")}),pt(),mt(),document.getElementById("btn-notif")?.addEventListener("click",r=>{r.preventDefault(),ut()})}async function gt(){B("/login",({main:e})=>Xe(e)),B("/dashboard",N(({main:e})=>ze(e))),B("/calendar",N(({main:e})=>Sa(e))),B("/employees",N(({main:e})=>sa(e))),B("/contracts",N(({main:e})=>la(e))),B("/sp",N(({main:e})=>$a(e))),B("/mutasi",N(({main:e})=>Pa(e))),B("/timeline",N(({main:e})=>oa(e))),B("/issues",N(({main:e})=>da(e))),B("/one-on-one",N(({main:e})=>ca(e))),B("/training",N(({main:e})=>pa(e))),B("/relievers",N(({main:e})=>ma(e))),B("/reports/inspection",N(({main:e})=>ua(e))),B("/reports/cleaning",N(({main:e})=>ga(e))),B("/reports/fogging",N(({main:e})=>ba(e))),B("/reports/basecamp",N(({main:e})=>ha(e))),B("/reports/supply",N(({main:e})=>Oe(e,"supply"))),B("/sop",N(({main:e})=>ya(e))),B("/checklist",N(({main:e})=>fa(e))),B("/forms",N(({main:e})=>Oe(e))),B("/users",N(({main:e})=>va(e))),B("/branches",N(({main:e})=>ka(e))),B("/profile",N(({main:e})=>_a(e))),B("/settings/import",N(({main:e})=>Ea(e)));let a=re();if(!a&&window.location.hash!=="#/login"&&te("/login"),a){let e=await f("/api/auth/me");e.ok?(oe(e.data.data),Ia()):(le(),te("/login"))}window.addEventListener("fm:login",()=>{Ia(),te("/dashboard")}),qe()}gt();
