var qa=Object.defineProperty;var Je=(a,e)=>()=>(a&&(e=a(a=0)),e);var Ge=(a,e)=>{for(var n in e)qa(a,n,{get:e[n],enumerable:!0})};var me={};Ge(me,{API:()=>Qe,apiFetch:()=>f,clearToken:()=>de,getToken:()=>ce,getUser:()=>ie,setToken:()=>Ee,setUser:()=>pe});function ce(){return localStorage.getItem("fm_token")}function Ee(a){localStorage.setItem("fm_token",a)}function de(){localStorage.removeItem("fm_token"),localStorage.removeItem("fm_user")}function ie(){try{return JSON.parse(localStorage.getItem("fm_user")||"null")}catch{return null}}function pe(a){localStorage.setItem("fm_user",JSON.stringify(a))}async function f(a,e={}){let n=ce(),i={"Content-Type":"application/json",...n?{Authorization:`Bearer ${n}`}:{},...e.headers||{}};try{let t=`cb=${Date.now()}`,c=a.includes("?")?"&":"?",u=`${Qe}${a}${c}${t}`,m=await fetch(u,{...e,headers:i}),o;try{let l=await m.text();try{o=JSON.parse(l)}catch{o={error:`Server Error (${m.status}): ${l.substring(0,80)}...`}}}catch{o={error:"Gagal membaca respon dari server"}}return m.status===401&&(de(),window.location.hash="#/login"),{ok:m.ok,status:m.status,data:o}}catch(t){return{ok:!1,status:0,data:{error:`Koneksi terputus. Periksa jaringan Anda. (${t.message})`}}}}var ja,Qe,P=Je(()=>{ja="",Qe=ja});var J={};Ge(J,{downloadExcel:()=>T,parseExcel:()=>Be,renderExcelButtons:()=>Ne});function Be(a){return new Promise((e,n)=>{let i=new FileReader;i.onload=t=>{try{let c=new Uint8Array(t.target.result),u=XLSX.read(c,{type:"array"}),m=u.SheetNames[0],o=u.Sheets[m],l=XLSX.utils.sheet_to_json(o,{defval:""});e(l)}catch(c){n(c)}},i.onerror=t=>n(t),i.readAsArrayBuffer(a)})}function T(a,e){try{let n=XLSX.utils.json_to_sheet(a),i=XLSX.utils.book_new();XLSX.utils.book_append_sheet(i,n,"Data"),XLSX.writeFile(i,`${e}.xlsx`)}catch(n){throw console.error("Error generating Excel file:",n),n}}function Ne(a){return`
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
  `}var L=Je(()=>{});P();var Pe={},ke=null;function B(a,e){Pe[a]=e}function se(a){window.location.hash=a}function Ve(){async function a(){let e=window.location.hash.replace("#","")||"/dashboard",[n,...i]=e.split("?"),t=Pe[n];if(!t){for(let[u,m]of Object.entries(Pe))if(u.endsWith("/*")&&n.startsWith(u.slice(0,-2))){t=m;break}}ke&&(ke(),ke=null);let c=document.getElementById("main-content");if(c&&(c.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>'),t){let u=new URLSearchParams(i.join("?")),m=n.split("/").filter(Boolean),o=await t({path:n,params:u,segments:m,main:c});o&&(ke=o)}else{let u=c||document.getElementById("app");u&&(u.innerHTML='<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>')}}window.addEventListener("hashchange",a),a()}var ue;function Ha(){return ue||(ue=document.createElement("div"),ue.id="toast-container",document.body.appendChild(ue)),ue}function ze(a,e="info",n=3500){let i=Ha(),t=document.createElement("div");t.className=`toast toast-${e}`;let c={success:"\u2713",error:"\u2715",warning:"\u26A0",info:"\u2139"};t.innerHTML=`<span class="toast-icon">${c[e]||"\u2139"}</span><span class="toast-msg">${a}</span>`,i.appendChild(t),requestAnimationFrame(()=>t.classList.add("show")),setTimeout(()=>{t.classList.remove("show"),setTimeout(()=>t.remove(),350)},n)}var K=a=>ze(a,"success"),j=a=>ze(a,"error");function ee({title:a,content:e,onConfirm:n,onCancel:i,confirmText:t="Simpan",cancelText:c="Batal",size:u="md",confirmClass:m="btn-primary"}){let o={sm:"400px",md:"560px",lg:"720px",xl:"900px"},l=document.createElement("div");l.className="modal-overlay",l.innerHTML=`
    <div class="modal" style="max-width:${o[u]||o.md}">
      <div class="modal-header">
        <h3 class="modal-title">${a}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof e=="string"?e:""}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${c}</button>
        ${n?`<button class="btn ${m} modal-confirm">${t}</button>`:""}
      </div>
    </div>
  `,e instanceof HTMLElement&&l.querySelector(".modal-body").appendChild(e);let s=()=>{l.classList.remove("show"),setTimeout(()=>l.remove(),250)};return l.querySelector(".modal-close").addEventListener("click",()=>{i&&i(),s()}),l.querySelector(".modal-cancel").addEventListener("click",()=>{i&&i(),s()}),n&&l.querySelector(".modal-confirm").addEventListener("click",()=>n(l,s)),l.addEventListener("click",r=>{r.target===l&&(i&&i(),s())}),document.body.appendChild(l),requestAnimationFrame(()=>l.classList.add("show")),{overlay:l,close:s}}function We(a,e,n="Konfirmasi"){return ee({title:n,content:`<p>${a}</p>`,onConfirm:(i,t)=>{e(),t()},confirmText:"Ya, Lanjutkan",confirmClass:"btn-danger"})}P();var ne={};function ge(a){if(ne[a]){try{ne[a].destroy()}catch{}delete ne[a]}}function Ua(){Object.keys(ne).forEach(ge)}var H=(a,e=0)=>{let n=Number(a);return isNaN(n)||a===null||a===void 0?e:n},U=(a,e="\u2014")=>{if(a==null||a==="")return e;let n=String(a).trim();return n===""||n==="[object Object]"?e:n},Xe=a=>{if(!a)return"\u2014";try{let e=new Date(a);return isNaN(e)?U(a):e.toLocaleDateString("id-ID",{day:"2-digit",month:"short",year:"numeric"})}catch{return U(a)}},Ja=a=>{if(!a)return"";try{let e=Date.now()-new Date(a).getTime();if(e<0)return"Baru saja";let n=Math.floor(e/6e4);if(n<1)return"Baru saja";if(n<60)return`${n} menit lalu`;let i=Math.floor(n/60);return i<24?`${i} jam lalu`:`${Math.floor(i/24)} hari lalu`}catch{return""}},Ze=a=>{if(!a||typeof a!="string")return"";try{let[e,n]=a.split("-");return new Date(Number(e),Number(n)-1).toLocaleDateString("id-ID",{month:"short",year:"2-digit"})}catch{return a}};function ea(a,e,n=900){if(!a)return;let i=Math.max(0,Math.round(H(e)));if(i===0){a.textContent="0";return}let t=Date.now(),c=()=>{let u=Math.min((Date.now()-t)/n,1),m=1-Math.pow(1-u,3);a.textContent=Math.round(m*i).toLocaleString("id-ID"),u<1?requestAnimationFrame(c):a.textContent=i.toLocaleString("id-ID")};requestAnimationFrame(c)}function Ga(a,e){if(a=H(a),e=H(e),e===0)return"";let n=a-e,i=Math.abs(Math.round(n/e*100));return n>0?`<span class="kpi-trend up">\u25B2 ${i}%</span>`:n<0?`<span class="kpi-trend down">\u25BC ${i}%</span>`:'<span class="kpi-trend neutral">= Sama</span>'}var Qa={Done:"pill-success",Aktif:"pill-success",Selesai:"pill-success",Open:"pill-danger",Pending:"pill-warning","In Progress":"pill-info","Tidak Aktif":"pill-neutral",Resign:"pill-neutral",Cut:"pill-neutral"},aa=a=>{let e=U(a,"\u2014");return`<span class="status-pill ${Qa[e]||"pill-neutral"}">${e}</span>`},Va=a=>{let e=H(a,999);return e<=7?`<span class="days-badge days-critical">${e} hari</span>`:e<=14?`<span class="days-badge days-warning">${e} hari</span>`:e<=30?`<span class="days-badge days-soon">${e} hari</span>`:`<span class="days-badge days-ok">${e} hari</span>`},za={issue:{emoji:"\u26A0\uFE0F",dot:"dot-danger",label:"Permasalahan"},contract:{emoji:"\u{1F4C4}",dot:"dot-info",label:"Kontrak"},employee:{emoji:"\u{1F464}",dot:"dot-success",label:"Karyawan"},one_on_one:{emoji:"\u{1F91D}",dot:"dot-purple",label:"One on One"},training:{emoji:"\u{1F393}",dot:"dot-primary",label:"Training"},supply:{emoji:"\u{1F4E6}",dot:"dot-warning",label:"Permintaan Barang"},reliever:{emoji:"\u{1F504}",dot:"dot-teal",label:"Reliefer"},inspection:{emoji:"\u{1F50D}",dot:"dot-blue",label:"Laporan Inspeksi"}},Wa=a=>za[a]||{emoji:"\u{1F4CC}",dot:"dot-neutral",label:U(a,"Aktivitas")},Q={family:"Inter",size:11},re="#94A3B8",Se="#F1F5F9",Ya=["#2563EB","#10B981","#F59E0B","#EF4444","#8B5CF6","#0EA5E9","#F97316","#14B8A6","#6366F1","#EC4899"],Xa=()=>window.innerWidth<768;function De(a={}){return{responsive:!0,maintainAspectRatio:!1,animation:{duration:700,easing:"easeOutQuart"},plugins:{legend:{position:Xa()?"bottom":"top",labels:{font:Q,color:"#64748B",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{mode:"index",intersect:!1,bodyFont:Q,titleFont:{...Q,weight:"700"}}},scales:{x:{grid:{color:Se},ticks:{font:Q,color:re,maxRotation:0}},y:{grid:{color:Se},ticks:{font:Q,color:re},beginAtZero:!0}},...a}}var Za=()=>Array(5).fill(0).map(()=>`
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
  </div>`).join("");function Ye(a=3){return Array(a).fill(0).map((e,n)=>`<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${n<a-1?"6px":"0"};border-radius:6px"></div>`).join("")}function at(){return Array(5).fill(0).map(()=>`
    <div class="activity-item">
      <div class="skeleton" style="width:34px;height:34px;border-radius:10px;flex-shrink:0"></div>
      <div style="flex:1">
        <div class="skeleton skeleton-text" style="width:65%;height:13px;margin-bottom:5px"></div>
        <div class="skeleton skeleton-text" style="width:35%;height:11px"></div>
      </div>
    </div>`).join("")}async function ae(a,e,n=8e3){try{let i=new AbortController,t=setTimeout(()=>i.abort(),n),c=await f(a,{signal:i.signal}).catch(()=>null);if(clearTimeout(t),!c||!c.ok)return e;let u=c.data;return u?u.data!==void 0?u.data??e:u:e}catch{return e}}function tt(){["skel-donut","skel-trend","skel-insp","skel-contract"].forEach(i=>{let t=document.getElementById(i);t&&(t.style.display="none")}),["chart-donut","chart-trend","chart-insp","chart-contract"].forEach(i=>{let t=document.getElementById(i);if(t&&t.style.display==="none"){t.style.display="block";let c=t.parentElement;if(c&&!c.querySelector(".chart-empty")){let u=document.createElement("div");u.className="chart-empty",u.textContent="Belum ada data",t.style.display="none",c.appendChild(u)}}});let a=document.getElementById("kpi-row");a&&a.querySelector(".skeleton")&&na({});let e=document.getElementById("mini-stats-row");e&&e.querySelector(".skeleton")&&ia({}),["table-contracts","table-issues"].forEach(i=>{let t=document.getElementById(i);t&&t.querySelector(".skeleton")&&(t.innerHTML='<div class="chart-empty">Belum ada data</div>')});let n=document.getElementById("activity-log");n&&n.querySelector(".skeleton")&&(n.innerHTML='<div class="chart-empty">Belum ada aktivitas</div>')}async function ta(a){Ua(),a._dashRefresh&&clearInterval(a._dashRefresh),a._skelTimeout&&clearTimeout(a._skelTimeout),a.innerHTML=`
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
        <div id="activity-log">${at()}</div>
      </div>

    </div>
  `,document.getElementById("btn-dash-refresh")?.addEventListener("click",()=>Ie(a)),a._skelTimeout=setTimeout(()=>tt(),5e3),await Ie(a),a._dashRefresh=setInterval(()=>{document.getElementById("dash-root")?Ie(a):clearInterval(a._dashRefresh)},6e4)}async function Ie(a){a._skelTimeout&&(clearTimeout(a._skelTimeout),a._skelTimeout=null);let[e,n,i,t,c,u,m,o]=await Promise.all([ae("/api/dashboard/kpi",{},8e3),ae("/api/dashboard/issues-trend",{},8e3),ae("/api/dashboard/contracts-chart",{},8e3),ae("/api/dashboard/issues-summary",{},8e3),ae("/api/dashboard/inspection-bar",{},8e3),ae("/api/dashboard/contracts-expiring",[],8e3),ae("/api/dashboard/stats",{},8e3),ae("/api/dashboard/activity-log",[],8e3)]);try{na(e)}catch(s){console.warn("KPI render:",s)}try{ia(e)}catch(s){console.warn("MiniStats render:",s)}try{nt(Array.isArray(t?.by_category)?t.by_category:[])}catch(s){console.warn("Donut render:",s),te("skel-donut","chart-donut")}try{it(n)}catch(s){console.warn("Trend render:",s),te("skel-trend","chart-trend")}try{st(c)}catch(s){console.warn("InspBar render:",s),te("skel-insp","chart-insp")}try{rt(i)}catch(s){console.warn("ContractBar render:",s),te("skel-contract","chart-contract")}try{let s=Array.isArray(u)?u:[];lt(s)}catch(s){console.warn("ContractsTable render:",s)}try{let s=Array.isArray(m)?m:Array.isArray(m?.recent_issues)?m.recent_issues:[];ot(s)}catch(s){console.warn("IssuesTable render:",s)}try{ct(Array.isArray(o)?o:[])}catch(s){console.warn("ActivityLog render:",s)}let l=document.getElementById("dash-updated");l&&(l.textContent=`Diperbarui: ${new Date().toLocaleTimeString("id-ID")}`)}function na(a){let e=document.getElementById("kpi-row");if(!e)return;a=a||{};let n=[{icon:"\u{1F465}",label:"Karyawan Aktif",sub:"Total karyawan aktif",href:"#/employees",color:"kpi-blue",key:"employees"},{icon:"\u{1F4C4}",label:"Kontrak Aktif",sub:"Kontrak yang masih berjalan",href:"#/contracts",color:"kpi-green",key:"contracts"},{icon:"\u23F0",label:"Kontrak Habis 30 Hari",sub:"",href:"#/contracts",color:"",key:"expiring30",warn:!0},{icon:"\u26A0\uFE0F",label:"Permasalahan Open",sub:"Belum diselesaikan",href:"#/issues",color:"",key:"issues",warnIfGT0:!0},{icon:"\u{1F91D}",label:"One on One Pending",sub:"Menunggu tindak lanjut",href:"#/one-on-one",color:"kpi-purple",key:"one_on_one"}];e.innerHTML=n.map(i=>{let t=H(a[i.key]?.current,0),c=a[i.key]?.prev,u=c!=null?Ga(t,c):"",m=i.color||"",o=i.sub||"";return i.warn&&(m=t>0?"kpi-amber":"kpi-green",o=t>0?`\u26A0\uFE0F ${t} kontrak segera berakhir`:"\u2705 Semua kontrak aman"),i.warnIfGT0&&(m=t>0?"kpi-red":"kpi-green"),`
      <a href="${i.href}" class="kpi-card ${m}" style="text-decoration:none">
        <div class="kpi-card-top">
          <div class="kpi-icon-wrap"><span class="kpi-icon-emoji">${i.icon}</span></div>
          ${u}
        </div>
        <div class="kpi-value" data-target="${t}">0</div>
        <div class="kpi-label">${i.label}</div>
        <div class="kpi-subtitle">${o}</div>
      </a>`}).join(""),e.querySelectorAll(".kpi-value").forEach(i=>ea(i,parseInt(i.dataset.target)||0))}function ia(a){let e=document.getElementById("mini-stats-row");if(!e)return;a=a||{};let n=[{icon:"\u{1F4C5}",label:"Jadwal Pending",val:a.schedule?.current,href:"#/schedule",color:"mini-blue"},{icon:"\u{1F393}",label:"Training Bulan Ini",val:a.training_month?.current,href:"#/training",color:"mini-indigo"},{icon:"\u{1F4E6}",label:"Permintaan Barang",val:a.supply?.current,href:"#/reports/supply",color:"mini-orange"},{icon:"\u{1F50D}",label:"Inspeksi Bulan Ini",val:a.inspection_month?.current,href:"#/reports/inspection",color:"mini-teal"},{icon:"\u{1F9F9}",label:"GC/DC Bulan Ini",val:a.cleaning_month?.current,href:"#/reports/cleaning",color:"mini-green"},{icon:"\u{1F99F}",label:"Fogging Bulan Ini",val:a.fogging_month?.current,href:"#/reports/fogging",color:"mini-purple"},{icon:"\u{1F3E2}",label:"Total Cabang",val:a.branches?.current,href:"#/branches",color:"mini-gray"}];e.innerHTML=n.map(i=>`
    <a href="${i.href}" class="mini-stat ${i.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${i.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${H(i.val)}">0</div>
        <div class="mini-stat-label">${i.label}</div>
      </div>
    </a>`).join(""),e.querySelectorAll(".mini-stat-value").forEach(i=>ea(i,parseInt(i.dataset.target)||0,700))}function nt(a){te("skel-donut","chart-donut");let e=document.getElementById("chart-donut");if(!e)return;ge("donut");let n=(a||[]).filter(c=>H(c.count)>0);if(!n.length){_e(e,"Belum ada data permasalahan");return}let i=n.map(c=>U(c.category,"Lainnya")),t=n.map(c=>H(c.count));ne.donut=new Chart(e,{type:"doughnut",data:{labels:i,datasets:[{data:t,backgroundColor:Ya.slice(0,n.length),borderWidth:2,borderColor:"#fff",hoverBorderColor:"#fff"}]},options:{responsive:!0,maintainAspectRatio:!1,animation:{duration:700},plugins:{legend:{position:window.innerWidth<768?"bottom":"right",labels:{font:Q,color:"#475569",usePointStyle:!0,padding:10,boxWidth:8,boxHeight:8}},tooltip:{bodyFont:Q,titleFont:{...Q,weight:"700"},callbacks:{label:c=>` ${c.label}: ${c.parsed} kasus`}}},cutout:"65%"}})}function it(a){te("skel-trend","chart-trend");let e=document.getElementById("chart-trend");if(!e)return;ge("trend"),a=a||{};let n=(a.labels||[]).map(Ze),i=(a.open||[]).map(c=>H(c)),t=(a.closed||[]).map(c=>H(c));if(!n.length){_e(e,"Belum ada data trend");return}ne.trend=new Chart(e,{type:"line",data:{labels:n,datasets:[{label:"Open",data:i,borderColor:"#EF4444",backgroundColor:"rgba(239,68,68,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#EF4444",borderWidth:2},{label:"Closed",data:t,borderColor:"#10B981",backgroundColor:"rgba(16,185,129,.08)",fill:!0,tension:.4,pointRadius:3,pointHoverRadius:5,pointBackgroundColor:"#10B981",borderWidth:2}]},options:De({plugins:{legend:{position:"top"}}})})}function st(a){te("skel-insp","chart-insp");let e=document.getElementById("chart-insp");if(!e)return;ge("inspBar"),a=a||{};let n=a.labels||[],i=(a.fc||[]).map(c=>H(c)),t=(a.spv||[]).map(c=>H(c));if(!n.length){_e(e,"Belum ada data inspeksi");return}ne.inspBar=new Chart(e,{type:"bar",data:{labels:n,datasets:[{label:"Skor FC",data:i,backgroundColor:"rgba(37,99,235,.75)",borderRadius:4,borderSkipped:!1},{label:"Skor SPV",data:t,backgroundColor:"rgba(16,185,129,.75)",borderRadius:4,borderSkipped:!1}]},options:De({plugins:{legend:{position:"top"}},scales:{x:{grid:{display:!1},ticks:{font:Q,color:re,maxRotation:45,minRotation:30}},y:{grid:{color:Se},ticks:{font:Q,color:re},min:0,max:100}}})})}function rt(a){te("skel-contract","chart-contract");let e=document.getElementById("chart-contract");if(!e)return;ge("contractBar"),a=a||{};let n=(a.labels||[]).map(Ze),i=(a.counts||[]).map(c=>H(c));if(!n.length){_e(e,"Belum ada data kontrak");return}let t=i.map(c=>c>5?"rgba(239,68,68,.75)":c>2?"rgba(245,158,11,.75)":"rgba(37,99,235,.65)");ne.contractBar=new Chart(e,{type:"bar",data:{labels:n,datasets:[{label:"Kontrak Berakhir",data:i,backgroundColor:t,borderRadius:6,borderSkipped:!1}]},options:De({plugins:{legend:{display:!1}},scales:{x:{grid:{display:!1},ticks:{font:Q,color:re}},y:{grid:{color:Se},ticks:{font:Q,color:re,precision:0},beginAtZero:!0}}})})}function lt(a){let e=document.getElementById("table-contracts");if(!e)return;let n=(a||[]).filter(i=>H(i.days_remaining,999)<=30).slice(0,10);if(!n.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada kontrak yang habis dalam 30 hari</div>';return}e.innerHTML=`
    <table class="dash-table">
      <thead><tr>
        <th>#</th><th>Nama Karyawan</th><th>Cabang</th><th>Berakhir</th><th>Sisa</th><th>Status</th>
      </tr></thead>
      <tbody>${n.map((i,t)=>`
        <tr>
          <td class="td-num">${t+1}</td>
          <td><strong>${U(i.emp_name||i.employee_name)}</strong></td>
          <td class="td-branch">${U(i.branch_name)}</td>
          <td style="white-space:nowrap;font-size:.8rem">${Xe(i.end_date)}</td>
          <td>${Va(i.days_remaining)}</td>
          <td>${aa(i.status)}</td>
        </tr>`).join("")}
      </tbody>
    </table>`}function ot(a){let e=document.getElementById("table-issues");if(!e)return;let n=(a||[]).slice(0,8);if(!n.length){e.innerHTML='<div class="chart-empty">\u2705 Tidak ada permasalahan terbuka</div>';return}e.innerHTML=`
    <table class="dash-table">
      <thead><tr>
        <th>Tanggal</th><th>Keluhan</th><th>Cabang</th><th>Kategori</th><th>Status</th>
      </tr></thead>
      <tbody>${n.map(i=>`
        <tr>
          <td style="white-space:nowrap;font-size:.78rem">${Xe(i.report_date)}</td>
          <td class="td-complaint" title="${U(i.complaint)}">${U(i.complaint)}</td>
          <td class="td-branch">${U(i.branch_name)}</td>
          <td><span class="category-tag">${U(i.category)}</span></td>
          <td>${aa(i.status)}</td>
        </tr>`).join("")}
      </tbody>
    </table>`}function ct(a){let e=document.getElementById("activity-log");if(!e)return;let n=(a||[]).slice(0,15);if(!n.length){e.innerHTML='<div class="chart-empty">Belum ada aktivitas tercatat</div>';return}e.innerHTML=`<div class="activity-list">${n.map(i=>{let t=Wa(i.type),c=U(i.label),u=i.branch?` \u2022 ${U(i.branch)}`:"",m=Ja(i.created_at);return`
      <div class="activity-item">
        <div class="activity-dot ${t.dot}">${t.emoji}</div>
        <div class="activity-body">
          <div class="activity-text"><strong>${t.label}</strong> \u2014 ${c}${u}</div>
          <div class="activity-time">${m||"\u2014"}</div>
        </div>
      </div>`}).join("")}</div>`}function te(a,e){let n=document.getElementById(a),i=document.getElementById(e);n&&(n.style.display="none",n.style.position=""),i&&(i.style.display="block")}function _e(a,e="Belum ada data"){if(!a)return;a.style.display="none";let n=a.parentElement;if(!n)return;if(!n.querySelector(".chart-empty")){let t=document.createElement("div");t.className="chart-empty",t.textContent=e,n.appendChild(t)}}P();async function sa(a){document.getElementById("app").innerHTML=`
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
  `;let e=document.getElementById("login-form"),n=document.getElementById("login-error"),i=document.getElementById("login-btn"),t=document.getElementById("toggle-password"),c=document.getElementById("login-password");t?.addEventListener("click",()=>{let u=c.type==="text";c.type=u?"password":"text",t.style.color=u?"":"var(--primary)"}),e?.addEventListener("submit",async u=>{u.preventDefault(),n.style.display="none";let m=e.username.value.trim(),o=e.password.value;if(!m||!o){n.textContent="Username dan password wajib diisi.",n.style.display="block";return}i.querySelector(".btn-text").style.display="none",i.querySelector(".btn-spinner").style.display="",i.disabled=!0;try{let l=await f("/api/auth/login",{method:"POST",body:JSON.stringify({username:m,password:o})});l.ok&&l.data.success?(Ee(l.data.data.token),pe(l.data.data.user),K("Login berhasil! Selamat datang \u{1F44B}"),window.dispatchEvent(new Event("fm:login"))):(n.textContent=l.data.error||"Username atau password salah.",n.style.display="block",i.classList.add("shake"),setTimeout(()=>i.classList.remove("shake"),600))}catch{n.textContent="Gagal terhubung ke server. Periksa koneksi internet.",n.style.display="block"}finally{i.querySelector(".btn-text").style.display="",i.querySelector(".btn-spinner").style.display="none",i.disabled=!1}})}P();function ra({columns:a,data:e,onEdit:n,onDelete:i,onView:t,actions:c=[],emptyText:u="Tidak ada data",bulkSelect:m=null}){let o=document.createElement("div");if(o.className="table-wrapper",!e||e.length===0)return o.innerHTML=`<div class="empty-state"><p>${u}</p></div>`,o;let l=document.createElement("table");l.className="data-table";let s=document.createElement("thead"),r=document.createElement("tr");if(m){let b=document.createElement("th");b.style.width="40px",b.style.textAlign="center";let g=document.createElement("input");g.type="checkbox",g.id="select-all-checkbox",g.title="Pilih semua",g.addEventListener("change",()=>{e.forEach(d=>{g.checked?m.selectedIds.add(d.id):m.selectedIds.delete(d.id)}),o.querySelectorAll(".row-checkbox").forEach(d=>d.checked=g.checked),m.onToggle()}),b.appendChild(g),r.appendChild(b)}if(a.forEach(b=>{let g=document.createElement("th");g.textContent=b.label,b.width&&(g.style.width=b.width),r.appendChild(g)}),n||i||t||c.length>0){let b=document.createElement("th");b.textContent="Aksi",b.style.width="120px",r.appendChild(b)}s.appendChild(r),l.appendChild(s);let h=document.createElement("tbody");return e.forEach(b=>{let g=document.createElement("tr");if(m){let d=document.createElement("td");d.style.textAlign="center",d.style.width="40px";let p=document.createElement("input");p.type="checkbox",p.className="row-checkbox",p.checked=m.selectedIds.has(b.id),p.addEventListener("change",()=>{if(p.checked)m.selectedIds.add(b.id);else{m.selectedIds.delete(b.id);let y=document.getElementById("select-all-checkbox");y&&(y.checked=!1)}m.onToggle()}),d.appendChild(p),g.appendChild(d)}if(a.forEach(d=>{let p=document.createElement("td");if(d.render){let y=d.render(b[d.key],b);y instanceof HTMLElement?p.appendChild(y):p.innerHTML=y||""}else p.textContent=b[d.key]!==null&&b[d.key]!==void 0&&b[d.key]!==""?b[d.key]:"";d.nowrap&&(p.style.whiteSpace="nowrap"),g.appendChild(p)}),n||i||t||c.length>0){let d=document.createElement("td");d.className="actions-cell";let p=document.createElement("div");if(p.className="btn-group",t){let y=document.createElement("button");y.className="btn btn-xs btn-ghost",y.innerHTML="\u{1F441}",y.title="Lihat",y.addEventListener("click",()=>t(b)),p.appendChild(y)}if(n){let y=document.createElement("button");y.className="btn btn-xs btn-secondary",y.innerHTML="\u270F\uFE0F",y.title="Edit",y.addEventListener("click",()=>n(b)),p.appendChild(y)}c.forEach(y=>{let v=document.createElement("button");v.className=`btn btn-xs ${y.class||"btn-ghost"}`,v.innerHTML=y.icon||y.label,v.title=y.label,v.addEventListener("click",()=>y.handler(b)),p.appendChild(v)}),d.appendChild(p),g.appendChild(d)}h.appendChild(g)}),l.appendChild(h),o.appendChild(l),o}function la({page:a,pages:e,total:n,limit:i,onPage:t}){if(e<=1)return null;let c=document.createElement("div");c.className="pagination";let u=document.createElement("span");u.className="pagination-info",u.textContent=`Total: ${n} data`,c.appendChild(u);let m=document.createElement("div");m.className="pagination-btns";let o=(r,h,b=!1,g=!1)=>{let d=document.createElement("button");d.className=`btn btn-sm ${g?"btn-primary":"btn-ghost"} pagination-btn`,d.textContent=r,d.disabled=b,d.addEventListener("click",()=>t(h)),m.appendChild(d)};o("\xAB",1,a===1),o("\u2039",a-1,a===1);let l=Math.max(1,a-2),s=Math.min(e,a+2);for(let r=l;r<=s;r++)o(r,r,!1,r===a);return o("\u203A",a+1,a===e),o("\xBB",e,a===e),c.appendChild(m),c}function Le(a){return a.map(e=>{if(e.type==="hidden")return`<input type="hidden" name="${e.name}" value="${e.value||""}">`;if(e.type==="row")return`<div class="form-row">${Le(e.fields)}</div>`;let n=e.required?"required":"",i=e.label?`<label class="form-label">${e.label}${e.required?' <span class="required">*</span>':""}</label>`:"",t="";switch(e.type){case"textarea":t=`<textarea name="${e.name}" class="form-control" placeholder="${e.placeholder||""}" ${n} rows="${e.rows||3}">${e.value||""}</textarea>`;break;case"select":let u=(e.options||[]).map(s=>{let r=typeof s=="object"?s.value:s,h=typeof s=="object"?s.label:s,b=e.value==r?"selected":"";return`<option value="${r}" ${b}>${h}</option>`}).join("");t=`<select name="${e.name}" class="form-control" ${n}><option value="">-- Pilih ${e.label||""} --</option>${u}</select>`;break;case"combobox":let m=`dl-${e.name}-${Math.random().toString(36).substring(7)}`,o=(e.options||[]).map(s=>{let r=typeof s=="object"?s.value:s;return`<option value="${typeof s=="object"?s.label:s}"></option>`}).join(""),l=e.value||"";if(e.value){let s=(e.options||[]).find(r=>(typeof r=="object"?r.value:r)==e.value);s&&(l=typeof s=="object"?s.label:s)}t=`
          <input type="text" name="${e.name}" list="${m}" class="form-control" value="${l}" placeholder="Pilih atau ketik baru..." ${n} autocomplete="off">
          <datalist id="${m}">${o}</datalist>
        `;break;case"checkbox":t=`<label class="checkbox-label"><input type="checkbox" name="${e.name}" value="1" ${e.value?"checked":""}> ${e.checkLabel||e.label}</label>`;break;case"date":t=`<input type="date" name="${e.name}" class="form-control" value="${e.value||""}" ${n}>`;break;case"number":t=`<input type="number" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" min="${e.min||""}" max="${e.max||""}" step="${e.step||"1"}" ${n}>`;break;case"email":t=`<input type="email" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${n}>`;break;case"url":t=`<input type="url" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||"https://..."}" ${n}>`;break;default:t=`<input type="${e.type||"text"}" name="${e.name}" class="form-control" value="${e.value||""}" placeholder="${e.placeholder||""}" ${n} autocomplete="off">`}let c=e.hint?`<div class="form-hint">${e.hint}</div>`:"";return`<div class="form-group ${e.class||""}">${i}${t}${c}</div>`}).join("")}function oa(a){let e={},n=new FormData(a);for(let[i,t]of n.entries())e[i]=t===""?null:t;return a.querySelectorAll("input[type=checkbox]").forEach(i=>{i.checked||(e[i.name]=null)}),e}function ca(a,e){e&&Object.entries(e).forEach(([n,i])=>{let t=a.querySelector(`[name="${n}"]`);t&&(t.type==="checkbox"?t.checked=!!i:t.value=i??"")})}L();function $({container:a,title:e,icon:n,apiPath:i,columns:t,formFields:c,filterFields:u,defaultFilters:m={},itemLabel:o="Data",canCreate:l=!0,canEdit:s=!0,canDelete:r=!0,onBeforeSubmit:h,onAfterLoad:b,extraActions:g=[],initialSearch:d="",exportOptions:p=null,bulkDelete:y=!1}){let v=1,S={...m};d&&(S.search=d);let k=new Set;a.innerHTML=`
    <div class="page-header">
      <h1 class="page-title">${n} ${e}</h1>
      <div class="page-actions">
        ${l?`<button class="btn btn-primary" id="btn-create">+ Tambah ${o}</button>`:""}
      </div>
    </div>

    ${y?`
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
  `;function _(){if(!document.getElementById("bulk-toolbar"))return;let x=document.getElementById("bulk-count"),D=document.getElementById("btn-bulk-delete"),R=document.getElementById("btn-bulk-cancel");x.textContent=`${k.size} item dipilih`,k.size>0?(D.disabled=!1,R.disabled=!1):(D.disabled=!0,R.disabled=!0)}document.getElementById("btn-bulk-cancel")?.addEventListener("click",()=>{k.clear(),document.querySelectorAll(".row-checkbox").forEach(x=>x.checked=!1);let C=document.getElementById("select-all-checkbox");C&&(C.checked=!1),_()}),document.getElementById("btn-bulk-delete")?.addEventListener("click",()=>{if(k.size===0)return;let C=[...k],x=document.createElement("div");x.style.cssText="position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center",x.innerHTML=`
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">\u26A0\uFE0F Hapus ${C.length} ${o}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${C.length} Data</button>
        </div>
      </div>
    `,document.body.appendChild(x),x.querySelector("#bulk-cancel-btn").addEventListener("click",()=>x.remove()),x.querySelector("#bulk-confirm-btn").addEventListener("click",async()=>{let D=x.querySelector("#bulk-confirm-btn");D.disabled=!0,D.textContent="Menghapus...";let R=await f(`${i}/bulk`,{method:"DELETE",body:JSON.stringify({ids:C})});x.remove(),R.ok?(K(`${C.length} ${o} berhasil dihapus.`),k.clear(),_(),w()):j(R.data?.error||"Gagal menghapus data.")})});let E=document.getElementById("filter-search"),A;if(E?.addEventListener("input",C=>{clearTimeout(A),A=setTimeout(()=>{S.search=C.target.value,v=1,w()},400)}),u?.forEach(C=>{C.type==="select"&&document.getElementById(`filter-${C.name}`)?.addEventListener("change",x=>{S[C.name]=x.target.value,v=1,w()})}),document.getElementById("btn-reset-filter")?.addEventListener("click",()=>{S={...m},E&&(E.value=""),u?.forEach(C=>{let x=document.getElementById(`filter-${C.name}`);x&&(x.value="")}),v=1,w()}),document.getElementById("btn-create")?.addEventListener("click",()=>V(null)),p){document.getElementById(`btn-export-${p.moduleName}`)?.addEventListener("click",async x=>{let D=x.target,R=D.innerHTML;D.innerHTML="\u23F3 Loading...",D.disabled=!0;try{await p.onExport()}catch{j("Gagal export data")}finally{D.innerHTML=R,D.disabled=!1}}),document.getElementById(`btn-template-${p.moduleName}`)?.addEventListener("click",()=>{p.onTemplate()});let C=document.getElementById(`input-import-${p.moduleName}`);C?.addEventListener("change",async x=>{let D=x.target.files[0];if(!D)return;let R=C.parentElement,G=R.innerHTML;R.innerHTML="\u23F3 Memproses...",R.style.pointerEvents="none",C.disabled=!0;try{let X=await Be(D);if(X.length===0)throw new Error("File kosong atau format salah");await p.onImport(X),K("Import berhasil!"),w()}catch(X){j(X.message||"Gagal import data")}finally{R.innerHTML=G,R.style.pointerEvents="auto",C.disabled=!1,C.value=""}})}async function w(){let C=document.getElementById("table-container");if(!C)return;C.innerHTML='<div class="loading-spinner"><div class="spinner"></div></div>';let x=new URLSearchParams({page:v,limit:20,...Object.fromEntries(Object.entries(S).filter(([,F])=>F))}),D=await f(`${i}?${x}`);if(!D.ok){C.innerHTML=`<div class="empty-state"><p class="text-danger">Gagal memuat data: ${D.data?.error||"Error"}</p></div>`;return}let R=D.data?.data||[],G=D.data?.pagination;b&&b(R);let X=ra({columns:t,data:R,onEdit:s?F=>V(F):null,actions:g.map(F=>({...F,handler:oe=>F.handler(oe,w)})),emptyText:`Tidak ada ${o.toLowerCase()}`,bulkSelect:y?{selectedIds:k,onToggle:_}:null});C.innerHTML="",C.appendChild(X);let z=document.getElementById("pagination-container");if(z&&(z.innerHTML="",G&&G.pages>1)){let F=la({page:G.page,pages:G.pages,total:G.total,limit:G.limit,onPage:oe=>{v=oe,w()}});F&&z.appendChild(F)}}function I(C){let x=typeof c=="function"?c(C):c;return Le(x)}function V(C){let x=!!C,D=document.createElement("form");if(D.noValidate=!0,D.innerHTML=I(C),x){let G=typeof c=="function"?c(C):c;ca(D,C)}let{close:R}=ee({title:x?`Edit ${o}`:`Tambah ${o}`,content:D,size:"lg",confirmText:x?"Simpan Perubahan":`Tambah ${o}`,onConfirm:async(G,X)=>{if(!D.reportValidity())return;let z=G.querySelector(".modal-confirm");z.disabled=!0,z.textContent="Menyimpan...";let F=oa(D),oe=typeof c=="function"?c(C):c,He=async $e=>{for(let q of $e)if(q.type==="row")await He(q.fields);else if(q.type==="combobox"&&F[q.name]){let fe=F[q.name],ve=(q.options||[]).find(W=>{let Z=String(typeof W=="object"?W.value:W),Ka=String(typeof W=="object"?W.label:W);return Z===fe||Ka===fe});if(ve)F[q.name]=typeof ve=="object"?ve.value:ve;else if(q.createApi){let W={};W[q.createApi.field]=fe,q.createApi.extra&&Object.assign(W,q.createApi.extra);let Z=await f(q.createApi.path,{method:"POST",body:JSON.stringify(W)});if(Z.ok&&Z.data?.id)F[q.name]=Z.data.id;else if(Z.ok&&!Z.data?.id)F[q.name]=fe;else throw new Error(`Gagal membuat master data: ${Z.data?.error||"Unknown error"}`)}}};try{await He(oe)}catch($e){j($e.message),z.disabled=!1,z.textContent=x?"Simpan Perubahan":`Tambah ${o}`;return}h&&(F=await h(F,C));let Fa=x?"PUT":"POST",Ra=x?`${i}/${C.id}`:i,Ue=await f(Ra,{method:Fa,body:JSON.stringify(F)});Ue.ok?(K(x?`${o} berhasil diperbarui.`:`${o} berhasil ditambahkan.`),X(),w()):(j(Ue.data?.error||"Gagal menyimpan data."),z.disabled=!1,z.textContent=x?"Simpan Perubahan":`Tambah ${o}`)}})}function Ct(C){We(`Hapus ${o} ini? Tindakan tidak dapat dibatalkan.`,async()=>{let x=await f(`${i}/${C.id}`,{method:"DELETE"});x.ok?(K(`${o} berhasil dihapus.`),w()):j(x.data?.error||"Gagal menghapus.")},`Hapus ${o}`)}return w(),w}P();function O(a){let e={Done:"badge-success",Aktif:"badge-success",Open:"badge-warning","In Progress":"badge-info",Pending:"badge-warning",Diproses:"badge-info",Selesai:"badge-success","Tidak Aktif":"badge-neutral",Resign:"badge-neutral",Cut:"badge-danger","Tidak Datang":"badge-danger"};return!a||a==="-"||String(a).trim()===""?"":`<span class="badge ${e[a]||"badge-neutral"}">${a}</span>`}function da(a){return a==null?'<span class="badge badge-neutral">-</span>':a<0?`<span class="badge badge-danger">Expired (${Math.abs(a)}h)</span>`:a<=14?`<span class="badge badge-danger">${a} hari</span>`:a<=30?`<span class="badge badge-warning">${a} hari</span>`:`<span class="badge badge-success">${a} hari</span>`}function we(a){return`<span class="badge ${{"FACILITY CARE":"badge-info",SECURITY:"badge-secondary"}[a]||"badge-neutral"}">${a||"-"}</span>`}function pa(a){return`<span class="badge ${{"Inspeksi Hygiene & Aset Bangunan":"badge-info","General Cleaning":"badge-success","Deep Cleaning":"badge-purple",Fogging:"badge-warning"}[a]||"badge-neutral"}">${a||"-"}</span>`}function Y(a){return`<span class="badge ${{Q1:"badge-info",Q2:"badge-success",Q3:"badge-warning",Q4:"badge-danger"}[a]||"badge-neutral"}">${a||"-"}</span>`}L();var Ae=[],Oe=[];async function dt(){Oe=(await f("/api/branches?all=1")).data?.data||[],Ae=Oe.map(e=>({value:e.id,label:e.full_name}))}async function ma(a){await dt(),$({container:a,title:"Karyawan",icon:"\u{1F465}",apiPath:"/api/employees",itemLabel:"Karyawan",bulkDelete:!0,columns:[{key:"full_name",label:"Nama Lengkap"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:e=>we(e)},{key:"phone",label:"No. HP",render:e=>e?`<a href="tel:${e}">${e}</a>`:"-"},{key:"join_date",label:"Tgl Masuk",render:e=>window.formatDate(e)},{key:"status",label:"Status",render:e=>O(e)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:Ae},{type:"select",name:"division",label:"Divisi",options:["FACILITY CARE","SECURITY"]},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif","Resign","Cut"]}],formFields:e=>[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap karyawan",value:e?.full_name},{name:"phone",label:"No. HP",placeholder:"08xx-xxxx-xxxx",value:e?.phone}]},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:Ae,createApi:{path:"/api/branches",field:"full_name"},value:e?.branch_id},{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:e?.division||"FACILITY CARE"}]},{type:"row",fields:[{name:"join_date",label:"Tanggal Masuk",type:"date",value:e?.join_date},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif","Resign","Cut"],value:e?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:e?.notes}],exportOptions:{moduleName:"employees",onExport:async()=>{let e=await f("/api/employees?limit=10000");if(e.ok){let n=e.data.data.map(i=>({"Nama Lengkap":i.full_name,Cabang:i.branch_name||"",Divisi:i.division||"","No. HP":i.phone||"","Tgl Masuk":i.join_date||"",Status:i.status||"",Catatan:i.notes||""}));T(n,"Data_Karyawan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","No. HP":"08123456789","Tgl Masuk":"2024-01-15",Status:"Aktif",Catatan:""},{"Nama Lengkap":"Andi Saputra",Cabang:"002. Bintaro",Divisi:"SECURITY","No. HP":"08987654321","Tgl Masuk":"2023-11-01",Status:"Aktif",Catatan:""}],"Template_Import_Karyawan")},onImport:async e=>{let n=c=>{if(!c)return null;let u=c.toLowerCase(),m=Oe.find(o=>o.full_name.toLowerCase()===u||o.code.toLowerCase()===u||o.name.toLowerCase()===u);return m?m.id:null},i=e.map(c=>({full_name:String(c["Nama Lengkap"]||"").trim(),branch_id:n(String(c.Cabang||"").trim()),division:String(c.Divisi||"").trim()||"FACILITY CARE",phone:String(c["No. HP"]||"").trim(),join_date:String(c["Tgl Masuk"]||"").trim(),status:String(c.Status||"").trim(),notes:String(c.Catatan||"").trim()})).filter(c=>c.full_name),t=await f("/api/employees/import",{method:"POST",body:JSON.stringify(i)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}}})}P();L();var Me=[],ua=[];async function pt(){let[a,e]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000&status=Aktif")]);Me=(a.data?.data||[]).map(n=>({value:n.id,label:n.full_name})),ua=(e.data?.data||[]).map(n=>({value:n.id,label:n.full_name}))}async function ga(a){await pt(),$({container:a,title:"Data Kontrak",icon:"\u{1F4CB}",apiPath:"/api/contracts",bulkDelete:!0,itemLabel:"Kontrak",columns:[{key:"employee_name",label:"Nama Karyawan"},{key:"branch_name",label:"Cabang"},{key:"division",label:"Divisi",render:e=>we(e)},{key:"start_date",label:"Tgl Mulai",nowrap:!0,render:e=>window.formatDate(e)},{key:"end_date",label:"Tgl Selesai",nowrap:!0,render:e=>window.formatDate(e)},{key:"days_remaining",label:"Sisa",render:e=>da(e)},{key:"contract_type",label:"Tipe Kontrak"},{key:"pkwt_number",label:"PKWT"},{key:"status",label:"Status",render:e=>O(e)}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:Me},{type:"select",name:"status",label:"Status",options:["Aktif","Tidak Aktif"]},{type:"select",name:"expiring_days",label:"Akan Habis",options:[{value:"7",label:"7 Hari"},{value:"14",label:"14 Hari"},{value:"30",label:"30 Hari"},{value:"60",label:"60 Hari"}]}],formFields:e=>[{type:"row",fields:[{name:"employee_id",label:"Karyawan",type:"select",required:!0,options:ua,value:e?.employee_id},{name:"branch_id",label:"Cabang",type:"combobox",options:Me,createApi:{path:"/api/branches",field:"full_name"},value:e?.branch_id}]},{type:"row",fields:[{name:"division",label:"Divisi",type:"select",required:!0,options:["FACILITY CARE","SECURITY"],value:e?.division||"FACILITY CARE"},{name:"status",label:"Status",type:"select",required:!0,options:["Aktif","Tidak Aktif"],value:e?.status||""}]},{type:"row",fields:[{name:"start_date",label:"Tanggal Mulai",type:"date",required:!0,value:e?.start_date},{name:"end_date",label:"Tanggal Selesai",type:"date",required:!0,value:e?.end_date}]},{type:"row",fields:[{name:"contract_type",label:"Tipe Kontrak",type:"select",options:["KONTRAK 6 BULAN","KONTRAK 1 TAHUN","KONTRAK 2 TAHUN"],value:e?.contract_type},{name:"pkwt_number",label:"No. PKWT",type:"select",options:["PKWT 1","PKWT 2","PKWT 3","PKWT 4","PKWT 5","PKWT 6"],value:e?.pkwt_number}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:e?.notes}],exportOptions:{moduleName:"contracts",onExport:async()=>{let e=await f("/api/contracts?limit=10000");if(e.ok){let n=e.data.data.map(i=>({"Nama Karyawan":i.employee_name,Cabang:i.branch_name||"",Divisi:i.division||"","Tgl Mulai":i.start_date||"","Tgl Selesai":i.end_date||"","Tipe Kontrak":i.contract_type||"",PKWT:i.pkwt_number||"",Status:i.status||"",Catatan:i.notes||""}));T(n,"Data_Kontrak")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Nama Karyawan":"Budi Santoso",Cabang:"001. Pondok Bambu",Divisi:"FACILITY CARE","Tgl Mulai":"2024-01-01","Tgl Selesai":"2024-12-31","Tipe Kontrak":"KONTRAK 1 TAHUN",PKWT:"PKWT 1",Status:"Aktif",Catatan:""}],"Template_Import_Kontrak")},onImport:async e=>{let[n,i]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000")]),t=n.data?.data||[],c=i.data?.data||[],u=s=>{if(!s)return null;let r=s.toLowerCase(),h=t.find(b=>b.full_name.toLowerCase()===r||b.code.toLowerCase()===r||b.name.toLowerCase()===r);return h?h.id:null},m=s=>{if(!s)return null;let r=s.toLowerCase(),h=c.find(b=>b.full_name.toLowerCase()===r);return h?h.id:null},o=e.map(s=>({employee_id:m(String(s["Nama Karyawan"]||"").trim()),branch_id:u(String(s.Cabang||"").trim()),division:String(s.Divisi||"").trim()||"FACILITY CARE",start_date:String(s["Tgl Mulai"]||"").trim(),end_date:String(s["Tgl Selesai"]||"").trim(),contract_type:String(s["Tipe Kontrak"]||"").trim(),pkwt_number:String(s.PKWT||"").trim(),status:String(s.Status||"").trim(),notes:String(s.Catatan||"").trim()})).filter(s=>s.employee_id&&s.start_date&&s.end_date),l=await f("/api/contracts/import",{method:"POST",body:JSON.stringify(o)});if(!l.ok)throw new Error(l.data?.error||"Import gagal")}}})}P();L();var Fe=[],Re=[];async function ba(a){let[e,n,i]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]);Fe=(e.data?.data||[]).map(o=>({value:o.id,label:o.full_name}));let t=(n.data?.data||[]).map(o=>({value:o.full_name,label:o.full_name}));Re=[...(i.data?.data||[]).filter(o=>o.role==="FC Spesialis").map(o=>({value:o.name,label:o.name}))];let u=o=>o&&!t.find(l=>l.value===o)?[...t,{value:o,label:o}]:t,m=o=>{if(!o||o==="-"||String(o).trim()==="")return"";let l=String(o).split("-");return l.length===3&&l[0].length===4?`${l[2]}-${l[1]}-${l[0]}`:o};$({container:a,title:"Jadwal Kegiatan",icon:"\u{1F5D3}\uFE0F",apiPath:"/api/schedule",bulkDelete:!0,itemLabel:"Jadwal",columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Kegiatan",render:o=>pa(o)},{key:"period",label:"Periode",render:o=>Y(o)},{key:"pic",label:"PIC"},{key:"opening_date",label:"Tgl Opening",nowrap:!0,render:o=>m(o)},{key:"target_date",label:"Tgl Target",nowrap:!0,render:o=>m(o)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:o=>m(o)},{key:"status",label:"Status",render:o=>O(o)}],filterFields:[{type:"select",name:"branch_id",label:"Cabang",options:Fe},{type:"select",name:"activity_type",label:"Kegiatan",options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","In Progress","Done"]},{type:"select",name:"pic",label:"PIC",options:Re}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Fe,createApi:{path:"/api/branches",field:"full_name"},value:o?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["Inspeksi Hygiene & Aset Bangunan","General Cleaning","Deep Cleaning","Fogging"],value:o?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:o?.period},{name:"pic",label:"PIC",type:"combobox",options:Re,createApi:{path:"/api/pic",field:"name"},value:o?.pic}]},{type:"row",fields:[{name:"opening_date",label:"Tanggal Opening",type:"date",value:o?.opening_date},{name:"target_date",label:"Tanggal Target",type:"date",value:o?.target_date}]},{type:"row",fields:[{name:"completion_date",label:"Tanggal Selesai",type:"date",value:o?.completion_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","In Progress","Done"],value:o?.status||""}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:o?.notes}],exportOptions:{moduleName:"schedule",onExport:async()=>{let o=await f("/api/schedule?limit=10000");if(o.ok){let l=o.data.data.map(s=>({Cabang:s.branch_name||"",Kegiatan:s.activity_type||"",Periode:s.period||"",PIC:s.pic||"","Tgl Opening":s.opening_date||"","Tgl Target":s.target_date||"","Tgl Selesai":s.completion_date||"",Status:s.status||"",Catatan:s.notes||""}));T(l,"Data_Jadwal_Kegiatan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Kegiatan:"General Cleaning",Periode:"Q1",PIC:"Fajar","Tgl Opening":"2024-02-01","Tgl Target":"2024-02-15","Tgl Selesai":"2024-02-14",Status:"Done",Catatan:""}],"Template_Import_Jadwal")},onImport:async o=>{let s=(await f("/api/branches?all=1")).data?.data||[],r=d=>{if(!d)return null;let p=d.toLowerCase(),y=s.find(v=>v.full_name.toLowerCase()===p||v.code.toLowerCase()===p||v.name.toLowerCase()===p);return y?y.id:null},h=d=>{if(d==null||d==="")return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let p=String(d).trim();if(p===""||p==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);if(/^\d{4,5}$/.test(p)){let v=Number(p);if(v>2e4&&v<99999){let S=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}let y=p.split(/[\/\-\.]/);if(y.length===3){let[v,S,k]=y.map(_=>_.trim());if(v.length===4&&S.length<=2&&k.length<=2)return`${v}-${S.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&S.length<=2&&v.length<=2)return`${k}-${S.padStart(2,"0")}-${v.padStart(2,"0")}`}return p},b=o.map(d=>({branch_id:r(String(d.Cabang||"").trim()),activity_type:String(d.Kegiatan||"").trim(),period:String(d.Periode||"").trim(),pic:String(d.PIC||d.Pic||"").trim(),opening_date:h(d["Tgl Opening"]||d["Tanggal Opening"]),target_date:h(d["Tgl Target"]||d["Tanggal Target"]),completion_date:h(d["Tgl Selesai"]||d["Tanggal Selesai"]),status:String(d.Status||"").trim(),notes:String(d.Catatan||d.Keterangan||"").trim()})).filter(d=>d.activity_type&&d.period),g=await f("/api/schedule/import",{method:"POST",body:JSON.stringify(b)});if(!g.ok)throw new Error(g.data?.error||"Import gagal")}}})}P();L();var Ke=[],Ce=[];async function ha(a){let[e,n,i]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]);Ke=(e.data?.data||[]).map(r=>({value:r.id,label:r.full_name})),Ce=(n.data?.data||[]).map(r=>({value:r.full_name,label:r.full_name}));let t=(i.data?.data||[]).filter(r=>r.role==="FC Spesialis").map(r=>({value:r.name,label:r.name})),c=(i.data?.data||[]).filter(r=>r.role==="Pelapor").map(r=>({value:r.name,label:r.name})),u=r=>r&&!Ce.find(h=>h.value===r)?[...Ce,{value:r,label:r}]:Ce,m=r=>r&&!t.find(h=>h.value===r)?[...t,{value:r,label:r}]:t,o=r=>r&&!c.find(h=>h.value===r)?[...c,{value:r,label:r}]:c,l=new Date().getFullYear(),s=Array.from({length:5},(r,h)=>String(l-h));$({container:a,title:"Permasalahan",icon:"\u26A0\uFE0F",apiPath:"/api/issues",bulkDelete:!0,itemLabel:"Permasalahan",columns:[{key:"report_date",label:"Tanggal",nowrap:!0,render:r=>window.formatDate(r)},{key:"branch_name",label:"Cabang"},{key:"category",label:"Kategori",render:r=>`<span class="badge badge-secondary">${r}</span>`},{key:"source",label:"Sumber"},{key:"complaint",label:"Keluhan",render:r=>`<span title="${r}">${r?.length>50?r.slice(0,50)+"\u2026":r}</span>`},{key:"employee_name",label:"Nama FC"},{key:"fc_specialist",label:"FC Spesialis"},{key:"solution",label:"Solusi",render:r=>`<span title="${r||""}">${r?.length>40?r.slice(0,40)+"\u2026":r||"-"}</span>`},{key:"status",label:"Status",render:r=>O(r)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:r=>window.formatDate(r)},{key:"day_count",label:"Hari",render:r=>r??"-"}],filterFields:[{type:"search",placeholder:"Cari keluhan / nama FC..."},{type:"select",name:"branch_id",label:"Cabang",options:Ke},{type:"select",name:"category",label:"Kategori",options:["SDM","Cleaning","Aset","K3","Lainnya"]},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]},{type:"select",name:"year",label:"Tahun",options:s}],formFields:r=>[{type:"row",fields:[{name:"report_date",label:"Tanggal Info",type:"date",required:!0,value:r?.report_date},{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:Ke,createApi:{path:"/api/branches",field:"full_name"},value:r?.branch_id}]},{type:"row",fields:[{name:"category",label:"Kategori",type:"select",required:!0,options:["SDM","Cleaning","Aset","K3","Lainnya"],value:r?.category},{name:"source",label:"Sumber Laporan",type:"select",options:[...o(r?.source),{value:"Lainnya",label:"Lainnya"}],value:r?.source}]},{name:"complaint",label:"Keluhan",type:"textarea",required:!0,rows:3,value:r?.complaint},{type:"row",fields:[{name:"employee_name",label:"Nama FC / Security",type:"combobox",options:u(r?.employee_name),createApi:{path:"/api/employees",field:"full_name",extra:{status:"Aktif"}},value:r?.employee_name},{name:"fc_specialist",label:"FC Spesialis",type:"select",options:m(r?.fc_specialist),value:r?.fc_specialist}]},{name:"solution",label:"Solusi / Tindakan",type:"textarea",rows:3,value:r?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","In Progress","Done"],value:r?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:r?.completion_date}]}],exportOptions:{moduleName:"issues",onExport:async()=>{let r=await f("/api/issues?limit=10000");if(r.ok){let h=r.data.data.map(b=>({Tanggal:b.report_date||"",Cabang:b.branch_name||"",Kategori:b.category||"",Sumber:b.source||"",Keluhan:b.complaint||"","Nama FC":b.employee_name||"","FC Spesialis":b.fc_specialist||"",Solusi:b.solution||"","Tgl Selesai":b.completion_date||"",Status:b.status||""}));T(h,"Data_Permasalahan")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Tanggal:"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning",Sumber:"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi Santoso","FC Spesialis":"Fajar",Solusi:"Teguran lisan","Tgl Selesai":"2024-03-02",Status:"Done"}],"Template_Import_Permasalahan")},onImport:async r=>{let b=(await f("/api/branches?all=1")).data?.data||[],g=y=>{if(!y)return null;let v=y.toLowerCase(),S=b.find(k=>k.full_name.toLowerCase()===v||k.code.toLowerCase()===v||k.name.toLowerCase()===v);return S?S.id:null},d=r.map(y=>({branch_id:g(String(y.Cabang||"").trim()),report_date:String(y.Tanggal||"").trim(),category:String(y.Kategori||"").trim(),source:String(y.Sumber||"").trim(),complaint:String(y.Keluhan||"").trim(),employee_name:String(y["Nama FC"]||"").trim(),fc_specialist:String(y["FC Spesialis"]||"").trim(),solution:String(y.Solusi||"").trim(),completion_date:String(y["Tgl Selesai"]||"").trim(),status:String(y.Status||"").trim()})).filter(y=>y.report_date&&y.complaint&&y.category),p=await f("/api/issues/import",{method:"POST",body:JSON.stringify(d)});if(!p.ok)throw new Error(p.data?.error||"Import gagal")}}})}P();async function ya(a){let[e,n,i]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]),t=(e.data?.data||[]).map(l=>({value:l.id,label:l.full_name})),c=(n.data?.data||[]).map(l=>({value:l.full_name,label:l.full_name})),u=(i.data?.data||[]).filter(l=>l.role==="FC Spesialis").map(l=>({value:l.name,label:l.name})),m=l=>l&&!c.find(s=>s.value===l)?[...c,{value:l,label:l}]:c,o=l=>l&&!u.find(s=>s.value===l)?[...u,{value:l,label:l}]:u;$({container:a,title:"One on One",icon:"\u{1F91D}",apiPath:"/api/one-on-one",bulkDelete:!0,itemLabel:"One on One",columns:[{key:"meeting_date",label:"Tanggal",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"employee_name",label:"Nama Karyawan"},{key:"pic",label:"PIC"},{key:"problem",label:"Masalah",render:l=>`<span title="${l||""}">${l?.length>50?l.slice(0,50)+"\u2026":l||"-"}</span>`},{key:"solution",label:"Solusi",render:l=>`<span title="${l||""}">${l?.length>40?l.slice(0,40)+"\u2026":l||"-"}</span>`},{key:"status",label:"Status",render:l=>O(l)},{key:"completion_date",label:"Tgl Selesai",nowrap:!0,render:l=>window.formatDate(l)},{key:"day_count",label:"Hari"},{key:"document_link",label:"Dokumen",render:l=>l?`<a href="${l}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama / masalah..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"status",label:"Status",options:["Open","Done"]}],exportOptions:{moduleName:"one_on_one",onExport:async l=>{let s=new URLSearchParams(l||{}).toString(),r=await f(`/api/one-on-one?limit=10000&${s}`);if(r.ok){let h=r.data.data.map(g=>({Tanggal:g.meeting_date||"",Cabang:g.branch_name||"","Nama Karyawan":g.employee_name||"",PIC:g.pic||"",Masalah:g.problem||"",Solusi:g.solution||"",Status:g.status||"","Tgl Selesai":g.completion_date||"",Dokumen:g.document_link||""})),{downloadExcel:b}=await Promise.resolve().then(()=>(L(),J));b(h,`Data_One_on_One_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let l=[{Tanggal:"2026-01-08",Cabang:"001. Pondok Bambu","Nama Karyawan":"Widya Astuti",PIC:"Rina",Masalah:"Terlambat terus",Solusi:"Teguran",Status:"Open","Tgl Selesai":"",Dokumen:"https://link.doc"}],{downloadExcel:s}=await Promise.resolve().then(()=>(L(),J));s(l,"Template_Import_OneOnOne")},onImport:async l=>{let s=g=>{if(!g)return null;let d=g.toLowerCase(),p=e.data?.data.find(y=>y.full_name.toLowerCase()===d||y.code.toLowerCase()===d||y.name.toLowerCase()===d);return p?p.id:null},r=g=>{if(!g)return"";if(g instanceof Date&&!isNaN(g.getTime()))return g.toISOString().slice(0,10);let d=String(g).trim();if(/^\d{4,5}$/.test(d)){let y=Number(d);if(y>2e4&&y<99999){let v=new Date(Date.UTC(1899,11,30)+y*864e5);return isNaN(v.getTime())?"":v.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(d))return d.slice(0,10);let p=d.split(/[\/\-\.]/);if(p.length===3){let[y,v,S]=p.map(k=>k.trim());if(y.length===4&&v.length<=2&&S.length<=2)return`${y}-${v.padStart(2,"0")}-${S.padStart(2,"0")}`;if(S.length===4&&v.length<=2&&y.length<=2)return`${S}-${v.padStart(2,"0")}-${y.padStart(2,"0")}`}return d},h=l.map(g=>({meeting_date:r(g.Tanggal),employee_name:String(g["Nama Karyawan"]||"").trim(),branch_id:s(String(g.Cabang||"").trim()),pic:String(g.PIC||"").trim(),problem:String(g.Masalah||"").trim(),solution:String(g.Solusi||"").trim(),status:String(g.Status||"").trim(),completion_date:r(g["Tgl Selesai"]),document_link:String(g.Dokumen||"").trim()})).filter(g=>g.meeting_date&&g.employee_name&&g.branch_id),b=await f("/api/one-on-one/import",{method:"POST",body:JSON.stringify(h)});if(!b.ok)throw new Error(b.data?.error||"Import gagal")}},formFields:l=>[{type:"row",fields:[{name:"meeting_date",label:"Tanggal",type:"date",required:!0,value:l?.meeting_date},{name:"branch_id",label:"Cabang",type:"combobox",options:t,createApi:{path:"/api/branches",field:"full_name"},value:l?.branch_id}]},{type:"row",fields:[{name:"employee_name",label:"Nama Karyawan",type:"combobox",required:!0,options:m(l?.employee_name),createApi:{path:"/api/employees",field:"full_name",extra:{status:"Aktif"}},value:l?.employee_name},{name:"pic",label:"PIC",type:"combobox",options:o(l?.pic),createApi:{path:"/api/pic",field:"name"},value:l?.pic}]},{name:"problem",label:"Masalah",type:"textarea",required:!0,rows:3,value:l?.problem},{name:"solution",label:"Solusi",type:"textarea",rows:3,value:l?.solution},{type:"row",fields:[{name:"status",label:"Status",type:"select",required:!0,options:["Open","Done"],value:l?.status||""},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:l?.completion_date}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:l?.document_link}]})}P();async function fa(a){let[e,n,i]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]),t=(e.data?.data||[]).map(s=>({value:s.id,label:s.full_name})),c=(n.data?.data||[]).map(s=>({value:s.full_name,label:s.full_name})),u=(i.data?.data||[]).filter(s=>s.role==="FC Spesialis").map(s=>({value:s.name,label:s.name})),m=s=>s&&!c.find(r=>r.value===s)?[...c,{value:s,label:s}]:c,o=s=>s&&!u.find(r=>r.value===s)?[...u,{value:s,label:s}]:u,l=Array.from({length:5},(s,r)=>String(new Date().getFullYear()-r));$({container:a,title:"Training",icon:"\u{1F393}",apiPath:"/api/training",bulkDelete:!0,itemLabel:"Training",columns:[{key:"training_date",label:"Tanggal",nowrap:!0,render:s=>window.formatDate(s)},{key:"batch",label:"Batch"},{key:"subject",label:"Materi"},{key:"branch_name",label:"Cabang"},{key:"trainer",label:"Trainer"},{key:"participants",label:"Peserta",render:s=>{try{let r=JSON.parse(s);return Array.isArray(r)?r.join(", "):s||"-"}catch{return s||"-"}}},{key:"score",label:"Nilai",render:s=>s!=null?`<strong>${s}</strong>`:"-"},{key:"document_link",label:"Dokumen",render:s=>s?`<a href="${s}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari materi / trainer..."},{type:"select",name:"year",label:"Tahun",options:l}],exportOptions:{moduleName:"training",onExport:async s=>{let r=new URLSearchParams(s||{}).toString(),h=await f(`/api/training?limit=10000&${r}`);if(h.ok){let b=h.data.data.map(d=>{let p=d.participants||"";try{let y=JSON.parse(p);p=Array.isArray(y)?y.join(", "):p}catch{}return{Tanggal:d.training_date||"",Batch:d.batch||"",Materi:d.subject||"",Cabang:d.branch_name||"",Trainer:d.trainer||"",Peserta:p,Nilai:d.score!==null&&d.score!==void 0?d.score:"",Dokumen:d.document_link||""}}),{downloadExcel:g}=await Promise.resolve().then(()=>(L(),J));g(b,`Data_Training_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let s=[{Tanggal:"2026-01-08",Batch:"Batch 1",Materi:"Standar Kebersihan",Cabang:"001. Pondok Bambu",Trainer:"Budi",Peserta:"Rina, Agus",Nilai:"85",Dokumen:"https://link.doc"}],{downloadExcel:r}=await Promise.resolve().then(()=>(L(),J));r(s,"Template_Import_Training")},onImport:async s=>{let r=d=>{if(!d)return null;let p=d.toLowerCase(),y=e.data?.data.find(v=>v.full_name.toLowerCase()===p||v.code.toLowerCase()===p||v.name.toLowerCase()===p);return y?y.id:null},h=d=>{if(!d)return"";if(d instanceof Date&&!isNaN(d.getTime()))return d.toISOString().slice(0,10);let p=String(d).trim();if(/^\d{4,5}$/.test(p)){let v=Number(p);if(v>2e4&&v<99999){let S=new Date(Date.UTC(1899,11,30)+v*864e5);return isNaN(S.getTime())?"":S.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(p))return p.slice(0,10);let y=p.split(/[\/\-\.]/);if(y.length===3){let[v,S,k]=y.map(_=>_.trim());if(v.length===4&&S.length<=2&&k.length<=2)return`${v}-${S.padStart(2,"0")}-${k.padStart(2,"0")}`;if(k.length===4&&S.length<=2&&v.length<=2)return`${k}-${S.padStart(2,"0")}-${v.padStart(2,"0")}`}return p},b=s.map(d=>({training_date:h(d.Tanggal),batch:String(d.Batch||"").trim(),subject:String(d.Materi||"").trim(),branch_id:r(String(d.Cabang||"").trim()),trainer:String(d.Trainer||"").trim(),participants:String(d.Peserta||"").trim(),score:d.Nilai?Number(d.Nilai):null,document_link:String(d.Dokumen||"").trim()})).filter(d=>d.training_date&&d.subject&&d.branch_id),g=await f("/api/training/import",{method:"POST",body:JSON.stringify(b)});if(!g.ok)throw new Error(g.data?.error||"Import gagal")}},formFields:s=>[{type:"row",fields:[{name:"training_date",label:"Tanggal Training",type:"date",required:!0,value:s?.training_date},{name:"batch",label:"Batch",placeholder:"Batch 1, Batch 2, ...",value:s?.batch}]},{name:"subject",label:"Materi / Topik Training",required:!0,placeholder:"Judul materi training",value:s?.subject},{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",options:t,createApi:{path:"/api/branches",field:"full_name"},value:s?.branch_id},{name:"trainer",label:"Trainer",type:"combobox",options:o(s?.trainer),createApi:{path:"/api/pic",field:"name"},value:s?.trainer}]},{name:"participants",label:"Peserta (pisahkan dengan koma)",type:"textarea",rows:3,placeholder:"Nama Peserta 1, Nama Peserta 2, ...",value:(()=>{try{let r=JSON.parse(s?.participants);return Array.isArray(r)?r.join(", "):s?.participants||""}catch{return s?.participants||""}})()},{type:"row",fields:[{name:"score",label:"Nilai / Score",type:"number",step:"0.1",min:"0",max:"100",value:s?.score},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:s?.document_link}]},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:s?.notes}],onBeforeSubmit:async s=>(s.participants&&(s.participants=JSON.stringify(s.participants.split(",").map(r=>r.trim()).filter(Boolean))),s)})}P();L();async function va(a){let[e,n]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000")]),i=(e.data?.data||[]).map(o=>({value:o.id,label:o.full_name})),t=(n.data?.data||[]).map(o=>({value:o.full_name,label:o.full_name})),c=o=>o&&!t.find(l=>l.value===o)?[...t,{value:o,label:o}]:t,u=["Krishna Aryaan Permana","Agung Septiadi","Indra Saputro","Wariskin","Iqbal"],m=o=>{let l=u.map(s=>({value:s,label:s}));return o&&!l.find(s=>s.value===o)?[...l,{value:o,label:o}]:l};$({container:a,title:"Jadwal Reliefer",icon:"\u{1F504}",apiPath:"/api/relievers",bulkDelete:!0,itemLabel:"Reliefer",columns:[{key:"backup_date",label:"Tanggal Backup",nowrap:!0,render:o=>window.formatDate(o)},{key:"branch_name",label:"Cabang"},{key:"original_fc_name",label:"FC Digantikan"},{key:"period",label:"Periode",render:o=>Y(o)},{key:"reliever_name",label:"Reliefer"},{key:"reason",label:"Keterangan"},{key:"shift",label:"Shift",render:o=>o?`<span class="badge badge-info">${o}</span>`:"-"},{key:"status",label:"Status",render:o=>O(o)}],filterFields:[{type:"search",placeholder:"Cari reliefer / FC..."},{type:"select",name:"branch_id",label:"Cabang",options:i},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done","Tidak Datang"]}],formFields:o=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:i,createApi:{path:"/api/branches",field:"full_name"},value:o?.branch_id},{name:"period",label:"Periode",type:"select",options:["Q1","Q2","Q3","Q4"],value:o?.period}]},{type:"row",fields:[{name:"original_fc_name",label:"FC yang Digantikan",type:"select",options:[{value:"",label:"BELUM ADA FC"},...c(o?.original_fc_name)],value:o?.original_fc_name},{name:"reliever_name",label:"Nama Reliefer",type:"select",required:!0,options:m(o?.reliever_name),value:o?.reliever_name}]},{type:"row",fields:[{name:"backup_date",label:"Tanggal Backup",type:"date",required:!0,value:o?.backup_date},{name:"completion_date",label:"Tanggal Selesai",type:"date",value:o?.completion_date}]},{type:"row",fields:[{name:"reason",label:"Keterangan",type:"select",options:["Cuti","Mengisi Kekosongan","Back Up Training","Deep Cleaning","Training Praktek Skill","Sakit","Lainnya"],value:o?.reason},{name:"shift",label:"Shift",type:"select",options:["Pagi","Siang","Full Shift","Middle"],value:o?.shift}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done","Tidak Datang"],value:o?.status||""}],exportOptions:{moduleName:"relievers",onExport:async()=>{let o=await f("/api/relievers?limit=10000");if(o.ok){let l=o.data.data.map(s=>({"Tanggal Backup":s.backup_date||"",Cabang:s.branch_name||"","FC Digantikan":s.original_fc_name||"",Periode:s.period||"",Reliefer:s.reliever_name||"",Keterangan:s.reason||"",Shift:s.shift||"","Tanggal Selesai":s.completion_date||"",Status:s.status||""}));T(l,"Data_Reliefer")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Tanggal Backup":"2024-03-10",Cabang:"001. Pondok Bambu","FC Digantikan":"Budi Santoso",Periode:"Q1",Reliefer:"Andi",Keterangan:"Sakit",Shift:"Pagi","Tanggal Selesai":"2024-03-10",Status:"Done"}],"Template_Import_Reliefer")},onImport:async o=>{let s=(await f("/api/branches?all=1")).data?.data||[],r=g=>{if(!g)return null;let d=g.toLowerCase(),p=s.find(y=>y.full_name.toLowerCase()===d||y.code.toLowerCase()===d||y.name.toLowerCase()===d);return p?p.id:null},h=o.map(g=>({branch_id:r(String(g.Cabang||"").trim()),backup_date:String(g["Tanggal Backup"]||"").trim(),original_fc_name:String(g["FC Digantikan"]||"").trim(),reliever_name:String(g.Reliefer||"").trim(),period:String(g.Periode||"").trim(),reason:String(g.Keterangan||"").trim(),shift:String(g.Shift||"").trim(),completion_date:String(g["Tanggal Selesai"]||"").trim(),status:String(g.Status||"").trim()})).filter(g=>g.reliever_name&&g.backup_date),b=await f("/api/relievers/import",{method:"POST",body:JSON.stringify(h)});if(!b.ok)throw new Error(b.data?.error||"Import gagal")}}})}P();L();async function ka(a){let n=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name})),i=Array.from({length:4},(t,c)=>String(new Date().getFullYear()-c));$({container:a,title:"Laporan Inspeksi Hygiene",icon:"\u{1F50D}",apiPath:"/api/reports/inspection",itemLabel:"Laporan Inspeksi",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"period",label:"Periode",render:t=>Y(t)},{key:"inspection_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"fc_score",label:"Point FC",render:t=>t!=null?`<strong class="${t>=80?"text-success":t>=60?"text-warning":"text-danger"}">${t}</strong>`:"-"},{key:"spv_score",label:"Point SPV",render:t=>t!=null?`<strong>${t}</strong>`:"-"},{key:"status",label:"Status",render:t=>O(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari cabang / PIC..."},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:n,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period}]},{type:"row",fields:[{name:"inspection_date",label:"Tanggal Inspeksi",type:"date",required:!0,value:t?.inspection_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""}]},{type:"row",fields:[{name:"fc_score",label:"Point FC",type:"number",step:"0.1",min:"0",max:"100",value:t?.fc_score},{name:"spv_score",label:"Point SPV",type:"number",step:"0.1",min:"0",max:"100",value:t?.spv_score}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"inspection_reports",onExport:async t=>{let c=new URLSearchParams(t||{}).toString(),u=await f(`/api/reports/inspection?limit=10000&${c}`);if(u.ok){let m=u.data.data.map(o=>({Cabang:o.branch_name||"",Periode:o.period||"",Tanggal:o.inspection_date||"","Point FC":o.fc_score!==null&&o.fc_score!==void 0?o.fc_score:"","Point SPV":o.spv_score!==null&&o.spv_score!==void 0?o.spv_score:"",Status:o.status||"","Link Dokumen":o.document_link||"",Catatan:o.notes||""}));T(m,`Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Periode:"Q1",Tanggal:"2026-01-08","Point FC":85,"Point SPV":90,Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:"Semua area bersih"}],"Template_Import_Inspeksi")},onImport:async t=>{let u=(await f("/api/branches?all=1")).data?.data||[],m=r=>{if(!r)return null;let h=r.toLowerCase(),b=u.find(g=>g.full_name.toLowerCase()===h||g.code.toLowerCase()===h||g.name.toLowerCase()===h);return b?b.id:null},o=r=>{if(r==null||r==="")return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let h=String(r).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let g=Number(h);if(g>2e4&&g<99999){let d=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let b=h.split(/[\/\-\.]/);if(b.length===3){let[g,d,p]=b.map(y=>y.trim());if(g.length===4&&d.length<=2&&p.length<=2)return`${g}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&d.length<=2&&g.length<=2)return`${p}-${d.padStart(2,"0")}-${g.padStart(2,"0")}`}return h},l=t.map(r=>({branch_id:m(String(r.Cabang||"").trim()),period:String(r.Periode||"").trim(),inspection_date:o(r.Tanggal),fc_score:r["Point FC"]!==void 0&&r["Point FC"]!==""?Number(r["Point FC"]):null,spv_score:r["Point SPV"]!==void 0&&r["Point SPV"]!==""?Number(r["Point SPV"]):null,status:String(r.Status||"").trim(),document_link:String(r["Link Dokumen"]||"").trim(),notes:String(r.Catatan||r.Keterangan||"").trim()})).filter(r=>r.branch_id&&r.period&&r.inspection_date),s=await f("/api/reports/inspection/import",{method:"POST",body:JSON.stringify(l)});if(!s.ok)throw new Error(s.data?.error||"Import gagal")}}})}P();L();async function Sa(a){let n=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name})),i=Array.from({length:4},(t,c)=>String(new Date().getFullYear()-c));$({container:a,title:"Laporan General Cleaning & Deep Cleaning",icon:"\u{1F9F9}",apiPath:"/api/reports/cleaning",itemLabel:"Laporan GC/DC",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:t=>`<span class="badge ${t==="Deep Cleaning"?"badge-purple":"badge-success"}">${t}</span>`},{key:"period",label:"Periode",render:t=>Y(t)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"status",label:"Status",render:t=>O(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"activity_type",label:"Jenis",options:["General Cleaning","Deep Cleaning"]},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:n,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id},{name:"activity_type",label:"Jenis Kegiatan",type:"select",required:!0,options:["General Cleaning","Deep Cleaning"],value:t?.activity_type}]},{type:"row",fields:[{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period},{name:"activity_date",label:"Tanggal",type:"date",required:!0,value:t?.activity_date}]},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://drive.google.com/...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"cleaning_reports",onExport:async t=>{let c=new URLSearchParams(t||{}).toString(),u=await f(`/api/reports/cleaning?limit=10000&${c}`);if(u.ok){let m=u.data.data.map(o=>({Cabang:o.branch_name||"",Jenis:o.activity_type||"",Periode:o.period||"",Tanggal:o.activity_date||"",Status:o.status||"","Link Dokumen":o.document_link||"",Catatan:o.notes||""}));T(m,`Laporan_GCDC_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Jenis:"General Cleaning",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:"Pembersihan lantai"}],"Template_Import_GCDC")},onImport:async t=>{let u=(await f("/api/branches?all=1")).data?.data||[],m=r=>{if(!r)return null;let h=r.toLowerCase(),b=u.find(g=>g.full_name.toLowerCase()===h||g.code.toLowerCase()===h||g.name.toLowerCase()===h);return b?b.id:null},o=r=>{if(r==null||r==="")return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let h=String(r).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let g=Number(h);if(g>2e4&&g<99999){let d=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let b=h.split(/[\/\-\.]/);if(b.length===3){let[g,d,p]=b.map(y=>y.trim());if(g.length===4&&d.length<=2&&p.length<=2)return`${g}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&d.length<=2&&g.length<=2)return`${p}-${d.padStart(2,"0")}-${g.padStart(2,"0")}`}return h},l=t.map(r=>({branch_id:m(String(r.Cabang||"").trim()),activity_type:String(r.Jenis||r.Kegiatan||"").trim(),period:String(r.Periode||"").trim(),activity_date:o(r.Tanggal),status:String(r.Status||"").trim(),document_link:String(r["Link Dokumen"]||"").trim(),notes:String(r.Catatan||r.Keterangan||"").trim()})).filter(r=>r.branch_id&&r.activity_type&&r.period&&r.activity_date),s=await f("/api/reports/cleaning/import",{method:"POST",body:JSON.stringify(l)});if(!s.ok)throw new Error(s.data?.error||"Import gagal")}}})}P();L();async function _a(a){let n=((await f("/api/branches?all=1")).data?.data||[]).map(t=>({value:t.id,label:t.full_name})),i=Array.from({length:4},(t,c)=>String(new Date().getFullYear()-c));$({container:a,title:"Rekap Fogging",icon:"\u{1F4A8}",apiPath:"/api/reports/fogging",itemLabel:"Fogging",bulkDelete:!0,columns:[{key:"branch_name",label:"Cabang"},{key:"activity_type",label:"Jenis",render:t=>`<span class="badge badge-warning">${t}</span>`},{key:"period",label:"Periode",render:t=>Y(t)},{key:"activity_date",label:"Tanggal",nowrap:!0,render:t=>window.formatDate(t)},{key:"status",label:"Status",render:t=>O(t)},{key:"document_link",label:"Dokumen",render:t=>t?`<a href="${t}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">\u{1F4C4} Buka</a>`:"-"},{key:"notes",label:"Catatan",render:t=>t||"-"}],filterFields:[{type:"search",placeholder:"Cari nama cabang/lokasi..."},{type:"select",name:"branch_id",label:"Cabang",options:n},{type:"select",name:"period",label:"Periode",options:["Q1","Q2","Q3","Q4"]},{type:"select",name:"status",label:"Status",options:["Pending","Done"]},{type:"select",name:"year",label:"Tahun",options:i}],formFields:t=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:n,createApi:{path:"/api/branches",field:"full_name"},value:t?.branch_id},{name:"period",label:"Periode",type:"select",required:!0,options:["Q1","Q2","Q3","Q4"],value:t?.period}]},{type:"row",fields:[{name:"activity_date",label:"Tanggal",type:"date",value:t?.activity_date},{name:"status",label:"Status",type:"select",required:!0,options:["Pending","Done"],value:t?.status||""}]},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:t?.document_link},{name:"notes",label:"Catatan",type:"textarea",rows:2,value:t?.notes}],exportOptions:{moduleName:"fogging_reports",onExport:async t=>{let c=new URLSearchParams(t||{}).toString(),u=await f(`/api/reports/fogging?limit=10000&${c}`);if(u.ok){let m=u.data.data.map(o=>({Cabang:o.branch_name||"",Jenis:o.activity_type||"Fogging",Periode:o.period||"",Tanggal:o.activity_date||"",Status:o.status||"","Link Dokumen":o.document_link||"",Catatan:o.notes||""}));T(m,`Laporan_Fogging_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Cabang:"001. Pondok Bambu",Jenis:"Fogging",Periode:"Q1",Tanggal:"2026-01-08",Status:"Done","Link Dokumen":"https://drive.google.com/...",Catatan:"Tuntas"}],"Template_Import_Fogging")},onImport:async t=>{let u=(await f("/api/branches?all=1")).data?.data||[],m=r=>{if(!r)return null;let h=r.toLowerCase(),b=u.find(g=>g.full_name.toLowerCase()===h||g.code.toLowerCase()===h||g.name.toLowerCase()===h);return b?b.id:null},o=r=>{if(r==null||r==="")return"";if(r instanceof Date&&!isNaN(r.getTime()))return r.toISOString().slice(0,10);let h=String(r).trim();if(h===""||h==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(h))return h.slice(0,10);if(/^\d{4,5}$/.test(h)){let g=Number(h);if(g>2e4&&g<99999){let d=new Date(Date.UTC(1899,11,30)+g*864e5);return isNaN(d.getTime())?"":d.toISOString().slice(0,10)}}let b=h.split(/[\/\-\.]/);if(b.length===3){let[g,d,p]=b.map(y=>y.trim());if(g.length===4&&d.length<=2&&p.length<=2)return`${g}-${d.padStart(2,"0")}-${p.padStart(2,"0")}`;if(p.length===4&&d.length<=2&&g.length<=2)return`${p}-${d.padStart(2,"0")}-${g.padStart(2,"0")}`}return h},l=t.map(r=>({branch_id:m(String(r.Cabang||"").trim()),activity_type:String(r.Jenis||r.Kegiatan||"Fogging").trim(),period:String(r.Periode||"").trim(),activity_date:o(r.Tanggal),status:String(r.Status||"").trim(),document_link:String(r["Link Dokumen"]||"").trim(),notes:String(r.Catatan||r.Keterangan||"").trim()})).filter(r=>r.branch_id&&r.period&&r.activity_date),s=await f("/api/reports/fogging/import",{method:"POST",body:JSON.stringify(l)});if(!s.ok)throw new Error(s.data?.error||"Import gagal")}}})}P();L();async function wa(a){let[e,n,i]=await Promise.all([f("/api/branches?all=1"),f("/api/employees?limit=10000"),f("/api/pic?limit=10000")]),t=(e.data?.data||[]).map(l=>({value:l.id,label:l.full_name})),c=(n.data?.data||[]).map(l=>({value:l.full_name,label:l.full_name})),u=(i.data?.data||[]).filter(l=>l.role==="FC Spesialis").map(l=>({value:l.name,label:l.name})),m=l=>l&&!c.find(s=>s.value===l)?[...c,{value:l,label:l}]:c,o=l=>l&&!u.find(s=>s.value===l)?[...u,{value:l,label:l}]:u;$({container:a,title:"Rekap Laporan Basecamp",icon:"\u{1F4DD}",apiPath:"/api/reports/basecamp",bulkDelete:!0,itemLabel:"Laporan Basecamp",columns:[{key:"info_date",label:"Tgl Info",nowrap:!0,render:l=>window.formatDate(l)},{key:"branch_name",label:"Cabang"},{key:"problem",label:"Permasalahan",render:l=>`<span title="${l||""}">${l?.length>60?l.slice(0,60)+"\u2026":l||"-"}</span>`},{key:"pic",label:"PIC"},{key:"done_date",label:"Tgl Done",nowrap:!0,render:l=>window.formatDate(l)},{key:"status",label:"Status",render:l=>O(l)},{key:"notes",label:"Keterangan",render:l=>l?.length>40?l.slice(0,40)+"\u2026":l||"-"}],filterFields:[{type:"search",placeholder:"Cari permasalahan / PIC..."},{type:"select",name:"branch_id",label:"Cabang",options:t},{type:"select",name:"status",label:"Status",options:["Open","In Progress","Done"]}],formFields:l=>[{type:"row",fields:[{name:"branch_id",label:"Cabang",type:"combobox",required:!0,options:t,createApi:{path:"/api/branches",field:"full_name"},value:l?.branch_id},{name:"pic",label:"PIC",type:"combobox",options:o(l?.pic),createApi:{path:"/api/pic",field:"name"},value:l?.pic}]},{name:"problem",label:"Permasalahan",type:"textarea",required:!0,rows:3,value:l?.problem},{type:"row",fields:[{name:"info_date",label:"Tanggal Info",type:"date",required:!0,value:l?.info_date},{name:"done_date",label:"Tanggal Done",type:"date",value:l?.done_date}]},{name:"status",label:"Status",type:"select",options:["Open","In Progress","Done"],value:l?.status||"Open"},{name:"notes",label:"Keterangan / Tindak Lanjut",type:"textarea",rows:2,value:l?.notes}],exportOptions:{moduleName:"basecamp_reports",onExport:async l=>{let s=new URLSearchParams(l||{}).toString(),r=await f(`/api/reports/basecamp?limit=10000&${s}`);if(r.ok){let h=r.data.data.map(b=>({"Tgl Info":b.info_date||"",Cabang:b.branch_name||"",Permasalahan:b.problem||"",PIC:b.pic||"","Tgl Done":b.done_date||"",Status:b.status||"",Keterangan:b.notes||""}));T(h,`Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Tgl Info":"2026-01-08",Cabang:"001. Pondok Bambu",Permasalahan:"Request fogging karena banyak nyamuk",PIC:"Fajar","Tgl Done":"2026-01-10",Status:"Done",Keterangan:"Sudah difogging"}],"Template_Import_Basecamp")},onImport:async l=>{let r=(await f("/api/branches?all=1")).data?.data||[],h=p=>{if(!p)return null;let y=p.toLowerCase(),v=r.find(S=>S.full_name.toLowerCase()===y||S.code.toLowerCase()===y||S.name.toLowerCase()===y);return v?v.id:null},b=p=>{if(p==null||p==="")return"";if(p instanceof Date&&!isNaN(p.getTime()))return p.toISOString().slice(0,10);let y=String(p).trim();if(y===""||y==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(y))return y.slice(0,10);if(/^\d{4,5}$/.test(y)){let S=Number(y);if(S>2e4&&S<99999){let k=new Date(Date.UTC(1899,11,30)+S*864e5);return isNaN(k.getTime())?"":k.toISOString().slice(0,10)}}let v=y.split(/[\/\-\.]/);if(v.length===3){let[S,k,_]=v.map(E=>E.trim());if(S.length===4&&k.length<=2&&_.length<=2)return`${S}-${k.padStart(2,"0")}-${_.padStart(2,"0")}`;if(_.length===4&&k.length<=2&&S.length<=2)return`${_}-${k.padStart(2,"0")}-${S.padStart(2,"0")}`}return y},g=l.map(p=>({info_date:b(p["Tgl Info"]||p["Tanggal Info"]),branch_id:h(String(p.Cabang||"").trim()),problem:String(p.Permasalahan||"").trim(),pic:String(p.PIC||"").trim(),done_date:b(p["Tgl Done"]||p["Tanggal Done"]),status:String(p.Status||"").trim(),notes:String(p.Keterangan||p.Catatan||"").trim()})).filter(p=>p.info_date&&p.branch_id&&p.problem),d=await f("/api/reports/basecamp/import",{method:"POST",body:JSON.stringify(g)});if(!d.ok)throw new Error(d.data?.error||"Import gagal")}}})}async function Ca(a){$({container:a,title:"SOP",icon:"\u{1F4D6}",apiPath:"/api/sop",bulkDelete:!0,itemLabel:"SOP",columns:[{key:"name",label:"Nama SOP"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"notes",label:"Catatan"}],filterFields:[{type:"search",placeholder:"Cari nama SOP..."}],exportOptions:{moduleName:"sop",onExport:async e=>{let n=new URLSearchParams(e||{}).toString(),{apiFetch:i}=await Promise.resolve().then(()=>(P(),me)),t=await i(`/api/sop?limit=10000&${n}`);if(t.ok){let c=t.data.data.map(m=>({"Nama SOP":m.name||"",Kategori:m.category||"",Dokumen:m.document_link||"",Catatan:m.notes||m.description||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(L(),J));u(c,`Master_SOP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama SOP":"SOP Cuci Tangan",Kategori:"Ketentuan & Basic",Dokumen:"https://link.com",Catatan:"Catatan singkat"}],{downloadExcel:n}=await Promise.resolve().then(()=>(L(),J));n(e,"Template_Import_SOP")},onImport:async e=>{let n=e.map(c=>({name:String(c["Nama SOP"]||"").trim(),category:String(c.Kategori||"").trim(),document_link:String(c.Dokumen||"").trim(),description:String(c.Catatan||"").trim()})).filter(c=>c.name),{apiFetch:i}=await Promise.resolve().then(()=>(P(),me)),t=await i("/api/sop/import",{method:"POST",body:JSON.stringify(n)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama SOP",required:!0,placeholder:"Nama SOP",value:e?.name},{name:"category",label:"Kategori",placeholder:"Ketentuan & Basic, Kualitas & Grooming, dst.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi / Catatan",type:"textarea",rows:3,value:e?.description}]})}async function xa(a){$({container:a,title:"Master Checklist",icon:"\u2705",apiPath:"/api/checklist",bulkDelete:!0,itemLabel:"Checklist",columns:[{key:"name",label:"Nama Checklist"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka Dokumen</a>`:"-"},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari checklist..."}],exportOptions:{moduleName:"checklist",onExport:async e=>{let n=new URLSearchParams(e||{}).toString(),{apiFetch:i}=await Promise.resolve().then(()=>(P(),me)),t=await i(`/api/checklist?limit=10000&${n}`);if(t.ok){let c=t.data.data.map(m=>({"Nama Checklist":m.name||"",Kategori:m.category||"",Dokumen:m.document_link||"",Deskripsi:m.description||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(L(),J));u(c,`Master_Checklist_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let e=[{"Nama Checklist":"Checklist Kebersihan Mingguan",Kategori:"Master Cleaning Program",Dokumen:"https://link.com",Deskripsi:"Deskripsi singkat"}],{downloadExcel:n}=await Promise.resolve().then(()=>(L(),J));n(e,"Template_Import_Checklist")},onImport:async e=>{let n=e.map(c=>({name:String(c["Nama Checklist"]||"").trim(),category:String(c.Kategori||"").trim(),document_link:String(c.Dokumen||"").trim(),description:String(c.Deskripsi||"").trim()})).filter(c=>c.name),{apiFetch:i}=await Promise.resolve().then(()=>(P(),me)),t=await i("/api/checklist/import",{method:"POST",body:JSON.stringify(n)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}},formFields:e=>[{name:"name",label:"Nama Checklist",required:!0,placeholder:"Nama checklist",value:e?.name},{name:"category",label:"Kategori",placeholder:"Master Cleaning Program, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://docs.google.com/...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:3,value:e?.description}]})}P();L();async function qe(a,e="forms"){if(e==="supply")return ut(a);mt(a)}function mt(a){$({container:a,title:"Master Form",icon:"\u{1F4C4}",apiPath:"/api/forms",bulkDelete:!0,itemLabel:"Form",columns:[{key:"name",label:"Nama Form"},{key:"category",label:"Kategori"},{key:"document_link",label:"Dokumen",render:e=>e?`<a href="${e}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">\u{1F4C4} Buka</a>`:"-"},{key:"is_public",label:"Publik",render:e=>e?'<span class="badge badge-success">Ya</span>':'<span class="badge badge-neutral">Tidak</span>'},{key:"description",label:"Deskripsi"}],filterFields:[{type:"search",placeholder:"Cari form..."}],formFields:e=>[{name:"name",label:"Nama Form",required:!0,placeholder:"Nama form",value:e?.name},{name:"category",label:"Kategori",placeholder:"Permintaan Barang, Penilaian, dll.",value:e?.category},{name:"document_link",label:"Link Dokumen",type:"url",placeholder:"https://...",value:e?.document_link},{name:"description",label:"Deskripsi",type:"textarea",rows:2,value:e?.description},{name:"is_public",label:"Akses Publik",type:"checkbox",checkLabel:"Form dapat diakses tanpa login",value:e?.is_public}]})}async function ut(a){let n=((await f("/api/branches?all=1")).data?.data||[]).map(i=>({value:i.id,label:i.full_name}));$({container:a,title:"Permintaan Barang & Chemical",icon:"\u{1F4E6}",apiPath:"/api/reports/supply",bulkDelete:!0,itemLabel:"Permintaan",canCreate:!0,columns:[{key:"submitted_at",label:"Waktu",nowrap:!0,render:i=>i?new Date(i).toLocaleString("id-ID"):"-"},{key:"submitter_name",label:"Pengirim"},{key:"branch_name",label:"Cabang",render:(i,t)=>t.branch_name_ref||t.branch_name||"-"},{key:"tools_items",label:"Alat/Barang",render:i=>{try{let t=JSON.parse(i);return Array.isArray(t)?t.join(", "):i}catch{return i||"-"}}},{key:"chemical_items",label:"Chemical",render:i=>{try{let t=JSON.parse(i);return Array.isArray(t)?t.join(", "):i}catch{return i||"-"}}},{key:"additional_notes",label:"Catatan",render:i=>i?.length>40?i.slice(0,40)+"\u2026":i||"-"},{key:"status",label:"Status",render:i=>O(i)},{key:"processed_by",label:"Diproses Oleh"}],filterFields:[{type:"select",name:"status",label:"Status",options:["Pending","Diproses","Selesai"]}],formFields:i=>{let t=i?.tools_items;try{t=Array.isArray(JSON.parse(t))?JSON.parse(t).join(", "):t}catch{}let c=i?.chemical_items;try{c=Array.isArray(JSON.parse(c))?JSON.parse(c).join(", "):c}catch{}return[{type:"row",fields:[{name:"submitter_name",label:"Nama Pengirim",required:!0,value:i?.submitter_name},{name:"branch_id",label:"Cabang",type:"combobox",options:n,createApi:{path:"/api/branches",field:"full_name"},value:i?.branch_id}]},{type:"row",fields:[{name:"tools_items",label:"Alat / Barang",placeholder:"Pisahkan dengan koma (Sapu, Mop)",value:t},{name:"tools_quantity",label:"Jumlah Alat",type:"number",value:i?.tools_quantity}]},{type:"row",fields:[{name:"chemical_items",label:"Chemical",placeholder:"Pisahkan dengan koma",value:c},{name:"chemical_quantity",label:"Jumlah Chemical",type:"number",value:i?.chemical_quantity}]},{name:"additional_notes",label:"Catatan",type:"textarea",rows:2,value:i?.additional_notes},{name:"status",label:"Status",type:"select",options:["Pending","Diproses","Selesai"],value:i?.status||""},{name:"processed_by",label:"Diproses Oleh",value:i?.processed_by}]},exportOptions:{moduleName:"supply_requests",onExport:async i=>{let t=new URLSearchParams(i||{}).toString(),c=await f(`/api/reports/supply?limit=10000&${t}`);if(c.ok){let u=c.data.data.map(m=>{let o=m.tools_items;try{o=Array.isArray(JSON.parse(o))?JSON.parse(o).join(", "):o}catch{}let l=m.chemical_items;try{l=Array.isArray(JSON.parse(l))?JSON.parse(l).join(", "):l}catch{}return{Waktu:m.submitted_at||"",Pengirim:m.submitter_name||"",Cabang:m.branch_name_ref||m.branch_name||"","Alat/Barang":o||"",Chemical:l||"",Catatan:m.additional_notes||"",Status:m.status||"","Diproses Oleh":m.processed_by||""}});T(u,`Permintaan_Barang_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{Waktu:"2026-01-08",Pengirim:"Fajar",Cabang:"001. Pondok Bambu","Alat/Barang":"Sapu, Mop",Chemical:"Karbol",Catatan:"Mendesak",Status:"Pending","Diproses Oleh":""}],"Template_Import_Permintaan")},onImport:async i=>{let c=(await f("/api/branches?all=1")).data?.data||[],u=s=>{if(!s)return null;let r=s.toLowerCase(),h=c.find(b=>b.full_name.toLowerCase()===r||b.code.toLowerCase()===r||b.name.toLowerCase()===r);return h?h.id:null},m=s=>{if(s==null||s==="")return"";if(s instanceof Date&&!isNaN(s.getTime()))return s.toISOString().slice(0,10);let r=String(s).trim();if(r===""||r==="0")return"";if(/^\d{4}-\d{2}-\d{2}/.test(r))return r.slice(0,10);if(/^\d{4,5}$/.test(r)){let b=Number(r);if(b>2e4&&b<99999){let g=new Date(Date.UTC(1899,11,30)+b*864e5);return isNaN(g.getTime())?"":g.toISOString().slice(0,10)}}let h=r.split(/[\/\-\.]/);if(h.length===3){let[b,g,d]=h.map(p=>p.trim());if(b.length===4&&g.length<=2&&d.length<=2)return`${b}-${g.padStart(2,"0")}-${d.padStart(2,"0")}`;if(d.length===4&&g.length<=2&&b.length<=2)return`${d}-${g.padStart(2,"0")}-${b.padStart(2,"0")}`}return r},o=i.map(s=>({submitted_at:m(s.Waktu||s.Tanggal),submitter_name:String(s.Pengirim||"").trim(),branch_id:u(String(s.Cabang||"").trim()),tools_items:String(s["Alat/Barang"]||s.Alat||"").trim(),chemical_items:String(s.Chemical||"").trim(),additional_notes:String(s.Catatan||s.Keterangan||"").trim(),status:String(s.Status||"").trim(),processed_by:String(s["Diproses Oleh"]||s.PIC||"").trim()})).filter(s=>s.submitted_at&&s.submitter_name&&s.branch_id),l=await f("/api/reports/supply/import",{method:"POST",body:JSON.stringify(o)});if(!l.ok)throw new Error(l.data?.error||"Import gagal")}},extraActions:[{label:"Update Status",icon:"\u{1F504}",class:"btn-secondary",handler:(i,t)=>{let c=ee({title:"Update Status Permintaan",content:`
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
            `,onConfirm:async(u,m)=>{let o=u.querySelector("#supply-status").value,l=u.querySelector("#supply-processed-by").value;(await f(`/api/reports/supply/${i.id}`,{method:"PUT",body:JSON.stringify({status:o,processed_by:l})})).ok?(K("Status diperbarui."),m(),t()):j("Gagal update status.")}})}}]})}P();L();async function Ta(a){let e=ie();if(!e||!["superadmin","admin"].includes(e.role)){a.innerHTML='<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';return}$({container:a,title:"Manajemen User",icon:"\u{1F510}",apiPath:"/api/users",bulkDelete:!0,itemLabel:"User",columns:[{key:"full_name",label:"Nama Lengkap"},{key:"username",label:"Username"},{key:"email",label:"Email"},{key:"role",label:"Role",render:n=>`<span class="badge ${{superadmin:"badge-danger",admin:"badge-purple",manager:"badge-info",spv:"badge-secondary",viewer:"badge-neutral"}[n]||"badge-neutral"}">${n}</span>`},{key:"is_active",label:"Status",render:n=>n?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'},{key:"created_at",label:"Dibuat",nowrap:!0,render:n=>n?new Date(n).toLocaleDateString("id-ID"):"-"}],filterFields:[{type:"search",placeholder:"Cari nama / username..."}],formFields:n=>{let i=!!n;return[{type:"row",fields:[{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"Nama lengkap",value:n?.full_name},{name:"username",label:"Username",required:!i,placeholder:"username",value:n?.username}]},{type:"row",fields:[{name:"email",label:"Email",type:"email",required:!i,placeholder:"email@contoh.com",value:n?.email},{name:"role",label:"Role",type:"select",required:!0,options:[{value:"superadmin",label:"Super Admin"},{value:"admin",label:"Admin"},{value:"manager",label:"Manager"},{value:"spv",label:"Supervisor"},{value:"viewer",label:"Viewer"}],value:n?.role||"viewer"}]},{type:"row",fields:[{name:"password",label:i?"Password Baru (kosongkan jika tidak diubah)":"Password",type:"password",required:!i,placeholder:"Min. 6 karakter"},{name:"is_active",label:"Status Aktif",type:"checkbox",checkLabel:"User aktif",value:i?n?.is_active:1}]}]},exportOptions:{moduleName:"users",onExport:async()=>{let n=await f("/api/users?limit=10000");if(n.ok){let i=n.data.data.map(t=>({"Nama Lengkap":t.full_name||"",Username:t.username||"",Email:t.email||"",Role:t.role||"",Status:t.is_active?"Aktif":"Nonaktif"}));T(i,"Data_Users")}else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Nama Lengkap":"Admin Cabang",Username:"admin01",Email:"admin@contoh.com",Role:"admin",Password:"password123"}],"Template_Import_Users")},onImport:async n=>{let i=n.map(c=>({full_name:String(c["Nama Lengkap"]||"").trim(),username:String(c.Username||"").trim(),email:String(c.Email||"").trim(),role:String(c.Role||"").trim()||"viewer",password:String(c.Password||"").trim()})).filter(c=>c.username&&c.password&&c.email&&c.full_name),t=await f("/api/users/import",{method:"POST",body:JSON.stringify(i)});if(!t.ok)throw new Error(t.data?.error||"Import gagal")}}})}P();L();async function $a(a){$({container:a,title:"Manajemen Cabang",icon:"\u{1F3E2}",apiPath:"/api/branches",itemLabel:"Cabang",bulkDelete:!0,columns:[{key:"code",label:"Kode",width:"60px"},{key:"full_name",label:"Nama Cabang"},{key:"city",label:"Kota"},{key:"is_active",label:"Status",render:e=>e?'<span class="badge badge-success">Aktif</span>':'<span class="badge badge-neutral">Nonaktif</span>'}],filterFields:[{type:"search",placeholder:"Cari nama / kode cabang..."}],formFields:e=>[{type:"row",fields:[{name:"code",label:"Kode Cabang",required:!0,placeholder:"001, A01, ...",value:e?.code},{name:"name",label:"Nama Pendek",required:!0,placeholder:"Pondok Bambu",value:e?.name}]},{name:"full_name",label:"Nama Lengkap",required:!0,placeholder:"001. Pondok Bambu",value:e?.full_name},{type:"row",fields:[{name:"city",label:"Kota",placeholder:"Jakarta",value:e?.city},{name:"is_active",label:"Status",type:"checkbox",checkLabel:"Cabang aktif",value:e?.is_active!==void 0?e.is_active:1}]}],exportOptions:{moduleName:"branches",onExport:async()=>{let e=await f("/api/branches?limit=10000");if(e.ok)T(e.data.data,"Data_Cabang");else throw new Error("Gagal mengambil data")},onTemplate:()=>{T([{"Kode Cabang":"001","Nama Pendek":"Pondok Bambu","Nama Lengkap":"001. Pondok Bambu",Kota:"Jakarta Timur"},{"Kode Cabang":"002","Nama Pendek":"Bintaro","Nama Lengkap":"002. Bintaro",Kota:"Tangerang Selatan"}],"Template_Import_Cabang")},onImport:async e=>{let n=e.map(t=>({code:String(t["Kode Cabang"]||"").trim(),name:String(t["Nama Pendek"]||"").trim(),full_name:String(t["Nama Lengkap"]||"").trim(),city:String(t.Kota||"").trim()})).filter(t=>t.code&&t.name),i=await f("/api/branches/import",{method:"POST",body:JSON.stringify(n)});if(!i.ok)throw new Error(i.data?.error||"Import gagal")}}})}P();async function Ea(a){let e=new Date,n=[];a.innerHTML=`
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
  `,document.getElementById("cal-prev").addEventListener("click",()=>{e.setMonth(e.getMonth()-1),t()}),document.getElementById("cal-next").addEventListener("click",()=>{e.setMonth(e.getMonth()+1),t()}),document.getElementById("cal-event-close").addEventListener("click",()=>{document.getElementById("cal-event-list").style.display="none"}),document.querySelectorAll(".cal-filter").forEach(c=>c.addEventListener("change",t));async function i(){try{let c=`${e.getFullYear()}-${String(e.getMonth()+1).padStart(2,"0")}`;n=(await f(`/api/dashboard/calendar?month=${c}`)).data?.data||[]}catch(c){console.warn("[Calendar] Failed to load events, rendering empty grid:",c),n=[]}}async function t(){let c=document.getElementById("calendar-grid");if(c){c.innerHTML=`<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join("")}
    </div>`,await i();try{let u=e.getFullYear(),m=e.getMonth(),o=e.toLocaleDateString("id-ID",{month:"long",year:"numeric"}),l=document.getElementById("cal-month-label");l&&(l.textContent=o);let s=new Set(Array.from(document.querySelectorAll(".cal-filter:checked")).map(k=>k.value)),r=n.filter(k=>s.has(k.type)),h={};r.forEach(k=>{let _=(k.event_date||"").slice(0,10);h[_]||(h[_]=[]),h[_].push(k)});let b=new Date(u,m,1).getDay(),g=new Date(u,m+1,0).getDate(),d=["Min","Sen","Sel","Rab","Kam","Jum","Sab"],p=new Date().toISOString().slice(0,10),y='<div class="calendar-grid">';d.forEach(k=>{y+=`<div class="cal-day-header">${k}</div>`});for(let k=0;k<b;k++)y+='<div class="cal-cell cal-cell-empty"></div>';for(let k=1;k<=g;k++){let _=`${u}-${String(m+1).padStart(2,"0")}-${String(k).padStart(2,"0")}`,E=h[_]||[],A=_===p;y+=`
          <div class="cal-cell ${A?"cal-today":""} ${E.length?"cal-has-events":""}"
               data-date="${_}" tabindex="0" role="button" aria-label="${_}">
            <div class="cal-day-num ${A?"today-num":""}">${k}</div>
            <div class="cal-events-preview">
              ${E.slice(0,3).map(w=>`
                <div class="cal-event-dot cal-color-${w.color||"gray"}" title="${xe(w.title||w.type)}">
                  <span class="cal-event-dot-label">${gt(w.title||w.branch_name||w.type,18)}</span>
                </div>
              `).join("")}
              ${E.length>3?`<div class="cal-more">+${E.length-3} lagi</div>`:""}
            </div>
          </div>`}let S=(b+g)%7;if(S!==0)for(let k=0;k<7-S;k++)y+='<div class="cal-cell cal-cell-empty"></div>';y+="</div>",c.innerHTML=y,c.querySelectorAll(".cal-cell[data-date]").forEach(k=>{k.addEventListener("click",()=>{let _=k.dataset.date,E=h[_]||[];if(!E.length)return;let A=document.getElementById("cal-event-list"),w=new Date(_+"T00:00:00").toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"});document.getElementById("cal-event-date").textContent=w,document.getElementById("cal-event-items").innerHTML=E.map(I=>`
            <div class="cal-event-item cal-color-border-${I.color||"gray"}">
              <div class="cal-event-type">${bt(I.type)}</div>
              <div class="cal-event-title">${xe(I.title||"-")}</div>
              <div class="cal-event-branch">${xe(I.branch_name||"")}</div>
              ${I.status?`<div class="cal-event-status">${xe(I.status)}</div>`:""}
              ${I.days_remaining!==void 0?`<div class="cal-event-extra">Sisa: ${I.days_remaining} hari</div>`:""}
            </div>
          `).join(""),A.style.display="block"})})}catch(u){console.error("[Calendar] Render error:",u),c&&(c.innerHTML=`
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">\u{1F4C5}</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`)}}}t()}function gt(a,e){return a?a.length>e?a.slice(0,e)+"\u2026":a:""}function xe(a){return a?String(a).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"):""}function bt(a){return{schedule:"\u{1F5D3} Jadwal",issue:"\u26A0\uFE0F Permasalahan",reliever:"\u{1F504} Reliefer",training:"\u{1F393} Training",contract_expiry:"\u{1F4CB} Kontrak Habis"}[a]||a}P();async function Pa(a){let e=ie(),n=(e?.full_name||e?.username||"U")[0].toUpperCase(),t={superadmin:"#7C3AED",admin:"#2563EB",manager:"#0891B2",spv:"#059669",viewer:"#64748B"}[e?.role]||"#64748B";a.innerHTML=`
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
  `;let c=localStorage.getItem("fm_token"),u=document.getElementById("session-info");if(c&&u)try{let m=JSON.parse(atob(c.split(".")[1])),o=new Date(m.exp*1e3);u.textContent=`Berakhir: ${o.toLocaleString("id-ID")}`}catch{u.textContent="Tidak tersedia"}document.getElementById("change-pwd-form")?.addEventListener("submit",async m=>{m.preventDefault();let o=document.getElementById("pwd-error"),l=document.getElementById("pwd-success"),s=document.getElementById("btn-save-pwd");o.style.display="none",l.style.display="none";let r=m.target,h=r.current_password.value,b=r.new_password.value,g=r.confirm_password.value;if(b!==g){o.textContent="\u274C Konfirmasi password tidak cocok.",o.style.display="block";return}if(b.length<6){o.textContent="\u274C Password baru minimal 6 karakter.",o.style.display="block";return}s.disabled=!0,s.textContent="\u23F3 Menyimpan...";let d=await f("/api/auth/change-password",{method:"POST",body:JSON.stringify({current_password:h,new_password:b})});s.disabled=!1,s.innerHTML='<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password',d.ok?(l.textContent="\u2705 Password berhasil diubah.",l.style.display="block",r.reset(),K("Password berhasil diubah.")):(o.textContent=d.data?.error||"Gagal mengubah password.",o.style.display="block")}),document.getElementById("btn-logout")?.addEventListener("click",()=>{confirm("Keluar dari semua sesi? Anda harus login ulang.")&&(localStorage.clear(),window.location.reload())})}P();var Te={Validasi:{module:"validation",label:"Master Referensi"},SOP:{module:"sop",label:"SOP"},"Master Karyawan":{module:"employees",label:"Karyawan"},"Data Kontrak":{module:"contracts",label:"Kontrak"},Permasalahan:{module:"issues",label:"Permasalahan"},"One on One":{module:"one_on_one",label:"One on One"},"Time Line":{module:"schedule",label:"Jadwal Kegiatan"},"Report Inspeksi Hygiene 2026":{module:"inspection",label:"Laporan Inspeksi"},"Report GC-DC 2026":{module:"cleaning",label:"Laporan GC/DC"},"Report Fogging 2026":{module:"fogging",label:"Laporan Fogging"},"Rekap Laporan Basecamp":{module:"basecamp",label:"Rekap Basecamp"},"Jadwal Reliefer":{module:"relievers",label:"Reliefer"},Training:{module:"training",label:"Training"},"Master Checklist":{module:"checklist",label:"Checklist"},"Master Form":{module:"forms",label:"Master Form"},"Permintaan Chemical":{module:"supply",label:"Inventory Chemical"}};function M(a){if(a==null||a==="")return null;if(a instanceof Date)return isNaN(a.getTime())?null:a.toISOString().slice(0,10);let e=String(a).trim();if(e===""||e==="0")return null;if(/^\d{4}-\d{2}-\d{2}/.test(e))return e.slice(0,10);if(/^\d{4,5}$/.test(e)){let t=Number(e);if(t>2e4&&t<99999){let c=new Date(Date.UTC(1899,11,30)+t*864e5);return isNaN(c.getTime())?null:c.toISOString().slice(0,10)}}let n=e.split(/[\/\-\.]/);if(n.length===3){let[t,c,u]=n.map(s=>s.trim()),m=Number(t),o=Number(c),l=Number(u);if(t.length===4&&m>1900)return`${t}-${c.padStart(2,"0")}-${u.padStart(2,"0")}`;if(u.length===4&&l>1900)return m>12?`${u}-${c.padStart(2,"0")}-${t.padStart(2,"0")}`:o>12?`${u}-${t.padStart(2,"0")}-${c.padStart(2,"0")}`:`${u}-${c.padStart(2,"0")}-${t.padStart(2,"0")}`;if(u.length===2&&!isNaN(l)){let s=l>=50?`19${u}`:`20${u}`;return m>12?`${s}-${c.padStart(2,"0")}-${t.padStart(2,"0")}`:`${s}-${c.padStart(2,"0")}-${t.padStart(2,"0")}`}}let i=new Date(e);return isNaN(i.getTime())?null:i.toISOString().slice(0,10)}function Ia(a){return Object.values(a).every(e=>e==null||String(e).trim()==="")}var ht={validation:{required:[],map:a=>({cabang:a.CABANG,pic:a.PIC,kegiatan:a.KEGIATAN,quartal:a.QUARTAL,masa_pkwt:a["MASA PKWT"],pic_pelapor:a["PIC PELAPOR"],kontrak:a.KONTRAK})},sop:{required:[{key:"Nama SOP",label:"Nama SOP"}],map:a=>({name:a["Nama SOP"],category:a.Kategori||"Umum",document_link:a["Link Document"],version:"1.0",effective_date:null,notes:""})},employees:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:a=>({full_name:a["Nama Lengkap"],branch_name:a.Cabang,division:a["Div / Bagian"]||"FACILITY CARE",phone:a["No. Hp"],join_date:M(a["Tanggal Masuk"]),status:a.Status||"",notes:""})},contracts:{required:[{key:"Nama Lengkap",label:"Nama Lengkap"}],map:a=>({employee_name:a["Nama Lengkap"],branch_name:a.Cabang,division:a["Div / Bagian"]||"FACILITY CARE",start_date:M(a["Tanggal Mulai"]),end_date:M(a["Tanggal Selesai"]),contract_type:a["Tipe Kontrak"]||"",pkwt_number:a.PKWT||"",status:a.Status||"",notes:a.keterangan})},issues:{required:[{key:"Keluhan",label:"Keluhan"}],map:a=>({report_date:M(a["Tanggal Info"]),branch_name:a.Cabang,category:a.Kategori,source:a["Sumber Laporan"],complaint:a.Keluhan,employee_name:a["Nama FC"],fc_specialist:a["FC Spesialis"],solution:a.Solusi,status:a.Status||"",completion_date:M(a["Tanggal Selesai"])})},one_on_one:{required:[],map:a=>({meeting_date:M(a.Tanggal),branch_name:a.Cabang,employee_name:a["Nama Karyawan"],pic:a.Pic,problem:a.Masalah,solution:a.Solusi,status:a.Status||"",completion_date:M(a["Tanggal Selesai"]),document_link:a["Link Document"]})},schedule:{required:[{key:"Kegiatan",label:"Kegiatan"}],map:a=>({branch_name:a.Cabang,activity_type:a.Kegiatan,period:a.Periode,pic:a.Pic||a.PIC,opening_date:M(a["Tanggal Opening"]||a["Tgl Opening"]),target_date:M(a["Tanggal Target"]||a["Tgl Target"]),completion_date:M(a["Tanggal Selesai"]||a["Tgl Selesai"]),status:a.Status||"",notes:a.Keterangan||a.Catatan})},inspection:{required:[],map:a=>({inspection_date:M(a.Tanggal),branch_name:a.Cabang,period:a.Periode,status:a.Status||"",fc_score:a["Point FC SP"]!==void 0&&a["Point FC SP"]!==null?parseFloat(String(a["Point FC SP"]).replace(",",".")):null,spv_score:a["Point SPV"]!==void 0&&a["Point SPV"]!==null?parseFloat(String(a["Point SPV"]).replace(",",".")):null,document_link:a.Link,notes:""})},cleaning:{required:[],map:a=>({activity_date:M(a.Tanggal),branch_name:a.Cabang,activity_type:a["Jenis Kegiatan"]||"General Cleaning",period:a.Periode,status:a.Status||"",document_link:a.Link,notes:""})},fogging:{required:[],map:a=>({activity_date:M(a.Tanggal),branch_name:a.Cabang,period:a.Periode,status:a.Status||"",document_link:a.Link,notes:""})},basecamp:{required:[{key:"Permasalahan",label:"Permasalahan"}],map:a=>({info_date:M(a["Tgl Info"]),branch_name:a.Cabang,problem:a.Permasalahan,pic:a.PIC,done_date:M(a["Tgl Done"]),status:a.Status||"",notes:a.Ket})},relievers:{required:[],map:a=>({branch_name:a.Cabang,original_fc_name:a["Nama Facility care"],period:a.Periode,reliever_name:a.Relifer,backup_date:M(a["Tanggal Back Up"]),completion_date:M(a["Tanggal Selesai"]),reason:a.Keterangan,shift:a.Shift,status:a.Status||""})},training:{required:[{key:"Materi",label:"Materi"}],map:a=>({training_date:M(a.Tanggal),batch:a.Batch,subject:a.Materi,participants:a.Peserta,branch_name:a.Cabang,trainer:a.Trainer,score:a.Nilai!==void 0&&a.Nilai!==null?parseFloat(String(a.Nilai).replace(",",".")):null,notes:""})},checklist:{required:[],map:a=>({name:a["Master Checklist"],category:"Umum",document_link:a["Link Document"],description:""})},forms:{required:[{key:"Master Form",label:"Master Form"}],map:a=>({name:a["Master Form"],category:"Umum",document_link:a["Link Document"],description:""})},supply:{required:[],map:a=>({submitted_at:M(a.Timestamp),submitter_name:a["Nama Lengkap"],branch_name:a["Kebutuhan Untuk Cabang"],tools_items:a["Alat - Alat / Barang"],tools_quantity:a["Jumlah Permintaan Alat / Barang"],chemical_items:a.Chemical,chemical_quantity:a["Jumlah Permintaan Chemical"],additional_notes:a["Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List."],status:a.Status||""})}};function yt(a,e){let n=Te[a];if(!n)return{valid:[],errors:[],mapped:[],skipped:!0};let i=ht[n.module];if(!i)return{valid:[],errors:[],mapped:[],skipped:!0};let t=[],c=[],u=[];return e.filter(o=>!Ia(o)).forEach((o,l)=>{let s=e.indexOf(o)+2,r=[];i.required.forEach(({key:b,label:g})=>{let d=o[b];if(d==null||String(d).trim()===""){let p=Object.keys(o).filter(y=>y.trim()).join(", ");r.push({column:g,originalValue:d||"",reason:`Kolom "${g}" wajib diisi dan tidak ditemukan`,hint:`Kolom yang tersedia: ${p.slice(0,120)}`})}});let h=i.map(o);r.length>0?c.push({row:s,data:h,raw:o,errors:r}):(t.push(o),u.push(h))}),{valid:t,errors:c,mapped:u}}function Da(a){let e=[];return a.SheetNames.forEach(n=>{let i=Te[n];if(!i)return;let t=a.Sheets[n],c=window.XLSX.utils.sheet_to_json(t,{defval:"",raw:!1,dateNF:"yyyy-mm-dd"}),u=yt(n,c),m=c.filter(o=>!Ia(o));e.push({sheetName:n,module:i.module,label:i.label,total:m.length,valid:u.mapped.length,errorCount:u.errors.length,errors:u.errors,mapped:u.mapped,skipped:!1})}),e}function La(){let a=window.XLSX,e=a.utils.book_new();Object.entries({Validasi:[{CABANG:"001. Pondok Bambu","NAMA KARYAWAN":"Budi Santoso",PIC:"Berlin",KEGIATAN:"General Cleaning",QUARTAL:"Q1","PIC PELAPOR":"Berlin",KONTRAK:"PKWT 1","MASA PKWT":"1 Tahun"}],SOP:[{"Nama SOP":"SOP Pembersihan Toilet",Kategori:"Cleaning","Link Document":"https://..."}],"Master Karyawan":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","No. Hp":"081234567890","Tanggal Masuk":"2024-01-15",Status:"Aktif"}],"Data Kontrak":[{"Nama Lengkap":"Budi Santoso",Cabang:"001. Pondok Bambu","Div / Bagian":"FACILITY CARE","Tanggal Mulai":"2024-01-01","Tanggal Selesai":"2024-12-31","Sisa Kontrak":"365",Status:"Aktif",keterangan:""}],Permasalahan:[{"Tanggal Info":"2024-03-01",Cabang:"001. Pondok Bambu",Kategori:"Cleaning","Sumber Laporan":"SPV",Keluhan:"Lantai kotor","Nama FC":"Budi","FC Spesialis":"Fajar",Solusi:"Teguran",Status:"Done","Tanggal Selesai":"2024-03-02",Day:"1"}],"One on One":[{Tanggal:"2024-03-05",Cabang:"001. Pondok Bambu","Nama Karyawan":"Budi Santoso",Pic:"Berlin",Masalah:"Keterlambatan",Solusi:"Coaching",Status:"Done","Tanggal Selesai":"2024-03-06",Day:"1","Link Document":""}]}).forEach(([i,t])=>{a.utils.book_append_sheet(e,a.utils.json_to_sheet(t),i)}),a.writeFile(e,"Template_Import_Data_Awal_FCMS.xlsx")}function Ba(a){let e=window.XLSX,n=e.utils.book_new(),i=!1;return a.forEach(t=>{if(!t.errors||t.errors.length===0)return;i=!0;let c=t.errors.map(m=>({"No. Baris":m.row,"Kolom Gagal":(m.errors||[]).map(o=>o.column||o).join("; "),"Alasan Error":(m.errors||[]).map(o=>o.reason||o).join("; "),...Object.fromEntries(Object.entries(m.data||{}).map(([o,l])=>[o,l??""]))})),u=e.utils.json_to_sheet(c);e.utils.book_append_sheet(n,u,t.sheetName.replace(/[\\\/\[\]*?:]/g,"_").slice(0,31))}),i?(e.writeFile(n,`Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`),!0):!1}var ft=["validation","employees","contracts","relievers","schedule","issues","one_on_one","training","checklist","forms","sop","inspection","cleaning","fogging","basecamp","supply"];function Na(a){a.innerHTML=`
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
  `;let e=null,n=null,i=0,t={upload:document.getElementById("step-upload"),validating:document.getElementById("step-validating"),preview:document.getElementById("step-preview"),importing:document.getElementById("step-importing"),summary:document.getElementById("step-summary")};function c(d){Object.entries(t).forEach(([p,y])=>{y.style.display=p===d?"":"none"})}document.getElementById("btn-backup-db")?.addEventListener("click",async()=>{let d=document.getElementById("btn-backup-db");d.disabled=!0,d.textContent="\u23F3 Memproses Backup...";try{let p=await f("/api/import/backup");if(p.ok){let y=new Blob([JSON.stringify(p.data,null,2)],{type:"application/json"}),v=URL.createObjectURL(y),S=document.createElement("a");S.href=v,S.download=`FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.json`,document.body.appendChild(S),S.click(),document.body.removeChild(S),URL.revokeObjectURL(v),K("Backup berhasil diunduh!")}else j("Gagal memproses backup: "+(p.data?.error||"Unknown error"))}catch(p){j("Gagal memproses backup: "+p.message)}finally{d.disabled=!1,d.textContent="\u{1F4E6} Backup Database"}});let u=document.getElementById("btn-sync-google");u&&u.addEventListener("click",async()=>{if(!confirm("Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?"))return;let d=u.innerHTML;u.innerHTML='<span class="spinner"></span> Menyinkronkan...',u.disabled=!0;try{let p=await f("/api/sync/google-sheets",{method:"POST"});p.ok?alert("Sinkronisasi Berhasil: "+(p.data?.message||"Data Karyawan & PIC telah diperbarui.")):alert("Gagal Sinkronisasi: "+(p.data?.error||"Unknown error"))}catch{alert("Terjadi kesalahan koneksi.")}finally{u.innerHTML=d,u.disabled=!1}}),document.getElementById("btn-download-template").addEventListener("click",()=>{La(),K("Template Excel berhasil didownload!")});let m=document.getElementById("file-input"),o=document.getElementById("upload-zone");document.getElementById("btn-browse").addEventListener("click",d=>{d.stopPropagation(),m.click()}),m.addEventListener("change",d=>{d.target.files[0]&&l(d.target.files[0])}),o.addEventListener("dragover",d=>{d.preventDefault(),o.classList.add("drag-over")}),o.addEventListener("dragleave",()=>o.classList.remove("drag-over")),o.addEventListener("drop",d=>{d.preventDefault(),o.classList.remove("drag-over");let p=d.dataTransfer.files[0];p&&p.name.match(/\.xlsx?$/i)?l(p):j("Hanya file .xlsx atau .xls yang didukung.")}),document.getElementById("btn-clear-file").addEventListener("click",()=>{e=null,m.value="",document.getElementById("file-info").style.display="none",o.style.display="",c("upload")});async function l(d){e=d,document.getElementById("file-name-display").textContent=`\u{1F4C4} ${d.name} (${(d.size/1024).toFixed(1)} KB)`,document.getElementById("file-info").style.display="flex",o.style.display="none",await s(d)}async function s(d){c("validating");let p=document.getElementById("validation-status"),y=document.getElementById("validation-bar");try{if(!window.XLSX)throw new Error("Library SheetJS belum termuat. Refresh halaman dan coba lagi.");p.textContent="Membaca file Excel...",y.style.width="20%",await be(200);let v=await d.arrayBuffer(),S=window.XLSX.read(v,{type:"array",cellDates:!0});p.textContent=`Memvalidasi ${S.SheetNames.length} sheet...`,y.style.width="50%",await be(100),n=Da(S),y.style.width="100%",p.textContent="Validasi selesai!",await be(300),r()}catch(v){c("upload"),j("Gagal memproses file: "+v.message),document.getElementById("file-info").style.display="flex",o.style.display="none"}}function r(){c("preview");let d=n.filter(w=>!w.skipped).length,p=n.reduce((w,I)=>w+I.total,0),y=n.reduce((w,I)=>w+I.valid,0),v=n.reduce((w,I)=>w+I.errorCount,0),S=p>0?Math.round(y/p*100):0;document.getElementById("preview-summary-badges").innerHTML=`
      <span class="badge badge-info">${d} sheet</span>
      <span class="badge badge-secondary">${p} baris</span>
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
    `,k.querySelectorAll(".btn-detail-error").forEach(w=>{w.addEventListener("click",()=>{let I=n[Number(w.dataset.idx)];h(I)})});let _=document.getElementById("error-detail-section"),E=document.getElementById("error-detail-container");E.innerHTML="",_.style.display="none";let A=document.getElementById("btn-start-import");y===0?(A.disabled=!0,A.innerHTML="\u26A0\uFE0F Tidak Ada Data Valid"):(A.disabled=!1,v>0?(A.innerHTML=`\u{1F680} Import ${y} Data Valid (${v} dilewati)`,A.title="Baris error akan dilewati, baris valid tetap diimport"):A.innerHTML=`\u{1F680} Mulai Import ${y} Data`)}function h(d){let p=document.getElementById("error-detail-section"),y=document.getElementById("error-detail-container");p.style.display="";let v=d.errors.slice(0,100).map(S=>(Array.isArray(S.errors)?S.errors:[]).map(_=>{let E=typeof _=="object";return`
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${S.row}</span></td>
            <td><strong>${E?_.column:"\u2014"}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${E&&_.originalValue!==void 0?_.originalValue||"(kosong)":"\u2014"}</code></td>
            <td class="error-msg">${E?_.reason:_}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${E&&_.aliases?`Gunakan salah satu nama kolom:<br><em>${_.aliases}</em>`:E&&_.hint?_.hint:""}
            </td>
          </tr>
        `}).join("")).join("");y.innerHTML=`
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
    `,p.scrollIntoView({behavior:"smooth",block:"start"})}document.getElementById("btn-back-to-upload").addEventListener("click",()=>{c("upload"),document.getElementById("file-info").style.display="none",o.style.display="",e=null,m.value=""}),document.getElementById("btn-download-log").addEventListener("click",()=>{if(!n)return;Ba(n)?K("Log error berhasil didownload."):K("Tidak ada error untuk didownload.")}),document.getElementById("btn-start-import").addEventListener("click",()=>{let d=document.querySelector('input[name="dup-strategy"]:checked')?.value||"skip";b(d)});async function b(d){c("importing"),i=Date.now();let p=[];ft.forEach(_=>{let E=n?.find(A=>A.module===_&&A.mapped?.length>0);E&&p.push(E)});let y=document.getElementById("import-steps-list");y.innerHTML=p.map(_=>`
      <div class="import-step-item" id="step-item-${_.module}">
        <span class="step-item-icon" id="step-icon-${_.module}">\u23F8\uFE0F</span>
        <span class="step-item-label">${_.label} <span class="step-item-count">(${_.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${_.module}"></span>
      </div>
    `).join("");let v=document.getElementById("import-bar"),S=document.getElementById("import-current-status"),k={totalSheets:p.length,totalRows:p.reduce((_,E)=>_+E.mapped.length,0),inserted:0,skipped:0,failed:0,moduleResults:[]};for(let _=0;_<p.length;_++){let E=p[_],A=document.getElementById(`step-icon-${E.module}`),w=document.getElementById(`step-status-${E.module}`);A.textContent="\u{1F504}",w.textContent="Mengimport...",S.textContent=`Mengimport ${E.label}...`,v.style.width=`${Math.round(_/p.length*100)}%`;try{let I=await f(`/api/import/${E.module}`,{method:"POST",body:JSON.stringify({rows:E.mapped,onDuplicate:d})});if(I.ok){let V=I.data;k.inserted+=V.inserted||0,k.skipped+=V.skipped||0,k.moduleResults.push({label:E.label,inserted:V.inserted||0,skipped:V.skipped||0,status:"ok"}),A.textContent="\u2705",w.innerHTML=`<span class="badge badge-success">${V.inserted||0} berhasil</span>${V.skipped>0?` <span class="badge badge-neutral">${V.skipped} skip</span>`:""}`}else k.failed++,k.moduleResults.push({label:E.label,inserted:0,skipped:0,status:"error",error:I.data?.error}),A.textContent="\u274C",w.innerHTML='<span class="badge badge-danger">Gagal</span>'}catch(I){k.failed++,k.moduleResults.push({label:E.label,inserted:0,skipped:0,status:"error",error:I.message}),A.textContent="\u274C",w.innerHTML='<span class="badge badge-danger">Gagal</span>'}await be(150)}v.style.width="100%",S.textContent="Selesai!",await be(400),g(k)}function g(d){c("summary");let p=((Date.now()-i)/1e3).toFixed(1),y=d.failed===0;document.getElementById("summary-status-icon").innerHTML=`
      <div class="summary-icon">${y?"\u{1F389}":"\u26A0\uFE0F"}</div>
      <h2 class="summary-title">${y?"Import Berhasil!":"Import Selesai dengan Beberapa Error"}</h2>
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
    `}document.getElementById("btn-import-again").addEventListener("click",()=>{e=null,n=null,m.value="",document.getElementById("file-info").style.display="none",o.style.display="",c("upload")}),document.getElementById("btn-go-to-dashboard").addEventListener("click",()=>{window.location.hash="/dashboard"})}function be(a){return new Promise(e=>setTimeout(e,a))}P();var je=[];async function Aa(a){je=((await f("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name})),$({container:a,title:"Data SP (Surat Peringatan)",icon:"\u2709\uFE0F",apiPath:"/api/sp",itemLabel:"SP",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:n=>n?new Date(n).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"branch_name",label:"Cabang"},{key:"sp_type",label:"Jenis SP",render:n=>`<span class="badge badge-warning">${n||"-"}</span>`},{key:"status",label:"Status",render:n=>`<span class="badge ${n==="Aktif"?"badge-danger":"badge-success"}">${n||"-"}</span>`},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"branch_id",label:"Cabang",options:je}],exportOptions:{moduleName:"sp_data",onExport:async n=>{let i=new URLSearchParams(n||{}).toString(),t=await f(`/api/sp?limit=10000&${i}`);if(t.ok){let c=t.data.data.map(m=>({Tanggal:m.tanggal||"","Nama Karyawan":m.employee_name||"",Cabang:m.branch_name||"","Jenis SP":m.sp_type||"",Status:m.status||"",Dokumen:m.document_link||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(L(),J));u(c,`Data_SP_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08","Nama Karyawan":"Budi Santoso",Cabang:"001. Pondok Bambu","Jenis SP":"SP 1",Status:"Aktif",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(L(),J));i(n,"Template_Import_SP")},onImport:async n=>{let t=(await f("/api/branches?all=1")).data?.data||[],c=l=>{if(!l)return null;let s=l.toLowerCase(),r=t.find(h=>h.full_name.toLowerCase()===s||h.code.toLowerCase()===s||h.name.toLowerCase()===s);return r?r.id:null},u=l=>{if(!l)return"";if(l instanceof Date&&!isNaN(l.getTime()))return l.toISOString().slice(0,10);let s=String(l).trim();if(/^\d{4,5}$/.test(s)){let h=Number(s);if(h>2e4&&h<99999){let b=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);let r=s.split(/[\/\-\.]/);if(r.length===3){let[h,b,g]=r.map(d=>d.trim());if(h.length===4&&b.length<=2&&g.length<=2)return`${h}-${b.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&b.length<=2&&h.length<=2)return`${g}-${b.padStart(2,"0")}-${h.padStart(2,"0")}`}return s},m=n.map(l=>({tanggal:u(l.Tanggal),employee_name:String(l["Nama Karyawan"]||"").trim(),branch_id:c(String(l.Cabang||"").trim()),sp_type:String(l["Jenis SP"]||"").trim(),status:String(l.Status||"").trim(),document_link:String(l.Dokumen||"").trim()})).filter(l=>l.tanggal&&l.employee_name&&l.branch_id),o=await f("/api/sp/import",{method:"POST",body:JSON.stringify(m)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"text",name:"employee_name",label:"Nama Karyawan",required:!0},{type:"select",name:"branch_id",label:"Cabang",required:!0,options:je},{type:"select",name:"sp_type",label:"Jenis Surat Peringatan",required:!0,options:["SP 1","SP 2","SP 3","Teguran Lisan"]},{type:"select",name:"status",label:"Status",required:!0,options:["Aktif","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}P();var he=[];async function Oa(a){he=((await f("/api/branches?all=1")).data?.data||[]).map(n=>({value:n.id,label:n.full_name})),$({container:a,title:"Data Mutasi",icon:"\u{1F504}",apiPath:"/api/mutasi",itemLabel:"Mutasi",bulkDelete:!0,columns:[{key:"tanggal",label:"Tanggal",render:n=>n?new Date(n).toLocaleDateString("id-ID",{year:"numeric",month:"short",day:"numeric"}):"-"},{key:"employee_name",label:"Nama Karyawan"},{key:"from_branch_name",label:"Cabang Asal"},{key:"to_branch_name",label:"Cabang Tujuan"},{key:"status",label:"Status",render:n=>`<span class="badge ${n==="Selesai"?"badge-success":"badge-warning"}">${n||"-"}</span>`},{key:"document_link",label:"Dokumen",render:n=>n?`<a href="${n}" target="_blank" class="text-primary hover-underline">Lihat</a>`:"-"}],filterFields:[{type:"search",placeholder:"Cari nama karyawan..."},{type:"select",name:"from_branch_id",label:"Cabang Asal",options:he},{type:"select",name:"to_branch_id",label:"Cabang Tujuan",options:he}],exportOptions:{moduleName:"mutasi_data",onExport:async n=>{let i=new URLSearchParams(n||{}).toString(),t=await f(`/api/mutasi?limit=10000&${i}`);if(t.ok){let c=t.data.data.map(m=>({Tanggal:m.tanggal||"","Nama Karyawan":m.employee_name||"","Cabang Asal":m.from_branch_name||"","Cabang Tujuan":m.to_branch_name||"",Status:m.status||"",Dokumen:m.document_link||""})),{downloadExcel:u}=await Promise.resolve().then(()=>(L(),J));u(c,`Data_Mutasi_${new Date().toISOString().slice(0,10)}`)}else throw new Error("Gagal mengambil data")},onTemplate:async()=>{let n=[{Tanggal:"2026-01-08","Nama Karyawan":"Widya Astuti","Cabang Asal":"001. Pondok Bambu","Cabang Tujuan":"007. Bekasi",Status:"Selesai",Dokumen:"https://link.doc"}],{downloadExcel:i}=await Promise.resolve().then(()=>(L(),J));i(n,"Template_Import_Mutasi")},onImport:async n=>{let t=(await f("/api/branches?all=1")).data?.data||[],c=l=>{if(!l)return null;let s=l.toLowerCase(),r=t.find(h=>h.full_name.toLowerCase()===s||h.code.toLowerCase()===s||h.name.toLowerCase()===s);return r?r.id:null},u=l=>{if(!l)return"";if(l instanceof Date&&!isNaN(l.getTime()))return l.toISOString().slice(0,10);let s=String(l).trim();if(/^\d{4,5}$/.test(s)){let h=Number(s);if(h>2e4&&h<99999){let b=new Date(Date.UTC(1899,11,30)+h*864e5);return isNaN(b.getTime())?"":b.toISOString().slice(0,10)}}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);let r=s.split(/[\/\-\.]/);if(r.length===3){let[h,b,g]=r.map(d=>d.trim());if(h.length===4&&b.length<=2&&g.length<=2)return`${h}-${b.padStart(2,"0")}-${g.padStart(2,"0")}`;if(g.length===4&&b.length<=2&&h.length<=2)return`${g}-${b.padStart(2,"0")}-${h.padStart(2,"0")}`}return s},m=n.map(l=>({tanggal:u(l.Tanggal),employee_name:String(l["Nama Karyawan"]||"").trim(),from_branch_id:c(String(l["Cabang Asal"]||"").trim()),to_branch_id:c(String(l["Cabang Tujuan"]||"").trim()),status:String(l.Status||"").trim(),document_link:String(l.Dokumen||"").trim()})).filter(l=>l.tanggal&&l.employee_name&&l.from_branch_id&&l.to_branch_id),o=await f("/api/mutasi/import",{method:"POST",body:JSON.stringify(m)});if(!o.ok)throw new Error(o.data?.error||"Import gagal")}},formFields:[{type:"date",name:"tanggal",label:"Tanggal",required:!0},{type:"text",name:"employee_name",label:"Nama Karyawan",required:!0},{type:"combobox",name:"from_branch_id",label:"Cabang Asal",required:!0,options:he,createApi:{path:"/api/branches",field:"full_name"}},{type:"combobox",name:"to_branch_id",label:"Cabang Tujuan",required:!0,options:he,createApi:{path:"/api/branches",field:"full_name"}},{type:"select",name:"status",label:"Status",required:!0,options:["Proses","Selesai"]},{type:"url",name:"document_link",label:"Link Dokumen (Opsional)"}]})}window.formatDate=a=>{if(!a||a==="-")return"";let e=a.split("-");return e.length===3&&e[0].length===4?`${e[2]}-${e[1]}-${e[0]}`:a};function N(a){return async e=>{if(!ce()){se("/login");return}return a(e)}}var ye=null;function vt(){ye&&clearInterval(ye);let a=()=>{let e=new Date,n=e.toLocaleTimeString("id-ID",{hour:"2-digit",minute:"2-digit",second:"2-digit"}),i=e.toLocaleDateString("id-ID",{weekday:"long",day:"numeric",month:"long",year:"numeric"}),t=document.getElementById("header-clock-time"),c=document.getElementById("header-clock-date");t&&(t.textContent=n),c&&(c.textContent=i)};a(),ye=setInterval(a,1e3)}async function kt(){try{let a=await f("/api/dashboard/kpi");if(!a.ok)return;let e=a.data?.data||a.data||{},n=(i,t)=>{let c=document.getElementById(i);c&&(c.textContent=t>0?t:"",c.style.display=t>0?"inline-flex":"none")};n("badge-issues",e.issues?.current||0),n("badge-contracts",e.expiring30?.current||0),n("badge-oo1",e.one_on_one?.current||0),n("badge-schedule",e.schedule?.current||0),n("badge-supply",e.supply?.current||0)}catch{}}var le=[];async function St(){try{let a=await f("/api/dashboard/notifications");if(!a.ok)return;le=a.data?.data||a.data||[];let e=document.getElementById("notif-dot");e&&(e.style.display=le.length>0?"block":"none",e.textContent=le.length)}catch{}}function _t(){if(!le.length){ee({title:"Notifikasi",content:'<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',confirmText:"Tutup",onConfirm:(e,n)=>n()});return}let a=`
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
  `;ee({title:`Notifikasi (${le.length})`,content:a,confirmText:"Tutup",onConfirm:(e,n)=>n()})}function Ma(){let a=ie(),e=(a?.full_name||"U")[0].toUpperCase();document.getElementById("app").innerHTML=`
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
  `;let n=document.getElementById("sidebar"),i=document.getElementById("sidebar-overlay"),t=document.getElementById("topbar-menu-btn"),c=document.getElementById("sidebar-close"),u=()=>{n.classList.add("open"),i.classList.add("show")},m=()=>{n.classList.remove("open"),i.classList.remove("show")};t?.addEventListener("click",u),c?.addEventListener("click",m),i?.addEventListener("click",m),document.querySelectorAll(".nav-item").forEach(l=>l.addEventListener("click",m));function o(){let l=window.location.hash.replace("#","")||"/dashboard";document.querySelectorAll(".nav-item").forEach(h=>{let b=h.dataset.route;h.classList.toggle("active",l===b||b!=="/dashboard"&&l.startsWith(b))});let s=document.getElementById("topbar-title"),r=document.querySelector(".nav-item.active .nav-label");s&&r&&(s.textContent=r.textContent)}window.addEventListener("hashchange",o),o(),vt(),document.getElementById("btn-fullscreen")?.addEventListener("click",()=>{document.fullscreenElement?document.exitFullscreen?.():document.documentElement.requestFullscreen?.()}),document.getElementById("logout-btn")?.addEventListener("click",async()=>{await f("/api/auth/logout",{method:"POST"}),de(),ye&&clearInterval(ye),se("/login")}),kt(),St(),document.getElementById("btn-notif")?.addEventListener("click",l=>{l.preventDefault(),_t()})}async function wt(){B("/login",({main:e})=>sa(e)),B("/dashboard",N(({main:e})=>ta(e))),B("/calendar",N(({main:e})=>Ea(e))),B("/employees",N(({main:e})=>ma(e))),B("/contracts",N(({main:e})=>ga(e))),B("/sp",N(({main:e})=>Aa(e))),B("/mutasi",N(({main:e})=>Oa(e))),B("/timeline",N(({main:e})=>ba(e))),B("/issues",N(({main:e})=>ha(e))),B("/one-on-one",N(({main:e})=>ya(e))),B("/training",N(({main:e})=>fa(e))),B("/relievers",N(({main:e})=>va(e))),B("/reports/inspection",N(({main:e})=>ka(e))),B("/reports/cleaning",N(({main:e})=>Sa(e))),B("/reports/fogging",N(({main:e})=>_a(e))),B("/reports/basecamp",N(({main:e})=>wa(e))),B("/reports/supply",N(({main:e})=>qe(e,"supply"))),B("/sop",N(({main:e})=>Ca(e))),B("/checklist",N(({main:e})=>xa(e))),B("/forms",N(({main:e})=>qe(e))),B("/users",N(({main:e})=>Ta(e))),B("/branches",N(({main:e})=>$a(e))),B("/profile",N(({main:e})=>Pa(e))),B("/settings/import",N(({main:e})=>Na(e)));let a=ce();if(!a&&window.location.hash!=="#/login"&&se("/login"),a){let e=await f("/api/auth/me");e.ok?(pe(e.data.data),Ma()):(de(),se("/login"))}window.addEventListener("fm:login",()=>{Ma(),se("/dashboard")}),Ve()}wt();
